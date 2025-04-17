"""
app.py

Main entry point for the Fleet Electrification Analyzer application.
Initializes the application, sets up logging, and launches the UI.
"""

import os
import sys
import logging
import argparse
import tkinter as tk
from pathlib import Path

# Add parent directory to path to ensure imports work properly
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import from local modules
from settings import (
    APP_NAME, 
    APP_VERSION, 
    DEFAULT_WINDOW_SIZE,
    LOG_FORMAT,
    LOG_LEVEL,
    LOG_DATE_FORMAT,
    DEFAULT_LOG_FILE
)
from utils import setup_logging
from ui.main_window import MainWindow

def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(description=f"{APP_NAME} v{APP_VERSION}")
    
    parser.add_argument(
        "--input", "-i",
        help="Path to input CSV file containing VINs"
    )
    
    parser.add_argument(
        "--output", "-o", 
        help="Path to output file for results"
    )
    
    parser.add_argument(
        "--threads", "-t", 
        type=int, 
        help="Number of worker threads for processing"
    )
    
    parser.add_argument(
        "--console-log", "-c",
        action="store_true",
        help="Log to console in addition to log file"
    )
    
    parser.add_argument(
        "--verbose", "-v", 
        action="store_true", 
        help="Enable verbose logging"
    )
    
    parser.add_argument(
        "--batch", "-b",
        action="store_true",
        help="Run in batch mode (no UI)"
    )
    
    return parser.parse_args()

def setup_application():
    """Set up directories and resources for the application."""
    # Ensure data directories exist
    from settings import DATA_DIR, CACHE_DIR, EXPORT_DIR, LOG_DIR, TEMP_DIR
    
    for dir_path in [DATA_DIR, CACHE_DIR, EXPORT_DIR, LOG_DIR, TEMP_DIR]:
        os.makedirs(dir_path, exist_ok=True)

def run_batch_mode(args):
    """Run in batch mode using command line arguments."""
    from data.processor import BatchProcessor
    
    logger = logging.getLogger(__name__)
    logger.info(f"Starting {APP_NAME} v{APP_VERSION} in batch mode")
    
    if not args.input:
        logger.error("Input file is required in batch mode")
        return 1
    
    if not args.output:
        logger.error("Output file is required in batch mode")
        return 1
    
    # Initialize processor
    processor = BatchProcessor(
        max_threads=args.threads or 10
    )
    
    # Set up logging callbacks for console output
    def log_message(msg):
        logger.info(msg)
    
    def progress_update(current, total):
        percent = int(current / total * 100)
        if percent % 10 == 0:
            logger.info(f"Progress: {percent}% ({current}/{total})")
    
    # Process file
    processor.process_file(
        input_path=args.input,
        output_path=args.output,
        log_callback=log_message,
        progress_callback=progress_update,
        done_callback=lambda vehicles: logger.info(f"Processing complete. Processed {len(vehicles)} vehicles.")
    )
    
    # Wait for processing to complete (preventing program exit)
    import time
    while processor.current_pipeline and processor.current_pipeline.processing_thread and processor.current_pipeline.processing_thread.is_alive():
        time.sleep(0.1)
    
    logger.info(f"Batch processing complete. Results saved to {args.output}")
    return 0

def run_gui_mode(args):
    """Run in GUI mode with the main application window."""
    logger = logging.getLogger(__name__)
    logger.info(f"Starting {APP_NAME} v{APP_VERSION} in GUI mode")
    
    # Create root window
    root = tk.Tk()
    root.title(f"{APP_NAME} v{APP_VERSION}")
    root.geometry(DEFAULT_WINDOW_SIZE)
    
    # Create main application window
    app = MainWindow(root)
    
    # If input file provided, load it
    if args.input:
        app.load_input_file(args.input)
    
    # Start the main event loop
    root.mainloop()
    
    return 0

def main():
    """Main entry point for the application."""
    # Parse command line arguments
    args = parse_arguments()
    
    # Set up logging
    log_level = logging.DEBUG if args.verbose else LOG_LEVEL
    logger = setup_logging(
        log_file=DEFAULT_LOG_FILE,
        console=args.console_log,
        level=log_level
    )
    
    # Set up application directories and resources
    setup_application()
    
    # Run in appropriate mode
    if args.batch:
        return run_batch_mode(args)
    else:
        return run_gui_mode(args)

if __name__ == "__main__":
    sys.exit(main())