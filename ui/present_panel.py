"""
present_panel.py

Present panel for the Fleet Electrification Analyzer application.
Provides UI for customizing and exporting PowerPoint presentations.
"""

import os
import logging
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Any, Optional, Callable

from settings import (
    PRIMARY_HEX_1,
    PRIMARY_HEX_2, 
    PRIMARY_HEX_3,
    SECONDARY_HEX_1,
    SECONDARY_HEX_2
)
from utils import SimpleTooltip, ProgressDialog
from powerpoint_export import export_prelim_deck
from powerpoint_customizer import (
    PowerPointCustomizer, executive_summary_config, technical_analysis_config,
    data_focused_config, timeline_focused_config, get_slide_selection_help
)

# Set up module logger
logger = logging.getLogger(__name__)

class PresentPanel:
    """
    Panel for PowerPoint presentation customization and export.
    Allows users to select slides, configure charts, and export presentations.
    """
    
    def __init__(self, parent_frame: ttk.Frame, sharing_data: dict):
        """
        Initialize the Present panel.
        
        Args:
            parent_frame: Parent frame to contain this panel
            sharing_data: Shared data dictionary for inter-panel communication
        """
        self.parent_frame = parent_frame
        self.sharing_data = sharing_data
        self.customizer = PowerPointCustomizer()
        self.template_path = None
        
        # Get root window reference
        self.root = parent_frame.winfo_toplevel()
        
        # Create main frame
        self.main_frame = ttk.Frame(parent_frame)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Initialize UI components
        self._create_header()
        self._create_preset_section()
        self._create_customization_section()
        self._create_template_section()
        self._create_preview_section()
        self._create_export_section()
        
        # Initialize data
        self._update_slide_options()
    
    def _create_header(self):
        """Create the header section with title and description."""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Title
        title_label = ttk.Label(
            header_frame, 
            text="PowerPoint Presentation Builder",
            font=("Segoe UI", 16, "bold")
        )
        title_label.pack(anchor=tk.W)
        
        # Description
        desc_label = ttk.Label(
            header_frame,
            text="Customize and export professional fleet electrification presentations with native PowerPoint charts",
            font=("Segoe UI", 10),
            foreground="#666666"
        )
        desc_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Workflow indicator
        workflow_frame = ttk.Frame(header_frame)
        workflow_frame.pack(fill=tk.X, pady=(10, 0))
        
        workflow_label = ttk.Label(
            workflow_frame,
            text="Workflow: Process → Results → Analysis → Present",
            font=("Segoe UI", 9, "italic"),
            foreground="#888888"
        )
        workflow_label.pack(anchor=tk.W)
    
    def _create_preset_section(self):
        """Create the preset configuration section."""
        preset_frame = ttk.LabelFrame(self.main_frame, text="Quick Start - Preset Configurations", padding=10)
        preset_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Preset buttons
        button_frame = ttk.Frame(preset_frame)
        button_frame.pack(fill=tk.X)
        
        presets = [
            ("Executive Summary", "executive_summary", "5 slides - High-level overview for leadership"),
            ("Technical Analysis", "technical_analysis", "10 slides - Comprehensive technical analysis"),
            ("Data-Focused", "data_focused", "7 slides - Emphasis on data quality and automated analysis"),
            ("Timeline-Focused", "timeline_focused", "6 slides - Focus on electrification timelines")
        ]
        
        for i, (name, preset_id, description) in enumerate(presets):
            btn_frame = ttk.Frame(button_frame)
            btn_frame.pack(fill=tk.X, pady=2)
            
            btn = ttk.Button(
                btn_frame,
                text=name,
                command=lambda p=preset_id: self._apply_preset(p),
                width=20
            )
            btn.pack(side=tk.LEFT, padx=(0, 10))
            
            desc_label = ttk.Label(
                btn_frame,
                text=description,
                font=("Segoe UI", 9),
                foreground="#666666"
            )
            desc_label.pack(side=tk.LEFT, anchor=tk.W)
            
            # Add tooltip
            SimpleTooltip(btn, f"Apply {name} preset configuration")
    
    def _create_customization_section(self):
        """Create the custom slide selection section."""
        custom_frame = ttk.LabelFrame(self.main_frame, text="Custom Slide Selection", padding=10)
        custom_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Instructions
        instructions = ttk.Label(
            custom_frame,
            text="Select which slides to include in your presentation:",
            font=("Segoe UI", 10, "bold")
        )
        instructions.pack(anchor=tk.W, pady=(0, 10))
        
        # Create scrollable frame for slide options
        canvas_frame = ttk.Frame(custom_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas and scrollbar
        canvas = tk.Canvas(canvas_frame, height=200)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Configure scrolling
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Store references
        self.slides_frame = scrollable_frame
        self.slide_checkboxes = {}
        
        # Selection controls
        controls_frame = ttk.Frame(custom_frame)
        controls_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            controls_frame,
            text="Select All",
            command=self._select_all_slides,
            width=12
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            controls_frame,
            text="Clear All",
            command=self._clear_all_slides,
            width=12
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            controls_frame,
            text="Validate Selection",
            command=self._validate_current_selection,
            width=15
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # Validation feedback
        self.validation_label = ttk.Label(
            controls_frame,
            text="",
            font=("Segoe UI", 9),
            foreground="#666666"
        )
        self.validation_label.pack(side=tk.LEFT, padx=(10, 0))
    
    def _create_template_section(self):
        """Create the template selection section."""
        template_frame = ttk.LabelFrame(self.main_frame, text="Presentation Template", padding=10)
        template_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Template selection
        template_inner = ttk.Frame(template_frame)
        template_inner.pack(fill=tk.X)
        
        ttk.Label(template_inner, text="Template:").pack(side=tk.LEFT)
        
        self.template_var = tk.StringVar(value="Default Template")
        template_entry = ttk.Entry(
            template_inner,
            textvariable=self.template_var,
            state="readonly",
            width=50
        )
        template_entry.pack(side=tk.LEFT, padx=(10, 10), fill=tk.X, expand=True)
        
        ttk.Button(
            template_inner,
            text="Browse...",
            command=self._browse_template,
            width=12
        ).pack(side=tk.RIGHT)
        
        # Template info
        template_info = ttk.Label(
            template_frame,
            text="Upload a custom .potx template file for branded presentations, or use the default template",
            font=("Segoe UI", 9),
            foreground="#666666"
        )
        template_info.pack(anchor=tk.W, pady=(5, 0))
    
    def _create_preview_section(self):
        """Create the preview section showing current configuration."""
        preview_frame = ttk.LabelFrame(self.main_frame, text="Presentation Preview", padding=10)
        preview_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Preview text
        self.preview_text = tk.Text(
            preview_frame,
            height=6,
            width=80,
            wrap=tk.WORD,
            font=("Consolas", 9),
            state=tk.DISABLED
        )
        self.preview_text.pack(fill=tk.X)
        
        # Preview controls
        preview_controls = ttk.Frame(preview_frame)
        preview_controls.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            preview_controls,
            text="Refresh Preview",
            command=self._update_preview,
            width=15
        ).pack(side=tk.LEFT)
        
        self.preview_status = ttk.Label(
            preview_controls,
            text="",
            font=("Segoe UI", 9),
            foreground="#666666"
        )
        self.preview_status.pack(side=tk.LEFT, padx=(10, 0))
    
    def _create_export_section(self):
        """Create the export section with export button and options."""
        export_frame = ttk.LabelFrame(self.main_frame, text="Export Presentation", padding=10)
        export_frame.pack(fill=tk.X)
        
        # Export controls
        controls_frame = ttk.Frame(export_frame)
        controls_frame.pack(fill=tk.X)
        
        # Export button
        self.export_button = ttk.Button(
            controls_frame,
            text="Generate PowerPoint",
            command=self._export_presentation,
            style="Accent.TButton"
        )
        self.export_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Export status
        self.export_status = ttk.Label(
            controls_frame,
            text="Ready to export",
            font=("Segoe UI", 10),
            foreground="#666666"
        )
        self.export_status.pack(side=tk.LEFT)
        
        # Progress bar (initially hidden)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            export_frame,
            variable=self.progress_var,
            mode='determinate'
        )
        
        # Export options
        options_frame = ttk.Frame(export_frame)
        options_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Auto-open checkbox
        self.auto_open_var = tk.BooleanVar(value=True)
        auto_open_cb = ttk.Checkbutton(
            options_frame,
            text="Open presentation after export",
            variable=self.auto_open_var
        )
        auto_open_cb.pack(side=tk.LEFT)
        
        # Export format (future expansion)
        ttk.Label(options_frame, text="Format:").pack(side=tk.LEFT, padx=(20, 5))
        format_combo = ttk.Combobox(
            options_frame,
            values=["PowerPoint (.pptx)"],
            state="readonly",
            width=15
        )
        format_combo.set("PowerPoint (.pptx)")
        format_combo.pack(side=tk.LEFT)
    
    def _update_slide_options(self):
        """Update the slide selection checkboxes."""
        # Clear existing checkboxes
        for widget in self.slides_frame.winfo_children():
            widget.destroy()
        
        self.slide_checkboxes.clear()
        
        # Get available slides
        options = self.customizer.get_customization_options()
        slides = options['slides']
        
        # Create checkboxes for each slide
        for slide_id, slide_info in slides.items():
            frame = ttk.Frame(self.slides_frame)
            frame.pack(fill=tk.X, pady=2)
            
            # Checkbox variable
            var = tk.BooleanVar()
            if slide_id in self.customizer.config.get_selected_slides():
                var.set(True)
            
            self.slide_checkboxes[slide_id] = var
            
            # Checkbox
            cb = ttk.Checkbutton(
                frame,
                text=slide_info['name'],
                variable=var,
                command=self._on_slide_selection_changed
            )
            cb.pack(side=tk.LEFT, anchor=tk.W)
            
            # Required indicator
            if slide_info.get('required', False):
                req_label = ttk.Label(
                    frame,
                    text="(Required)",
                    font=("Segoe UI", 8),
                    foreground="#FF6B6B"
                )
                req_label.pack(side=tk.LEFT, padx=(5, 0))
                # Disable required slides
                cb.configure(state="disabled")
            
            # Description
            desc_label = ttk.Label(
                frame,
                text=f" - {slide_info['description']}",
                font=("Segoe UI", 9),
                foreground="#666666"
            )
            desc_label.pack(side=tk.LEFT, padx=(10, 0))
            
            # Charts indicator
            if slide_info.get('charts'):
                chart_label = ttk.Label(
                    frame,
                    text=f"📊 {len(slide_info['charts'])} chart(s)",
                    font=("Segoe UI", 8),
                    foreground="#4ECDC4"
                )
                chart_label.pack(side=tk.RIGHT)
    
    def _apply_preset(self, preset_name: str):
        """Apply a preset configuration."""
        try:
            success = self.customizer.apply_preset(preset_name)
            if success:
                self._update_slide_options()
                self._update_preview()
                self.export_status.configure(text=f"Applied {preset_name.replace('_', ' ').title()} preset")
            else:
                messagebox.showerror("Error", f"Failed to apply preset: {preset_name}")
        except Exception as e:
            logger.error(f"Failed to apply preset {preset_name}: {e}")
            messagebox.showerror("Error", f"Failed to apply preset: {e}")
    
    def _on_slide_selection_changed(self):
        """Handle slide selection changes."""
        # Get current selection
        selected_slides = [
            slide_id for slide_id, var in self.slide_checkboxes.items() 
            if var.get()
        ]
        
        # Update customizer
        self.customizer.customize_slides(selected_slides)
        
        # Update preview
        self._update_preview()
        
        # Update validation
        self._validate_current_selection()
    
    def _select_all_slides(self):
        """Select all available slides."""
        for var in self.slide_checkboxes.values():
            var.set(True)
        self._on_slide_selection_changed()
    
    def _clear_all_slides(self):
        """Clear all slide selections (except required)."""
        options = self.customizer.get_customization_options()
        
        for slide_id, var in self.slide_checkboxes.items():
            slide_info = options['slides'].get(slide_id, {})
            if not slide_info.get('required', False):
                var.set(False)
        
        self._on_slide_selection_changed()
    
    def _validate_current_selection(self):
        """Validate current slide selection and show feedback."""
        selected_slides = [
            slide_id for slide_id, var in self.slide_checkboxes.items() 
            if var.get()
        ]
        
        validation = self.customizer.validate_selection(selected_slides)
        
        # Update validation label
        if validation['valid']:
            status_text = f"✓ Valid selection: {validation['final_slide_count']} slides"
            if validation['estimated_generation_time']:
                status_text += f" (~{validation['estimated_generation_time']})"
            self.validation_label.configure(text=status_text, foreground="#28a745")
        else:
            self.validation_label.configure(text="✗ Invalid selection", foreground="#dc3545")
        
        # Show warnings if any
        if validation['warnings']:
            warning_text = "Warnings: " + "; ".join(validation['warnings'])
            # You could show this in a separate label or tooltip
    
    def _browse_template(self):
        """Browse for a PowerPoint template file."""
        file_path = filedialog.askopenfilename(
            title="Select PowerPoint Template",
            filetypes=[
                ("PowerPoint Template", "*.potx"),
                ("PowerPoint Presentation", "*.pptx"),
                ("All Files", "*.*")
            ]
        )
        
        if file_path:
            self.template_path = file_path
            self.template_var.set(os.path.basename(file_path))
            self.export_status.configure(text=f"Template selected: {os.path.basename(file_path)}")
    
    def _update_preview(self):
        """Update the preview text showing current configuration."""
        try:
            selected_slides = [
                slide_id for slide_id, var in self.slide_checkboxes.items() 
                if var.get()
            ]
            
            options = self.customizer.get_customization_options()
            validation = self.customizer.validate_selection(selected_slides)
            
            # Build preview text
            preview_lines = []
            preview_lines.append(f"Presentation Configuration Preview")
            preview_lines.append("=" * 40)
            preview_lines.append(f"Total Slides: {validation['final_slide_count']}")
            preview_lines.append(f"Estimated Generation Time: {validation['estimated_generation_time']}")
            preview_lines.append(f"Template: {self.template_var.get()}")
            preview_lines.append("")
            preview_lines.append("Selected Slides:")
            
            for slide_id in selected_slides:
                slide_info = options['slides'].get(slide_id, {})
                name = slide_info.get('name', slide_id)
                charts = slide_info.get('charts', [])
                chart_info = f" (📊 {len(charts)} chart(s))" if charts else ""
                preview_lines.append(f"  • {name}{chart_info}")
            
            if validation['warnings']:
                preview_lines.append("")
                preview_lines.append("Warnings:")
                for warning in validation['warnings']:
                    preview_lines.append(f"  ⚠ {warning}")
            
            # Update preview text
            self.preview_text.configure(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, "\n".join(preview_lines))
            self.preview_text.configure(state=tk.DISABLED)
            
            # Update status
            self.preview_status.configure(text=f"Preview updated - {len(selected_slides)} slides selected")
            
        except Exception as e:
            logger.error(f"Failed to update preview: {e}")
            self.preview_status.configure(text="Preview update failed")
    
    def _export_presentation(self):
        """Export the PowerPoint presentation with current configuration."""
        try:
            # Check if we have fleet data
            fleet_data = self.sharing_data.get('fleet')
            if not fleet_data or not hasattr(fleet_data, 'vehicles') or not fleet_data.vehicles:
                messagebox.showerror(
                    "No Data", 
                    "No fleet data available for export.\n\nPlease process vehicle data in the Process tab first."
                )
                return
            
            # Get current configuration
            config = self.customizer.get_configuration()
            
            # Validate selection
            selected_slides = config.get_selected_slides()
            validation = self.customizer.validate_selection(selected_slides)
            
            if not validation['valid']:
                messagebox.showerror("Invalid Selection", "Current slide selection is invalid. Please check your selection.")
                return
            
            # Show warnings if any
            if validation['warnings']:
                warning_msg = "Warnings about your selection:\n\n" + "\n".join(f"• {w}" for w in validation['warnings'])
                warning_msg += "\n\nDo you want to continue anyway?"
                
                if not messagebox.askyesno("Selection Warnings", warning_msg):
                    return
            
            # Disable export button
            self.export_button.configure(state="disabled", text="Generating...")
            self.progress_bar.pack(fill=tk.X, pady=(5, 0))
            self.progress_var.set(0)
            
            # Prepare export data
            export_data = {
                'fleet': fleet_data,
                'vehicles': fleet_data.vehicles,
                'fleet_name': fleet_data.name,
                'client_name': self.sharing_data.get('client_name', 'Fleet Analysis Client'),
                'stage': 'Preliminary Analysis'
            }
            
            # Start export in background thread
            def export_thread():
                try:
                    self.progress_var.set(25)
                    
                    # Export presentation
                    output_path = export_prelim_deck(
                        data=export_data,
                        template_path=self.template_path,
                        slide_config=config
                    )
                    
                    self.progress_var.set(100)
                    
                    # Update UI in main thread
                    self.root.after(0, lambda: self._export_completed(output_path))
                    
                except Exception as e:
                    logger.error(f"Export failed: {e}")
                    self.root.after(0, lambda: self._export_failed(str(e)))
            
            # Start export thread
            threading.Thread(target=export_thread, daemon=True).start()
            
        except Exception as e:
            logger.error(f"Failed to start export: {e}")
            messagebox.showerror("Export Error", f"Failed to start export: {e}")
            self._reset_export_ui()
    
    def _export_completed(self, output_path: str):
        """Handle successful export completion."""
        try:
            # Reset UI
            self._reset_export_ui()
            
            # Update status
            file_size = os.path.getsize(output_path)
            self.export_status.configure(
                text=f"✓ Export completed: {os.path.basename(output_path)} ({file_size:,} bytes)"
            )
            
            # Show success message
            msg = f"PowerPoint presentation exported successfully!\n\nFile: {output_path}\nSize: {file_size:,} bytes"
            
            if self.auto_open_var.get():
                msg += "\n\nOpening presentation..."
                messagebox.showinfo("Export Successful", msg)
                
                # Try to open the file
                try:
                    if os.name == 'nt':  # Windows
                        os.startfile(output_path)
                    elif os.name == 'posix':  # macOS/Linux
                        os.system(f'open "{output_path}"')
                except Exception as e:
                    logger.warning(f"Could not open presentation: {e}")
            else:
                messagebox.showinfo("Export Successful", msg)
            
        except Exception as e:
            logger.error(f"Error in export completion: {e}")
    
    def _export_failed(self, error_message: str):
        """Handle export failure."""
        self._reset_export_ui()
        self.export_status.configure(text="✗ Export failed")
        messagebox.showerror("Export Failed", f"Failed to export presentation:\n\n{error_message}")
    
    def _reset_export_ui(self):
        """Reset the export UI to normal state."""
        self.export_button.configure(state="normal", text="Generate PowerPoint")
        self.progress_bar.pack_forget()
        self.progress_var.set(0)
    
    def refresh_data(self):
        """Refresh the panel when new data is available."""
        try:
            fleet_data = self.sharing_data.get('fleet')
            
            if fleet_data and hasattr(fleet_data, 'vehicles'):
                vehicle_count = len(fleet_data.vehicles)
                self.export_status.configure(
                    text=f"Ready to export - {vehicle_count} vehicles available"
                )
                self.export_button.configure(state="normal")
            else:
                self.export_status.configure(text="No fleet data available")
                self.export_button.configure(state="disabled")
            
            # Update preview
            self._update_preview()
            
        except Exception as e:
            logger.error(f"Failed to refresh Present panel data: {e}")
    
    def get_panel_frame(self) -> ttk.Frame:
        """Get the main panel frame."""
        return self.main_frame

# Helper function for integration
def create_present_panel(parent_frame: ttk.Frame, sharing_data: dict) -> PresentPanel:
    """
    Create and return a PresentPanel instance.
    
    Args:
        parent_frame: Parent frame to contain the panel
        sharing_data: Shared data dictionary
        
    Returns:
        PresentPanel instance
    """
    return PresentPanel(parent_frame, sharing_data)
