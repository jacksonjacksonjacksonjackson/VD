I'm rebuilding my Fleet Electrification Analyzer application with an improved structure. Please write complete implementation files according to this structure:

fleet_analyzer/
├── app.py                      # Main entry point and application setup
├── settings.py                 # All configuration in one place
├── data/
│   ├── models.py               # Data models and structures
│   ├── providers.py            # Data retrieval services (API clients)
│   └── processor.py            # Data processing pipeline
├── analysis/
│   ├── calculations.py         # Core analysis calculations
│   ├── charts.py               # Chart generation and visualization
│   └── reports.py              # Report generation and exports
├── ui/
│   ├── main_window.py          # Main application window
│   ├── process_panel.py        # Data processing panel
│   ├── results_panel.py        # Results display panel
│   └── analysis_panel.py       # Analysis tools panel
└── utils.py                    # Utility functions and helpers

The application is a tool for analyzing fleet vehicles with these core features:
1. VIN decoding to identify vehicles
2. Fuel economy data retrieval
3. Fleet analysis, visualization, and reporting
4. Fleet modeling (electrification timeline, emissions inventory, TCO, etc..)
5. Fleet electriciation
6. Charging needs modeling


Key improvements to implement:
- Improved user customizability and control
- User field selection
- data filtering
- improved vehicle data enrichment and matching critera
- Better data models with validation
- Improved API clients with caching and error handling
- Enhanced threading for parallel processing
- Better separation of UI from business logic
- More modular structure for easier feature additions
- Enhanced visualization capabilities
- Improved error handling throughout
- Thread-safe caching mechanism
- Enhanced exporters for different formats
- More polished UI with better organization
- User should be able to upload custom fields with the VIN numbers, like asset ID.
- All possible vehicle data is stored for future use, but not shown unless the user selects the additional fields
- Core vehicle data table/spreadsheet should be more usable for analysis (sorting, filtering, adding mpg information for vehicles missing it)

Ensure the implementation maintains compatibility with the original application's data and workflows while adding these new features.

Please generate each file completely, one at a time, with full implementations. Don't skip any important code, and ensure the files work together cohesively. Focus on production-quality, maintainable code with good documentation. After you provide each file, I'll confirm before you proceed to the next one.

First, please create settings.py with all necessary configuration settings from the original app.