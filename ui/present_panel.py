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
from utils import SimpleTooltip, ProgressDialog, ScrollableFrame
from powerpoint_export import export_prelim_deck
from ui.theme import Colors, Fonts, Spacing
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
        
        # Create scrollable container
        self._sf_main = ScrollableFrame(self.parent_frame)
        self._sf_main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.main_frame = self._sf_main.scrollable_frame
        
        # Initialize UI components
        self._create_header()
        self._create_details_section()
        self._create_vehicle_filter_section()
        self._create_preset_section()
        self._create_customization_section()
        self._create_scenario_section()
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
    
    def _create_details_section(self):
        """Create editable presentation details: client name, subtitle, stage."""
        details_frame = ttk.LabelFrame(self.main_frame, text="Presentation Details", padding=10)
        details_frame.pack(fill=tk.X, pady=(0, 15))

        # Row 1: Client name
        row1 = ttk.Frame(details_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="Client Name:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.client_name_var = tk.StringVar(value="")
        ttk.Entry(row1, textvariable=self.client_name_var, width=50).pack(
            side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True
        )

        # Row 2: Subtitle / context note
        row2 = ttk.Frame(details_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="Subtitle:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.subtitle_var = tk.StringVar(value="Fleet Electrification Analysis")
        ttk.Entry(row2, textvariable=self.subtitle_var, width=50).pack(
            side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True
        )

        # Row 3: Stage
        row3 = ttk.Frame(details_frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="Stage:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.stage_var = tk.StringVar(value="Preliminary Analysis")
        stage_combo = ttk.Combobox(
            row3, textvariable=self.stage_var, width=30,
            values=["Preliminary Analysis", "Final Analysis", "Draft", "Revised"],
            state="normal"
        )
        stage_combo.pack(side=tk.LEFT, padx=(5, 0))

        hint = ttk.Label(
            details_frame,
            text="These fields appear on the cover slide and headers",
            font=("Segoe UI", 9), foreground="#888888"
        )
        hint.pack(anchor=tk.W, pady=(5, 0))

    def _create_vehicle_filter_section(self):
        """Create vehicle filter controls for subsetting which vehicles go into slides."""
        filter_frame = ttk.LabelFrame(self.main_frame, text="Vehicle Filter (optional)", padding=10)
        filter_frame.pack(fill=tk.X, pady=(0, 15))

        # Department filter
        dept_row = ttk.Frame(filter_frame)
        dept_row.pack(fill=tk.X, pady=2)
        ttk.Label(dept_row, text="Department:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.dept_filter_var = tk.StringVar(value="All Departments")
        self.dept_filter_combo = ttk.Combobox(
            dept_row, textvariable=self.dept_filter_var, width=35,
            values=["All Departments"], state="readonly"
        )
        self.dept_filter_combo.pack(side=tk.LEFT, padx=(5, 0))

        # ACF category filter
        acf_row = ttk.Frame(filter_frame)
        acf_row.pack(fill=tk.X, pady=2)
        ttk.Label(acf_row, text="ACF Category:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.acf_filter_var = tk.StringVar(value="All Categories")
        self.acf_filter_combo = ttk.Combobox(
            acf_row, textvariable=self.acf_filter_var, width=35,
            values=["All Categories"], state="readonly"
        )
        self.acf_filter_combo.pack(side=tk.LEFT, padx=(5, 0))

        # Payback filter
        payback_row = ttk.Frame(filter_frame)
        payback_row.pack(fill=tk.X, pady=2)
        ttk.Label(payback_row, text="Max Payback:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.payback_filter_var = tk.StringVar(value="No Limit")
        ttk.Combobox(
            payback_row, textvariable=self.payback_filter_var, width=35,
            values=["No Limit", "< 3 years", "< 5 years", "< 7 years", "< 10 years"],
            state="readonly"
        ).pack(side=tk.LEFT, padx=(5, 0))

        # Filter summary label
        self.filter_summary_label = ttk.Label(
            filter_frame, text="", font=("Segoe UI", 9), foreground="#666666"
        )
        self.filter_summary_label.pack(anchor=tk.W, pady=(5, 0))

    def _get_filtered_vehicles(self):
        """Return the subset of fleet vehicles that pass current filter settings."""
        fleet_data = self.sharing_data.get('fleet')
        if not fleet_data or not hasattr(fleet_data, 'vehicles'):
            return []

        vehicles = list(fleet_data.vehicles)

        # Department filter
        dept = self.dept_filter_var.get()
        if dept and dept != "All Departments":
            vehicles = [
                v for v in vehicles
                if getattr(v, 'custom_fields', {}).get('Department', '') == dept
                or getattr(v, 'fleet_management_fields', {}).get('department', '') == dept
            ]

        # ACF category filter
        acf = self.acf_filter_var.get()
        if acf and acf != "All Categories":
            vehicles = [
                v for v in vehicles
                if getattr(v, 'custom_fields', {}).get('ACF Category', '') == acf
            ]

        # Payback filter
        payback = self.payback_filter_var.get()
        if payback and payback != "No Limit":
            try:
                max_years = int(payback.split("<")[1].split("year")[0].strip())
                filtered = []
                for v in vehicles:
                    pb = getattr(v, 'custom_fields', {}).get('_payback_years')
                    if pb is None:
                        filtered.append(v)  # include vehicles without payback data
                    elif pb < max_years:
                        filtered.append(v)
                vehicles = filtered
            except (ValueError, IndexError):
                pass

        return vehicles

    def _populate_filter_dropdowns(self):
        """Populate department and ACF filter dropdowns from current fleet data."""
        fleet_data = self.sharing_data.get('fleet')
        if not fleet_data or not hasattr(fleet_data, 'vehicles'):
            return

        departments = set()
        acf_categories = set()
        for v in fleet_data.vehicles:
            cf = getattr(v, 'custom_fields', {})
            dept = cf.get('Department', '') or getattr(v, 'fleet_management_fields', {}).get('department', '')
            if dept:
                departments.add(dept)
            acf = cf.get('ACF Category', '')
            if acf:
                acf_categories.add(acf)

        self.dept_filter_combo['values'] = ["All Departments"] + sorted(departments)
        self.acf_filter_combo['values'] = ["All Categories"] + sorted(acf_categories)

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
        sf_slides = ScrollableFrame(custom_frame)
        sf_slides.pack(fill=tk.BOTH, expand=True)
        sf_slides.canvas.configure(height=200)

        # Store references
        self.slides_frame = sf_slides.scrollable_frame
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
    
    def _create_scenario_section(self):
        """Create scenario selector for the Scenario Comparison slide.

        Only visible when the 'scenario_comparison' slide is in the selected slides.
        Lets consultants choose which of the 4 preset scenarios to include.
        """
        self._scenario_section_frame = ttk.LabelFrame(
            self.main_frame,
            text="Scenario Comparison — Select Scenarios to Include",
            padding=10
        )
        self._scenario_section_frame.pack(fill=tk.X, pady=(0, 15))

        hint = ttk.Label(
            self._scenario_section_frame,
            text="Choose which electrification scenarios appear on the Scenario Comparison slide.\n"
                 "At least one scenario must be selected. (Only applies when that slide is included.)",
            font=("Segoe UI", 9),
            foreground="#555555",
            justify=tk.LEFT,
            wraplength=600,
        )
        hint.pack(anchor=tk.W, pady=(0, 8))

        # Scenario definitions: (display_name, scenario_key, description)
        SCENARIO_DEFS = [
            ("Aggressive (2030)",      "aggressive",      "Replace all eligible vehicles by 2030 — maximum speed"),
            ("Moderate (2035)",        "moderate",        "Balanced timeline, most fleets use this as their baseline"),
            ("Conservative (2040)",    "conservative",    "Gradual transition — minimises annual budget impact"),
            ("ACF Compliance Only",    "acf_compliance",  "Targets only CARB-regulated vehicles; ignores light-duty"),
        ]

        self._scenario_vars: dict = {}
        checkboxes_frame = ttk.Frame(self._scenario_section_frame)
        checkboxes_frame.pack(fill=tk.X)

        for key, (display_name, scenario_key, description) in enumerate(SCENARIO_DEFS):
            row_frame = ttk.Frame(checkboxes_frame)
            row_frame.pack(fill=tk.X, pady=2)

            var = tk.BooleanVar(value=True)  # all selected by default
            self._scenario_vars[scenario_key] = var

            cb = ttk.Checkbutton(
                row_frame,
                text=display_name,
                variable=var,
                command=self._on_scenario_selection_changed,
                width=22,
            )
            cb.pack(side=tk.LEFT)

            desc_lbl = ttk.Label(
                row_frame,
                text=f"— {description}",
                font=("Segoe UI", 9),
                foreground="#666666",
            )
            desc_lbl.pack(side=tk.LEFT, padx=(5, 0))

        # Validation label
        self._scenario_validation_label = ttk.Label(
            self._scenario_section_frame,
            text="",
            font=("Segoe UI", 9),
            foreground="#CC3300",
        )
        self._scenario_validation_label.pack(anchor=tk.W, pady=(6, 0))

    def _on_scenario_selection_changed(self):
        """Validate that at least one scenario is selected."""
        selected = [k for k, v in self._scenario_vars.items() if v.get()]
        if not selected:
            self._scenario_validation_label.configure(
                text="⚠ At least one scenario must be selected."
            )
        else:
            self._scenario_validation_label.configure(text="")

    def _get_selected_scenarios(self) -> list:
        """Return list of selected scenario keys; falls back to all if none checked."""
        selected = [k for k, v in self._scenario_vars.items() if v.get()]
        if not selected:
            return ["aggressive", "moderate", "conservative", "acf_compliance"]
        return selected

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
            height=10,
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
        
        # Export button (PRIMARY GREEN - main action)
        self.export_button = ttk.Button(
            controls_frame,
            text="📊 Generate PowerPoint",
            command=self._export_presentation,
            style="Primary.TButton"
        )
        self.export_button.pack(side=tk.LEFT, padx=(0, Spacing.MARGIN_ELEMENT))
        SimpleTooltip(self.export_button, "Generate professional PowerPoint presentation\n• Editable native charts with real fleet data\n• Customized slides based on your selection\n• Ready for stakeholder presentations")
        
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
        
        # Format dropdown removed — only PowerPoint is supported
    
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
        """Update the preview text showing current configuration with data-aware content."""
        try:
            selected_slides = [
                slide_id for slide_id, var in self.slide_checkboxes.items()
                if var.get()
            ]

            options = self.customizer.get_customization_options()
            validation = self.customizer.validate_selection(selected_slides)

            # Get filtered vehicle count for data-aware preview
            filtered = self._get_filtered_vehicles()
            vehicle_count = len(filtered)
            filter_active = (self.dept_filter_var.get() != "All Departments"
                             or self.acf_filter_var.get() != "All Categories"
                             or self.payback_filter_var.get() != "No Limit")

            # Build data-aware slide descriptions
            slide_context = self._build_slide_context(filtered)

            # Build preview text
            lines = []
            client = self.client_name_var.get().strip() or "(no client name)"
            lines.append(f"  Client:    {client}")
            lines.append(f"  Subtitle:  {self.subtitle_var.get()}")
            lines.append(f"  Stage:     {self.stage_var.get()}")
            lines.append(f"  Template:  {self.template_var.get()}")
            if filter_active:
                lines.append(f"  Vehicles:  {vehicle_count} (filtered)")
            elif vehicle_count > 0:
                lines.append(f"  Vehicles:  {vehicle_count}")
            else:
                lines.append(f"  Vehicles:  No fleet data loaded")
            lines.append("")
            lines.append(f"  Slides ({validation['final_slide_count']}):")

            for slide_id in selected_slides:
                slide_info = options['slides'].get(slide_id, {})
                name = slide_info.get('name', slide_id)
                charts = slide_info.get('charts', [])
                chart_tag = f"  [{len(charts)} chart]" if charts else ""
                context = slide_context.get(slide_id, "")
                detail = f" — {context}" if context else ""
                lines.append(f"    {name}{chart_tag}{detail}")

            if validation.get('warnings'):
                lines.append("")
                for warning in validation['warnings']:
                    lines.append(f"  ! {warning}")

            # Update filter summary
            if filter_active:
                parts = []
                if self.dept_filter_var.get() != "All Departments":
                    parts.append(self.dept_filter_var.get())
                if self.acf_filter_var.get() != "All Categories":
                    parts.append(self.acf_filter_var.get())
                if self.payback_filter_var.get() != "No Limit":
                    parts.append(f"payback {self.payback_filter_var.get()}")
                self.filter_summary_label.configure(
                    text=f"Filter active: {', '.join(parts)} ({vehicle_count} vehicles)"
                )
            else:
                self.filter_summary_label.configure(text="")

            # Update preview text widget
            self.preview_text.configure(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, "\n".join(lines))
            self.preview_text.configure(state=tk.DISABLED)

            self.preview_status.configure(text=f"Preview updated — {len(selected_slides)} slides selected")

        except Exception as e:
            logger.error(f"Failed to update preview: {e}")
            self.preview_status.configure(text="Preview update failed")

    def _build_slide_context(self, vehicles):
        """Build per-slide context strings from vehicle data for preview."""
        ctx = {}
        if not vehicles:
            return ctx

        n = len(vehicles)
        success = sum(1 for v in vehicles if getattr(v, 'processing_success', False))

        # Avg MPG
        mpg_vals = []
        for v in vehicles:
            fe = getattr(v, 'fuel_economy', None)
            if fe:
                comb = getattr(fe, 'combined_mpg', 0) or 0
                if comb > 0:
                    mpg_vals.append(comb)
        avg_mpg = sum(mpg_vals) / len(mpg_vals) if mpg_vals else 0

        # Departments
        depts = set()
        acf_b_count = 0
        ev_years = {}
        for v in vehicles:
            cf = getattr(v, 'custom_fields', {})
            d = cf.get('Department', '') or getattr(v, 'fleet_management_fields', {}).get('department', '')
            if d:
                depts.add(d)
            if cf.get('_acf_code') == 'B':
                acf_b_count += 1
            yr = cf.get('Proposed EV Year', '')
            if yr and yr not in ('N/A', 'Exempt', ''):
                try:
                    ev_years[int(yr)] = ev_years.get(int(yr), 0) + 1
                except (ValueError, TypeError):
                    pass

        ctx['cover'] = f"{n} vehicles, {self.client_name_var.get().strip() or 'client TBD'}"
        ctx['fleet_snapshot'] = f"{n} vehicles, {len(depts)} departments, avg {avg_mpg:.1f} MPG" if avg_mpg else f"{n} vehicles"
        ctx['fleet_composition'] = f"Body class & make breakdown of {n} vehicles"
        ctx['financial_summary'] = "TCO comparison, payback timeline"
        ctx['emissions_timeline'] = "CO2 reduction projection by replacement year"
        ctx['emissions_by_weight'] = "Emissions by weight class"
        ctx['electrification_timeline_weight'] = f"Replacements by year & weight class"
        ctx['electrification_timeline_body'] = f"Replacements by year & body type"
        if acf_b_count:
            ctx['executive_recommendations'] = f"{acf_b_count} ACF-regulated vehicles, top 5 priorities"
        else:
            ctx['executive_recommendations'] = "Top 5 priority replacements"
        ctx['replacement_schedule'] = f"Top 12 vehicles sorted by target year"
        ctx['scenario_comparison'] = "4 electrification scenarios compared"
        ctx['age_analysis'] = "Fleet age distribution"
        ctx['data_quality'] = f"{success}/{n} successfully processed"
        ctx['next_steps'] = "Data-driven action items"

        return ctx
    
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
            
            # Prompt user for save location
            fleet_name = fleet_data.name or "fleet"
            safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in fleet_name).strip()
            default_filename = f"{safe_name}_presentation.pptx"

            user_out_path = filedialog.asksaveasfilename(
                title="Save PowerPoint Presentation",
                defaultextension=".pptx",
                filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")],
                initialfile=default_filename
            )
            if not user_out_path:
                return  # User cancelled

            # Disable export button
            self.export_button.configure(state="disabled", text="Generating...")
            self.progress_bar.pack(fill=tk.X, pady=(5, 0))
            self.progress_var.set(0)

            # Apply vehicle filter
            filtered_vehicles = self._get_filtered_vehicles()
            if not filtered_vehicles:
                messagebox.showerror(
                    "No Vehicles",
                    "No vehicles match the current filter settings.\n\n"
                    "Adjust the filters in the Vehicle Filter section or select 'All'."
                )
                self._reset_export_ui()
                return

            # Prepare export data with user-editable fields
            client_name = self.client_name_var.get().strip() or fleet_data.name or 'Fleet Analysis Client'
            subtitle = self.subtitle_var.get().strip()
            stage = self.stage_var.get().strip() or 'Preliminary Analysis'
            export_data = {
                'fleet': fleet_data,
                'vehicles': filtered_vehicles,
                'fleet_name': fleet_data.name,
                'client_name': client_name,
                'stage': stage,
                'subtitle': subtitle if subtitle else f"{stage} Fleet Electrification Analysis",
                'selected_scenarios': self._get_selected_scenarios(),
            }

            # Start export in background thread
            def export_thread():
                try:
                    self.progress_var.set(25)

                    # Export presentation to user-chosen path
                    output_path = export_prelim_deck(
                        data=export_data,
                        template_path=self.template_path,
                        out_path=user_out_path,
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
        self.export_button.configure(state="normal", text="📊 Generate PowerPoint")
        self.progress_bar.pack_forget()
        self.progress_var.set(0)
    
    def refresh_data(self):
        """Refresh the panel when new data is available."""
        try:
            fleet_data = self.sharing_data.get('fleet')

            if fleet_data and hasattr(fleet_data, 'vehicles'):
                vehicle_count = len(fleet_data.vehicles)
                self.export_status.configure(
                    text=f"Ready to export — {vehicle_count} vehicles available"
                )
                self.export_button.configure(state="normal")

                # Populate filter dropdowns from fleet data
                self._populate_filter_dropdowns()

                # Pre-fill client name from fleet name if empty
                if not self.client_name_var.get().strip() and fleet_data.name:
                    self.client_name_var.set(fleet_data.name)
            else:
                self.export_status.configure(text="No fleet data available")
                self.export_button.configure(state="disabled")

            # Update preview
            self._update_preview()

        except Exception as e:
            logger.error(f"Failed to refresh Present panel data: {e}")
    
    def get_panel_frame(self) -> ttk.Frame:
        """Get the main panel frame."""
        # Return the parent frame that contains the scrollable canvas
        return self.parent_frame

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
