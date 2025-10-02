"""
ui/theme.py

Professional UI theme for Fleet Electrification Analyzer.
Provides consistent colors, fonts, spacing for polish and professionalism.
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict, Tuple, Optional

###############################################################################
# Color Palette - Professional Engineering Tool
###############################################################################

class Colors:
    """Modern color palette with light/dark variants."""
    
    # Primary colors (from existing settings)
    PRIMARY_DARK = "#3C465A"      # Charcoal - main text, headers
    PRIMARY_LIGHT = "#FFFFFF"     # White - backgrounds
    PRIMARY_GREEN = "#6B9E78"     # Reseda green - accents, success
    
    # Secondary colors
    SECONDARY_ORANGE = "#D45D1E"  # Deep orange - warnings, highlights
    SECONDARY_GREY = "#DEDEE0"    # Light grey - borders, disabled
    
    # Extended palette for UI
    BACKGROUND = "#F5F7FA"        # Soft blue-grey background
    SURFACE = "#FFFFFF"           # Card/panel surface
    SURFACE_HOVER = "#F8F9FB"     # Hover state
    BORDER = "#E1E4E8"            # Default borders
    BORDER_FOCUS = "#6B9E78"      # Focused input borders
    
    # Text hierarchy
    TEXT_PRIMARY = "#2D3748"      # Main text
    TEXT_SECONDARY = "#4A5568"    # Secondary text
    TEXT_TERTIARY = "#718096"     # Tertiary text, hints
    TEXT_DISABLED = "#A0AEC0"     # Disabled text
    
    # Status colors
    SUCCESS = "#48BB78"           # Success states
    WARNING = "#ED8936"           # Warning states
    ERROR = "#F56565"             # Error states
    INFO = "#4299E1"              # Info states
    
    # Button variants
    BTN_PRIMARY = "#6B9E78"       # Primary action buttons
    BTN_PRIMARY_HOVER = "#5A8A66"
    BTN_PRIMARY_ACTIVE = "#4A7555"
    
    BTN_SECONDARY = "#4A5568"     # Secondary actions
    BTN_SECONDARY_HOVER = "#2D3748"
    BTN_SECONDARY_ACTIVE = "#1A202C"
    
    BTN_DANGER = "#F56565"        # Destructive actions
    BTN_DANGER_HOVER = "#E53E3E"
    BTN_DANGER_ACTIVE = "#C53030"
    
    BTN_GHOST = "transparent"     # Ghost/text buttons
    BTN_GHOST_HOVER = "#F7FAFC"

###############################################################################
# Typography
###############################################################################

class Fonts:
    """Consistent font hierarchy."""
    
    # Font families (fallback to system defaults)
    FAMILY_SANS = "Segoe UI"      # Windows/Linux
    FAMILY_MONO = "Consolas"      # Monospace
    
    # Font sizes
    SIZE_DISPLAY = 24             # Page titles
    SIZE_H1 = 20                  # Section headers
    SIZE_H2 = 16                  # Subsection headers
    SIZE_H3 = 14                  # Card titles
    SIZE_BODY = 11                # Body text
    SIZE_SMALL = 10               # Small text, captions
    SIZE_TINY = 9                 # Tiny text, footnotes
    
    # Font weights
    WEIGHT_NORMAL = "normal"
    WEIGHT_BOLD = "bold"
    
    @staticmethod
    def configure(style: ttk.Style):
        """Configure tkinter font settings."""
        # Detect platform for font families
        import platform
        if platform.system() == "Darwin":  # macOS
            Fonts.FAMILY_SANS = "SF Pro Text"
            Fonts.FAMILY_MONO = "SF Mono"
        elif platform.system() == "Linux":
            Fonts.FAMILY_SANS = "Ubuntu"
            Fonts.FAMILY_MONO = "Ubuntu Mono"

###############################################################################
# Spacing System
###############################################################################

class Spacing:
    """Consistent spacing scale (8pt grid system)."""
    
    XXS = 2   # 2px  - Tiny gaps
    XS = 4    # 4px  - Minimal spacing
    SM = 8    # 8px  - Small spacing
    MD = 12   # 12px - Medium spacing
    LG = 16   # 16px - Large spacing
    XL = 24   # 24px - Extra large spacing
    XXL = 32  # 32px - Section spacing
    XXXL = 48 # 48px - Page spacing
    
    # Common padding presets
    PADDING_NONE = 0
    PADDING_TIGHT = 4
    PADDING_NORMAL = 8
    PADDING_COMFORTABLE = 12
    PADDING_SPACIOUS = 16
    
    # Common margins
    MARGIN_SECTION = 16
    MARGIN_SUBSECTION = 12
    MARGIN_ELEMENT = 8

###############################################################################
# Component Styles
###############################################################################

class Styles:
    """Configure ttk styles for modern appearance."""
    
    @staticmethod
    def configure_theme(style: ttk.Style):
        """Apply modern theme to ttk widgets."""
        
        # Configure base theme
        style.theme_use('clam')  # Start with clam theme
        
        # Configure fonts
        Fonts.configure(style)
        
        # Frame styles
        style.configure(
            "TFrame",
            background=Colors.BACKGROUND
        )
        
        style.configure(
            "Card.TFrame",
            background=Colors.SURFACE,
            relief="flat",
            borderwidth=1,
            bordercolor=Colors.BORDER
        )
        
        # Label styles
        style.configure(
            "TLabel",
            background=Colors.BACKGROUND,
            foreground=Colors.TEXT_PRIMARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY)
        )
        
        style.configure(
            "Title.TLabel",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_DISPLAY, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY
        )
        
        style.configure(
            "Heading1.TLabel",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H1, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY
        )
        
        style.configure(
            "Heading2.TLabel",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H2, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY
        )
        
        style.configure(
            "Heading3.TLabel",
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD),
            foreground=Colors.TEXT_PRIMARY
        )
        
        style.configure(
            "Secondary.TLabel",
            foreground=Colors.TEXT_SECONDARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY)
        )
        
        style.configure(
            "Caption.TLabel",
            foreground=Colors.TEXT_TERTIARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL)
        )
        
        # Button styles
        style.configure(
            "Primary.TButton",
            background=Colors.BTN_PRIMARY,
            foreground=Colors.PRIMARY_LIGHT,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
            borderwidth=0,
            focuscolor=Colors.BORDER_FOCUS,
            padding=(Spacing.LG, Spacing.SM)
        )
        
        style.map(
            "Primary.TButton",
            background=[
                ('active', Colors.BTN_PRIMARY_ACTIVE),
                ('pressed', Colors.BTN_PRIMARY_ACTIVE),
                ('hover', Colors.BTN_PRIMARY_HOVER)
            ]
        )
        
        style.configure(
            "Secondary.TButton",
            background=Colors.BTN_SECONDARY,
            foreground=Colors.PRIMARY_LIGHT,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
            borderwidth=0,
            padding=(Spacing.MD, Spacing.SM)
        )
        
        style.map(
            "Secondary.TButton",
            background=[
                ('active', Colors.BTN_SECONDARY_ACTIVE),
                ('pressed', Colors.BTN_SECONDARY_ACTIVE),
                ('hover', Colors.BTN_SECONDARY_HOVER)
            ]
        )
        
        style.configure(
            "Danger.TButton",
            background=Colors.BTN_DANGER,
            foreground=Colors.PRIMARY_LIGHT,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
            borderwidth=0,
            padding=(Spacing.MD, Spacing.SM)
        )
        
        # Entry styles
        style.configure(
            "TEntry",
            fieldbackground=Colors.SURFACE,
            foreground=Colors.TEXT_PRIMARY,
            bordercolor=Colors.BORDER,
            lightcolor=Colors.BORDER_FOCUS,
            darkcolor=Colors.BORDER,
            insertcolor=Colors.TEXT_PRIMARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
            padding=Spacing.SM
        )
        
        # Combobox styles
        style.configure(
            "TCombobox",
            fieldbackground=Colors.SURFACE,
            background=Colors.SURFACE,
            foreground=Colors.TEXT_PRIMARY,
            bordercolor=Colors.BORDER,
            arrowcolor=Colors.TEXT_SECONDARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
            padding=Spacing.SM
        )
        
        # LabelFrame styles
        style.configure(
            "TLabelframe",
            background=Colors.SURFACE,
            bordercolor=Colors.BORDER,
            relief="solid",
            borderwidth=1
        )
        
        style.configure(
            "TLabelframe.Label",
            background=Colors.SURFACE,
            foreground=Colors.TEXT_PRIMARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_H3, Fonts.WEIGHT_BOLD)
        )
        
        # Notebook (tabs) styles
        style.configure(
            "TNotebook",
            background=Colors.BACKGROUND,
            borderwidth=0
        )
        
        style.configure(
            "TNotebook.Tab",
            background=Colors.SURFACE,
            foreground=Colors.TEXT_SECONDARY,
            padding=(Spacing.LG, Spacing.MD),
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY)
        )
        
        style.map(
            "TNotebook.Tab",
            background=[
                ('selected', Colors.SURFACE),
                ('active', Colors.SURFACE_HOVER)
            ],
            foreground=[
                ('selected', Colors.PRIMARY_GREEN),
                ('active', Colors.TEXT_PRIMARY)
            ]
        )
        
        # Treeview styles
        style.configure(
            "Treeview",
            background=Colors.SURFACE,
            foreground=Colors.TEXT_PRIMARY,
            fieldbackground=Colors.SURFACE,
            bordercolor=Colors.BORDER,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY),
            rowheight=28
        )
        
        style.configure(
            "Treeview.Heading",
            background=Colors.BACKGROUND,
            foreground=Colors.TEXT_PRIMARY,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_BODY, Fonts.WEIGHT_BOLD),
            relief="flat"
        )
        
        style.map(
            "Treeview",
            background=[('selected', Colors.PRIMARY_GREEN)],
            foreground=[('selected', Colors.PRIMARY_LIGHT)]
        )
        
        # Progressbar styles
        style.configure(
            "TProgressbar",
            background=Colors.PRIMARY_GREEN,
            troughcolor=Colors.SURFACE,
            bordercolor=Colors.BORDER,
            lightcolor=Colors.PRIMARY_GREEN,
            darkcolor=Colors.PRIMARY_GREEN,
            thickness=8
        )
        
        # Scrollbar styles
        style.configure(
            "TScrollbar",
            background=Colors.BORDER,
            troughcolor=Colors.BACKGROUND,
            bordercolor=Colors.BACKGROUND,
            arrowcolor=Colors.TEXT_SECONDARY
        )

###############################################################################
# Layout Helpers
###############################################################################

class Layout:
    """Common layout patterns and helpers."""
    
    @staticmethod
    def create_card(parent, title: Optional[str] = None, padding: int = Spacing.PADDING_NORMAL) -> ttk.Frame:
        """
        Create a card-style frame with optional title.
        
        Args:
            parent: Parent widget
            title: Optional title for the card
            padding: Internal padding
            
        Returns:
            Card frame
        """
        if title:
            card = ttk.LabelFrame(parent, text=title, style="TLabelframe", padding=padding)
        else:
            card = ttk.Frame(parent, style="Card.TFrame", padding=padding)
        
        return card
    
    @staticmethod
    def create_scrollable_frame(parent, height: Optional[int] = None) -> Tuple[tk.Canvas, ttk.Frame]:
        """
        Create a scrollable frame using canvas and scrollbar.
        
        Args:
            parent: Parent widget
            height: Optional fixed height for canvas
            
        Returns:
            Tuple of (canvas, scrollable_frame)
        """
        # Container
        container = ttk.Frame(parent)
        
        # Canvas and scrollbar
        canvas = tk.Canvas(container, bg=Colors.BACKGROUND, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Configure scrolling
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Set height if specified
        if height:
            canvas.configure(height=height)
        
        # Pack
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        container.pack(fill="both", expand=True)
        
        # Enable mousewheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        return canvas, scrollable_frame
    
    @staticmethod
    def create_sticky_footer(parent, height: int = 60) -> ttk.Frame:
        """
        Create a sticky footer frame for actions.
        
        Args:
            parent: Parent widget
            height: Height of footer
            
        Returns:
            Footer frame
        """
        footer = ttk.Frame(parent, height=height, style="Card.TFrame")
        footer.pack(side=tk.BOTTOM, fill=tk.X, padx=0, pady=0)
        footer.pack_propagate(False)  # Maintain fixed height
        
        # Add a subtle top border
        separator = ttk.Separator(footer, orient='horizontal')
        separator.pack(side=tk.TOP, fill=tk.X)
        
        return footer
    
    @staticmethod
    def create_filter_bar(parent) -> ttk.Frame:
        """
        Create a horizontal filter bar.
        
        Args:
            parent: Parent widget
            
        Returns:
            Filter bar frame
        """
        filter_bar = ttk.Frame(parent, style="Card.TFrame", padding=Spacing.PADDING_NORMAL)
        filter_bar.pack(side=tk.TOP, fill=tk.X, padx=Spacing.MARGIN_ELEMENT, pady=Spacing.MARGIN_ELEMENT)
        
        return filter_bar
    
    @staticmethod
    def create_button_group(parent, orientation: str = "horizontal", spacing: int = Spacing.SM) -> ttk.Frame:
        """
        Create a button group container.
        
        Args:
            parent: Parent widget
            orientation: "horizontal" or "vertical"
            spacing: Spacing between buttons
            
        Returns:
            Button group frame
        """
        frame = ttk.Frame(parent)
        frame.spacing = spacing
        frame.orientation = orientation
        
        return frame
    
    @staticmethod
    def add_separator(parent, orient: str = "horizontal", pady: int = Spacing.MD):
        """Add a visual separator."""
        sep = ttk.Separator(parent, orient=orient)
        if orient == "horizontal":
            sep.pack(fill=tk.X, pady=pady)
        else:
            sep.pack(fill=tk.Y, padx=pady)
        return sep

###############################################################################
# Status Badge Helper
###############################################################################

class StatusBadge:
    """Create status badges with consistent styling."""
    
    COLORS = {
        "success": (Colors.SUCCESS, Colors.PRIMARY_LIGHT),
        "warning": (Colors.WARNING, Colors.PRIMARY_LIGHT),
        "error": (Colors.ERROR, Colors.PRIMARY_LIGHT),
        "info": (Colors.INFO, Colors.PRIMARY_LIGHT),
        "default": (Colors.BORDER, Colors.TEXT_PRIMARY)
    }
    
    @staticmethod
    def create(parent, text: str, status: str = "default") -> ttk.Label:
        """
        Create a status badge label.
        
        Args:
            parent: Parent widget
            text: Badge text
            status: Badge status (success, warning, error, info, default)
            
        Returns:
            Badge label
        """
        bg_color, fg_color = StatusBadge.COLORS.get(status, StatusBadge.COLORS["default"])
        
        badge = tk.Label(
            parent,
            text=text,
            background=bg_color,
            foreground=fg_color,
            font=(Fonts.FAMILY_SANS, Fonts.SIZE_SMALL, Fonts.WEIGHT_BOLD),
            padx=Spacing.SM,
            pady=Spacing.XXS,
            relief="flat"
        )
        
        return badge

###############################################################################
# Initialization
###############################################################################

def initialize_theme():
    """Initialize the modern UI theme."""
    style = ttk.Style()
    Styles.configure_theme(style)
    return style

