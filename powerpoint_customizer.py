"""
powerpoint_customizer.py

User interface for customizing PowerPoint presentations.
Allows users to select which slides and charts to include.
"""

import logging
from typing import Dict, List, Any, Optional
from powerpoint_charts import SlideConfiguration

logger = logging.getLogger(__name__)

class PowerPointCustomizer:
    """Interactive customizer for PowerPoint presentations."""
    
    def __init__(self):
        self.config = SlideConfiguration()
    
    def get_customization_options(self) -> Dict[str, Any]:
        """
        Get all available customization options for the user interface.
        
        Returns:
            Dictionary with slide and chart options
        """
        return {
            'slides': self.config.get_slide_options(),
            'charts': self.config.get_chart_options(),
            'current_selection': self.config.get_selected_slides(),
            'presets': self._get_preset_configurations()
        }
    
    def _get_preset_configurations(self) -> Dict[str, Dict[str, Any]]:
        """Get predefined presentation configurations."""
        return {
            'executive_summary': {
                'name': 'Executive Summary',
                'description': 'High-level overview with key metrics and timelines',
                'slides': [
                    'cover',
                    'fleet_snapshot',
                    'emissions_timeline',
                    'electrification_timeline_weight',
                    'next_steps'
                ]
            },
            'technical_analysis': {
                'name': 'Technical Analysis',
                'description': 'Detailed technical analysis with all charts',
                'slides': [
                    'cover',
                    'fleet_snapshot',
                    'fleet_composition',
                    'emissions_timeline',
                    'emissions_by_weight',
                    'electrification_timeline_weight',
                    'electrification_timeline_body',
                    'age_analysis',
                    'data_quality',
                    'next_steps'
                ]
            },
            'data_focused': {
                'name': 'Data-Focused',
                'description': 'Emphasis on automated data analysis and quality',
                'slides': [
                    'cover',
                    'fleet_snapshot',
                    'fleet_composition',
                    'age_analysis',
                    'data_quality',
                    'emissions_by_weight',
                    'next_steps'
                ]
            },
            'timeline_focused': {
                'name': 'Timeline-Focused',
                'description': 'Focus on electrification timelines and projections',
                'slides': [
                    'cover',
                    'fleet_snapshot',
                    'emissions_timeline',
                    'electrification_timeline_weight',
                    'electrification_timeline_body',
                    'next_steps'
                ]
            }
        }
    
    def apply_preset(self, preset_name: str) -> bool:
        """
        Apply a predefined configuration preset.
        
        Args:
            preset_name: Name of the preset to apply
            
        Returns:
            True if successful, False if preset not found
        """
        presets = self._get_preset_configurations()
        
        if preset_name not in presets:
            logger.error(f"Unknown preset: {preset_name}")
            return False
        
        preset = presets[preset_name]
        return self.config.set_selected_slides(preset['slides'])
    
    def customize_slides(self, selected_slides: List[str]) -> bool:
        """
        Set custom slide selection.
        
        Args:
            selected_slides: List of slide IDs to include
            
        Returns:
            True if successful, False if invalid slides
        """
        return self.config.set_selected_slides(selected_slides)
    
    def get_configuration(self) -> SlideConfiguration:
        """Get the current slide configuration."""
        return self.config
    
    def validate_selection(self, selected_slides: List[str]) -> Dict[str, Any]:
        """
        Validate a slide selection and provide feedback.
        
        Args:
            selected_slides: List of slide IDs to validate
            
        Returns:
            Dictionary with validation results
        """
        available_slides = self.config.get_slide_options()
        
        # Check for invalid slides
        invalid_slides = [s for s in selected_slides if s not in available_slides]
        
        # Check for missing required slides
        required_slides = [sid for sid, info in available_slides.items() if info['required']]
        missing_required = [s for s in required_slides if s not in selected_slides]
        
        # Calculate estimated slide count
        final_slides = list(set(selected_slides + required_slides))
        
        return {
            'valid': len(invalid_slides) == 0,
            'invalid_slides': invalid_slides,
            'missing_required': missing_required,
            'final_slide_count': len(final_slides),
            'estimated_generation_time': self._estimate_generation_time(len(final_slides)),
            'warnings': self._get_selection_warnings(selected_slides)
        }
    
    def _estimate_generation_time(self, slide_count: int) -> str:
        """Estimate presentation generation time based on slide count."""
        base_time = 5  # Base time in seconds
        time_per_slide = 2  # Additional seconds per slide
        
        total_time = base_time + (slide_count * time_per_slide)
        
        if total_time < 60:
            return f"{total_time} seconds"
        else:
            minutes = total_time // 60
            seconds = total_time % 60
            return f"{minutes}m {seconds}s"
    
    def _get_selection_warnings(self, selected_slides: List[str]) -> List[str]:
        """Get warnings about the current slide selection."""
        warnings = []
        
        # Check if presentation might be too short
        if len(selected_slides) < 4:
            warnings.append("Presentation may be too short for a comprehensive analysis")
        
        # Check if presentation might be too long
        if len(selected_slides) > 12:
            warnings.append("Presentation may be too long for executive audiences")
        
        # Check for missing key analysis slides
        key_slides = ['fleet_snapshot', 'emissions_timeline']
        missing_key = [s for s in key_slides if s not in selected_slides]
        if missing_key:
            warnings.append(f"Consider including key analysis slides: {', '.join(missing_key)}")
        
        # Check for data-heavy selection without data quality slide
        data_heavy_slides = ['fleet_composition', 'age_analysis', 'emissions_by_weight']
        has_data_heavy = any(s in selected_slides for s in data_heavy_slides)
        if has_data_heavy and 'data_quality' not in selected_slides:
            warnings.append("Consider including 'Data Quality' slide when using data-heavy charts")
        
        return warnings

def create_presentation_config(preset: str = None, custom_slides: List[str] = None) -> SlideConfiguration:
    """
    Convenience function to create a presentation configuration.
    
    Args:
        preset: Name of preset configuration to use
        custom_slides: List of custom slide IDs (overrides preset)
        
    Returns:
        SlideConfiguration object
    """
    customizer = PowerPointCustomizer()
    
    if custom_slides:
        customizer.customize_slides(custom_slides)
    elif preset:
        customizer.apply_preset(preset)
    # Otherwise use default configuration
    
    return customizer.get_configuration()

# Convenience functions for common configurations
def executive_summary_config() -> SlideConfiguration:
    """Create configuration for executive summary presentation."""
    return create_presentation_config(preset='executive_summary')

def technical_analysis_config() -> SlideConfiguration:
    """Create configuration for technical analysis presentation."""
    return create_presentation_config(preset='technical_analysis')

def data_focused_config() -> SlideConfiguration:
    """Create configuration for data-focused presentation."""
    return create_presentation_config(preset='data_focused')

def timeline_focused_config() -> SlideConfiguration:
    """Create configuration for timeline-focused presentation."""
    return create_presentation_config(preset='timeline_focused')

# Example usage functions
def get_slide_selection_help() -> str:
    """Get help text for slide selection."""
    customizer = PowerPointCustomizer()
    options = customizer.get_customization_options()
    
    help_text = "Available Slides:\n\n"
    
    for slide_id, slide_info in options['slides'].items():
        required = " (Required)" if slide_info['required'] else ""
        help_text += f"• {slide_id}: {slide_info['name']}{required}\n"
        help_text += f"  {slide_info['description']}\n\n"
    
    help_text += "Available Presets:\n\n"
    
    for preset_id, preset_info in options['presets'].items():
        help_text += f"• {preset_id}: {preset_info['name']}\n"
        help_text += f"  {preset_info['description']}\n\n"
    
    return help_text

if __name__ == "__main__":
    # Example usage
    print("PowerPoint Customization Options:")
    print("=" * 40)
    print(get_slide_selection_help())
    
    # Test preset configurations
    config = executive_summary_config()
    print(f"Executive Summary slides: {config.get_selected_slides()}")
    
    config = technical_analysis_config()
    print(f"Technical Analysis slides: {config.get_selected_slides()}")
