"""
Slide Handlers

This directory contains individual handlers for different slide types.

Each handler is responsible for:
- Processing one specific slide type
- Updating slide content and formatting
- Returning the modified presentation

Available Handlers:
- points.py - Handles bullet point slides
- image_text.py - Handles image + text slides  
- table.py - Handles table slides
- phases.py - Handles phase/timeline slides
- statistics.py - Handles statistics/metrics slides
- people.py - Handles people/team slides
- cover.py - Handles cover slides
- contact.py - Handles contact information slides
- images.py - Handles multi-image gallery slides

Handler Interface:
Each handler must implement:

def handle_slide_type(presentation_bytes: BytesIO, slide_data: Dict[str, Any], image_data: BytesIO = None) -> BytesIO:
    Handle slide modification
    
    Args:
        presentation_bytes: BytesIO containing the presentation
        slide_data: Dictionary with slide configuration
        image_data: Optional BytesIO containing image data
        
    Returns:
        BytesIO containing the modified presentation
"""

# Import all handlers for easy access
from .points import handle_points_slide
from .image_text import handle_image_text_slide
from .table import handle_table_slide
from .phases import handle_phases_slide
from .statistics import handle_statistics_slide
from .people import handle_people_slide
from .cover import handle_cover_slide
from .contact import handle_contact_slide
from .images import handle_images_slide

__all__ = [
    'handle_points_slide',
    'handle_image_text_slide', 
    'handle_table_slide',
    'handle_phases_slide',
    'handle_statistics_slide',
    'handle_people_slide',
    'handle_cover_slide',
    'handle_contact_slide',
    'handle_images_slide'
]