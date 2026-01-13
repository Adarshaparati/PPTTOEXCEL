import os
import io
import requests
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from typing import Dict, List, Any, Optional
from PIL import Image
from services.s3_service import s3_service
from datetime import datetime

# Import slide handlers
from services.handlers.points import handle_points_slide
from services.handlers.image_text import handle_image_text_slide
from services.handlers.table import handle_table_slide
from services.handlers.phases import handle_phases_slide


class SlideDataService:
    """Service for managing and generating different slide types"""
    
    @staticmethod
    def hex_to_rgb(hex_str):
        """Convert #RRGGBB string to RGBColor safely."""
        if not hex_str:
            return RGBColor(0, 0, 0)
        
        hex_str = str(hex_str).strip().replace("#", "")
        if len(hex_str) != 6 or any(c not in "0123456789ABCDEFabcdef" for c in hex_str):
            print(f"⚠️ Invalid hex color '{hex_str}', defaulting to black")
            return RGBColor(0, 0, 0)
        
        return RGBColor(int(hex_str[0:2], 16),
                        int(hex_str[2:4], 16),
                        int(hex_str[4:6], 16))
    
    def download_image_from_url(self, image_url: str) -> BytesIO:
        """
        Download image from URL and return as BytesIO
        
        Args:
            image_url: URL of the image to download
            
        Returns:
            BytesIO object containing the image data
        """
        try:
            print(f"🖼️ Downloading image from: {image_url}")
            response = requests.get(image_url, timeout=10)
            response.raise_for_status()
            print(f"✅ Image downloaded successfully ({len(response.content)} bytes)")
            return BytesIO(response.content)
        except Exception as e:
            error_msg = f"Failed to download image from {image_url}: {str(e)}"
            print(f"❌ {error_msg}")
            raise RuntimeError(error_msg)
    
    def download_template_from_s3(self, s3_url: str) -> BytesIO:
        """
        Download PowerPoint template from S3
        
        Args:
            s3_url: S3 URL of the template or just the S3 key
            
        Returns:
            BytesIO object containing the template data
        """
        try:
            from urllib.parse import unquote, urlparse
            
            # Extract key from S3 URL
            if s3_url.startswith('http'):
                # Try direct HTTP download first (works for public S3 URLs)
                print(f"📥 Attempting direct download from S3 URL: {s3_url}")
                try:
                    response = requests.get(s3_url, timeout=30)
                    response.raise_for_status()
                    print(f"✅ Template downloaded directly via HTTP ({len(response.content)} bytes)")
                    return BytesIO(response.content)
                except Exception as http_error:
                    print(f"⚠️ Direct HTTP download failed: {http_error}")
                    print(f"🔄 Falling back to boto3 download...")
                    
                    # Fall back to boto3 download
                    parsed = urlparse(s3_url)
                    # Get the path without leading slash and decode any URL encoding
                    key = unquote(parsed.path.lstrip('/'))
                    
                    if not key:
                        raise ValueError(f"Invalid S3 URL format - no key found: {s3_url}")
                        
                    print(f"📥 Parsed S3 URL - Key: {key}")
            else:
                # Assume it's just the key
                key = s3_url
                print(f"📥 Using provided key: {key}")
            
            # Download from S3 using boto3
            file_data = s3_service.download_file(key)

            if not file_data:
                raise RuntimeError(f"S3 download failed or returned no data for key: {key}")

            print(f"✅ Template downloaded successfully ({len(file_data)} bytes)")
            return BytesIO(file_data)
            
        except Exception as e:
            error_msg = f"Failed to download template from S3: {str(e)}"
            print(f"❌ {error_msg}")
            raise RuntimeError(error_msg)
    
    def generate_slide(
        self,
        slide_type: str,
        presentation_s3_url: str,
        slide_data: Dict[str, Any]
    ) -> BytesIO:
        """
        Universal slide generator - routes to appropriate handler based on slide type
        
        Args:
            slide_type: Type of slide ('points', 'image_text', 'table', 'phases', etc.)
            presentation_s3_url: S3 URL of the presentation to modify
            slide_data: Dictionary containing slide configuration
            
        Returns:
            BytesIO object containing the modified presentation
        """
        print(f"� Generating {slide_type.upper()} slide...")
        
        # Download the presentation
        presentation_bytes = self.download_template_from_s3(presentation_s3_url)
        
        # Download image if provided
        image_data = None
        if slide_data.get('image_url'):
            try:
                image_data = self.download_image_from_url(slide_data['image_url'])
            except Exception as e:
                print(f"⚠️ Warning: Could not download image: {e}")
        
        # Route to appropriate handler
        handler_map = {
            'points': handle_points_slide,
            'image_text': handle_image_text_slide,
            'table': handle_table_slide,
            'phases': handle_phases_slide,
        }
        
        handler = handler_map.get(slide_type)
        if not handler:
            available_types = ', '.join(handler_map.keys())
            raise ValueError(f"Unsupported slide type '{slide_type}'. Available types: {available_types}")
        
        # Call the appropriate handler
        try:
            result = handler(presentation_bytes, slide_data, image_data)
            print(f"✅ {slide_type.upper()} slide generated successfully")
            return result
        except Exception as e:
            error_msg = f"Error in {slide_type} handler: {str(e)}"
            print(f"❌ {error_msg}")
            raise RuntimeError(error_msg)
    
    # Convenience methods for backward compatibility and specific slide types
    def generate_points_slide(self, template_s3_url: str, slide_data: Dict[str, Any]) -> BytesIO:
        """Generate points slide - backwards compatible method"""
        return self.generate_slide('points', template_s3_url, slide_data)
    
    def generate_image_text_slide(self, template_s3_url: str, slide_data: Dict[str, Any]) -> BytesIO:
        """Generate image+text slide - backwards compatible method"""
        return self.generate_slide('image_text', template_s3_url, slide_data)
    
    def generate_table_slide(self, template_s3_url: str, slide_data: Dict[str, Any]) -> BytesIO:
        """Generate table slide - backwards compatible method"""
        return self.generate_slide('table', template_s3_url, slide_data)
    
    def generate_phases_slide(self, template_s3_url: str, slide_data: Dict[str, Any]) -> BytesIO:
        """Generate phases slide - new slide type"""
        return self.generate_slide('phases', template_s3_url, slide_data)
    
    def generate_multi_slide_presentation(
        self,
        template_s3_url: str,
        slides_config: List[Dict[str, Any]]
    ) -> BytesIO:
        """
        Generate multiple slides in a single presentation using the new handler-based approach
        
        Args:
            template_s3_url: S3 URL of the PowerPoint template
            slides_config: List of slide configurations, each containing:
                - slide_type: str - Type of slide ('points', 'image_text', 'table', 'phases')
                - slide_data: Dict - Data specific to the slide type
                
        Returns:
            BytesIO object containing the complete presentation
        """
        print(f"🎨 Generating multi-slide presentation...")
        
        # Download template once
        presentation_bytes = self.download_template_from_s3(template_s3_url)
        current_presentation = presentation_bytes
        
        # Process each slide configuration
        for config in slides_config:
            slide_type = config.get('slide_type')
            slide_data = config.get('slide_data', {})
            
            if not slide_type:
                print(f"⚠️ Skipping config without slide_type")
                continue
            
            try:
                # Use the universal handler for each slide
                current_presentation = self._process_slide_in_presentation(
                    current_presentation, slide_type, slide_data
                )
            except Exception as e:
                print(f"⚠️ Error processing {slide_type} slide: {e}")
                continue
        
        print(f"✅ Multi-slide presentation generated successfully")
        return current_presentation
    
    def _process_slide_in_presentation(self, presentation_bytes: BytesIO, slide_type: str, slide_data: Dict[str, Any]) -> BytesIO:
        """Process a single slide in a presentation"""
        # Download image if needed
        image_data = None
        if slide_data.get('image_url'):
            try:
                image_data = self.download_image_from_url(slide_data['image_url'])
            except Exception as e:
                print(f"⚠️ Warning: Could not download image for {slide_type}: {e}")
        
        # Route to appropriate handler
        handler_map = {
            'points': handle_points_slide,
            'image_text': handle_image_text_slide,
            'table': handle_table_slide,
            'phases': handle_phases_slide,
        }
        
        handler = handler_map.get(slide_type)
        if not handler:
            print(f"⚠️ Unknown slide type: {slide_type}")
            return presentation_bytes
        
        return handler(presentation_bytes, slide_data, image_data)


# Create singleton instance
slide_data_service = SlideDataService()