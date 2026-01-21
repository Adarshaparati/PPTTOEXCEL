import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from typing import List, Dict, Any
from dotenv import load_dotenv

load_dotenv()

class GoogleSheetsService:
    """Service for handling Google Sheets operations"""
    
    def __init__(self):
        self.scopes = ['https://www.googleapis.com/auth/spreadsheets']
        self.spreadsheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.credentials_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS')
        
        if not self.spreadsheet_id:
            raise ValueError("GOOGLE_SHEET_ID not found in environment variables")
        
        if not self.credentials_json:
            raise ValueError("GOOGLE_SHEETS_CREDENTIALS not found in environment variables")
        
        # Parse credentials from JSON string
        try:
            credentials_dict = json.loads(self.credentials_json)
            self.credentials = Credentials.from_service_account_info(
                credentials_dict,
                scopes=self.scopes
            )
        except json.JSONDecodeError:
            raise ValueError("Invalid GOOGLE_SHEETS_CREDENTIALS JSON format")
        
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.sheet = self.service.spreadsheets()
    
    def append_rows(self, values: List[List[Any]], range_name: str = "Sheet1") -> Dict:
        """
        Append rows to Google Sheet
        
        Args:
            values: List of rows to append
            range_name: Sheet name or range (default: "Sheet1")
            
        Returns:
            dict with operation details
        """
        try:
            body = {
                'values': values
            }
            
            result = self.sheet.values().append(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            
            return {
                'success': True,
                'updated_range': result.get('updates', {}).get('updatedRange'),
                'updated_rows': result.get('updates', {}).get('updatedRows'),
                'updated_cells': result.get('updates', {}).get('updatedCells')
            }
            
        except HttpError as error:
            return {
                'success': False,
                'error': str(error)
            }
    
    def append_ppt_data(
        self,
        template_id: str,
        s3_url: str,
        slide_data: List[Dict],
        
    ) -> Dict:
        """
        Append PowerPoint extraction data to Google Sheets
        EXACT same data as Excel, just with Template ID prepended
        
        Args:
            template_id: Template identifier for this presentation
            ppt_filename: Original PowerPoint filename (not used, kept for compatibility)
            s3_url: S3 URL of the uploaded PowerPoint (not used, kept for compatibility)
            slide_data: List of slide data dictionaries
            excel_url: URL of generated Excel file (not used, kept for compatibility)
            
        Returns:
            dict with operation details
        """
        rows = []
        
        for slide in slide_data:
            shapes = slide.get('shapes', [])
            
            # Each shape becomes a row - exact same as Excel but with Template ID first
            for shape in shapes:
                # Build row exactly as Excel does, just prepend Template ID
                row = [
                    template_id,  # ONLY ADDITION - Template ID at the start
                    # Everything below matches Excel exactly (45 columns)
                    shape.get('slide_number', ''),           # Slide No
                    shape.get('shape_name', ''),             # Shape Name
                    shape.get('shape_type', ''),             # Shape Type
                    shape.get('content', ''),                # Content
                    shape.get('left_emu', ''),               # Left (EMU)
                    shape.get('top_emu', ''),                # Top (EMU)
                    shape.get('width_emu', ''),              # Width (EMU)
                    shape.get('height_emu', ''),             # Height (EMU)
                    shape.get('left_inches', ''),            # Left (Inches)
                    shape.get('top_inches', ''),             # Top (Inches)
                    shape.get('width_inches', ''),           # Width (Inches)
                    shape.get('height_inches', ''),          # Height (Inches)
                    shape.get('font_name', ''),              # Font Name
                    shape.get('font_size', ''),              # Font Size
                    shape.get('font_bold', ''),              # Font Bold
                    shape.get('font_italic', ''),            # Font Italic
                    shape.get('font_underline', ''),         # Font Underline
                    shape.get('font_color', ''),             # Font Color
                    shape.get('text_alignment', ''),         # Text Alignment
                    shape.get('line_spacing', ''),           # Line Spacing
                    shape.get('paragraph_spacing', ''),      # Paragraph Spacing
                    shape.get('fill_color', ''),             # Fill Color
                    shape.get('fill_type', ''),              # Fill Type
                    shape.get('transparency', ''),           # Transparency
                    shape.get('line_color', ''),             # Line Color
                    shape.get('line_width', ''),             # Line Width
                    shape.get('line_style', ''),             # Line Style
                    shape.get('rotation', ''),               # Rotation
                    shape.get('has_image', ''),              # Has Image
                    shape.get('image_format', ''),           # Image Format
                    shape.get('image_width', ''),            # Image Width
                    shape.get('image_height', ''),           # Image Height
                    shape.get('image_file_size', ''),        # Image File Size
                    shape.get('image_url', ''),              # Image URL
                    shape.get('image_s3_url', ''),           # Image S3 URL
                    shape.get('image_base64', ''),           # Image Base64
                    shape.get('chart_type', ''),             # Chart Type
                    shape.get('chart_title', ''),            # Chart Title
                    shape.get('chart_data', ''),             # Chart Data
                    shape.get('chart_categories', ''),       # Chart Categories
                    shape.get('chart_series', ''),           # Chart Series
                    shape.get('hyperlink', ''),              # Hyperlink
                    shape.get('z_order', ''),                # Z-Order
                    shape.get('hidden', ''),                 # Hidden
                    shape.get('shadow', ''),                 # Shadow
                    shape.get('glow_effect', ''),            # Glow Effect
                    shape.get('reflection', ''),             # Reflection
                    shape.get('3d_effects', ''),             # 3D Effects
                    shape.get('placeholder_type', ''),       # Placeholder Type
                    shape.get('animation_effects', '')       # Animation Effects
                ]
                rows.append(row)
        
        if rows:
            result = self.append_rows(rows)
            if result.get('success'):
                result['template_id'] = template_id
                result['rows_added'] = len(rows)
            return result
        else:
            return {
                'success': False,
                'error': 'No data to append'
            }
    
    def get_sheet_data(self, range_name: str = "Sheet1") -> Dict:
        """
        Get data from Google Sheet
        
        Args:
            range_name: Sheet name or range (default: "Sheet1")
            
        Returns:
            dict with sheet data
        """
        try:
            result = self.sheet.values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            return {
                'success': True,
                'data': values,
                'row_count': len(values)
            }
            
        except HttpError as error:
            return {
                'success': False,
                'error': str(error)
            }
    
    def clear_sheet(self, range_name: str = "Sheet1") -> Dict:
        """
        Clear all data from a sheet range
        
        Args:
            range_name: Sheet name or range to clear
            
        Returns:
            dict with operation details
        """
        try:
            result = self.sheet.values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            
            return {
                'success': True,
                'cleared_range': result.get('clearedRange')
            }
            
        except HttpError as error:
            return {
                'success': False,
                'error': str(error)
            }
    
    def initialize_headers(self, range_name: str = "Sheet1!A1:AT1") -> Dict:
        """
        Initialize sheet with headers - Template ID + all Excel columns (46 total)
        
        Args:
            range_name: Range for headers
            
        Returns:
            dict with operation details
        """
        headers = [[
            # ONLY ADDITION - Template ID
            "Template ID",
            
            # EXACT SAME AS EXCEL (45 columns)
            "Slide No",
            "Shape Name",
            "Shape Type",
            "Content",
            "Left (EMU)",
            "Top (EMU)",
            "Width (EMU)",
            "Height (EMU)",
            "Left (Inches)",
            "Top (Inches)",
            "Width (Inches)",
            "Height (Inches)",
            "Font Name",
            "Font Size",
            "Font Bold",
            "Font Italic",
            "Font Underline",
            "Font Color",
            "Text Alignment",
            "Line Spacing",
            "Paragraph Spacing",
            "Fill Color",
            "Fill Type",
            "Transparency",
            "Line Color",
            "Line Width",
            "Line Style",
            "Rotation",
            "Has Image",
            "Image Format",
            "Image Width",
            "Image Height",
            "Image File Size",
            "Image URL",
            "Image S3 URL",
            "Image Base64",
            "Chart Type",
            "Chart Title",
            "Chart Data",
            "Chart Categories",
            "Chart Series",
            "Hyperlink",
            "Z-Order",
            "Hidden",
            "Shadow",
            "Glow Effect",
            "Reflection",
            "3D Effects",
            "Placeholder Type",
            "Animation Effects"
        ]]
        
        try:
            body = {
                'values': headers
            }
            
            result = self.sheet.values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            return {
                'success': True,
                'updated_cells': result.get('updatedCells')
            }
            
        except HttpError as error:
            return {
                'success': False,
                'error': str(error)
            }

# Create singleton instance (will be initialized when needed)
sheets_service = None

def get_sheets_service():
    """Get or create Google Sheets service instance"""
    global sheets_service
    if sheets_service is None:
        try:
            sheets_service = GoogleSheetsService()
        except Exception as e:
            print(f"Warning: Could not initialize Google Sheets service: {e}")
            return None
    return sheets_service
