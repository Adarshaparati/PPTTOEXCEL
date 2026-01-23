from fastapi import APIRouter, HTTPException, Body
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional
from datetime import datetime
from services.slide_data_service import slide_data_service
from services.s3_service import s3_service

router = APIRouter()


# ============================================
# Pydantic Models for Request Validation
# ============================================

class PointData(BaseModel):
    """Model for individual point/bullet item"""
    text: str
    color: Optional[str] = None
    font_size: Optional[int] = None


class PointsSlideData(BaseModel):
    """Model for points slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    header: str = Field(..., description="Slide header/title")
    description: str = Field(..., description="Main description text")
    image_url: Optional[str] = Field(None, description="URL of the image to insert")
    points: Optional[List[PointData]] = Field(None, description="List of bullet points")
    header_color: Optional[str] = Field(None, description="Hex color for header (e.g., #3667B2)")
    description_color: Optional[str] = Field(None, description="Hex color for description")
    background_color: Optional[str] = Field(None, description="Hex color for background")


class ImageTextSlideData(BaseModel):
    """Model for image+text slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Slide title")
    text: str = Field(..., description="Main text content")
    image_url: Optional[str] = Field(None, description="URL of the image")
    title_color: Optional[str] = Field(None, description="Hex color for title")
    text_color: Optional[str] = Field(None, description="Hex color for text")


class PhaseData(BaseModel):
    """Model for individual phase/timeline item"""
    name: str
    description: str
    status: Optional[str] = None
    color: Optional[str] = None


class PhasesSlideData(BaseModel):
    """Model for phases slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Slide title")
    phases: List[PhaseData] = Field(..., description="List of phases/timeline items")
    timeline_color: Optional[str] = Field(None, description="Hex color for timeline elements")


class StatData(BaseModel):
    """Model for individual statistic item"""
    label: str = Field(..., description="Statistic label/name")
    value: str = Field(..., description="Statistic value")
    color: Optional[str] = Field(None, description="Hex color for this stat")
    font_size: Optional[int] = Field(None, description="Font size for this stat")


class StatisticsSlideData(BaseModel):
    """Model for statistics slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Slide title")
    description: Optional[str] = Field(None, description="Optional description text")
    stat_data: List[StatData] = Field(..., description="List of statistics to display")
    title_color: Optional[str] = Field(None, description="Hex color for title")
    description_color: Optional[str] = Field(None, description="Hex color for description")
    background_color: Optional[str] = Field(None, description="Hex color for background")


class PeopleSlideData(BaseModel):
    """Model for people slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Slide title")
    description: Optional[str] = Field(None, description="Optional description text")
    names: List[str] = Field(..., description="List of people names")
    designations: List[str] = Field(..., description="List of designations/titles")
    descriptions: List[str] = Field(..., description="List of descriptions for each person")
    title_color: Optional[str] = Field(None, description="Hex color for title")
    description_color: Optional[str] = Field(None, description="Hex color for description")
    background_color: Optional[str] = Field(None, description="Hex color for background")


class CoverSlideData(BaseModel):
    """Model for cover slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Main title of the presentation")
    subtitle: str = Field(default="Generated via Template Flow", description="Subtitle text")
    slide_name: str = Field(..., description="Name of the slide")
    slide_data_id: Optional[str] = Field(None, description="Slide data identifier")
    slide_type: str = Field(default="Cover", description="Type of slide")
    company_name: str = Field(..., description="Company name")
    image: List[str] = Field(default_factory=list, description="Array of image URLs")
    colors: Optional[Dict[str, str]] = Field(None, description="Color scheme dictionary")


class ContactSlideData(BaseModel):
    """Model for contact slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Contact slide title")
    slide_name: str = Field(..., description="Name of the slide")
    website_link: str = Field(..., description="Website URL")
    linkedin_link: str = Field(default="", description="LinkedIn profile URL")
    contact_email: str = Field(default="", description="Contact email address")
    contact_phone: str = Field(default="", description="Contact phone number")
    image: List[str] = Field(default_factory=list, description="Array of image URLs")
    colors: Optional[Dict[str, str]] = Field(None, description="Color scheme dictionary")


class ImagesSlideData(BaseModel):
    """Model for images slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    slide_name: str = Field(..., description="Name of the slide")
    title: str = Field(..., description="Slide title")
    headers: Optional[List[str]] = Field(None, description="Array of headers for each image")
    descriptions: Optional[List[str]] = Field(None, description="Array of descriptions for each image")
    images: Optional[List[str]] = Field(None, description="Array of image URLs")


class TableSlideData(BaseModel):
    """Model for table slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Slide title")
    table_data: List[List[str]] = Field(..., description="2D array of table data")
    header_row: Optional[bool] = Field(True, description="First row is header")
    header_color: Optional[str] = Field(None, description="Hex color for header")


class ChartDataset(BaseModel):
    """Model for chart dataset"""
    label: str = Field(..., description="Dataset label/name")
    data: List[str] = Field(..., description="Array of data values")


class ChartData(BaseModel):
    """Model for chart data structure"""
    labels: List[str] = Field(..., description="Chart labels/categories")
    datasets: List[ChartDataset] = Field(..., description="Chart datasets")


class Chart(BaseModel):
    """Model for individual chart"""
    title: str = Field(..., description="Chart title")
    chart_data: ChartData = Field(..., description="Chart data with labels and datasets")


class ChartOptions(BaseModel):
    """Model for chart options"""
    responsive: Optional[bool] = Field(True, description="Responsive chart")
    maintainAspectRatio: Optional[bool] = Field(False, description="Maintain aspect ratio")
    backgroundColor: Optional[str] = Field("#FFFFFF", description="Background color")


class GraphSlideData(BaseModel):
    """Model for graph/chart slide data"""
    slide_number: int = Field(..., description="Slide number to update (1-indexed)", ge=1)
    title: str = Field(..., description="Slide title")
    chart_type: Optional[str] = Field("bar", description="Chart type: bar, line, pie, column, area")
    charts: List[Chart] = Field(..., description="List of charts to display")
    chart_options: Optional[ChartOptions] = Field(None, description="Chart configuration options")


class PointsSlideRequest(BaseModel):
    """Request model for generating points slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: PointsSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class ImageTextSlideRequest(BaseModel):
    """Request model for generating image+text slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: ImageTextSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class TableSlideRequest(BaseModel):
    """Request model for generating table slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: TableSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class GraphSlideRequest(BaseModel):
    """Request model for generating graph slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    templateId: Optional[str] = Field(None, description="Template identifier")
    slide_data: GraphSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class PhasesSlideRequest(BaseModel):
    """Request model for generating phases slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint presentation")
    slide_data: PhasesSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class StatisticsSlideRequest(BaseModel):
    """Request model for generating statistics slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: StatisticsSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class PeopleSlideRequest(BaseModel):
    """Request model for generating people slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: PeopleSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class CoverSlideRequest(BaseModel):
    """Request model for generating cover slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: CoverSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class ContactSlideRequest(BaseModel):
    """Request model for generating contact slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: ContactSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class ImagesSlideRequest(BaseModel):
    """Request model for generating images slide"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slide_data: ImagesSlideData
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


class SlideConfig(BaseModel):
    """Model for individual slide configuration in multi-slide generation"""
    slide_type: str = Field(..., description="Type of slide: 'points', 'image_text', 'table', 'phases', 'statistics', 'people', 'cover', 'contact', 'images', or 'graphs'")
    slide_data: Dict[str, Any] = Field(..., description="Data specific to the slide type")


class MultiSlideRequest(BaseModel):
    """Request model for generating multiple slides"""
    template_s3_url: str = Field(..., description="S3 URL or key of the PowerPoint template")
    slides_config: List[SlideConfig] = Field(..., description="List of slide configurations")
    upload_to_s3: bool = Field(True, description="Upload generated PPT to S3")
    output_filename: Optional[str] = Field(None, description="Custom output filename")


# ============================================
# API Endpoints
# ============================================

@router.post("/generate-points-slide")
async def generate_points_slide(request: PointsSlideRequest):
    """
    Modify a points/bullet slide in an existing presentation
    
    This endpoint modifies an existing PowerPoint presentation by updating
    a specific slide with new content (header, description, points, image).
    
    Example request:
    ```json
    {
        "template_s3_url": "https://bucket.s3.region.amazonaws.com/path/presentation.pptx",
        "slide_data": {
            "slide_number": 2,
            "header": "Overview",
            "description": "This document compares the use of Google's API...",
            "image_url": "https://example.com/image.png",
            "points": [
                {"text": "Point 1", "color": "#000000"},
                {"text": "Point 2", "color": "#3667B2"}
            ],
            "header_color": "#3667B2",
            "description_color": "#000000"
        },
        "upload_to_s3": true,
        "output_filename": "modified_presentation.pptx"
    }
    ```
    
    Note: Pass the S3 URL of your existing presentation. The API will:
    1. Download the presentation from S3
    2. Modify the specified slide
    3. Upload the modified presentation back to S3 (or return it directly)
    """
    try:
        # Convert Pydantic model to dict
        slide_data_dict = request.slide_data.dict()
        
        # Generate the slide
        output_ppt = slide_data_service.generate_points_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        # Prepare filename
        filename = request.output_filename or f"points_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        # Upload to S3 if requested
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Points slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "points",
                "timestamp": datetime.now().isoformat()
            }
        else:
            # Return as streaming response
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        # Check if it's an S3 404 error
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating points slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating points slide: {str(e)}"
        )


@router.post("/generate-image-text-slide")
async def generate_image_text_slide(request: ImageTextSlideRequest):
    """
    Generate an image+text slide
    
    Example request:
    ```json
    {
        "template_s3_url": "presentations/template.pptx",
        "slide_data": {
            "slide_number": 3,
            "title": "Feature Overview",
            "text": "This feature allows users to...",
            "image_url": "https://example.com/feature.png",
            "title_color": "#3667B2",
            "text_color": "#000000"
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_image_text_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"image_text_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Image+Text slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "image_text",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating image+text slide: {str(e)}"
        )


@router.post("/generate-table-slide")
async def generate_table_slide(request: TableSlideRequest):
    """
    Generate a table slide
    
    Example request:
    ```json
    {
        "template_s3_url": "presentations/template.pptx",
        "slide_data": {
            "slide_number": 4,
            "title": "Cost Comparison",
            "table_data": [
                ["Feature", "Plan A", "Plan B"],
                ["Price", "$10", "$20"],
                ["Storage", "10GB", "50GB"]
            ],
            "header_row": true,
            "header_color": "#3667B2"
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_table_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"table_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Table slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "table",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating table slide: {str(e)}"
        )


@router.post("/generate-graph-slide")
async def generate_graph_slide(request: GraphSlideRequest):
    """
    Generate a graph/chart slide
    
    Example request:
    ```json
    {
        "template_s3_url": "presentations/template.pptx",
        "templateId": "template_123",
        "slide_data": {
            "slide_number": 5,
            "title": "Sales Performance Q4",
            "chart_type": "bar",
            "charts": [
                {
                    "title": "Monthly Revenue",
                    "chart_data": {
                        "labels": ["Oct", "Nov", "Dec", "Jan", "Feb"],
                        "datasets": [
                            {
                                "label": "Sales 2024",
                                "data": ["100000", "120000", "150000", "135000", "160000"]
                            },
                            {
                                "label": "Sales 2025",
                                "data": ["110000", "130000", "165000", "145000", "175000"]
                            }
                        ]
                    }
                }
            ],
            "chart_options": {
                "responsive": true,
                "maintainAspectRatio": false,
                "backgroundColor": "#FFFFFF"
            }
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_graph_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"graph_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Graph slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "graph",
                "chart_type": slide_data_dict.get("chart_type", "bar"),
                "charts_count": len(slide_data_dict.get("charts", [])),
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating graph slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating graph slide: {str(e)}"
        )


@router.post("/generate-phases-slide")
async def generate_phases_slide(request: PhasesSlideRequest):
    """
    Generate a phases/timeline slide
    
    Example request:
    ```json
    {
        "template_s3_url": "https://bucket.s3.region.amazonaws.com/path/presentation.pptx",
        "slide_data": {
            "slide_number": 4,
            "title": "Project Timeline",
            "phases": [
                {
                    "name": "Phase 1: Planning",
                    "description": "Initial project setup and requirements gathering",
                    "status": "Completed",
                    "color": "#28a745"
                },
                {
                    "name": "Phase 2: Development", 
                    "description": "Core development and implementation",
                    "status": "In Progress",
                    "color": "#ffc107"
                },
                {
                    "name": "Phase 3: Testing",
                    "description": "Quality assurance and testing",
                    "status": "Planned",
                    "color": "#6c757d"
                }
            ],
            "timeline_color": "#3667B2"
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_phases_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"phases_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Phases slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "phases",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating phases slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating phases slide: {str(e)}"
        )


@router.post("/generate-statistics-slide")
async def generate_statistics_slide(request: StatisticsSlideRequest):
    """
    Generate a statistics slide
    
    Example request:
    ```json
    {
        "template_s3_url": "https://bucket.s3.region.amazonaws.com/path/presentation.pptx",
        "slide_data": {
            "slide_number": 5,
            "title": "Key Metrics",
            "description": "Performance statistics for Q4 2024",
            "stat_data": [
                {
                    "label": "Total Users",
                    "value": "1.2M",
                    "color": "#28a745",
                    "font_size": 24
                },
                {
                    "label": "Revenue Growth",
                    "value": "+35%",
                    "color": "#007bff",
                    "font_size": 24
                },
                {
                    "label": "Customer Satisfaction",
                    "value": "4.8/5",
                    "color": "#ffc107",
                    "font_size": 24
                },
                {
                    "label": "Market Share",
                    "value": "23%",
                    "color": "#dc3545",
                    "font_size": 24
                }
            ],
            "title_color": "#2E86AB",
            "description_color": "#333333"
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_statistics_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"statistics_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Statistics slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "statistics",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating statistics slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating statistics slide: {str(e)}"
        )


@router.post("/generate-people-slide")
async def generate_people_slide(request: PeopleSlideRequest):
    """
    Generate a people slide
    
    Example request:
    ```json
    {
        "template_s3_url": "https://bucket.s3.region.amazonaws.com/path/presentation.pptx",
        "slide_data": {
            "slide_number": 6,
            "title": "Our Team",
            "description": "Meet the key members of our organization",
            "names": [
                "John Smith",
                "Sarah Johnson",
                "Michael Chen",
                "Emily Davis"
            ],
            "designations": [
                "CEO & Founder",
                "CTO",
                "Lead Developer",
                "Product Manager"
            ],
            "descriptions": [
                "10+ years experience in tech leadership",
                "Expert in cloud architecture and AI",
                "Full-stack developer with 8 years experience",
                "Product strategy and user experience specialist"
            ],
            "title_color": "#2E86AB",
            "description_color": "#333333"
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_people_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"people_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "People slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "people",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating people slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating people slide: {str(e)}"
        )


@router.post("/generate-cover-slide")
async def generate_cover_slide(request: CoverSlideRequest):
    """
    Generate a cover slide
    
    Example request:
    ```json
    {
        "template_s3_url": "https://bucket.s3.region.amazonaws.com/path/presentation.pptx",
        "slide_data": {
            "slide_number": 1,
            "title": "Annual Business Report 2024",
            "subtitle": "Generated via Template Flow",
            "slide_name": "Annual Business Report 2024",
            "slide_data_id": "cover_001",
            "slide_type": "Cover",
            "company_name": "TechCorp Solutions",
            "image": [
                "https://example.com/logo.png",
                "https://example.com/background.jpg"
            ],
            "colors": {
                "primary": "#2E86AB",
                "secondary": "#A23B72",
                "accent": "#F18F01",
                "background": "#C73E1D"
            }
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_cover_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"cover_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Cover slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "cover",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating cover slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating cover slide: {str(e)}"
        )


@router.post("/generate-contact-slide")
async def generate_contact_slide(request: ContactSlideRequest):
    """
    Generate a contact slide
    
    Example request:
    ```json
    {
        "template_s3_url": "https://bucket.s3.region.amazonaws.com/path/presentation.pptx",
        "slide_data": {
            "slide_number": 10,
            "title": "Contact Us",
            "slide_name": "Contact Information",
            "website_link": "https://www.techcorp.com",
            "linkedin_link": "https://linkedin.com/company/techcorp",
            "contact_email": "info@techcorp.com",
            "contact_phone": "+1 (555) 123-4567",
            "image": [
                "https://example.com/contact-bg.jpg",
                "https://example.com/qr-code.png"
            ],
            "colors": {
                "primary": "#2E86AB",
                "secondary": "#A23B72",
                "text": "#333333",
                "background": "#f8f9fa"
            }
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_contact_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"contact_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Contact slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "contact",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating contact slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating contact slide: {str(e)}"
        )


@router.post("/generate-images-slide")
async def generate_images_slide(request: ImagesSlideRequest):
    """
    Generate an images slide
    
    Example request:
    ```json
    {
        "template_s3_url": "https://bucket.s3.region.amazonaws.com/path/presentation.pptx",
        "slide_data": {
            "slide_number": 7,
            "slide_name": "Product Gallery",
            "title": "Our Products",
            "headers": [
                "Mobile App",
                "Web Platform",
                "Desktop Software",
                "API Integration"
            ],
            "descriptions": [
                "Cross-platform mobile application with intuitive design",
                "Responsive web platform for all devices",
                "Powerful desktop software for professionals",
                "Seamless API integration for developers"
            ],
            "images": [
                "https://example.com/mobile-app.png",
                "https://example.com/web-platform.jpg",
                "https://example.com/desktop-software.png",
                "https://example.com/api-docs.jpg"
            ]
        },
        "upload_to_s3": true
    }
    ```
    """
    try:
        slide_data_dict = request.slide_data.dict()
        
        output_ppt = slide_data_service.generate_images_slide(
            template_s3_url=request.template_s3_url,
            slide_data=slide_data_dict
        )
        
        filename = request.output_filename or f"images_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Images slide generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_type": "images",
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except RuntimeError as re:
        error_msg = str(re)
        if "404" in error_msg or "Not Found" in error_msg:
            raise HTTPException(
                status_code=404,
                detail=f"Template not found in S3. Please check: 1) The S3 key/URL is correct, 2) The file exists in S3, 3) AWS credentials are valid. Error: {error_msg}"
            )
        raise HTTPException(status_code=500, detail=f"Error generating images slide: {error_msg}")
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating images slide: {str(e)}"
        )


@router.post("/generate-multi-slide")
async def generate_multi_slide_presentation(request: MultiSlideRequest):
    """
    Generate multiple slides in a single presentation
    
    Example request:
    ```json
    {
        "template_s3_url": "presentations/template.pptx",
        "slides_config": [
            {
                "slide_type": "points",
                "slide_data": {
                    "slide_number": 2,
                    "header": "Overview",
                    "description": "Key points...",
                    "image_url": "https://example.com/image1.png"
                }
            },
            {
                "slide_type": "image_text",
                "slide_data": {
                    "slide_number": 3,
                    "title": "Features",
                    "text": "Description...",
                    "image_url": "https://example.com/image2.png"
                }
            },
            {
                "slide_type": "table",
                "slide_data": {
                    "slide_number": 4,
                    "title": "Comparison",
                    "table_data": [["Col1", "Col2"], ["Data1", "Data2"]]
                }
            }
        ],
        "upload_to_s3": true,
        "output_filename": "complete_presentation.pptx"
    }
    ```
    """
    try:
        # Convert to dict format expected by service
        slides_config = [config.dict() for config in request.slides_config]
        
        output_ppt = slide_data_service.generate_multi_slide_presentation(
            template_s3_url=request.template_s3_url,
            slides_config=slides_config
        )
        
        filename = request.output_filename or f"multi_slide_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
        
        if request.upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=output_ppt.getvalue(),
                filename=filename,
                folder="generated_presentations"
            )
            
            if not upload_result.get('success'):
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed to upload to S3: {upload_result.get('error')}"
                )
            
            return {
                "success": True,
                "message": "Multi-slide presentation generated successfully",
                "s3_url": upload_result['s3_url'],
                "s3_key": upload_result['s3_key'],
                "filename": filename,
                "slide_count": len(slides_config),
                "timestamp": datetime.now().isoformat()
            }
        else:
            output_ppt.seek(0)
            return StreamingResponse(
                output_ppt,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": f"attachment; filename={filename}"}
            )
    
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating multi-slide presentation: {str(e)}"
        )


@router.get("/slide-types")
async def get_supported_slide_types():
    """
    Get list of supported slide types and their required fields
    """
    return {
        "supported_slide_types": [
            {
                "type": "points",
                "description": "Slide with header, description, bullet points, and image",
                "required_fields": ["slide_number", "header", "description"],
                "optional_fields": ["image_url", "points", "header_color", "description_color", "background_color"]
            },
            {
                "type": "image_text",
                "description": "Slide with title, text content, and image",
                "required_fields": ["slide_number", "title", "text"],
                "optional_fields": ["image_url", "title_color", "text_color"]
            },
            {
                "type": "table",
                "description": "Slide with title and data table",
                "required_fields": ["slide_number", "title", "table_data"],
                "optional_fields": ["header_row", "header_color"]
            },
            {
                "type": "phases",
                "description": "Slide with timeline/phases information",
                "required_fields": ["slide_number", "title", "phases"],
                "optional_fields": ["timeline_color"]
            },
            {
                "type": "statistics",
                "description": "Slide with key statistics and metrics",
                "required_fields": ["slide_number", "title", "stat_data"],
                "optional_fields": ["description", "title_color", "description_color", "background_color"]
            },
            {
                "type": "people",
                "description": "Slide with team members or people information",
                "required_fields": ["slide_number", "title", "names", "designations", "descriptions"],
                "optional_fields": ["description", "title_color", "description_color", "background_color"]
            },
            {
                "type": "cover",
                "description": "Cover slide for presentation",
                "required_fields": ["slide_number", "title", "company_name"],
                "optional_fields": ["subtitle", "slide_name", "slide_data_id", "slide_type", "image", "colors"]
            },
            {
                "type": "contact",
                "description": "Contact information slide",
                "required_fields": ["slide_number", "title", "slide_name", "website_link"],
                "optional_fields": ["linkedin_link", "contact_email", "contact_phone", "image", "colors"]
            },
            {
                "type": "images",
                "description": "Slide with multiple images, headers, and descriptions",
                "required_fields": ["slide_number", "slide_name", "title"],
                "optional_fields": ["headers", "descriptions", "images"]
            },
            {
                "type": "graphs",
                "description": "Slide with charts and graphs",
                "required_fields": ["slide_number", "title", "charts"],
                "optional_fields": ["chart_type", "chart_options"]
            }
        ],
        "note": "More slide types can be added easily using the handler pattern"
    }
