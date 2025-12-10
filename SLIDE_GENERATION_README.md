# Slide Generation Service - Quick Start

## Overview

The Slide Generation Service provides a flexible API to programmatically create PowerPoint presentations with different slide types. You can generate:

- **Points Slides**: Header, description, bullet points, and image
- **Image+Text Slides**: Title, text content, and image
- **Table Slides**: Title and data tables
- **Multi-Slide Presentations**: Combine multiple slide types

## Features

✅ **S3 Integration**: Upload templates from S3 and save generated presentations back to S3  
✅ **Image URL Support**: Insert images from any public URL  
✅ **Custom Formatting**: Control colors, fonts, and styles  
✅ **Multiple Slide Types**: Different layouts for different content types  
✅ **Batch Generation**: Create multiple slides in one API call  
✅ **Extensible**: Easy to add new slide types (charts, diagrams, etc.)

## Quick Start

### 1. Prerequisites

- Python 3.8+
- FastAPI server running
- AWS S3 access (configured in `.env`)
- PowerPoint template uploaded to S3

### 2. Environment Setup

Make sure your `.env` file includes S3 credentials:

```bash
AWS_ACCESS_KEY_ID=your_access_key
AWS_SECRET_ACCESS_KEY=your_secret_key
AWS_REGION=us-east-1
S3_BUCKET=your-bucket-name
```

### 3. Start the Server

```bash
python main.py
```

The API will be available at `http://localhost:8000`

### 4. Test the API

Visit the interactive documentation:
```
http://localhost:8000/docs
```

Or use the example script:
```bash
python example_slide_generation.py
```

## Basic Usage

### Example 1: Generate a Points Slide

```python
import requests

response = requests.post(
    "http://localhost:8000/api/slides/generate-points-slide",
    json={
        "template_s3_url": "presentations/my_template.pptx",
        "slide_data": {
            "slide_number": 2,
            "header": "Project Overview",
            "description": "Key project highlights and milestones",
            "image_url": "https://example.com/image.png",
            "points": [
                {"text": "Scalable architecture", "color": "#3667B2"},
                {"text": "Real-time processing", "color": "#000000"},
                {"text": "Cloud-native design", "color": "#3667B2"}
            ],
            "header_color": "#3667B2",
            "description_color": "#000000"
        },
        "upload_to_s3": True,
        "output_filename": "project_overview.pptx"
    }
)

print(response.json())
```

### Example 2: Generate Multiple Slides

```python
response = requests.post(
    "http://localhost:8000/api/slides/generate-multi-slide",
    json={
        "template_s3_url": "presentations/template.pptx",
        "slides_config": [
            {
                "slide_type": "points",
                "slide_data": {
                    "slide_number": 2,
                    "header": "Overview",
                    "description": "Introduction to our solution",
                    "points": [{"text": "Point 1"}, {"text": "Point 2"}]
                }
            },
            {
                "slide_type": "table",
                "slide_data": {
                    "slide_number": 3,
                    "title": "Comparison",
                    "table_data": [
                        ["Feature", "Value"],
                        ["Users", "10,000"],
                        ["Revenue", "$100K"]
                    ]
                }
            }
        ],
        "upload_to_s3": True
    }
)
```

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/slides/generate-points-slide` | POST | Generate points/bullet slide |
| `/api/slides/generate-image-text-slide` | POST | Generate image+text slide |
| `/api/slides/generate-table-slide` | POST | Generate table slide |
| `/api/slides/generate-multi-slide` | POST | Generate multiple slides |
| `/api/slides/slide-types` | GET | Get supported slide types |

## Template Requirements

### For Points Slides

Your PowerPoint template needs shapes with these names:
- `Header1` - Main header text
- `Description1` - Description text
- `Description1_BG` - Bullet points container
- `Image` - Image placeholder

### For Image+Text Slides

Required shapes:
- `P100` or similar with "title" - Title text
- `S100` or similar with "text" - Content text
- Shape with "image" in name - Image placeholder

### For Table Slides

Required:
- Shape with "title" in name - Title text
- A table shape - Will be auto-detected and populated

## Response Format

### Success Response (upload_to_s3 = true)

```json
{
  "success": true,
  "message": "Points slide generated successfully",
  "s3_url": "https://bucket.s3.amazonaws.com/path/to/file.pptx",
  "s3_key": "generated_presentations/file.pptx",
  "filename": "file.pptx",
  "slide_type": "points",
  "timestamp": "2025-12-05T18:06:00.000000"
}
```

### Success Response (upload_to_s3 = false)

Returns the PowerPoint file as a binary stream for direct download.

### Error Response

```json
{
  "detail": "Error message describing what went wrong"
}
```

## Common Use Cases

### 1. Automated Report Generation

Generate monthly reports with data from your database:

```python
# Fetch data from database
data = get_monthly_metrics()

# Generate slides
response = requests.post(url, json={
    "template_s3_url": "templates/monthly_report.pptx",
    "slides_config": [
        {
            "slide_type": "table",
            "slide_data": {
                "slide_number": 2,
                "title": "Monthly Metrics",
                "table_data": data.to_table()
            }
        }
    ]
})
```

### 2. Dynamic Presentation Creation

Create presentations based on user input:

```python
def create_custom_presentation(user_data):
    slides = []
    
    # Add overview slide
    slides.append({
        "slide_type": "points",
        "slide_data": {
            "slide_number": 2,
            "header": user_data['title'],
            "description": user_data['description'],
            "points": user_data['key_points']
        }
    })
    
    # Add data slides
    for idx, dataset in enumerate(user_data['datasets']):
        slides.append({
            "slide_type": "table",
            "slide_data": {
                "slide_number": idx + 3,
                "title": dataset['title'],
                "table_data": dataset['data']
            }
        })
    
    return generate_presentation(slides)
```

### 3. Batch Processing

Generate multiple presentations in parallel:

```python
from concurrent.futures import ThreadPoolExecutor

presentations = [
    {"template": "template1.pptx", "data": data1},
    {"template": "template2.pptx", "data": data2},
    {"template": "template3.pptx", "data": data3},
]

def generate_single(pres):
    return requests.post(url, json=pres)

with ThreadPoolExecutor(max_workers=5) as executor:
    results = list(executor.map(generate_single, presentations))
```

## Troubleshooting

### Issue: "Slide X not found in template"

**Solution**: Verify your template has the correct number of slides. Slide numbers are 1-indexed.

### Issue: "Failed to download image from URL"

**Solution**: 
- Ensure the image URL is publicly accessible
- Check if the URL returns a valid image format
- Use pre-signed S3 URLs for private images

### Issue: "Shape not found in slide"

**Solution**: 
- Check that your template has shapes with the expected names
- Use PowerPoint's Selection Pane to view all shape names
- Names are case-insensitive and partial matches work

### Issue: "Invalid hex color"

**Solution**: 
- Use format: `#RRGGBB` (e.g., `#3667B2`)
- Don't use color names like "blue" or RGB format

## Advanced Configuration

### Custom Shape Names

If your template uses different shape names, modify the service:

```python
# In slide_data_service.py
# Look for shape name matching logic and adjust:
if shape.name == 'YourCustomHeaderName':
    # Update header
```

### Adding New Slide Types

To add a new slide type (e.g., chart):

1. Add method to `SlideDataService`:
```python
def generate_chart_slide(self, template_s3_url, slide_data):
    # Implementation
    pass
```

2. Add Pydantic model in `slide_routes.py`:
```python
class ChartSlideData(BaseModel):
    slide_number: int
    title: str
    chart_data: Dict[str, List]
    chart_type: str
```

3. Add endpoint in `slide_routes.py`:
```python
@router.post("/generate-chart-slide")
async def generate_chart_slide(request: ChartSlideRequest):
    # Implementation
    pass
```

## Best Practices

1. **Template Organization**: Keep templates organized in S3 folders by type
2. **Error Handling**: Always check response status and handle errors
3. **Image URLs**: Use CDN or S3 for images to ensure availability
4. **Batch Operations**: Use multi-slide endpoint for better performance
5. **Testing**: Test with small templates before processing large presentations
6. **Naming Conventions**: Use clear, descriptive filenames for outputs
7. **S3 Upload**: Set `upload_to_s3: true` for production, `false` for testing

## Performance Tips

- Use `generate-multi-slide` instead of multiple individual calls
- Cache frequently used templates locally
- Pre-process and validate data before API calls
- Use async requests for parallel generation
- Optimize image sizes before uploading

## Next Steps

- Review the [Complete API Guide](SLIDE_API_GUIDE.md)
- Check out [Example Scripts](example_slide_generation.py)
- Explore the [Interactive Docs](http://localhost:8000/docs)
- Add custom slide types for your needs

## Support

For issues or questions:
- Check the API documentation at `/docs`
- Review the example scripts
- Verify your template structure
- Check server logs for detailed error messages

---

**Version**: 2.0.0  
**Last Updated**: December 5, 2025
