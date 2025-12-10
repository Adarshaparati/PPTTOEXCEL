# Slide Generation API Guide

This guide explains how to use the new Slide Generation API to create different types of PowerPoint slides programmatically.

## Overview

The Slide Generation API allows you to create customized PowerPoint presentations with different slide types:
- **Points Slide**: Slide with header, description, bullet points, and image
- **Image+Text Slide**: Slide with title, text content, and image
- **Table Slide**: Slide with title and data table
- **Multi-Slide Presentation**: Combine multiple slide types in one presentation

## Base URL

```
http://localhost:8000/api/slides
```

## Authentication

Currently, no authentication is required. In production, consider adding API key authentication.

---

## API Endpoints

### 1. Generate Points Slide

**Endpoint:** `POST /api/slides/generate-points-slide`

**Description:** Create a slide with header, description, bullet points, and an image.

**Request Body:**
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slide_data": {
    "slide_number": 2,
    "header": "Overview",
    "description": "This document compares the use of Google's API with third-party alternatives for Veo 3 integration.",
    "image_url": "https://example.com/overview-image.png",
    "points": [
      {
        "text": "Key difference: API integration approach",
        "color": "#000000",
        "font_size": 14
      },
      {
        "text": "Consideration: Cost and scalability",
        "color": "#3667B2",
        "font_size": 14
      },
      {
        "text": "Performance benchmarks",
        "color": "#000000"
      }
    ],
    "header_color": "#3667B2",
    "description_color": "#000000",
    "background_color": "#FFFFFF"
  },
  "upload_to_s3": true,
  "output_filename": "points_slide_output.pptx"
}
```

**Parameters:**
- `template_s3_url` (required): S3 URL or key of your PowerPoint template
- `slide_data` (required): Object containing slide content
  - `slide_number` (required): Which slide to update (1-indexed)
  - `header` (required): Slide header text
  - `description` (required): Main description text
  - `image_url` (optional): URL of image to insert
  - `points` (optional): Array of bullet points with optional formatting
  - `header_color` (optional): Hex color for header (e.g., "#3667B2")
  - `description_color` (optional): Hex color for description
  - `background_color` (optional): Hex color for background
- `upload_to_s3` (optional, default: true): Whether to upload result to S3
- `output_filename` (optional): Custom filename for the output

**Response (when upload_to_s3 = true):**
```json
{
  "success": true,
  "message": "Points slide generated successfully",
  "s3_url": "https://bucket.s3.amazonaws.com/generated_presentations/points_slide_output.pptx",
  "s3_key": "generated_presentations/points_slide_output.pptx",
  "filename": "points_slide_output.pptx",
  "slide_type": "points",
  "timestamp": "2025-12-05T18:06:00.000000"
}
```

**Response (when upload_to_s3 = false):**
Returns the PowerPoint file as a downloadable stream.

---

### 2. Generate Image+Text Slide

**Endpoint:** `POST /api/slides/generate-image-text-slide`

**Description:** Create a slide with title, text content, and an image.

**Request Body:**
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slide_data": {
    "slide_number": 3,
    "title": "Feature Overview",
    "text": "This feature allows users to integrate seamlessly with our platform using modern API patterns and best practices.",
    "image_url": "https://example.com/feature-diagram.png",
    "title_color": "#3667B2",
    "text_color": "#000000"
  },
  "upload_to_s3": true,
  "output_filename": "feature_slide.pptx"
}
```

**Parameters:**
- `template_s3_url` (required): S3 URL or key of your PowerPoint template
- `slide_data` (required):
  - `slide_number` (required): Which slide to update (1-indexed)
  - `title` (required): Slide title
  - `text` (required): Main text content
  - `image_url` (optional): URL of image to insert
  - `title_color` (optional): Hex color for title
  - `text_color` (optional): Hex color for text
- `upload_to_s3` (optional, default: true): Whether to upload result to S3
- `output_filename` (optional): Custom filename

**Response:**
```json
{
  "success": true,
  "message": "Image+Text slide generated successfully",
  "s3_url": "https://bucket.s3.amazonaws.com/generated_presentations/feature_slide.pptx",
  "s3_key": "generated_presentations/feature_slide.pptx",
  "filename": "feature_slide.pptx",
  "slide_type": "image_text",
  "timestamp": "2025-12-05T18:06:00.000000"
}
```

---

### 3. Generate Table Slide

**Endpoint:** `POST /api/slides/generate-table-slide`

**Description:** Create a slide with a title and data table.

**Request Body:**
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slide_data": {
    "slide_number": 4,
    "title": "Cost Comparison",
    "table_data": [
      ["Feature", "Google API", "Third-Party Alternative"],
      ["Monthly Cost", "$100", "$50"],
      ["API Calls/Month", "100,000", "50,000"],
      ["Support Level", "Enterprise", "Standard"],
      ["SLA", "99.9%", "99.5%"]
    ],
    "header_row": true,
    "header_color": "#3667B2"
  },
  "upload_to_s3": true,
  "output_filename": "comparison_table.pptx"
}
```

**Parameters:**
- `template_s3_url` (required): S3 URL or key of your PowerPoint template
- `slide_data` (required):
  - `slide_number` (required): Which slide to update (1-indexed)
  - `title` (required): Slide title
  - `table_data` (required): 2D array of table data (rows and columns)
  - `header_row` (optional, default: true): Whether first row is a header
  - `header_color` (optional): Hex color for header row
- `upload_to_s3` (optional, default: true): Whether to upload result to S3
- `output_filename` (optional): Custom filename

**Response:**
```json
{
  "success": true,
  "message": "Table slide generated successfully",
  "s3_url": "https://bucket.s3.amazonaws.com/generated_presentations/comparison_table.pptx",
  "s3_key": "generated_presentations/comparison_table.pptx",
  "filename": "comparison_table.pptx",
  "slide_type": "table",
  "timestamp": "2025-12-05T18:06:00.000000"
}
```

---

### 4. Generate Multi-Slide Presentation

**Endpoint:** `POST /api/slides/generate-multi-slide`

**Description:** Generate multiple slides of different types in a single presentation.

**Request Body:**
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slides_config": [
    {
      "slide_type": "points",
      "slide_data": {
        "slide_number": 2,
        "header": "Overview",
        "description": "Key comparison points",
        "image_url": "https://example.com/overview.png",
        "points": [
          {"text": "Integration approach"},
          {"text": "Cost analysis"},
          {"text": "Performance metrics"}
        ]
      }
    },
    {
      "slide_type": "image_text",
      "slide_data": {
        "slide_number": 3,
        "title": "Architecture",
        "text": "System architecture overview showing API integration patterns.",
        "image_url": "https://example.com/architecture.png"
      }
    },
    {
      "slide_type": "table",
      "slide_data": {
        "slide_number": 4,
        "title": "Detailed Comparison",
        "table_data": [
          ["Metric", "Option A", "Option B"],
          ["Cost", "$100", "$50"],
          ["Performance", "High", "Medium"]
        ]
      }
    }
  ],
  "upload_to_s3": true,
  "output_filename": "complete_presentation.pptx"
}
```

**Parameters:**
- `template_s3_url` (required): S3 URL or key of your PowerPoint template
- `slides_config` (required): Array of slide configurations
  - Each config contains:
    - `slide_type`: "points", "image_text", or "table"
    - `slide_data`: Data specific to that slide type (see individual endpoints)
- `upload_to_s3` (optional, default: true): Whether to upload result to S3
- `output_filename` (optional): Custom filename

**Response:**
```json
{
  "success": true,
  "message": "Multi-slide presentation generated successfully",
  "s3_url": "https://bucket.s3.amazonaws.com/generated_presentations/complete_presentation.pptx",
  "s3_key": "generated_presentations/complete_presentation.pptx",
  "filename": "complete_presentation.pptx",
  "slide_count": 3,
  "timestamp": "2025-12-05T18:06:00.000000"
}
```

---

### 5. Get Supported Slide Types

**Endpoint:** `GET /api/slides/slide-types`

**Description:** Get information about all supported slide types and their required fields.

**Response:**
```json
{
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
    }
  ],
  "note": "More slide types can be added in the future (charts, diagrams, etc.)"
}
```

---

## Usage Examples

### Example 1: Create a Points Slide using cURL

```bash
curl -X POST "http://localhost:8000/api/slides/generate-points-slide" \
  -H "Content-Type: application/json" \
  -d '{
    "template_s3_url": "presentations/my_template.pptx",
    "slide_data": {
      "slide_number": 2,
      "header": "Project Overview",
      "description": "This project aims to revolutionize the way we handle data processing.",
      "image_url": "https://example.com/project-image.png",
      "points": [
        {"text": "Scalable architecture", "color": "#3667B2"},
        {"text": "Real-time processing", "color": "#000000"},
        {"text": "Cloud-native design", "color": "#3667B2"}
      ],
      "header_color": "#3667B2"
    },
    "upload_to_s3": true
  }'
```

### Example 2: Create Multiple Slides using Python

```python
import requests

url = "http://localhost:8000/api/slides/generate-multi-slide"

payload = {
    "template_s3_url": "presentations/template.pptx",
    "slides_config": [
        {
            "slide_type": "points",
            "slide_data": {
                "slide_number": 2,
                "header": "Introduction",
                "description": "Welcome to our presentation",
                "points": [
                    {"text": "Point 1"},
                    {"text": "Point 2"}
                ]
            }
        },
        {
            "slide_type": "table",
            "slide_data": {
                "slide_number": 3,
                "title": "Data Analysis",
                "table_data": [
                    ["Metric", "Value"],
                    ["Users", "10,000"],
                    ["Revenue", "$100K"]
                ]
            }
        }
    ],
    "upload_to_s3": True,
    "output_filename": "my_presentation.pptx"
}

response = requests.post(url, json=payload)
print(response.json())
```

### Example 3: Download Slide Directly (without S3)

```bash
curl -X POST "http://localhost:8000/api/slides/generate-points-slide" \
  -H "Content-Type: application/json" \
  -d '{
    "template_s3_url": "presentations/template.pptx",
    "slide_data": {
      "slide_number": 2,
      "header": "Test Slide",
      "description": "Testing direct download"
    },
    "upload_to_s3": false
  }' \
  --output downloaded_slide.pptx
```

---

## Template Requirements

### Points Slide Template
Your PowerPoint template should have shapes with these names (case-insensitive):
- `Header1` or shape with "header" in name - for the main header
- `Description1` or shape with "description" in name - for the description text
- `Description1_BG` or shape with "points" in name - for bullet points
- `Image` or shape with "image" in name - for the image placeholder

### Image+Text Slide Template
Required shape names:
- `P100` or shape with "title" in name - for the title
- `S100` or shape with "text" in name - for the text content
- Shape with "image" or "picture" in name - for the image

### Table Slide Template
Required shape names:
- Shape with "title" in name - for the title
- A table shape - the service will automatically find and update the first table

---

## Error Handling

The API returns standard HTTP status codes:

- `200 OK`: Successful operation
- `400 Bad Request`: Invalid input data (missing required fields, invalid slide number, etc.)
- `500 Internal Server Error`: Server error (failed to download template, image download failed, etc.)

**Error Response Format:**
```json
{
  "detail": "Error message describing what went wrong"
}
```

---

## Color Formatting

Colors should be provided in hex format:
- Valid: `"#3667B2"`, `"#000000"`, `"#FFFFFF"`
- Invalid: `"blue"`, `"rgb(54,103,178)"`, `"3667B2"`

---

## Future Enhancements

Planned slide types for future releases:
- Chart slides (bar, line, pie charts)
- Diagram slides (flowcharts, process diagrams)
- Split-screen slides (multiple images with captions)
- Timeline slides
- Custom shape manipulation

---

## Best Practices

1. **Template Setup**: Ensure your PowerPoint template has properly named shapes matching the expected names
2. **Image URLs**: Use publicly accessible image URLs or pre-signed S3 URLs
3. **Slide Numbers**: Always use 1-indexed slide numbers (first slide = 1, not 0)
4. **File Naming**: Use descriptive filenames with `.pptx` extension
5. **Error Handling**: Always check the response for errors and handle them appropriately
6. **Testing**: Test with a single slide type before generating multi-slide presentations

---

## Support

For issues or questions:
1. Check the API documentation at `/docs` (Swagger UI)
2. Review this guide for common use cases
3. Check the shape names in your PowerPoint template match the expected names

---

## API Testing with Swagger

The API includes interactive documentation at `http://localhost:8000/docs` where you can:
- Test all endpoints directly from your browser
- See detailed request/response schemas
- View example payloads
- Understand parameter requirements
