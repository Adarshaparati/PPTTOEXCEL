# Handler Quick Reference Guide

This guide provides quick examples for using each slide handler via the API.

## Base URL
```
POST /api/<handler-endpoint>
```

## 1. Points Slide Handler

**Endpoint:** `/api/generate-points-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 2,
    "header": "Key Features",
    "description": "Our product offers comprehensive solutions",
    "image_url": "https://example.com/feature-image.png",
    "points": [
      {"text": "Easy to use", "color": "#000000", "font_size": 18},
      {"text": "Fast performance", "color": "#3667B2", "font_size": 18},
      {"text": "Secure by design", "color": "#000000", "font_size": 18}
    ],
    "header_color": "#3667B2",
    "description_color": "#000000"
  },
  "upload_to_s3": true
}
```

## 2. Image+Text Slide Handler

**Endpoint:** `/api/generate-image-text-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 3,
    "title": "Product Overview",
    "text": "Our innovative solution transforms the way you work...",
    "image_url": "https://example.com/product.png",
    "title_color": "#3667B2",
    "text_color": "#000000"
  },
  "upload_to_s3": true
}
```

## 3. Table Slide Handler

**Endpoint:** `/api/generate-table-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 4,
    "title": "Pricing Comparison",
    "table_data": [
      ["Feature", "Basic", "Pro", "Enterprise"],
      ["Users", "1-5", "6-50", "Unlimited"],
      ["Storage", "10GB", "100GB", "1TB"],
      ["Support", "Email", "Priority", "24/7 Dedicated"],
      ["Price", "$10/mo", "$50/mo", "Custom"]
    ],
    "header_row": true,
    "header_color": "#3667B2"
  },
  "upload_to_s3": true
}
```

## 4. Phases Slide Handler

**Endpoint:** `/api/generate-phases-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 5,
    "title": "Project Timeline",
    "phases": [
      {
        "name": "Phase 1: Discovery",
        "description": "Requirements gathering and analysis",
        "status": "Completed",
        "color": "#28a745"
      },
      {
        "name": "Phase 2: Development",
        "description": "Core system implementation",
        "status": "In Progress",
        "color": "#ffc107"
      },
      {
        "name": "Phase 3: Testing",
        "description": "QA and user acceptance testing",
        "status": "Planned",
        "color": "#6c757d"
      },
      {
        "name": "Phase 4: Deployment",
        "description": "Production release and monitoring",
        "status": "Planned",
        "color": "#6c757d"
      }
    ],
    "timeline_color": "#3667B2"
  },
  "upload_to_s3": true
}
```

## 5. Statistics Slide Handler

**Endpoint:** `/api/generate-statistics-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 6,
    "title": "Key Metrics Q4 2024",
    "description": "Outstanding performance across all indicators",
    "stat_data": [
      {
        "label": "Total Users",
        "value": "1.2M",
        "color": "#28a745",
        "font_size": 32
      },
      {
        "label": "Revenue Growth",
        "value": "+35%",
        "color": "#007bff",
        "font_size": 32
      },
      {
        "label": "Customer Satisfaction",
        "value": "4.8/5",
        "color": "#ffc107",
        "font_size": 32
      },
      {
        "label": "Market Share",
        "value": "23%",
        "color": "#dc3545",
        "font_size": 32
      }
    ],
    "title_color": "#2E86AB",
    "description_color": "#333333"
  },
  "upload_to_s3": true
}
```

## 6. People Slide Handler

**Endpoint:** `/api/generate-people-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 7,
    "title": "Leadership Team",
    "description": "Meet the experts driving our success",
    "names": [
      "John Smith",
      "Sarah Johnson",
      "Michael Chen",
      "Emily Davis"
    ],
    "designations": [
      "Chief Executive Officer",
      "Chief Technology Officer",
      "VP of Engineering",
      "Product Manager"
    ],
    "descriptions": [
      "15+ years in tech leadership, former VP at Fortune 500",
      "Cloud architecture expert, AI/ML specialist",
      "Full-stack developer, led teams at major startups",
      "Product strategy expert, user experience champion"
    ],
    "title_color": "#2E86AB",
    "description_color": "#333333"
  },
  "upload_to_s3": true
}
```

## 7. Cover Slide Handler

**Endpoint:** `/api/generate-cover-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 1,
    "title": "Annual Business Report 2024",
    "subtitle": "Driving Innovation and Growth",
    "slide_name": "Cover",
    "slide_data_id": "cover_001",
    "slide_type": "Cover",
    "company_name": "TechCorp Solutions Inc.",
    "image": [
      "https://example.com/company-logo.png",
      "https://example.com/cover-background.jpg"
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

## 8. Contact Slide Handler

**Endpoint:** `/api/generate-contact-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 10,
    "title": "Get In Touch",
    "slide_name": "Contact Information",
    "website_link": "https://www.techcorp.com",
    "linkedin_link": "https://linkedin.com/company/techcorp",
    "contact_email": "info@techcorp.com",
    "contact_phone": "+1 (555) 123-4567",
    "image": [
      "https://example.com/contact-qr-code.png",
      "https://example.com/office-photo.jpg"
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

## 9. Images Slide Handler

**Endpoint:** `/api/generate-images-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slide_data": {
    "slide_number": 8,
    "slide_name": "Product Gallery",
    "title": "Our Product Suite",
    "headers": [
      "Mobile Application",
      "Web Platform",
      "Desktop Software",
      "API & Integrations"
    ],
    "descriptions": [
      "Cross-platform mobile app with intuitive design and offline support",
      "Responsive web platform accessible from any device",
      "Powerful desktop software for advanced users and professionals",
      "Seamless API integration with popular third-party tools"
    ],
    "images": [
      "https://example.com/mobile-app-screenshot.png",
      "https://example.com/web-platform-screenshot.jpg",
      "https://example.com/desktop-software-screenshot.png",
      "https://example.com/api-documentation.jpg"
    ]
  },
  "upload_to_s3": true
}
```

## Multi-Slide Generation

**Endpoint:** `/api/generate-multi-slide`

**Example Request:**
```json
{
  "template_s3_url": "https://bucket.s3.region.amazonaws.com/template.pptx",
  "slides_config": [
    {
      "slide_type": "cover",
      "slide_data": {
        "slide_number": 1,
        "title": "Q4 Business Review",
        "subtitle": "2024 Performance Report",
        "company_name": "TechCorp",
        "colors": {"primary": "#2E86AB"}
      }
    },
    {
      "slide_type": "statistics",
      "slide_data": {
        "slide_number": 2,
        "title": "Key Metrics",
        "stat_data": [
          {"label": "Revenue", "value": "$2.5M", "color": "#28a745"},
          {"label": "Growth", "value": "+45%", "color": "#007bff"}
        ]
      }
    },
    {
      "slide_type": "people",
      "slide_data": {
        "slide_number": 3,
        "title": "Our Team",
        "names": ["John Smith", "Sarah Johnson"],
        "designations": ["CEO", "CTO"],
        "descriptions": ["15+ years experience", "Cloud expert"]
      }
    },
    {
      "slide_type": "contact",
      "slide_data": {
        "slide_number": 4,
        "title": "Contact Us",
        "website_link": "https://techcorp.com",
        "contact_email": "info@techcorp.com"
      }
    }
  ],
  "upload_to_s3": true,
  "output_filename": "Q4_Business_Review.pptx"
}
```

## Response Format

All endpoints return:

**Success Response (upload_to_s3 = true):**
```json
{
  "success": true,
  "message": "<Slide type> slide generated successfully",
  "s3_url": "https://bucket.s3.region.amazonaws.com/generated_presentations/...",
  "s3_key": "generated_presentations/...",
  "filename": "...",
  "slide_type": "<type>",
  "timestamp": "2024-01-15T10:30:00.000Z"
}
```

**Success Response (upload_to_s3 = false):**
Returns the PowerPoint file directly for download.

**Error Response:**
```json
{
  "detail": "Error message describing what went wrong"
}
```

## Common Parameters

- `template_s3_url`: Full S3 URL or S3 key of the template
- `slide_number`: Slide index (1-based) to modify
- `upload_to_s3`: Upload result to S3 (true) or return file directly (false)
- `output_filename`: Optional custom filename

## Tips

1. **Slide Numbers**: Use 1-based indexing (slide 1, slide 2, etc.)
2. **Colors**: Always use hex format with # (e.g., "#3667B2")
3. **Images**: Provide full HTTP/HTTPS URLs
4. **Optional Fields**: Most fields are optional except slide_number and core content
5. **Multi-Slide**: Processes slides in order, continues on errors
6. **Template**: Use same template for consistency across slides

## Shape Naming Conventions (for templates)

For handlers to find elements, name shapes in your templates:
- **Points**: `Header1`, `Description1`, `Description1_BG`, `Image`
- **Image+Text**: `P100` (title), `S100` (text), `Image`
- **Table**: `Title`, `Table 1`
- **Phases**: `Title`, `Phase1`, `Phase2`, etc.
- **Statistics**: `Title`, `Description`, `Stat1`, `Stat2`, etc. OR `Label1`, `Value1`, etc.
- **People**: `Title`, `Description`, `Person1`, `Person2`, etc. OR `Name1`, `Designation1`, etc.
- **Cover**: `Title`, `Subtitle`, `Company`, `Image`, `Logo`
- **Contact**: `Title`, `Website`, `LinkedIn`, `Email`, `Phone`, `Image`
- **Images**: `Title`, `Image1`, `Header1`, `Description1`, etc.
