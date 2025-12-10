# API Flow Diagram

## Overall Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         CLIENT APPLICATION                       │
│  (curl, Python requests, Postman, Browser, etc.)                │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             │ HTTP Request (JSON)
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                    FastAPI Application (main.py)                 │
│                                                                   │
│  ┌────────────────────┐        ┌──────────────────────┐        │
│  │  PPT Routes        │        │   Slide Routes       │        │
│  │  /api/*            │        │   /api/slides/*      │        │
│  └────────────────────┘        └──────────────────────┘        │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 │ Route to Handler
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Request Handler & Validation                  │
│                  (routes/slide_routes.py)                        │
│                                                                   │
│  • Validate request with Pydantic models                        │
│  • Extract parameters                                           │
│  • Call appropriate service method                              │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 │ Validated Data
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Slide Data Service                            │
│              (services/slide_data_service.py)                    │
│                                                                   │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  1. Download Template from S3                            │  │
│  │     ↓                                                     │  │
│  │  2. Download Images from URLs                            │  │
│  │     ↓                                                     │  │
│  │  3. Load PowerPoint with python-pptx                     │  │
│  │     ↓                                                     │  │
│  │  4. Find and Update Shapes                               │  │
│  │     - Text shapes (header, description, title)           │  │
│  │     - Image shapes (replace with new images)             │  │
│  │     - Table shapes (populate with data)                  │  │
│  │     ↓                                                     │  │
│  │  5. Save to BytesIO                                      │  │
│  └──────────────────────────────────────────────────────────┘  │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 │ Generated PPT (BytesIO)
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Output Handler                                │
│                                                                   │
│  If upload_to_s3 = true:                                        │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  Upload to S3 (services/s3_service.py)                   │  │
│  │  Return JSON with S3 URL                                 │  │
│  └──────────────────────────────────────────────────────────┘  │
│                                                                   │
│  If upload_to_s3 = false:                                       │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  Stream file directly to client                          │  │
│  │  Return PPT as binary download                           │  │
│  └──────────────────────────────────────────────────────────┘  │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 │ Response
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                         CLIENT APPLICATION                       │
│  Receives either:                                                │
│  • JSON response with S3 URL                                    │
│  • Binary PowerPoint file stream                                │
└─────────────────────────────────────────────────────────────────┘
```

## Detailed Flow for Points Slide Generation

```
POST /api/slides/generate-points-slide
│
├─ Request Body:
│  {
│    "template_s3_url": "presentations/template.pptx",
│    "slide_data": {
│      "slide_number": 2,
│      "header": "Overview",
│      "description": "Description text...",
│      "image_url": "https://example.com/image.png",
│      "points": [
│        {"text": "Point 1", "color": "#3667B2"}
│      ]
│    },
│    "upload_to_s3": true
│  }
│
▼
Pydantic Validation
│  - PointsSlideRequest model
│  - PointsSlideData validation
│  - PointData validation
│
▼
SlideDataService.generate_points_slide()
│
├─ Step 1: Download Template
│  │  S3Service.download_file(s3_key)
│  │  → Returns BytesIO with PPT content
│  │
├─ Step 2: Download Image (if provided)
│  │  requests.get(image_url)
│  │  → Returns BytesIO with image content
│  │
├─ Step 3: Load Presentation
│  │  Presentation(template_bytes)
│  │  → python-pptx Presentation object
│  │
├─ Step 4: Get Target Slide
│  │  prs.slides[slide_number - 1]
│  │
├─ Step 5: Update Shapes
│  │
│  ├─ Find "Header1" shape
│  │  │  Set text to slide_data['header']
│  │  │  Apply header_color if provided
│  │  │
│  ├─ Find "Description1" shape
│  │  │  Set text to slide_data['description']
│  │  │  Apply description_color if provided
│  │  │
│  ├─ Find "Image" shape
│  │  │  Remove old image
│  │  │  Insert new image from URL
│  │  │
│  └─ Find "Description1_BG" shape (points)
│     │  Format points as bullets
│     │  Apply individual point colors
│     │  Apply font sizes
│
├─ Step 6: Save to BytesIO
│  │  prs.save(output)
│  │  → BytesIO with modified PPT
│  │
▼
Output Handling
│
├─ If upload_to_s3 = true:
│  │
│  ├─ S3Service.upload_file()
│  │  │  Upload to "generated_presentations/" folder
│  │  │
│  └─ Return JSON:
│     {
│       "success": true,
│       "s3_url": "https://...",
│       "s3_key": "...",
│       "filename": "...",
│       "timestamp": "..."
│     }
│
└─ If upload_to_s3 = false:
   │
   └─ Return StreamingResponse
      │  media_type: "application/vnd.openxmlformats-..."
      │  Binary PPT file stream
```

## Multi-Slide Generation Flow

```
POST /api/slides/generate-multi-slide
│
├─ Request Body:
│  {
│    "template_s3_url": "...",
│    "slides_config": [
│      {"slide_type": "points", "slide_data": {...}},
│      {"slide_type": "image_text", "slide_data": {...}},
│      {"slide_type": "table", "slide_data": {...}}
│    ]
│  }
│
▼
SlideDataService.generate_multi_slide_presentation()
│
├─ Step 1: Download Template ONCE
│  │  (Reuse same template for all slides)
│  │
├─ Step 2: Load Presentation ONCE
│  │  prs = Presentation(template_bytes)
│  │
├─ Step 3: Process Each Slide Config
│  │
│  ├─ For slide_type = "points":
│  │  │  Call _update_points_slide_in_prs(prs, slide_data)
│  │  │
│  ├─ For slide_type = "image_text":
│  │  │  Call _update_image_text_slide_in_prs(prs, slide_data)
│  │  │
│  └─ For slide_type = "table":
│     │  Call _update_table_slide_in_prs(prs, slide_data)
│
├─ Step 4: Save Final Presentation
│  │  prs.save(output)
│  │
▼
Return result (S3 or stream)
```

## Service Dependencies

```
┌──────────────────────────────────────────────────────────┐
│                   slide_data_service.py                   │
│                                                            │
│  Dependencies:                                            │
│  ├─ s3_service.py (download templates, upload results)   │
│  ├─ python-pptx (PowerPoint manipulation)                │
│  ├─ requests (download images from URLs)                 │
│  ├─ Pillow (image processing)                            │
│  └─ BytesIO (in-memory file handling)                    │
└──────────────────────────────────────────────────────────┘
         │
         │ calls
         ▼
┌──────────────────────────────────────────────────────────┐
│                      s3_service.py                        │
│                                                            │
│  Methods used:                                            │
│  ├─ download_file(key) → Returns file bytes              │
│  └─ upload_file(data, filename, folder) → Returns URL    │
└──────────────────────────────────────────────────────────┘
```

## Error Flow

```
Request
  │
  ▼
Try:
  ├─ Validate Request (Pydantic)
  │  └─ If invalid → 400 Bad Request
  │
  ├─ Download Template
  │  └─ If fails → 500 Internal Server Error
  │
  ├─ Download Images
  │  └─ If fails → 500 Internal Server Error
  │
  ├─ Process Slides
  │  ├─ If slide_number invalid → 400 Bad Request
  │  └─ If shape not found → 500 (continues gracefully)
  │
  ├─ Upload to S3 (if requested)
  │  └─ If fails → 500 Internal Server Error
  │
  └─ Return Response
     └─ 200 OK
```

## Data Flow Example

```
User provides:
  template_s3_url: "presentations/template.pptx"
  slide_number: 2
  header: "Overview"
  image_url: "https://example.com/image.png"
  
        ↓
        
Service fetches:
  1. Template from S3:
     presentations/template.pptx → BytesIO(PPT data)
     
  2. Image from URL:
     https://example.com/image.png → BytesIO(Image data)
     
        ↓
        
Service processes:
  1. Load template into python-pptx
  2. Navigate to slide #2
  3. Find shape named "Header1"
  4. Update text to "Overview"
  5. Find shape named "Image"
  6. Replace with new image
  7. Save modified presentation
  
        ↓
        
Service returns:
  If upload_to_s3=true:
    Upload to: generated_presentations/points_slide_20251205.pptx
    Return: {s3_url: "https://...", s3_key: "..."}
    
  If upload_to_s3=false:
    Stream file directly to client
```

---

This diagram shows the complete flow from client request to final response, including all the intermediate steps and service interactions.
