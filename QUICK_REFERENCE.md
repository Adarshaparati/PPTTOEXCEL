# Slide Generation API - Quick Reference Card

## 🚀 Quick Start

```bash
# Start server
python3 main.py

# Access docs
http://localhost:8000/docs
```

## 📍 Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/slides/generate-points-slide` | POST | Points/bullet slide |
| `/api/slides/generate-image-text-slide` | POST | Image+text slide |
| `/api/slides/generate-table-slide` | POST | Table slide |
| `/api/slides/generate-multi-slide` | POST | Multiple slides |
| `/api/slides/slide-types` | GET | List slide types |

## 📝 Minimal Request Examples

### Points Slide
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slide_data": {
    "slide_number": 2,
    "header": "Title",
    "description": "Description"
  }
}
```

### Image+Text Slide
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slide_data": {
    "slide_number": 3,
    "title": "Title",
    "text": "Content"
  }
}
```

### Table Slide
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slide_data": {
    "slide_number": 4,
    "title": "Title",
    "table_data": [
      ["Col1", "Col2"],
      ["Data1", "Data2"]
    ]
  }
}
```

### Multi-Slide
```json
{
  "template_s3_url": "presentations/template.pptx",
  "slides_config": [
    {
      "slide_type": "points",
      "slide_data": {
        "slide_number": 2,
        "header": "Title",
        "description": "Desc"
      }
    }
  ]
}
```

## 🎨 Optional Fields

### Points Slide
```json
{
  "image_url": "https://example.com/image.png",
  "points": [
    {"text": "Point 1", "color": "#3667B2", "font_size": 14}
  ],
  "header_color": "#3667B2",
  "description_color": "#000000",
  "background_color": "#FFFFFF"
}
```

### Image+Text Slide
```json
{
  "image_url": "https://example.com/image.png",
  "title_color": "#3667B2",
  "text_color": "#000000"
}
```

### Table Slide
```json
{
  "header_row": true,
  "header_color": "#3667B2"
}
```

## 📤 Output Options

```json
{
  "upload_to_s3": true,           // true = S3 upload, false = direct download
  "output_filename": "custom.pptx" // Optional custom filename
}
```

## ✅ Success Response (S3)

```json
{
  "success": true,
  "message": "Points slide generated successfully",
  "s3_url": "https://bucket.s3.amazonaws.com/path/file.pptx",
  "s3_key": "generated_presentations/file.pptx",
  "filename": "file.pptx",
  "slide_type": "points",
  "timestamp": "2025-12-05T18:06:00.000000"
}
```

## ❌ Error Response

```json
{
  "detail": "Error message"
}
```

**Common Errors:**
- `400` - Invalid input (missing fields, bad slide number)
- `500` - Server error (template not found, image download failed)

## 🎯 Template Shape Names

### Points Slide
- `Header1` - Main header
- `Description1` - Description
- `Description1_BG` - Bullet points
- `Image` - Image placeholder

### Image+Text Slide
- `P100` or "title" - Title text
- `S100` or "text" - Content text
- "image" or "picture" - Image placeholder

### Table Slide
- "title" - Title text
- Table shape - Auto-detected

## 🎨 Color Format

```
✅ Correct:  "#3667B2", "#000000", "#FFFFFF"
❌ Wrong:    "blue", "rgb(54,103,178)", "3667B2"
```

## 🐍 Python Example

```python
import requests

response = requests.post(
    "http://localhost:8000/api/slides/generate-points-slide",
    json={
        "template_s3_url": "presentations/template.pptx",
        "slide_data": {
            "slide_number": 2,
            "header": "Overview",
            "description": "Key points"
        },
        "upload_to_s3": True
    }
)

result = response.json()
print(f"S3 URL: {result['s3_url']}")
```

## 💻 cURL Example

```bash
curl -X POST "http://localhost:8000/api/slides/generate-points-slide" \
  -H "Content-Type: application/json" \
  -d '{
    "template_s3_url": "presentations/template.pptx",
    "slide_data": {
      "slide_number": 2,
      "header": "Title",
      "description": "Description"
    }
  }'
```

## 🔧 Common Tasks

### Download Instead of S3 Upload
```bash
curl -X POST "http://localhost:8000/api/slides/generate-points-slide" \
  -H "Content-Type: application/json" \
  -d '{"template_s3_url": "...", "slide_data": {...}, "upload_to_s3": false}' \
  --output output.pptx
```

### Generate Multiple Slides at Once
Use `/api/slides/generate-multi-slide` with `slides_config` array

### Get Slide Type Info
```bash
curl http://localhost:8000/api/slides/slide-types
```

## 📚 Documentation Files

- **SLIDE_API_GUIDE.md** - Full API docs
- **SLIDE_GENERATION_README.md** - User guide
- **API_FLOW_DIAGRAM.md** - Architecture diagram
- **example_slide_generation.py** - Working examples

## 🔍 Debugging Tips

1. Check shape names: Use PowerPoint's Selection Pane
2. Verify slide numbers: 1-indexed (first slide = 1)
3. Test image URLs: Must be publicly accessible
4. Check S3 credentials: Verify `.env` file
5. View logs: Check terminal output for errors

## ⚡ Performance Tips

- Use `generate-multi-slide` for multiple slides
- Optimize image sizes before uploading
- Cache frequently used templates
- Use async requests for parallel generation

## 🎓 Learning Path

1. ✅ Read SLIDE_GENERATION_README.md
2. ✅ Run example_slide_generation.py
3. ✅ Test single slide generation
4. ✅ Try multi-slide generation
5. ✅ Add custom slide types

---

**Version**: 2.0.0  
**Last Updated**: December 5, 2025  
**Support**: Check `/docs` for interactive API testing
