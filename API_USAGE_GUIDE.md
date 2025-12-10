# PowerPoint Generation API Usage Guide

## 🚀 New Feature: Generate PPT from Template + Google Sheets Data

### Overview
This API generates a PowerPoint presentation by:
1. Taking a template PPT from S3
2. Finding data in Google Sheets by Template ID
3. Replacing shape content based on shape names
4. Applying colors and images from the data

---

## 📡 API Endpoint

### **POST** `/api/generate-ppt`

Generate a PowerPoint presentation from template and Google Sheets data.

#### Parameters:
| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `template_s3_url` | string | ✅ Yes | Full S3 URL of the template PowerPoint |
| `template_id` | string | ✅ Yes | Template ID to fetch data from Google Sheets |
| `upload_to_s3` | boolean | ❌ No | Upload result to S3 (default: `true`) |
| `output_filename` | string | ❌ No | Custom filename for output |

#### Example Request:
```
POST http://localhost:8000/api/generate-ppt?template_s3_url=https://zynthimage.s3.us-east-1.amazonaws.com/presentations/parati%20-%20API%20Integration%20Comparison_20251127_190055_062df089.pptx&template_id=avcvdgd5353&upload_to_s3=true
```

#### Example Response:
```json
{
  "success": true,
  "message": "Presentation generated successfully",
  "stats": {
    "template_id": "avcvdgd5353",
    "slides_count": 9,
    "data_rows": 98,
    "text_replacements": 45,
    "output_filename": "Generated_avcvdgd5353_20251201_123456.pptx",
    "file_size": 245678
  },
  "s3_url": "https://zynthimage.s3.us-east-1.amazonaws.com/generated_presentations/Generated_avcvdgd5353_20251201_123456.pptx",
  "s3_key": "generated_presentations/Generated_avcvdgd5353_20251201_123456.pptx",
  "local_path": "output/Generated_avcvdgd5353_20251201_123456.pptx",
  "download_url": "/api/download/Generated_avcvdgd5353_20251201_123456.pptx",
  "timestamp": "2025-12-01T12:34:56.789"
}
```

---

## 📥 Download Endpoint

### **GET** `/api/generate-ppt-download`

Generate and directly download the PowerPoint file (no S3 upload).

#### Parameters:
Same as above, but `upload_to_s3` is ignored.

#### Example Request:
```
GET http://localhost:8000/api/generate-ppt-download?template_s3_url=https://zynthimage.s3.us-east-1.amazonaws.com/presentations/parati%20-%20API%20Integration%20Comparison_20251127_190055_062df089.pptx&template_id=avcvdgd5353
```

This will return the PowerPoint file directly for download.

---

## 🎯 How It Works

### 1. **Shape Name Matching**
The API matches shape names in your template with the "Shape Name" column in Google Sheets.

**Example:**
- Template has a shape named: `Google Shape;58;p15`
- Google Sheets has a row with:
  - Shape Name: `Google Shape;58;p15`
  - Content: `parati - API Integration`
- **Result:** The shape's text is replaced with `parati - API Integration`

### 2. **Placeholder Replacement**
You can also use `{{ShapeName}}` placeholders in your template text.

**Example:**
- Template text: `"Welcome to {{Google Shape;58;p15}}"`
- **Result:** `"Welcome to parati - API Integration"`

### 3. **Color Application**
- Fill colors from "Fill Color" column are applied to shapes
- Font colors from "Font Color" column are applied to text

### 4. **Image Replacement**
- If "Image S3 URL" or "Image URL" is provided, the shape is replaced with that image

---

## 📊 Google Sheets Data Format

Your Google Sheets should have these columns:
- **Template ID** (Column A): Filter rows by this ID
- **Slide No** (Column B): Which slide the shape is on
- **Shape Name** (Column C): Exact name of the shape in PPT
- **Content** (Column E): Text to replace in the shape
- **Fill Color** (Column V): Hex color for shape background (e.g., `#3667B2`)
- **Font Color** (Column R): Hex color for text (e.g., `#FFFFFF`)
- **Image S3 URL** (Column AG): URL of image to replace shape

---

## 🧪 Testing with Thunder Client

### Request Configuration:
```
Method: POST
URL: http://localhost:8000/api/generate-ppt

Query Parameters:
  - template_s3_url: https://zynthimage.s3.us-east-1.amazonaws.com/presentations/parati - API Integration Comparison_20251127_190055_062df089.pptx
  - template_id: avcvdgd5353
  - upload_to_s3: true

Headers: (none needed)
Body: (none needed)
```

---

## 🎨 Example Use Case

You have a presentation template with shapes named:
- `Google Shape;58;p15` - Title slide
- `Google Shape;59;p16` - Subtitle
- `Google Shape;60;p17` - Content box

Your Google Sheets has data for template_id `AD12`:
```
Template ID | Shape Name              | Content
AD12        | Google Shape;58;p15     | "Product Overview"
AD12        | Google Shape;59;p16     | "Q4 2025 Results"
AD12        | Google Shape;60;p17     | "Revenue increased by 25%"
```

**API Call:**
```
POST /api/generate-ppt?template_s3_url=...&template_id=AD12
```

**Result:** A new PPT with all shapes filled with your data!

---

## 💡 Tips

1. **Shape Names Must Match Exactly**: Ensure shape names in your template match those in Google Sheets
2. **Use Template IDs**: Group related data using the same template_id
3. **Color Format**: Use hex colors like `#3667B2` or `#FFFFFF`
4. **Image URLs**: Must be publicly accessible or from S3
5. **Large Files**: For large presentations, use the direct download endpoint to avoid S3 delays

---

## 🔧 Troubleshooting

### No replacements made?
- Check that Template ID matches exactly (case-sensitive)
- Verify shape names in PPT match Google Sheets exactly
- Ensure Google Sheets has data for that template_id

### Colors not applied?
- Check that color values are in hex format: `#RRGGBB`
- Ensure "Fill Color" and "Font Color" columns have valid data

### Images not replaced?
- Verify image URLs are accessible
- Check that shape names match exactly
- Ensure images are in supported formats (JPG, PNG)

---

## 📞 Support

For issues or questions, check the server logs for detailed debug information.
