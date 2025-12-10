# PowerPoint to Excel Extractor - API Documentation

## 🚀 Quick Start

### Installation

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure environment variables** in `.env`:
   ```env
   AWS_ACCESS_KEY_ID=your_access_key
   AWS_SECRET_ACCESS_KEY=your_secret_key
   AWS_REGION=us-east-1
   S3_BUCKET_NAME=zynth-ppt-to-excel
   ```

3. **Start the API server**:
   ```bash
   python main.py
   ```
   
   Or with uvicorn directly:
   ```bash
   uvicorn main:app --reload --port 8000
   ```

4. **Access the API**:
   - API: http://localhost:8000
   - Interactive Docs: http://localhost:8000/docs
   - Alternative Docs: http://localhost:8000/redoc

---

## 📋 API Endpoints

### 1. **Upload PowerPoint to S3**
Upload a PowerPoint file to S3 and get the S3 URL.

**Endpoint**: `POST /api/upload-ppt`

**Request**:
```bash
curl -X POST "http://localhost:8000/api/upload-ppt" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@presentation.pptx"
```

**Response**:
```json
{
  "success": true,
  "message": "File uploaded successfully",
  "s3_url": "https://zynth-ppt-to-excel.s3.us-east-1.amazonaws.com/presentations/presentation_20251127_143052_a1b2c3d4.pptx",
  "s3_key": "presentations/presentation_20251127_143052_a1b2c3d4.pptx",
  "filename": "presentation_20251127_143052_a1b2c3d4.pptx",
  "original_filename": "presentation.pptx",
  "file_size": 524288,
  "timestamp": "2025-11-27T14:30:52.123456"
}
```

---

### 2. **Extract Data from S3 PowerPoint**
Extract data from a PowerPoint file already in S3.

**Endpoint**: `POST /api/extract-ppt`

**Parameters**:
- `s3_key` (required): S3 object key
- `upload_images_to_s3` (optional): Upload images to S3 (default: true)

**Request**:
```bash
curl -X POST "http://localhost:8000/api/extract-ppt?s3_key=presentations/presentation_20251127_143052_a1b2c3d4.pptx&upload_images_to_s3=true"
```

**Response**:
```json
{
  "success": true,
  "message": "Data extracted successfully",
  "extracted_data": {
    "total_slides": 5,
    "total_images": 3,
    "total_charts": 2,
    "total_tables": 1,
    "slides": [
      {
        "slide_number": 1,
        "shapes": [
          {
            "shape_name": "Title 1",
            "shape_type": "Placeholder",
            "content": "Welcome to Our Presentation",
            "has_image": false,
            "image_s3_url": null
          }
        ]
      }
    ]
  },
  "excel_filename": "PPT_Analysis_presentation_20251127_143052.xlsx",
  "local_excel_path": "output/PPT_Analysis_presentation_20251127_143052.xlsx",
  "excel_s3_url": "https://zynth-ppt-to-excel.s3.us-east-1.amazonaws.com/excel_outputs/PPT_Analysis_presentation_20251127_143052.xlsx",
  "download_url": "/api/download/PPT_Analysis_presentation_20251127_143052.xlsx",
  "timestamp": "2025-11-27T14:31:05.789012"
}
```

---

### 3. **Upload and Extract (All-in-One)** ⭐
Upload a PowerPoint file and extract data in one request.

**Endpoint**: `POST /api/upload-and-extract`

**Parameters**:
- `file` (required): PowerPoint file
- `upload_images_to_s3` (optional): Upload images to S3 (default: true)

**Request**:
```bash
curl -X POST "http://localhost:8000/api/upload-and-extract?upload_images_to_s3=true" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@presentation.pptx"
```

**Response**:
```json
{
  "success": true,
  "message": "File uploaded and data extracted successfully",
  "ppt_s3_url": "https://zynth-ppt-to-excel.s3.us-east-1.amazonaws.com/presentations/presentation_20251127_143052_a1b2c3d4.pptx",
  "ppt_s3_key": "presentations/presentation_20251127_143052_a1b2c3d4.pptx",
  "original_filename": "presentation.pptx",
  "file_size": 524288,
  "extracted_data": {
    "total_slides": 5,
    "total_images": 3,
    "total_charts": 2,
    "total_tables": 1,
    "slides": [...]
  },
  "excel_filename": "PPT_Analysis_presentation_20251127_143052.xlsx",
  "local_excel_path": "output/PPT_Analysis_presentation_20251127_143052.xlsx",
  "excel_s3_url": "https://zynth-ppt-to-excel.s3.us-east-1.amazonaws.com/excel_outputs/PPT_Analysis_presentation_20251127_143052.xlsx",
  "download_url": "/api/download/PPT_Analysis_presentation_20251127_143052.xlsx",
  "timestamp": "2025-11-27T14:30:52.123456"
}
```

---

### 4. **Download Excel File**
Download the generated Excel file.

**Endpoint**: `GET /api/download/{filename}`

**Request**:
```bash
curl -X GET "http://localhost:8000/api/download/PPT_Analysis_presentation_20251127_143052.xlsx" \
  -o analysis.xlsx
```

**Response**: Excel file download

---

### 5. **List S3 Files**
List files in S3 bucket by prefix.

**Endpoint**: `GET /api/list-files`

**Parameters**:
- `prefix` (optional): S3 prefix/folder (default: "presentations")

**Request**:
```bash
curl -X GET "http://localhost:8000/api/list-files?prefix=presentations"
```

**Response**:
```json
{
  "success": true,
  "prefix": "presentations",
  "file_count": 3,
  "files": [
    "presentations/presentation1_20251127_143052_a1b2c3d4.pptx",
    "presentations/presentation2_20251127_144023_e5f6g7h8.pptx",
    "presentations/presentation3_20251127_145101_i9j0k1l2.pptx"
  ]
}
```

---

### 6. **Delete S3 File**
Delete a file from S3.

**Endpoint**: `DELETE /api/delete-file`

**Parameters**:
- `s3_key` (required): S3 object key

**Request**:
```bash
curl -X DELETE "http://localhost:8000/api/delete-file?s3_key=presentations/presentation_20251127_143052_a1b2c3d4.pptx"
```

**Response**:
```json
{
  "success": true,
  "message": "File deleted successfully: presentations/presentation_20251127_143052_a1b2c3d4.pptx",
  "s3_key": "presentations/presentation_20251127_143052_a1b2c3d4.pptx",
  "timestamp": "2025-11-27T14:35:22.456789"
}
```

---

### 7. **Generate Presigned URL**
Generate a temporary presigned URL for S3 object.

**Endpoint**: `GET /api/presigned-url`

**Parameters**:
- `s3_key` (required): S3 object key
- `expiration` (optional): URL expiration in seconds (default: 3600)

**Request**:
```bash
curl -X GET "http://localhost:8000/api/presigned-url?s3_key=presentations/presentation_20251127_143052_a1b2c3d4.pptx&expiration=7200"
```

**Response**:
```json
{
  "success": true,
  "presigned_url": "https://zynth-ppt-to-excel.s3.us-east-1.amazonaws.com/presentations/presentation_20251127_143052_a1b2c3d4.pptx?X-Amz-Algorithm=...",
  "s3_key": "presentations/presentation_20251127_143052_a1b2c3d4.pptx",
  "expires_in": 7200,
  "timestamp": "2025-11-27T14:36:10.123456"
}
```

---

### 8. **Health Check**
Check API health status.

**Endpoint**: `GET /api/health`

**Request**:
```bash
curl -X GET "http://localhost:8000/api/health"
```

**Response**:
```json
{
  "status": "healthy",
  "service": "ppt-extractor",
  "timestamp": "2025-11-27T14:30:00.000000"
}
```

---

## 💻 Usage Examples

### Python Example

```python
import requests

# Upload and extract in one step
files = {'file': open('presentation.pptx', 'rb')}
response = requests.post(
    'http://localhost:8000/api/upload-and-extract',
    files=files,
    params={'upload_images_to_s3': True}
)

data = response.json()
print(f"Success: {data['success']}")
print(f"S3 URL: {data['ppt_s3_url']}")
print(f"Excel URL: {data['excel_s3_url']}")
print(f"Download: {data['download_url']}")
print(f"Total Slides: {data['extracted_data']['total_slides']}")
print(f"Total Images: {data['extracted_data']['total_images']}")

# Download Excel file
excel_filename = data['excel_filename']
excel_response = requests.get(f'http://localhost:8000/api/download/{excel_filename}')
with open('local_analysis.xlsx', 'wb') as f:
    f.write(excel_response.content)
```

### JavaScript Example

```javascript
// Upload and extract
const formData = new FormData();
formData.append('file', fileInput.files[0]);

fetch('http://localhost:8000/api/upload-and-extract?upload_images_to_s3=true', {
  method: 'POST',
  body: formData
})
  .then(response => response.json())
  .then(data => {
    console.log('Success:', data.success);
    console.log('S3 URL:', data.ppt_s3_url);
    console.log('Excel S3 URL:', data.excel_s3_url);
    console.log('Total Slides:', data.extracted_data.total_slides);
    console.log('Download URL:', data.download_url);
  })
  .catch(error => console.error('Error:', error));
```

### cURL Example (Postman/Insomnia)

```bash
# Step 1: Upload and extract
curl -X POST "http://localhost:8000/api/upload-and-extract" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@/path/to/presentation.pptx" \
  -F "upload_images_to_s3=true" \
  | jq .

# Step 2: Download Excel
curl -X GET "http://localhost:8000/api/download/PPT_Analysis_presentation_20251127_143052.xlsx" \
  -o analysis.xlsx
```

---

## 📊 Extracted Data Structure

The `extracted_data` object contains:

```json
{
  "total_slides": 5,
  "total_images": 3,
  "total_charts": 2,
  "total_tables": 1,
  "slides": [
    {
      "slide_number": 1,
      "shapes": [
        {
          "shape_name": "Title 1",
          "shape_type": "Placeholder",
          "content": "Welcome",
          "has_image": false,
          "image_s3_url": null
        },
        {
          "shape_name": "Picture 1",
          "shape_type": "Picture",
          "content": "[IMAGE]",
          "has_image": true,
          "image_s3_url": "https://zynth-ppt-to-excel.s3.us-east-1.amazonaws.com/extracted_images/slide_1_shape_3_a1b2c3d4e5f6.jpg"
        }
      ]
    }
  ]
}
```

### Excel File Columns (45+ columns)

The generated Excel file contains comprehensive data:

| Category | Columns |
|----------|---------|
| **Basic Info** | Slide No, Shape Name, Shape Type, Content |
| **Position & Size** | Left (EMU), Top (EMU), Width (EMU), Height (EMU), Inches |
| **Font Details** | Font Name, Size, Bold, Italic, Underline, Color |
| **Text Format** | Alignment, Line Spacing, Paragraph Spacing |
| **Fill Properties** | Fill Color, Type, Transparency |
| **Line Properties** | Line Color, Width, Style, Rotation |
| **Images** | Has Image, Format, Width, Height, File Size, URL, S3 URL, Base64 |
| **Charts** | Chart Type, Title, Data, Categories, Series |
| **Advanced** | Hyperlink, Z-Order, Hidden, Shadow, Glow, Reflection, 3D Effects, Placeholder Type, Animation |

---

## 🔧 Configuration

### Environment Variables

```env
# AWS Configuration (Required)
AWS_ACCESS_KEY_ID=your_access_key
AWS_SECRET_ACCESS_KEY=your_secret_key
AWS_REGION=us-east-1
S3_BUCKET_NAME=zynth-ppt-to-excel

# Server Configuration (Optional)
PORT=8000

# Local Storage (Optional)
LOCAL_OUTPUT_FOLDER=output
```

### S3 Bucket Structure

```
zynth-ppt-to-excel/
├── presentations/          # Uploaded PowerPoint files
│   ├── file1_20251127_143052_a1b2c3d4.pptx
│   └── file2_20251127_144023_e5f6g7h8.pptx
├── excel_outputs/          # Generated Excel files
│   ├── PPT_Analysis_file1_20251127_143052.xlsx
│   └── PPT_Analysis_file2_20251127_144023.xlsx
└── extracted_images/       # Extracted images from presentations
    ├── slide_1_shape_0_abc123.jpg
    └── slide_2_shape_3_def456.png
```

---

## 🚨 Error Handling

All endpoints return standardized error responses:

```json
{
  "detail": "Error message describing what went wrong"
}
```

**Common HTTP Status Codes**:
- `200`: Success
- `400`: Bad Request (invalid file type, empty file)
- `404`: Not Found (file not found in S3)
- `500`: Internal Server Error

---

## 🔐 Security Best Practices

1. **Never commit `.env` file** - Add to `.gitignore`
2. **Use IAM roles** when deploying to AWS (EC2, Lambda)
3. **Enable CORS** for specific origins in production
4. **Use presigned URLs** for temporary access
5. **Rotate AWS keys** regularly
6. **Implement rate limiting** for production
7. **Add authentication** (JWT, API keys) for public APIs

---

## 📈 Performance Tips

1. **Large files**: Use streaming for files > 100MB
2. **Async processing**: Consider background jobs for batch processing
3. **Caching**: Cache frequently accessed files
4. **CDN**: Use CloudFront for S3 content delivery
5. **Compression**: Enable gzip compression

---

## 🐛 Troubleshooting

### API won't start
```bash
# Check if port is in use
lsof -i :8000

# Use different port
uvicorn main:app --port 8001
```

### S3 upload fails
- Verify AWS credentials in `.env`
- Check S3 bucket name is correct
- Ensure IAM user has S3 write permissions

### File not found
- Check S3 key is correct (case-sensitive)
- Verify file exists: `GET /api/list-files`

### Large file timeout
- Increase uvicorn timeout
- Use async processing for large files

---

## 📚 Additional Resources

- **FastAPI Docs**: https://fastapi.tiangolo.com/
- **AWS S3 Docs**: https://docs.aws.amazon.com/s3/
- **python-pptx Docs**: https://python-pptx.readthedocs.io/

---

## 🎯 Next Steps

1. ✅ Start API: `python main.py`
2. ✅ Test endpoints: http://localhost:8000/docs
3. ✅ Upload a PowerPoint file
4. ✅ Download and review Excel analysis
5. ✅ Integrate with your frontend application

**Happy coding! 🚀**
