# PowerPoint to Google Sheets API - Quick Start

## 🎉 New Feature Added!

Data from extracted PowerPoint presentations is now automatically saved to Google Sheets!

---

## 🚀 What's New?

### 1. **Template ID Support**
Each presentation can be tagged with a `template_id` for easy tracking.

### 2. **Automatic Google Sheets Sync**
All extracted data (slides, shapes, content, images) is appended to your Google Sheet as new rows.

### 3. **Sequential Data Storage**
- Presentation 1 → Rows 2-20
- Presentation 2 → Rows 21-45  
- Presentation 3 → Rows 46-70
- And so on...

---

## 📋 Setup (3 Steps)

### 1. Install New Dependencies
```bash
pip3 install -r requirements.txt
```

### 2. Set Up Google Service Account
Follow the detailed guide in: `GOOGLE_SHEETS_SETUP.md`

Quick version:
- Create Google Cloud project
- Enable Google Sheets API
- Create Service Account
- Download JSON key
- Share your Google Sheet with the service account email

### 3. Update .env File
```env
GOOGLE_SHEET_ID=12F312EN5svGtdMWASyhpmAJeGhe-OIazVZkyFMEwy_I
GOOGLE_SHEETS_CREDENTIALS={"type":"service_account",...paste entire JSON here...}
```

---

## 🎯 API Usage

### Upload & Extract with Template ID

**Thunder Client:**

**URL:** `POST http://localhost:8000/api/upload-and-extract`

**Query Parameters:**
| Key | Value |
|-----|-------|
| `template_id` | `template_001` |
| `upload_images_to_s3` | `true` |
| `save_to_sheets` | `true` |

**Body (Form):**
| Key | Value |
|-----|-------|
| `file` | Select your .pptx file |

---

**cURL:**
```bash
curl -X POST "http://localhost:8000/api/upload-and-extract?template_id=template_001&save_to_sheets=true" \
  -F "file=@presentation.pptx"
```

---

**Python:**
```python
import requests

url = "http://localhost:8000/api/upload-and-extract"
files = {'file': open('presentation.pptx', 'rb')}
params = {
    'template_id': 'my_template_123',
    'upload_images_to_s3': True,
    'save_to_sheets': True
}

response = requests.post(url, files=files, params=params)
print(response.json())
```

---

## 📊 Response Format

```json
{
  "success": true,
  "template_id": "template_001",
  "ppt_s3_url": "https://...",
  "extracted_data": {
    "total_slides": 10,
    "total_images": 5,
    "slides": [...]
  },
  "excel_s3_url": "https://...",
  "sheets_saved": true,
  "sheets_result": {
    "success": true,
    "updated_rows": 25,
    "updated_cells": 250,
    "updated_range": "Sheet1!A2:J26"
  }
}
```

---

## 📈 Google Sheet Columns

| Column | Example |
|--------|---------|
| Template ID | `template_001` |
| PPT Filename | `parati - API Integration Comparison.pptx` |
| S3 URL | `https://zynthimage.s3...` |
| Slide No | `1` |
| Shape Name | `Title 1` |
| Shape Type | `TextBox` |
| Content | `Welcome to Our Presentation` |
| Has Image | `No` |
| Image S3 URL | `https://zynthimage.s3.../image.jpg` |
| Excel URL | `https://zynthimage.s3.../analysis.xlsx` |

---

## 💡 Use Cases

### Track Multiple Templates
```bash
# Template 1
curl -X POST "http://localhost:8000/api/upload-and-extract?template_id=intro_v1" -F "file=@intro.pptx"

# Template 2
curl -X POST "http://localhost:8000/api/upload-and-extract?template_id=sales_v2" -F "file=@sales.pptx"

# Template 3
curl -X POST "http://localhost:8000/api/upload-and-extract?template_id=product_demo" -F "file=@demo.pptx"
```

All data is stored in one Google Sheet with template_id for filtering!

---

## 🎨 Features

✅ Upload PowerPoint to S3  
✅ Extract comprehensive data (text, images, charts, tables)  
✅ Generate Excel analysis file  
✅ Save images to S3  
✅ **NEW:** Append data to Google Sheets  
✅ **NEW:** Template ID tracking  
✅ Sequential row addition  
✅ No data overwriting  

---

## 🔄 Workflow

```
1. Upload PPT → API
2. API uploads to S3
3. API extracts all data
4. API saves images to S3
5. API generates Excel → S3
6. API appends rows to Google Sheets ✨
7. Return complete response
```

---

## 🛠️ Server Commands

```bash
# Install dependencies
pip3 install -r requirements.txt

# Start server
python3 main.py

# Server runs on
http://localhost:8000

# API Docs
http://localhost:8000/docs
```

---

## 📝 Notes

- Each presentation adds rows sequentially
- No data is overwritten
- Template IDs help organize different presentation types
- Service account email must have Editor access to the Google Sheet
- Credentials are stored in `.env` (never commit this file!)

---

**Ready to test! Upload a presentation and watch the data appear in your Google Sheet! 🎊**
