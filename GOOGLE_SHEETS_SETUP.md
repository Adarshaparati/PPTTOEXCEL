# Google Sheets Integration - Setup Guide

## 📋 Overview

This feature automatically saves extracted PowerPoint data to Google Sheets. Each slide's content is appended as new rows, allowing you to track multiple presentations in one sheet.

---

## 🔧 Setup Steps

### Step 1: Create a Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Note your project ID

### Step 2: Enable Google Sheets API

1. In Google Cloud Console, go to **APIs & Services** → **Library**
2. Search for "Google Sheets API"
3. Click **Enable**

### Step 3: Create a Service Account

1. Go to **IAM & Admin** → **Service Accounts**
2. Click **Create Service Account**
3. Enter details:
   - Name: `ppt-extractor-service`
   - Description: `Service account for PPT to Sheets integration`
4. Click **Create and Continue**
5. Skip role assignment (click **Continue**)
6. Click **Done**

### Step 4: Generate Service Account Key

1. Click on your newly created service account
2. Go to **Keys** tab
3. Click **Add Key** → **Create new key**
4. Select **JSON** format
5. Click **Create**
6. A JSON file will download - **keep this safe!**

### Step 5: Share Your Google Sheet

1. Open your Google Sheet: `https://docs.google.com/spreadsheets/d/12F312EN5svGtdMWASyhpmAJeGhe-OIazVZkyFMEwy_I/edit`
2. Click the **Share** button
3. Add the service account email (found in the JSON file, looks like: `ppt-extractor-service@your-project.iam.gserviceaccount.com`)
4. Give it **Editor** permission
5. Click **Send**

### Step 6: Configure Environment Variables

1. Open the downloaded JSON key file
2. Copy its entire content (it should look like a JSON object)
3. Update your `.env` file:

```env
# === Google Sheets Configuration ===
GOOGLE_SHEET_ID=12F312EN5svGtdMWASyhpmAJeGhe-OIazVZkyFMEwy_I
GOOGLE_SHEETS_CREDENTIALS={"type":"service_account","project_id":"your-project-id","private_key_id":"abc123...","private_key":"-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBg...\n-----END PRIVATE KEY-----\n","client_email":"ppt-extractor-service@your-project.iam.gserviceaccount.com","client_id":"123456789","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_x509_cert_url":"https://www.googleapis.com/robot/v1/metadata/x509/ppt-extractor-service%40your-project.iam.gserviceaccount.com"}
```

**Important:** 
- The `GOOGLE_SHEETS_CREDENTIALS` must be a single-line JSON string
- Do NOT add line breaks in the `.env` file
- Keep quotes around the value

### Step 7: Install Dependencies

```bash
pip3 install -r requirements.txt
```

---

## 📊 Google Sheet Structure

The data will be appended with these columns:

| Column | Description |
|--------|-------------|
| **Template ID** | Your custom template identifier |
| **PPT Filename** | Original PowerPoint filename |
| **S3 URL** | S3 link to uploaded PowerPoint |
| **Slide No** | Slide number |
| **Shape Name** | Name of the shape/object |
| **Shape Type** | Type (TextBox, Picture, Chart, etc.) |
| **Content** | Text content or description |
| **Has Image** | Yes/No |
| **Image S3 URL** | S3 link to extracted image |
| **Excel URL** | S3 link to generated Excel file |

---

## 🚀 API Usage

### Upload and Extract with Template ID

**Endpoint:** `POST /api/upload-and-extract`

**Parameters:**
```
file: PowerPoint file
template_id: Your template identifier (e.g., "template_001", "intro_v1")
upload_images_to_s3: true
save_to_sheets: true
```

**Example - Thunder Client:**

| Field | Value |
|-------|-------|
| **file** | Select your .pptx file |
| **template_id** | `template_001` |
| **upload_images_to_s3** | `true` |
| **save_to_sheets** | `true` |

**Example - cURL:**
```bash
curl -X POST "http://localhost:8000/api/upload-and-extract?template_id=template_001&upload_images_to_s3=true&save_to_sheets=true" \
  -F "file=@presentation.pptx"
```

**Example - Python:**
```python
import requests

url = "http://localhost:8000/api/upload-and-extract"
files = {'file': open('presentation.pptx', 'rb')}
params = {
    'template_id': 'template_001',
    'upload_images_to_s3': True,
    'save_to_sheets': True
}

response = requests.post(url, files=files, params=params)
data = response.json()

print(f"Template ID: {data['template_id']}")
print(f"Sheets Saved: {data['sheets_saved']}")
print(f"Rows Added: {data['sheets_result']['updated_rows']}")
```

**Response:**
```json
{
  "success": true,
  "message": "File uploaded and data extracted successfully",
  "template_id": "template_001",
  "ppt_s3_url": "https://...",
  "extracted_data": {...},
  "excel_s3_url": "https://...",
  "sheets_saved": true,
  "sheets_result": {
    "success": true,
    "updated_rows": 15,
    "updated_cells": 150
  }
}
```

---

## 📖 How It Works

### Multiple Presentations

**Presentation 1:**
```
Row 1: Template_001, presentation1.pptx, S3_URL, Slide 1, Title, TextBox, "Welcome", No, "", Excel_URL
Row 2: Template_001, presentation1.pptx, S3_URL, Slide 1, Image1, Picture, "[IMAGE]", Yes, Image_S3_URL, Excel_URL
Row 3: Template_001, presentation1.pptx, S3_URL, Slide 2, Title, TextBox, "About Us", No, "", Excel_URL
...
```

**Presentation 2** (appended after Presentation 1):
```
Row 16: Template_002, presentation2.pptx, S3_URL, Slide 1, Title, TextBox, "Introduction", No, "", Excel_URL
Row 17: Template_002, presentation2.pptx, S3_URL, Slide 1, Logo, Picture, "[IMAGE]", Yes, Image_S3_URL, Excel_URL
...
```

Each new presentation adds rows sequentially, creating a complete history of all processed presentations.

---

## 🎯 Optional: Initialize Sheet Headers

To add headers to your Google Sheet, use this endpoint:

**Endpoint:** `POST /api/sheets/initialize-headers`

```bash
curl -X POST "http://localhost:8000/api/sheets/initialize-headers"
```

This will add column headers to Row 1 of your sheet.

---

## 🔍 Troubleshooting

### Error: "GOOGLE_SHEETS_CREDENTIALS not found"
- Check your `.env` file has the credentials as a single-line JSON
- Restart the server after updating `.env`

### Error: "Permission denied" or "The caller does not have permission"
- Ensure you shared the Google Sheet with the service account email
- Give it **Editor** permission, not just Viewer

### Error: "Invalid JSON format"
- The credentials must be valid JSON
- No line breaks within the JSON string in `.env`
- All quotes must be properly escaped

### Data not appearing in sheet
- Check `sheets_saved` in API response
- Verify `GOOGLE_SHEET_ID` matches your sheet URL
- Ensure the sheet name is "Sheet1" or update range_name parameter

---

## 🎨 Customization

### Change Sheet Name

Edit `services/sheets_service.py`:

```python
def append_rows(self, values: List[List[Any]], range_name: str = "YourSheetName"):
```

### Add More Columns

Edit the `append_ppt_data` method in `services/sheets_service.py` to include additional data fields.

---

## 📚 Additional Resources

- [Google Sheets API Documentation](https://developers.google.com/sheets/api)
- [Service Account Authentication](https://cloud.google.com/iam/docs/service-accounts)
- [Google Cloud Console](https://console.cloud.google.com/)

---

**Your data is now automatically saved to Google Sheets with each API call!** 🎉
