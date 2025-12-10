# PowerPoint Processing & Generation Suite

This project provides comprehensive PowerPoint processing capabilities:
1. **Excel → PowerPoint Generator** - Creates PowerPoint presentations from Excel data using Microsoft Graph
2. **PowerPoint Extraction API** - Extracts data from PowerPoint files to Excel (via S3)
3. **🆕 Slide Generation API** - Programmatically generate custom slides (Points, Image+Text, Tables)

## 🆕 NEW: Slide Generation API (v2.0)

Generate custom PowerPoint slides programmatically with our REST API! Perfect for:
- Automated report generation
- Dynamic presentations from data
- Template-based content creation
- Bulk slide generation

**Quick Example:**
```bash
curl -X POST "http://localhost:8000/api/slides/generate-points-slide" \
  -H "Content-Type: application/json" \
  -d '{
    "template_s3_url": "presentations/template.pptx",
    "slide_data": {
      "slide_number": 2,
      "header": "Overview",
      "description": "Key project highlights",
      "points": [
        {"text": "Feature 1", "color": "#3667B2"},
        {"text": "Feature 2", "color": "#000000"}
      ]
    }
  }'
```

**📚 Documentation:**
- [Quick Start Guide](SLIDE_GENERATION_README.md) - Get started in 5 minutes
- [Complete API Guide](SLIDE_API_GUIDE.md) - Detailed API documentation
- [Quick Reference](QUICK_REFERENCE.md) - Cheat sheet for developers
- [API Flow Diagram](API_FLOW_DIAGRAM.md) - Architecture overview
- [Examples](example_slide_generation.py) - Working code examples

**🎯 Supported Slide Types:**
- **Points Slides**: Header, description, bullet points, and images
- **Image+Text Slides**: Title, text content, and images
- **Table Slides**: Title and data tables
- **Multi-Slide**: Combine multiple types in one presentation

**Start the API Server:**
```bash
python3 main.py
# Visit http://localhost:8000/docs for interactive API documentation
```

---

## Excel → PowerPoint Generator (Original Feature)

This creates a **new PowerPoint presentation** in OneDrive using your **Framework Master Deck (.pptx)** as the template. 
It reads the presentation title from cell **A1** of a local Excel file and names the generated deck accordingly.

## What it does
1. Reads `A1` from a local Excel file (default: `Book.xlsx`).
2. Authenticates you to Microsoft Graph via **Device Code** flow (interactive in terminal).
3. Downloads your template PowerPoint from **OneDrive for Business**.
4. Uploads a **new copy** named after `A1` into `/Documents/GeneratedPresentations` in your OneDrive.

> Works with Microsoft 365 Work/School accounts (Business/Enterprise).

---

## Setup

### 1) App Registration (once)
1. Go to **Entra ID (Azure AD) → App registrations → New registration**.
2. Name: `ExcelToPptGenerator`.
3. Account type: `Single tenant` (your org) is fine.
4. Add a **Redirect URI (Public client/native)**: `http://localhost`.
5. After creating the app, open **Authentication** → enable **"Allow public client flows"** (for device code).
6. **API permissions → Microsoft Graph (Delegated)** add:
   - `Files.ReadWrite.All`
   - `offline_access`
7. Click **Grant admin consent** (recommended).

> This project uses **delegated** permissions and the **Device Code** flow — easiest to run from a local machine.

### 2) Local environment
- Install Python 3.9+
- In a terminal, run:
```bash
pip install -r requirements.txt
cp .env.example .env
```
- Open `.env` and fill your **TENANT_ID** and **CLIENT_ID**.
- Put your local Excel file (default `Book.xlsx`) next to the script, or update `LOCAL_EXCEL_PATH` in `.env`.

### 3) OneDrive paths
- Upload your master deck, e.g. to: `/Documents/FrameworkMasterDeck.pptx`
- Update `.env` → `TEMPLATE_ONEDRIVE_PATH=/Documents/FrameworkMasterDeck.pptx`
- The generated file is uploaded to `/Documents/GeneratedPresentations` (auto-created if missing).

---

## Run
```bash
python -m src.generate_ppt_from_excel
```
You will see a device code in the terminal — open the URL, paste the code, sign in with your **work account**.

On success, check your OneDrive folder:
```
/Documents/GeneratedPresentations/<Title from A1>.pptx
```

---

## Config (.env)
See `.env.example` for all options:

- `TENANT_ID` — your Entra ID tenant GUID
- `CLIENT_ID` — your app registration Client (Application) ID
- `LOCAL_EXCEL_PATH` — local path to Excel file (default `Book.xlsx`)
- `EXCEL_TITLE_CELL` — cell address to read (default `A1`)
- `TEMPLATE_ONEDRIVE_PATH` — OneDrive path to template (default `/Documents/FrameworkMasterDeck.pptx`)
- `DEST_FOLDER_ONEDRIVE` — OneDrive folder for generated decks (default `/Documents/GeneratedPresentations`)

---

## Notes
- This sample uses `/me/drive/...` endpoints (delegated user context).
- If you want **application-only** (daemon) auth or SharePoint Site/Drive targets, that’s doable — ping me and I’ll extend this.
