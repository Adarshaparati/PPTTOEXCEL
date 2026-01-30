# S3 Bucket Override Fix

## Problem
When passing a full S3 URL from a different bucket (e.g., `https://zynthimage.s3.amazonaws.com/presentations/1769781142431_adarshatest.pptx`), the `/extract-ppt` endpoint would fail with:
```
{"detail":"File not found in S3: presentations/1769781142431_adarshatest.pptx"}
```

This happened because:
1. The code correctly extracted the S3 key: `presentations/1769781142431_adarshatest.pptx`
2. But then tried to download from the **configured bucket** (`S3_BUCKET_NAME` env variable)
3. The file actually exists in the `zynthimage` bucket, not the configured bucket

## Solution

### 1. Updated `routes/ppt_routes.py` - `/extract-ppt` endpoint
Added bucket detection from S3 URL:
```python
# Extract bucket name from URL if provided
# Format: bucket.s3.amazonaws.com or bucket.s3.region.amazonaws.com
bucket_override = None
if s3_key.startswith('http://') or s3_key.startswith('https://'):
    parsed_url = urlparse(s3_key)
    netloc = parsed_url.netloc
    if netloc.endswith('.s3.amazonaws.com'):
        bucket_override = netloc.split('.')[0]
    
    s3_key = unquote(parsed_url.path.lstrip('/'))

# Use bucket override if detected from URL
if bucket_override:
    ppt_bytes = s3_service.download_file(s3_key, bucket_name=bucket_override)
else:
    ppt_bytes = s3_service.download_file(s3_key)
```

### 2. Updated `services/s3_service.py` - `download_file()` method
Added optional `bucket_name` parameter:
```python
def download_file(self, s3_key: str, bucket_name: Optional[str] = None) -> Optional[bytes]:
    """
    Download file from S3 and return raw bytes.

    Args:
        s3_key: S3 object key (should be unquoted/decoded)
        bucket_name: Optional bucket name override (defaults to configured bucket)
    """
    bucket = bucket_name if bucket_name else self.bucket_name
    # ... rest of implementation
```

### 3. Fixed Excel File Handling
Uncommented and fixed the code for saving extracted Excel files:
- Creates `output/` folder locally
- Saves Excel file with timestamp
- Uploads to S3 in `extracted_excel/` folder
- Returns proper URLs in response

## How It Works Now

1. **With Full S3 URL** (from different bucket):
   ```
   GET /extract-ppt?s3_key=https://zynthimage.s3.amazonaws.com/presentations/1769781142431_adarshatest.pptx
   ```
   - Detects bucket: `zynthimage`
   - Extracts key: `presentations/1769781142431_adarshatest.pptx`
   - Downloads from `zynthimage` bucket ✅

2. **With S3 Key Only** (default bucket):
   ```
   GET /extract-ppt?s3_key=presentations/myfile.pptx
   ```
   - Uses configured `S3_BUCKET_NAME` env variable ✅

3. **With Full S3 URL** (same bucket):
   ```
   GET /extract-ppt?s3_key=https://mybucket.s3.amazonaws.com/presentations/myfile.pptx
   ```
   - Detects bucket: `mybucket`
   - Works regardless of `S3_BUCKET_NAME` config ✅

## Response Format
The endpoint now returns:
```json
{
    "success": true,
    "message": "Data extracted successfully",
    "template_id": "template-123",
    "excel_filename": "PPT_Analysis_filename_timestamp.xlsx",
    "local_excel_path": "output/PPT_Analysis_filename_timestamp.xlsx",
    "excel_s3_url": "https://bucket.s3.region.amazonaws.com/extracted_excel/...",
    "download_url": "/api/download/PPT_Analysis_filename_timestamp.xlsx",
    "sheets_saved": true/false,
    "sheets_result": {...},
    "timestamp": "2026-01-30T12:34:56.789Z"
}
```

## Benefits
✅ Works with URLs from different S3 buckets
✅ Backward compatible with S3 keys only
✅ Properly saves and uploads extracted Excel files
✅ Clear error messages for debugging
✅ No breaking changes to existing API
