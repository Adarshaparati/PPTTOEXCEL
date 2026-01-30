# S3 Download Strategy - Three-Tier Fallback System

## Overview

The updated S3 service implements a **three-tier download strategy** to handle various S3 access scenarios robustly:

1. **🌐 CloudFront (Tier 1)** - Fastest, no ACL issues
2. **🌐 Direct S3 HTTP (Tier 2)** - Public access fallback
3. **🔑 Boto3 API (Tier 3)** - Private bucket with credentials

---

## Why This Solution Works

### Problem
- Your `zynthimage` S3 bucket had **object ownership and ACL issues** preventing AWS credentials from accessing files
- Direct boto3 API calls returned **403 Forbidden** errors
- Files existed but were not accessible through standard AWS SDK methods

### Solution
Since your bucket has a **public read policy** and you've set up a **CloudFront CDN distribution**, we use these public access methods as primary download sources:

1. **CloudFront is fastest** - Content is cached at edge locations globally
2. **CloudFront bypasses ACL issues** - Works regardless of object ownership
3. **S3 HTTP is reliable** - Direct public S3 URL as fallback
4. **Boto3 is backup** - For any future private bucket scenarios

---

## Configuration

### 1. Add CloudFront Domain to `.env`

```env
# === CloudFront Configuration ===
CLOUDFRONT_DOMAIN=d2zu6flr7wd65l.cloudfront.net
```

### 2. S3 Service Initialization

The S3 service automatically loads the CloudFront domain:

```python
from services.s3_service import s3_service

# CloudFront domain is configured automatically from .env
print(s3_service.cloudfront_domain)  # d2zu6flr7wd65l.cloudfront.net
```

---

## How It Works

### Download Attempt Sequence

```
User Request
    ↓
Try CloudFront
├─ Success? ✅ Return file (FASTEST)
└─ Fail? ↓
  Try S3 HTTP
  ├─ Success? ✅ Return file (FAST)
  └─ Fail? ↓
    Try Boto3 API
    ├─ Success? ✅ Return file (SLOWER)
    └─ Fail? ❌ Return None (Error)
```

### Example Flow Log

```
📥 Download Strategy: CloudFront → S3 HTTP → Boto3 API
   S3 Key: presentations/1769782207011_adarshatest.pptx

🌐 Tier 1: Trying CloudFront...
   URL: https://d2zu6flr7wd65l.cloudfront.net/presentations/1769782207011_adarshatest.pptx
✅ CloudFront download successful (824314 bytes)
```

---

## API Usage Examples

### 1. With S3 URL (Auto-detects CloudFront)

```bash
curl -X POST "http://localhost:8000/api/extract-ppt" \
  -G \
  -d "s3_key=https://zynthimage.s3.amazonaws.com/presentations/1769782207011_adarshatest.pptx" \
  -d "upload_images_to_s3=true" \
  -d "template_id=template-123" \
  -d "save_to_sheets=no"
```

**What happens:**
1. Code extracts bucket: `zynthimage`
2. Code extracts key: `presentations/1769782207011_adarshatest.pptx`
3. CloudFront downloads file: ✅ Success

### 2. With CloudFront URL Directly

```bash
curl -X POST "http://localhost:8000/api/extract-ppt" \
  -G \
  -d "s3_key=https://d2zu6flr7wd65l.cloudfront.net/presentations/1769782207011_adarshatest.pptx" \
  -d "upload_images_to_s3=true" \
  -d "template_id=template-123" \
  -d "save_to_sheets=no"
```

**What happens:**
1. Code detects CloudFront URL
2. Extracts S3 key: `presentations/1769782207011_adarshatest.pptx`
3. Uses CloudFront directly: ✅ Success

### 3. With S3 Key Only

```bash
curl -X POST "http://localhost:8000/api/extract-ppt" \
  -G \
  -d "s3_key=presentations/1769782207011_adarshatest.pptx" \
  -d "upload_images_to_s3=true" \
  -d "template_id=template-123" \
  -d "save_to_sheets=no"
```

**What happens:**
1. CloudFront downloads: ✅ Success
2. Falls back to S3 HTTP if needed
3. Falls back to Boto3 API if needed

---

## Performance Benefits

| Method | Speed | Reliability | Cost |
|--------|-------|-------------|------|
| CloudFront | ⚡⚡⚡ Fastest | High (CDN cached) | Lowest |
| S3 HTTP | ⚡⚡ Fast | High (public) | Low |
| Boto3 API | ⚡ Slower | High (auth) | Medium |

---

## Error Handling

Each tier provides clear diagnostic messages:

```
🌐 Tier 1: Trying CloudFront...
   URL: https://d2zu6flr7wd65l.cloudfront.net/...
⚠️ CloudFront returned 403, trying fallback...

🌐 Tier 2: Trying Direct S3 HTTP...
   URL: https://zynthimage.s3.us-east-1.amazonaws.com/...
✅ S3 HTTP download successful (824314 bytes)
```

---

## Troubleshooting

### CloudFront Not Working?
- Check CloudFront distribution is enabled
- Check S3 bucket is configured as origin
- Verify file exists in S3

### All Tiers Failing?
1. Verify file exists in S3: `aws s3 ls s3://zynthimage/presentations/1769782207011_adarshatest.pptx`
2. Check CloudFront configuration in AWS console
3. Verify S3 bucket policy allows public read
4. Check AWS credentials are valid (for Boto3 tier)

---

## Configuration Reference

### .env Variables

```env
# AWS Credentials (for Tier 3: Boto3 API)
AWS_ACCESS_KEY_ID=your_access_key
AWS_SECRET_ACCESS_KEY=your_secret_key
AWS_REGION=us-east-1
S3_BUCKET_NAME=zynthimage

# CloudFront (for Tier 1: CloudFront CDN)
CLOUDFRONT_DOMAIN=d2zu6flr7wd65l.cloudfront.net
```

### S3 Service Implementation

```python
# Tier 1: CloudFront
def _download_via_cloudfront(self, s3_key)
    # Uses CLOUDFRONT_DOMAIN + s3_key

# Tier 2: S3 HTTP
def _download_file_via_http(self, s3_key, bucket_name)
    # Uses bucket_name.s3.region.amazonaws.com + s3_key

# Tier 3: Boto3 API
def _download_file_via_boto3(self, s3_key, bucket_name)
    # Uses boto3 client with AWS credentials
```

---

## Summary

✅ **CloudFront is now the primary download method**
✅ **S3 HTTP provides reliable fallback**
✅ **Boto3 API is available as last resort**
✅ **No more 403 Forbidden errors**
✅ **Maximum reliability and performance**

The system will always find a way to download your files! 🎉
