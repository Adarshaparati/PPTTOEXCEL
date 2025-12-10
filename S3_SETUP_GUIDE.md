# S3 PowerPoint Extractor - Setup Guide

## Overview
This script downloads PowerPoint files from AWS S3, extracts comprehensive data, and saves an Excel analysis file locally.

## Prerequisites

### 1. AWS Account Setup
- Create an AWS account at https://aws.amazon.com
- Access the AWS Management Console

### 2. Create S3 Bucket
1. Go to S3 service in AWS Console
2. Click "Create bucket"
3. Choose a unique bucket name (e.g., `my-ppt-files`)
4. Select your preferred region
5. Keep default settings or customize as needed
6. Click "Create bucket"

### 3. Upload PowerPoint File
1. Open your S3 bucket
2. Click "Upload"
3. Add your PowerPoint file(s)
4. Note the file path (e.g., `presentations/myfile.pptx`)

### 4. Create IAM User with S3 Access
1. Go to IAM service in AWS Console
2. Click "Users" → "Add users"
3. Enter username (e.g., `ppt-extractor`)
4. Select "Access key - Programmatic access"
5. Click "Next: Permissions"
6. Choose "Attach existing policies directly"
7. Search and select `AmazonS3ReadOnlyAccess` (or create custom policy)
8. Complete user creation
9. **IMPORTANT**: Save the Access Key ID and Secret Access Key

## Installation

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure Environment Variables
Edit the `.env` file in your project root:

```env
# === AWS S3 Configuration ===
AWS_ACCESS_KEY_ID=AKIAIOSFODNN7EXAMPLE
AWS_SECRET_ACCESS_KEY=wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY
AWS_REGION=us-east-1
S3_BUCKET_NAME=my-ppt-files
S3_PPT_KEY=presentations/myfile.pptx
LOCAL_OUTPUT_FOLDER=output
```

**Configuration Details:**
- `AWS_ACCESS_KEY_ID`: Your IAM user's access key
- `AWS_SECRET_ACCESS_KEY`: Your IAM user's secret key
- `AWS_REGION`: AWS region where your bucket is located (e.g., us-east-1, eu-west-1)
- `S3_BUCKET_NAME`: Name of your S3 bucket
- `S3_PPT_KEY`: Full path to your PPT file in the bucket (e.g., `folder/subfolder/file.pptx`)
- `LOCAL_OUTPUT_FOLDER`: Local folder to save Excel files (default: `output`)

## Usage

### Run the Extractor
```bash
python src/extract_ppt_from_s3.py
```

### Expected Output
```
============================================================
📊 PowerPoint to Excel Extractor (S3 Version)
============================================================
🗂️  S3 Bucket: my-ppt-files
📄 PPT File: presentations/myfile.pptx
🌍 Region: us-east-1
============================================================
⬇️ Downloading from S3: s3://my-ppt-files/presentations/myfile.pptx
✅ Downloaded 524288 bytes
🔍 Extracting data from PowerPoint...
============================================================
✅ Extraction completed successfully!
============================================================
💾 Excel saved to: /path/to/output/PPT_Analysis_myfile_20251126_143052.xlsx
📁 Images extracted to: ./extracted_images/
============================================================
```

## Output Files

### 1. Excel Analysis File
**Location**: `output/PPT_Analysis_[filename]_[timestamp].xlsx`

**Contains 45+ columns with:**
- Slide numbers and shape names
- Position and size (EMU and inches)
- Font details (name, size, color, bold, italic)
- Text alignment and spacing
- Fill properties (color, type, transparency)
- Line properties (color, width, style)
- Images (format, dimensions, URLs, base64)
- Charts (type, title, data, series, categories)
- Hyperlinks and effects
- And much more...

### 2. Extracted Images
**Location**: `extracted_images/`

All images from the presentation are extracted with filenames like:
- `slide_1_shape_0_a1b2c3d4e5f6.jpg`
- `slide_2_shape_3_f6e5d4c3b2a1.png`

## AWS Regions Reference

Common AWS regions:
- `us-east-1` - US East (N. Virginia)
- `us-west-2` - US West (Oregon)
- `eu-west-1` - Europe (Ireland)
- `ap-southeast-1` - Asia Pacific (Singapore)
- `ap-south-1` - Asia Pacific (Mumbai)

Find your region in the S3 bucket properties or AWS Console URL.

## Security Best Practices

### 1. Never Commit Credentials
- Keep `.env` file in `.gitignore`
- Never share your AWS keys publicly

### 2. Use Minimal Permissions
Instead of `AmazonS3ReadOnlyAccess`, create a custom policy:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "s3:GetObject",
        "s3:ListBucket"
      ],
      "Resource": [
        "arn:aws:s3:::my-ppt-files",
        "arn:aws:s3:::my-ppt-files/*"
      ]
    }
  ]
}
```

### 3. Rotate Keys Regularly
- Change your AWS access keys periodically
- Delete unused IAM users

### 4. Use IAM Roles (for EC2/Lambda)
If running on AWS infrastructure, use IAM roles instead of access keys.

## Troubleshooting

### Error: "Missing required environment variables"
- Check that all AWS variables are set in `.env`
- Ensure no typos in variable names

### Error: "NoSuchBucket" or "AccessDenied"
- Verify bucket name is correct
- Check IAM permissions include S3 read access
- Ensure region matches bucket location

### Error: "NoSuchKey"
- Verify the S3_PPT_KEY path is correct
- Check file exists in S3 bucket
- Path should NOT start with `/` (use `folder/file.pptx`, not `/folder/file.pptx`)

### Error: "The security token included in the request is invalid"
- Access keys may be incorrect or expired
- Re-check AWS_ACCESS_KEY_ID and AWS_SECRET_ACCESS_KEY
- Regenerate keys in IAM if needed

## Comparison: OneDrive vs S3

| Feature | OneDrive (Original) | S3 (New) |
|---------|-------------------|----------|
| **Authentication** | Device code flow | Access keys |
| **Setup Complexity** | Medium | Easy |
| **Storage** | Microsoft OneDrive | AWS S3 |
| **Output** | Uploads to OneDrive | Saves locally |
| **Cost** | OneDrive subscription | AWS S3 pricing |
| **Speed** | Slower (auth flow) | Faster |
| **Automation** | Requires interaction | Fully automated |

## Next Steps

1. ✅ Install dependencies: `pip install -r requirements.txt`
2. ✅ Configure `.env` with your AWS credentials
3. ✅ Upload PPT to S3 bucket
4. ✅ Run: `python src/extract_ppt_from_s3.py`
5. ✅ Find your Excel in the `output/` folder

## Support

For AWS-specific issues:
- AWS S3 Documentation: https://docs.aws.amazon.com/s3/
- AWS IAM Documentation: https://docs.aws.amazon.com/iam/

For script issues:
- Check your Python version (3.7+)
- Verify all dependencies are installed
- Review error messages carefully
