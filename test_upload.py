#!/usr/bin/env python3
"""
Quick test script to upload a PowerPoint file
"""
import requests
import sys
import os

# API endpoint
API_URL = "http://localhost:8000/api/upload-ppt"

# Check if file path is provided
if len(sys.argv) < 2:
    print("Usage: python3 test_upload.py <path_to_pptx_file>")
    print("\nExample:")
    print("  python3 test_upload.py presentation.pptx")
    sys.exit(1)

file_path = sys.argv[1]

# Check if file exists
if not os.path.exists(file_path):
    print(f"❌ Error: File not found: {file_path}")
    sys.exit(1)

# Check file extension
if not file_path.lower().endswith(('.ppt', '.pptx')):
    print(f"❌ Error: File must be .ppt or .pptx")
    sys.exit(1)

print(f"📤 Uploading {file_path}...")
print(f"📍 To: {API_URL}\n")

try:
    # Open and upload file
    with open(file_path, 'rb') as f:
        files = {'file': f}
        response = requests.post(API_URL, files=files)
    
    print(f"Status Code: {response.status_code}")
    print(f"Response:\n")
    
    if response.status_code == 200:
        data = response.json()
        print("✅ SUCCESS!")
        print(f"📁 Original Filename: {data['original_filename']}")
        print(f"📦 S3 Filename: {data['filename']}")
        print(f"🔗 S3 URL: {data['s3_url']}")
        print(f"🔑 S3 Key: {data['s3_key']}")
        print(f"📊 File Size: {data['file_size']} bytes")
        print(f"⏰ Timestamp: {data['timestamp']}")
    else:
        print(f"❌ ERROR!")
        print(response.json())
        
except Exception as e:
    print(f"❌ Exception: {str(e)}")
