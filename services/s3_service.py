import os
import boto3
from io import BytesIO
from datetime import datetime
from typing import BinaryIO, Optional
import hashlib
from dotenv import load_dotenv

load_dotenv()

class S3Service:
    """Service for handling S3 operations"""
    
    def __init__(self):
        self.aws_access_key = os.getenv("AWS_ACCESS_KEY_ID")
        self.aws_secret_key = os.getenv("AWS_SECRET_ACCESS_KEY")
        self.aws_region = os.getenv("AWS_REGION", "us-east-1")
        self.bucket_name = os.getenv("S3_BUCKET_NAME")
        self.cloudfront_domain = os.getenv("CLOUDFRONT_DOMAIN", None)  # e.g., d2zu6flr7wd65l.cloudfront.net
        
        if not all([self.aws_access_key, self.aws_secret_key, self.bucket_name]):
            raise ValueError("Missing AWS credentials or bucket name in environment variables")
        
        self.s3_client = boto3.client(
            's3',
            aws_access_key_id=self.aws_access_key,
            aws_secret_access_key=self.aws_secret_key,
            region_name=self.aws_region
        )
        
        if self.cloudfront_domain:
            print(f"✅ CloudFront domain configured: {self.cloudfront_domain}")
    
    def upload_file(
        self,
        file_data: bytes,
        filename: str,
        folder: str = "presentations",
        content_type: str = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ) -> dict:
        """
        Upload file to S3
        
        Args:
            file_data: File content as bytes
            filename: Original filename
            folder: S3 folder path
            content_type: MIME type
            
        Returns:
            dict with s3_key, s3_url, bucket_name
        """
        try:
            # Generate unique filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_hash = hashlib.md5(file_data).hexdigest()[:8]
            base_name = os.path.splitext(filename)[0]
            extension = os.path.splitext(filename)[1]
            
            unique_filename = f"{base_name}_{timestamp}_{file_hash}{extension}"
            s3_key = f"{folder}/{unique_filename}"
            
            # Upload to S3
            self.s3_client.put_object(
                Bucket=self.bucket_name,
                Key=s3_key,
                Body=file_data,
                ContentType=content_type
            )
            
            # Generate S3 URL
            s3_url = f"https://{self.bucket_name}.s3.{self.aws_region}.amazonaws.com/{s3_key}"
            
            return {
                "success": True,
                "s3_key": s3_key,
                "s3_url": s3_url,
                "bucket_name": self.bucket_name,
                "filename": unique_filename
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def download_file(self, s3_key: str, bucket_name: Optional[str] = None) -> Optional[bytes]:
        """
        Download file from S3 using a three-tier strategy for maximum reliability:
        1. CloudFront (fastest, no ACL issues)
        2. Direct S3 HTTP (public access fallback)
        3. Boto3 API (for private buckets with proper credentials)

        Args:
            s3_key: S3 object key (should be unquoted/decoded)
            bucket_name: Optional bucket name override (defaults to configured bucket)

        Returns:
            file content as bytes on success, or None on failure
        """
        print(f"📥 Download Strategy: CloudFront → S3 HTTP → Boto3 API")
        print(f"   S3 Key: {s3_key}")
        
        # Try CloudFront first (if configured)
        if self.cloudfront_domain:
            result = self._download_via_cloudfront(s3_key)
            if result:
                return result
        
        # Try direct S3 HTTP (public access)
        bucket = bucket_name if bucket_name else self.bucket_name
        result = self._download_file_via_http(s3_key, bucket)
        if result:
            return result
        
        # Fall back to Boto3 API (requires valid credentials)
        return self._download_file_via_boto3(s3_key, bucket)
    
    def _download_via_cloudfront(self, s3_key: str) -> Optional[bytes]:
        """
        Download file via CloudFront CDN.
        Works even with S3 ACL/ownership issues - recommended approach.
        
        Args:
            s3_key: S3 object key
        
        Returns:
            file content as bytes on success, or None on failure
        """
        try:
            import requests
            
            # URL encode the key to handle special characters
            from urllib.parse import quote
            encoded_key = quote(s3_key, safe='/')
            cloudfront_url = f"https://{self.cloudfront_domain}/{encoded_key}"
            
            print(f"🌐 Tier 1: Trying CloudFront...")
            print(f"   URL: {cloudfront_url}")
            
            response = requests.get(cloudfront_url, timeout=30)
            
            if response.status_code == 200:
                file_size = len(response.content)
                print(f"✅ CloudFront download successful ({file_size} bytes)")
                return response.content
            else:
                print(f"⚠️ CloudFront returned {response.status_code}, trying fallback...")
                return None
                
        except Exception as e:
            print(f"⚠️ CloudFront failed ({type(e).__name__}), trying fallback...")
            return None
    
    def _download_file_via_http(self, s3_key: str, bucket_name: str) -> Optional[bytes]:
        """
        Download file from S3 via HTTP public URL.
        Works when bucket has public read policy.
        
        Args:
            s3_key: S3 object key
            bucket_name: Bucket name
        
        Returns:
            file content as bytes on success, or None on failure
        """
        try:
            import requests
            
            # Construct public S3 URL
            # Format: https://bucket.s3.region.amazonaws.com/key
            s3_url = f"https://{bucket_name}.s3.{self.aws_region}.amazonaws.com/{s3_key}"
            
            print(f"🌐 Tier 2: Trying Direct S3 HTTP...")
            print(f"   URL: {s3_url}")
            
            response = requests.get(s3_url, timeout=30)
            
            if response.status_code == 200:
                file_size = len(response.content)
                print(f"✅ S3 HTTP download successful ({file_size} bytes)")
                return response.content
            elif response.status_code == 403:
                print(f"⚠️ S3 HTTP returned 403 (Access Denied), trying fallback...")
                return None
            elif response.status_code == 404:
                print(f"⚠️ S3 HTTP returned 404 (Not Found), trying fallback...")
                return None
            else:
                print(f"⚠️ S3 HTTP returned {response.status_code}, trying fallback...")
                return None
                
        except Exception as e:
            print(f"⚠️ S3 HTTP failed ({type(e).__name__}), trying fallback...")
            return None
    
    def _download_file_via_boto3(self, s3_key: str, bucket_name: str) -> Optional[bytes]:
        """
        Download file using Boto3 API.
        Requires valid AWS credentials with S3 permissions.
        
        Args:
            s3_key: S3 object key
            bucket_name: Bucket name
        
        Returns:
            file content as bytes on success, or None on failure
        """
        try:
            print(f"🔑 Tier 3: Trying Boto3 API...")
            print(f"   Bucket: {bucket_name}, Key: {s3_key}")
            
            buffer = BytesIO()
            self.s3_client.download_fileobj(bucket_name, s3_key, buffer)
            buffer.seek(0)
            data = buffer.getvalue()
            file_size = len(data)
            print(f"✅ Boto3 API download successful ({file_size} bytes)")
            return data
            
        except self.s3_client.exceptions.NoSuchKey:
            print(f"❌ Boto3 API: File not found")
            return None
        except Exception as e:
            error_msg = f"{type(e).__name__}: {str(e)}"
            print(f"❌ Boto3 API failed: {error_msg}")
            return None
    
    def get_file_url(self, s3_key: str, expiration: int = 3600) -> Optional[str]:
        """
        Generate presigned URL for S3 object
        
        Args:
            s3_key: S3 object key
            expiration: URL expiration time in seconds (default: 1 hour)
            
        Returns:
            Presigned URL or None if failed
        """
        try:
            url = self.s3_client.generate_presigned_url(
                'get_object',
                Params={
                    'Bucket': self.bucket_name,
                    'Key': s3_key
                },
                ExpiresIn=expiration
            )
            return url
        except Exception as e:
            print(f"Error generating presigned URL: {e}")
            return None
    
    def upload_excel_to_s3(
        self,
        excel_data: bytes,
        filename: str,
        folder: str = "excel_outputs"
    ) -> dict:
        """
        Upload Excel file to S3
        
        Args:
            excel_data: Excel file content as bytes
            filename: Excel filename
            folder: S3 folder path
            
        Returns:
            dict with upload details
        """
        return self.upload_file(
            file_data=excel_data,
            filename=filename,
            folder=folder,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    def upload_image_to_s3(
        self,
        image_data: bytes,
        filename: str,
        content_type: str = "image/jpeg",
        folder: str = "extracted_images"
    ) -> dict:
        """
        Upload image to S3
        
        Args:
            image_data: Image content as bytes
            filename: Image filename
            content_type: Image MIME type
            folder: S3 folder path
            
        Returns:
            dict with upload details
        """
        return self.upload_file(
            file_data=image_data,
            filename=filename,
            folder=folder,
            content_type=content_type
        )
    
    def list_files(self, prefix: str = "") -> list:
        """
        List files in S3 bucket
        
        Args:
            prefix: S3 prefix/folder to filter
            
        Returns:
            List of file keys
        """
        try:
            response = self.s3_client.list_objects_v2(
                Bucket=self.bucket_name,
                Prefix=prefix
            )
            
            if 'Contents' in response:
                return [obj['Key'] for obj in response['Contents']]
            return []
            
        except Exception as e:
            print(f"Error listing S3 files: {e}")
            return []
    
    def delete_file(self, s3_key: str) -> bool:
        """
        Delete file from S3
        
        Args:
            s3_key: S3 object key
            
        Returns:
            True if successful, False otherwise
        """
        try:
            self.s3_client.delete_object(
                Bucket=self.bucket_name,
                Key=s3_key
            )
            return True
        except Exception as e:
            print(f"Error deleting S3 file: {e}")
            return False

# Create singleton instance
s3_service = S3Service()