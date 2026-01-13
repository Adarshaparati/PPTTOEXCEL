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
        
        if not all([self.aws_access_key, self.aws_secret_key, self.bucket_name]):
            raise ValueError("Missing AWS credentials or bucket name in environment variables")
        
        self.s3_client = boto3.client(
            's3',
            aws_access_key_id=self.aws_access_key,
            aws_secret_access_key=self.aws_secret_key,
            region_name=self.aws_region
        )
    
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
    
    def download_file(self, s3_key: str) -> Optional[bytes]:
        """
        Download file from S3 and return raw bytes.

        Args:
            s3_key: S3 object key (should be unquoted/decoded)

        Returns:
            file content as bytes on success, or None on failure
        """
        try:
            print(f"📥 S3 Download - Bucket: {self.bucket_name}, Key: {s3_key}")
            buffer = BytesIO()
            self.s3_client.download_fileobj(self.bucket_name, s3_key, buffer)
            buffer.seek(0)
            data = buffer.getvalue()
            file_size = len(data)
            print(f"✅ Downloaded {file_size} bytes from S3")
            return data
        except self.s3_client.exceptions.NoSuchKey:
            error_msg = f"File not found in S3. Key: {s3_key}"
            print(f"❌ {error_msg}")
            return None
        except Exception as e:
            error_msg = f"{type(e).__name__}: {str(e)}"
            print(f"❌ Error downloading from S3: {error_msg}")
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