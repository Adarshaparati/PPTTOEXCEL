from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from datetime import datetime
from io import BytesIO
import os
from services.s3_service import s3_service
from services.extraction_service import extraction_service
from services.sheets_service import get_sheets_service
from services.ppt_generation_service import ppt_generation_service

router = APIRouter()

@router.get("/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "service": "ppt-extractor",
        "timestamp": datetime.now().isoformat()
    }

@router.post("/upload-ppt")
async def upload_ppt(file: UploadFile = File(...)):
    """
    Upload PowerPoint file to S3
    
    Returns:
        - s3_url: URL to access the uploaded file
        - s3_key: S3 object key
        - filename: Uploaded filename
    """
    try:
        # Validate file type
        if not file.filename.endswith(('.ppt', '.pptx')):
            raise HTTPException(
                status_code=400,
                detail="Invalid file type. Only .ppt and .pptx files are allowed."
            )
        
        # Read file content
        file_content = await file.read()
        
        if len(file_content) == 0:
            raise HTTPException(
                status_code=400,
                detail="Uploaded file is empty."
            )
        
        # Upload to S3
        upload_result = s3_service.upload_file(
            file_data=file_content,
            filename=file.filename,
            folder="presentations"
        )
        
        if not upload_result.get('success'):
            raise HTTPException(
                status_code=500,
                detail=f"Failed to upload file to S3: {upload_result.get('error')}"
            )
        
        return {
            "success": True,
            "message": "File uploaded successfully",
            "s3_url": upload_result['s3_url'],
            "s3_key": upload_result['s3_key'],
            "filename": upload_result['filename'],
            "original_filename": file.filename,
            "file_size": len(file_content),
            "timestamp": datetime.now().isoformat()
        }
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error uploading file: {str(e)}"
        )

@router.post("/extract-ppt")
async def extract_ppt(
    s3_key: str, 
    template_id: str = "default",
    upload_images_to_s3: str = "true",
    save_to_sheets: str = "true"
):
    """
    Extract data from PowerPoint file in S3
    
    Args:
        - s3_key: S3 object key OR full S3 URL of the PowerPoint file
        - template_id: Template identifier for this presentation (default: "default")
        - upload_images_to_s3: Whether to upload extracted images to S3 ("true"/"false", default: "true")
        - save_to_sheets: Whether to save data to Google Sheets ("yes"/"no"/"true"/"false", default: "true")
    
    Returns:
        - extracted_data: JSON with slide information
        - excel_url: URL to download the Excel file
        - local_excel_path: Local path to the Excel file
        - sheets_result: Result of Google Sheets operation
    """
    try:
        # Convert string parameters to boolean
        upload_images = upload_images_to_s3.lower() in ["true", "yes", "1"]
        save_sheets = save_to_sheets.lower() in ["true", "yes", "1"]
        # Convert S3 URL to S3 key if full URL is provided
        if s3_key.startswith('http://') or s3_key.startswith('https://'):
            # Extract key from URL
            # Format: https://bucket.s3.region.amazonaws.com/key or https://s3.region.amazonaws.com/bucket/key
            from urllib.parse import urlparse, unquote
            parsed_url = urlparse(s3_key)
            # Remove leading slash and decode URL encoding
            s3_key = unquote(parsed_url.path.lstrip('/'))
            print(f"Extracted S3 key from URL: {s3_key}")
        
        # Download PPT from S3
        ppt_bytes = s3_service.download_file(s3_key)
        
        if not ppt_bytes:
            raise HTTPException(
                status_code=404,
                detail=f"File not found in S3: {s3_key}"
            )
        
        # Extract data
        wb, extracted_data = extraction_service.extract_ppt_to_excel(
            ppt_bytes,
            upload_images_to_s3=upload_images
        )
        
        # Save Excel locally
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_base = os.path.basename(s3_key).replace('.pptx', '').replace('.ppt', '')
        excel_filename = f"PPT_Analysis_{filename_base}_{timestamp}.xlsx"
        
        output_folder = "output"
        os.makedirs(output_folder, exist_ok=True)
        local_excel_path = os.path.join(output_folder, excel_filename)
        
        wb.save(local_excel_path)
        
        # Optionally upload Excel to S3
        excel_s3_url = None
        if upload_images:  # Use same flag for consistency
            with open(local_excel_path, 'rb') as f:
                excel_bytes = f.read()
            excel_upload_result = s3_service.upload_excel_to_s3(
                excel_data=excel_bytes,
                filename=excel_filename
            )
            if excel_upload_result.get('success'):
                excel_s3_url = excel_upload_result.get('s3_url')
        
        # Save to Google Sheets
        sheets_result = None
        if save_sheets:
            sheets_svc = get_sheets_service()
            if sheets_svc:
                # Generate S3 URL for the source PPT
                ppt_s3_url = f"https://{s3_service.bucket_name}.s3.{s3_service.aws_region}.amazonaws.com/{s3_key}"
                
                sheets_result = sheets_svc.append_ppt_data(
                    template_id=template_id,
                    ppt_filename=filename_base,
                    s3_url=ppt_s3_url,
                    slide_data=extracted_data.get('slides', []),
                    excel_url=excel_s3_url
                )
        
        return {
            "success": True,
            "message": "Data extracted successfully",
            "template_id": template_id,
            "extracted_data": extracted_data,
            "excel_filename": excel_filename,
            "local_excel_path": local_excel_path,
            "excel_s3_url": excel_s3_url,
            "download_url": f"/api/download/{excel_filename}",
            "sheets_saved": sheets_result.get('success') if sheets_result else False,
            "sheets_result": sheets_result,
            "timestamp": datetime.now().isoformat()
        }
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error extracting data: {str(e)}"
        )

@router.post("/upload-and-extract")
async def upload_and_extract(
    file: UploadFile = File(...),
    template_id: str = "default",
    upload_images_to_s3: bool = True,
    save_to_sheets: bool = True
):
    """
    Upload PowerPoint file and extract data in one step
    
    Args:
        - file: PowerPoint file to upload
        - template_id: Template identifier for this presentation (default: "default")
        - upload_images_to_s3: Whether to upload extracted images to S3 (default: True)
        - save_to_sheets: Whether to save data to Google Sheets (default: True)
    
    Returns:
        - s3_url: URL of uploaded PowerPoint
        - extracted_data: JSON with slide information
        - excel_url: URL to download the Excel file
        - sheets_result: Result of Google Sheets operation
    """
    try:
        # Validate file type
        if not file.filename.endswith(('.ppt', '.pptx')):
            raise HTTPException(
                status_code=400,
                detail="Invalid file type. Only .ppt and .pptx files are allowed."
            )
        
        # Read file content
        file_content = await file.read()
        
        if len(file_content) == 0:
            raise HTTPException(
                status_code=400,
                detail="Uploaded file is empty."
            )
        
        # Upload to S3
        upload_result = s3_service.upload_file(
            file_data=file_content,
            filename=file.filename,
            folder="presentations"
        )
        
        if not upload_result.get('success'):
            raise HTTPException(
                status_code=500,
                detail=f"Failed to upload file to S3: {upload_result.get('error')}"
            )
        
        # Extract data
        wb, extracted_data = extraction_service.extract_ppt_to_excel(
            file_content,
            upload_images_to_s3=upload_images_to_s3
        )
        
        # Save Excel locally
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_base = file.filename.replace('.pptx', '').replace('.ppt', '')
        excel_filename = f"PPT_Analysis_{filename_base}_{timestamp}.xlsx"
        
        output_folder = "output"
        os.makedirs(output_folder, exist_ok=True)
        local_excel_path = os.path.join(output_folder, excel_filename)
        
        wb.save(local_excel_path)
        
        # Optionally upload Excel to S3
        excel_s3_url = None
        if upload_images_to_s3:
            with open(local_excel_path, 'rb') as f:
                excel_bytes = f.read()
            excel_upload_result = s3_service.upload_excel_to_s3(
                excel_data=excel_bytes,
                filename=excel_filename
            )
            if excel_upload_result.get('success'):
                excel_s3_url = excel_upload_result.get('s3_url')
        
        # Save to Google Sheets
        sheets_result = None
        if save_to_sheets:
            sheets_svc = get_sheets_service()
            if sheets_svc:
                sheets_result = sheets_svc.append_ppt_data(
                    template_id=template_id,
                    ppt_filename=file.filename,
                    s3_url=upload_result['s3_url'],
                    slide_data=extracted_data.get('slides', []),
                    excel_url=excel_s3_url
                )
        
        return {
            "success": True,
            "message": "File uploaded and data extracted successfully",
            "template_id": template_id,
            "ppt_s3_url": upload_result['s3_url'],
            "ppt_s3_key": upload_result['s3_key'],
            "original_filename": file.filename,
            "file_size": len(file_content),
            "extracted_data": extracted_data,
            "excel_filename": excel_filename,
            "local_excel_path": local_excel_path,
            "excel_s3_url": excel_s3_url,
            "download_url": f"/api/download/{excel_filename}",
            "sheets_saved": sheets_result.get('success') if sheets_result else False,
            "sheets_result": sheets_result,
            "timestamp": datetime.now().isoformat()
        }
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing file: {str(e)}"
        )

@router.get("/download/{filename}")
async def download_excel(filename: str):
    """
    Download Excel file by filename
    
    Args:
        - filename: Name of the Excel file to download
    
    Returns:
        - Excel file download
    """
    try:
        file_path = os.path.join("output", filename)
        
        if not os.path.exists(file_path):
            raise HTTPException(
                status_code=404,
                detail=f"File not found: {filename}"
            )
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error downloading file: {str(e)}"
        )

@router.get("/list-files")
async def list_files(prefix: str = "presentations"):
    """
    List files in S3 bucket
    
    Args:
        - prefix: S3 prefix/folder to filter (default: "presentations")
    
    Returns:
        - List of files in S3
    """
    try:
        files = s3_service.list_files(prefix=prefix)
        
        return {
            "success": True,
            "prefix": prefix,
            "file_count": len(files),
            "files": files
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error listing files: {str(e)}"
        )

@router.delete("/delete-file")
async def delete_file(s3_key: str):
    """
    Delete file from S3
    
    Args:
        - s3_key: S3 object key to delete
    
    Returns:
        - Success message
    """
    try:
        success = s3_service.delete_file(s3_key)
        
        if not success:
            raise HTTPException(
                status_code=500,
                detail="Failed to delete file from S3"
            )
        
        return {
            "success": True,
            "message": f"File deleted successfully: {s3_key}",
            "s3_key": s3_key,
            "timestamp": datetime.now().isoformat()
        }
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error deleting file: {str(e)}"
        )

@router.get("/presigned-url")
async def get_presigned_url(s3_key: str, expiration: int = 3600):
    """
    Generate presigned URL for S3 object
    
    Args:
        - s3_key: S3 object key
        - expiration: URL expiration time in seconds (default: 3600 = 1 hour)
    
    Returns:
        - Presigned URL
    """
    try:
        url = s3_service.get_file_url(s3_key, expiration)
        
        if not url:
            raise HTTPException(
                status_code=500,
                detail="Failed to generate presigned URL"
            )
        
        return {
            "success": True,
            "presigned_url": url,
            "s3_key": s3_key,
            "expires_in": expiration,
            "timestamp": datetime.now().isoformat()
        }
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating presigned URL: {str(e)}"
        )

@router.post("/generate-ppt")
async def generate_ppt_from_data(
    template_s3_url: str,
    template_id: str,
    upload_to_s3: bool = True,
    output_filename: str = None
):
    """
    Generate PowerPoint from template and Google Sheets data
    
    Args:
        - template_s3_url: Full S3 URL of template PowerPoint file
        - template_id: Template ID to fetch data from Google Sheets
        - upload_to_s3: Whether to upload generated PPT to S3 (default: True)
        - output_filename: Optional custom filename
    
    Returns:
        - Generated PowerPoint file or S3 URL
    
    Example:
        POST /api/generate-ppt?template_s3_url=https://...&template_id=AD12&upload_to_s3=true
    """
    try:
        # Generate presentation
        ppt_bytes, filename, stats = ppt_generation_service.generate_presentation(
            template_s3_url=template_s3_url,
            template_id=template_id,
            output_filename=output_filename
        )
        
        response_data = {
            "success": True,
            "message": "Presentation generated successfully",
            "stats": stats,
            "timestamp": datetime.now().isoformat()
        }
        
        # Upload to S3 if requested
        if upload_to_s3:
            upload_result = s3_service.upload_file(
                file_data=ppt_bytes,
                filename=filename,
                folder="generated_presentations"
            )
            
            if upload_result.get('success'):
                response_data['s3_url'] = upload_result['s3_url']
                response_data['s3_key'] = upload_result['s3_key']
                print(f"✅ Uploaded to S3: {upload_result['s3_url']}")
        
        # Save locally for download
        output_folder = "output"
        os.makedirs(output_folder, exist_ok=True)
        local_path = os.path.join(output_folder, filename)
        
        with open(local_path, 'wb') as f:
            f.write(ppt_bytes)
        
        response_data['local_path'] = local_path
        response_data['download_url'] = f"/api/download/{filename}"
        
        return response_data
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating presentation: {str(e)}"
        )

@router.get("/generate-ppt-download")
async def generate_and_download_ppt(
    template_s3_url: str,
    template_id: str,
    output_filename: str = None
):
    """
    Generate PowerPoint and return as downloadable file
    
    Args:
        - template_s3_url: Full S3 URL of template PowerPoint file
        - template_id: Template ID to fetch data from Google Sheets
        - output_filename: Optional custom filename
    
    Returns:
        - PowerPoint file for direct download
    """
    try:
        # Generate presentation
        ppt_bytes, filename, stats = ppt_generation_service.generate_presentation(
            template_s3_url=template_s3_url,
            template_id=template_id,
            output_filename=output_filename
        )
        
        # Return as streaming response
        return StreamingResponse(
            BytesIO(ppt_bytes),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error generating presentation: {str(e)}"
        )
