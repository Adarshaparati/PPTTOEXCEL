import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv
from routes.ppt_routes import router as ppt_router
from routes.slide_routes import router as slide_router

# Load environment variables
load_dotenv()

# Create FastAPI app
app = FastAPI(
    title="PowerPoint to Excel Extractor API",
    description="Upload PowerPoint files to S3 and extract comprehensive data to Excel. Generate custom slides with different types.",
    version="2.0.0"
)

# CORS configuration - adjust origins for production
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Change to specific origins in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create necessary directories
os.makedirs("output", exist_ok=True)
os.makedirs("extracted_images", exist_ok=True)
os.makedirs("temp", exist_ok=True)

# Mount static files for serving Excel and images
app.mount("/files", StaticFiles(directory="output"), name="files")
app.mount("/images", StaticFiles(directory="extracted_images"), name="images")

# Include routers
app.include_router(ppt_router, prefix="/api", tags=["PowerPoint Processing"])
app.include_router(slide_router, prefix="/api/slides", tags=["Slide Generation"])

@app.get("/")
async def root():
    """Root endpoint with API information"""
    return {
        "message": "PowerPoint to Excel Extractor API",
        "version": "2.0.0",
        "endpoints": {
            "upload": "POST /api/upload-ppt",
            "extract": "POST /api/extract-ppt",
            "upload_and_extract": "POST /api/upload-and-extract",
            "download": "GET /api/download/{filename}",
            "health": "GET /api/health",
            "generate_points_slide": "POST /api/slides/generate-points-slide",
            "generate_image_text_slide": "POST /api/slides/generate-image-text-slide",
            "generate_table_slide": "POST /api/slides/generate-table-slide",
            "generate_phases_slide": "POST /api/slides/generate-phases-slide",
            "generate_multi_slide": "POST /api/slides/generate-multi-slide",
            "slide_types": "GET /api/slides/slide-types"
        },
        "docs": "/docs",
        "redoc": "/redoc"
    }

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "service": "ppt-extractor-api"
    }

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=True  # Disable in production
    )
