"""
GDPR-Compliant DOCX Anonymization API

This API accepts a DOCX file and returns a version with author-identifiable
information anonymized to comply with GDPR requirements.
"""

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
import tempfile
import os
from pathlib import Path
from anonymizer import DocxAnonymizer
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DOCX Anonymization API",
    description="Anonymize author information in DOCX files for GDPR compliance",
    version="1.0.0"
)


@app.post("/anonymise-docx")
async def anonymise_docx(file: UploadFile = File(...)):
    """
    Anonymize author information in a DOCX file.
    
    Args:
        file: The DOCX file to anonymize (multipart/form-data)
        
    Returns:
        The anonymized DOCX file ready for download
        
    Raises:
        HTTPException: If the file is invalid or processing fails
    """
    # Validate file extension
    if not file.filename.endswith('.docx'):
        raise HTTPException(
            status_code=400,
            detail="Invalid file format. Only .docx files are accepted."
        )
    
    # Create temporary files for processing
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    
    try:
        # Save uploaded file
        content = await file.read()
        temp_input.write(content)
        temp_input.flush()
        temp_input.close()
        
        logger.info(f"Processing file: {file.filename}")
        
        # Initialize anonymizer and process document
        anonymizer = DocxAnonymizer()
        success = anonymizer.anonymize_document(temp_input.name, temp_output.name)
        
        if not success:
            raise HTTPException(
                status_code=500,
                detail="Failed to process document. The file may be corrupted or invalid."
            )
        
        # Generate output filename
        original_name = Path(file.filename).stem
        output_filename = f"{original_name}_anonymized.docx"
        
        logger.info(f"Successfully anonymized: {file.filename}")
        
        # Return the anonymized file
        return FileResponse(
            path=temp_output.name,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=output_filename,
            background=None  # Keep file until response is sent
        )
        
    except HTTPException:
        # Re-raise HTTP exceptions
        raise
        
    except Exception as e:
        logger.error(f"Error processing document: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Internal server error: {str(e)}"
        )
        
    finally:
        # Cleanup: schedule temp files for deletion
        # Note: temp_output is cleaned up after FileResponse sends it
        try:
            if os.path.exists(temp_input.name):
                os.unlink(temp_input.name)
        except Exception as e:
            logger.warning(f"Failed to cleanup temp file: {e}")


@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "service": "docx-anonymizer"}


@app.get("/")
async def root():
    """Root endpoint with API information"""
    return {
        "service": "DOCX Anonymization API",
        "version": "1.0.0",
        "endpoints": {
            "anonymise": "POST /anonymise-docx",
            "health": "GET /health"
        }
    }
