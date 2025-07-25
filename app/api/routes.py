import os
import shutil
from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from app.services.translation_service import TranslationService
import aiofiles
from typing import Dict

router = APIRouter()
translation_service = TranslationService()

# Define supported languages
SUPPORTED_LANGUAGES: Dict[str, str] = {
    'modern_english': 'Modern English',
    'spanish': 'Spanish',
    'german': 'German',
    'dutch': 'Dutch',
}

@router.get("/languages")
async def get_languages():
    """Get list of supported languages"""
    return SUPPORTED_LANGUAGES

@router.post("/translate")
async def translate_document(
    file: UploadFile = File(...),
    target_language: str = Form(...),
    client_id: str = Form(None),
):
    # Validate target language
    if target_language not in SUPPORTED_LANGUAGES:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported language. Please choose from: {', '.join(SUPPORTED_LANGUAGES.keys())}"
        )

    # Validate file type
    if not file.filename.endswith(('.doc', '.docx')):
        raise HTTPException(
            status_code=400,
            detail="Only .doc and .docx files are supported"
        )

    input_filename = f"uploads/{file.filename}"
    
    try:
        # Ensure uploads directory exists
        os.makedirs("uploads", exist_ok=True)

        # Save uploaded file
        try:
            async with aiofiles.open(input_filename, 'wb') as out_file:
                content = await file.read()
                await out_file.write(content)
        except Exception as e:
            raise HTTPException(
                status_code=500,
                detail=f"Error saving uploaded file: {str(e)}"
            )

        # Process the document
        try:
            output_path = await translation_service.process_document(
                input_filename,
                SUPPORTED_LANGUAGES[target_language],
                client_id
            )
            
            if not os.path.exists(output_path):
                raise HTTPException(
                    status_code=500,
                    detail="Failed to generate translated document"
                )
            
            filename = os.path.basename(output_path)
            return {
                "status": "success",
                "fileName": filename,
                "downloadUrl": f"http://localhost:8000/api/download/{filename}"
            }
            
        except Exception as e:
            error_message = str(e)
            if "Token limit exceeded" in error_message:
                raise HTTPException(
                    status_code=402,
                    detail="Translation token limit exceeded. Please try with a smaller document or upgrade your account."
                )
            print(f"Translation error: {error_message}")
            raise HTTPException(
                status_code=500,
                detail=f"Error processing document: {error_message}"
            )

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error processing file: {str(e)}"
        )

    finally:
        # Clean up uploaded files
        if os.path.exists(input_filename):
            os.remove(input_filename)

@router.get("/download/{filename}")
async def download_file(filename: str):
    """Download a translated file"""
    file_path = os.path.join("uploads", filename)
    
    if not os.path.exists(file_path):
        raise HTTPException(
            status_code=404,
            detail="File not found"
        )
    
    try:
        headers = {
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers=headers
        )
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error downloading file: {str(e)}"
        )
