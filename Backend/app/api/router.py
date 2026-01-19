import io
from fastapi import APIRouter, UploadFile, File
from fastapi.responses import StreamingResponse
from services.WordETL import ConverterService
from services.TimeSummary import get_summary

router = APIRouter()


@router.post("/convert")
async def convert_word_to_excel(file: UploadFile = File(...)):
    """
    Convert Word document tables to Excel file
    
    Args:
        file: Uploaded Word document (.docx)
        
    Returns:
        StreamingResponse: Excel file download
    """
    # Read the uploaded file
    contents = await file.read()
    
    # Convert using service
    output = await ConverterService.convert_word_to_excel(contents, file.filename)
    
    # Return as downloadable file
    return StreamingResponse(
        io.BytesIO(output.read()),
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={
            'Content-Disposition': 'attachment; filename=converted_tables.xlsx'
        }
    )


@router.post("/summary")
async def excel_summary(file: UploadFile = File(...)):
    """
    Convert Word document tables to Excel file
    
    Args:
        file: Uploaded Word document (.docx)
        
    Returns:
        StreamingResponse: Excel file download
    """
    # Read the uploaded file
    contents = await file.read()
    
    # Convert using service
    output_file = r"志工時數結算_已處理.xlsx"
    output = await get_summary(contents)
    
    # Return as downloadable file
    return StreamingResponse(
        io.BytesIO(output.read()),
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={
            'Content-Disposition': f'attachment; filename={output_file}'
        }
    )