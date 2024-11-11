from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
import base64
import io
from typing import Dict, Any

from src.schemas.models import ExcelRequest, ExcelResponse
from src.excel.processor import ExcelProcessor

app = FastAPI(title="Excel Processing API")

@app.exception_handler(HTTPException)
async def exception_handler(request: Request, exc: HTTPException):

    return JSONResponse(
        status_code=exc.status_code,
        content={
            "output": f"Bad Request: {exc.detail}",
            "status": "Error",
            "error_code": exc.status_code,
            "status_code": exc.status_code,
        }
    )

@app.post("/transform_excel", response_model=ExcelResponse)
async def transform_excel(request: ExcelRequest) -> Dict[str, Any]:
    try:
        # Decode base64 Excel file
        excel_data = base64.b64decode(request.file)
        excel_buffer = io.BytesIO(excel_data)
        
        # Create Excel processor instance
        processor = ExcelProcessor(excel_buffer)
        
        # Process each operation
        for operation in request.operations:
            processor.process_operations(operation.sheet_name, operation.processing)
        
        # Get processed Excel file
        output_buffer = processor.save()
        output_base64 = base64.b64encode(output_buffer.read()).decode()
        
        return {
            "output": output_base64,
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "status": "Success",
            "error_code": 200,
            "status_code": 200,
        }
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail={str(e)}
        )