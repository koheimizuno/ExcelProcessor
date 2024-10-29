from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
import base64
import io
from typing import Dict, Any

from src.schemas.models import ExcelRequest, ExcelResponse
from src.excel.processor import ExcelProcessor

app = FastAPI(title="Excel Processing API")

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
            detail={
                "output": f"Bad Request: {str(e)}",
                "status": "Error",
                "error_code": 400,
                "status_code": 400,
            }
        )