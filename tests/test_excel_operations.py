import pytest
import io
import openpyxl
from src.operations.excel_operations import xlsx_operation
from src.schemas.models import Processing, ProcessingTarget, Cell, CellRange

def test_set_cells_operation():
    # Create test workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    # Test data
    processing = [
        Processing(
            processing_type="set_cells",
            target=ProcessingTarget(
                cells=CellRange(
                    start_cell=Cell(col_letter="A", row=1),
                    end_cell=Cell(col_letter="B", row=2)
                ),
                values=[["Test1", "Test2"], ["Test3", "Test4"]]
            )
        )
    ]
    
    # Process operation
    result = xlsx_operation(buffer, "Sheet1", processing)
    
    # Verify results
    result_wb = openpyxl.load_workbook(result)
    result_ws = result_wb["Sheet1"]
    
    assert result_ws['A1'].value == "Test1"
    assert result_ws['B1'].value == "Test2"
    assert result_ws['A2'].value == "Test3"
    assert result_ws['B2'].value == "Test4"