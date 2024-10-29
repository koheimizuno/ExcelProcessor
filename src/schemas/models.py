from pydantic import BaseModel
from typing import List, Optional, Union, Dict, Any

class Cell(BaseModel):
    col_letter: Optional[str] = None
    row: Optional[int] = None

class CellRange(BaseModel):
    start_cell: Cell
    end_cell: Optional[Cell] = None

class PasteTarget(BaseModel):
    sheet_name: str
    cells: Dict[str, Cell]
    is_insert: bool = False

class ProcessingTarget(BaseModel):
    cells: CellRange
    values: Optional[List[List[Any]]] = None
    styles: Optional[Dict[str, Any]] = None

class Processing(BaseModel):
    processing_type: str
    target: Optional[ProcessingTarget] = None
    paste_target: Optional[PasteTarget] = None

class Operation(BaseModel):
    sheet_name: str
    processing: List[Processing]

class ExcelRequest(BaseModel):
    file: str  # Base64 encoded Excel file
    mimetype: str
    operations: List[Operation]

class ExcelResponse(BaseModel):
    output: str  # Base64 encoded Excel file
    mimetype: str
    status: str
    error_code: int
    status_code: int