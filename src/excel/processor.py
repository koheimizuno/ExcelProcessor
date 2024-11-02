from typing import List, Dict, Any, Optional
import io
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from src.schemas.models import Processing
from src.excel.operations import xlsx_operation

class ExcelProcessor:
    def __init__(self, excel_file: io.BytesIO):
        """Initialize ExcelProcessor with an Excel file."""
        self.workbook = openpyxl.load_workbook(excel_file)
        self.operations = xlsx_operation(self.workbook)

    def process_operations(self, sheet_name: str, processing: List[Processing]) -> None:
        """Process a list of operations for a specific sheet."""
        for process in processing:
            operation_method = self._get_operation_method(process.processing_type)
            operation_method(sheet_name, process)

    def _get_operation_method(self, processing_type: str):
        """Get the corresponding operation method based on processing type."""
        operation_map = {
            "copy": self.operations.copy_cells,
            "copy_sheet": self.operations.copy_sheet,
            "copy_style": self.operations.copy_style,
            "insert_sheet": self.operations.insert_sheet,
            "delete_sheet": self.operations.delete_sheet,
            "insert": self.operations.insert_rows_or_cols,
            "delete": self.operations.delete_rows_or_cols,
            "hidden": self.operations.hide_rows_or_cols,
            "set_cells": self.operations.set_cells,
            "join_cells": self.operations.join_cells
        }
        return operation_map.get(processing_type)

    def save(self) -> io.BytesIO:
        """Save the workbook to a buffer and return it."""
        output = io.BytesIO()
        self.workbook.save(output)
        output.seek(0)
        return output