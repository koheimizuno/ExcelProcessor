from typing import Dict, Any, Optional, List
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from src.schemas.models import CellRange

def apply_styles(source_cell: Optional[openpyxl.cell.Cell], 
                target_cell: openpyxl.cell.Cell, 
                styles: Optional[Dict[str, Any]] = None) -> None:
    """Apply styles to a target cell either from a source cell or styles dict."""
    if source_cell:
        target_cell.font = source_cell.font
        target_cell.border = source_cell.border
        target_cell.fill = source_cell.fill
        target_cell.number_format = source_cell.number_format
        target_cell.protection = source_cell.protection
        target_cell.alignment = source_cell.alignment
    
    if styles:
        if 'font' in styles:
            target_cell.font = openpyxl.styles.Font(**styles['font'])
        if 'border' in styles:
            target_cell.border = openpyxl.styles.Border(**styles['border'])
        if 'fill' in styles:
            target_cell.fill = openpyxl.styles.PatternFill(**styles['fill'])
        if 'alignment' in styles:
            target_cell.alignment = openpyxl.styles.Alignment(**styles['alignment'])

def get_cell_range(cell_range: CellRange) -> Dict[str, List[int]]:
    """Convert CellRange to lists of row and column indices."""
    start_col = column_index_from_string(cell_range.start_cell.col_letter) if cell_range.start_cell.col_letter else None
    start_row = cell_range.start_cell.row if cell_range.start_cell.row else None
    
    if cell_range.end_cell:
        end_col = column_index_from_string(cell_range.end_cell.col_letter) if cell_range.end_cell.col_letter else start_col
        end_row = cell_range.end_cell.row if cell_range.end_cell.row else start_row
    else:
        end_col = start_col
        end_row = start_row
    
    return {
        'rows': list(range(start_row, end_row + 1)) if start_row else [],
        'cols': list(range(start_col, end_col + 1)) if start_col else []
    }