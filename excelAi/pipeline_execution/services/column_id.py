from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from .errors import PipelineValidationError


@dataclass(frozen=True)
class ColumnId:
    sheet_name: str
    column_order: int  # 0-based (API), converted to 1-based for openpyxl
    column_name: str


def parse_column_id(column_id: Any) -> ColumnId:
    """
    Parse "<sheetName>:<columnOrder>:<columnName>".

    Strict:
    - exactly 3 colon-separated parts
    - columnOrder must be an int >= 0 (0-based column order as provided by the client)
    """
    if not isinstance(column_id, str):
        raise PipelineValidationError("columnId must be a string")
    raw = column_id.strip()
    parts = raw.split(":")
    if len(parts) != 3:
        raise PipelineValidationError(
            "columnId must be in format '<sheetName>:<columnOrder>:<columnName>'"
        )

    sheet_name, col_order_raw, col_name = parts[0], parts[1], parts[2]
    if not sheet_name:
        raise PipelineValidationError("columnId sheetName must be non-empty")
    if not col_name:
        raise PipelineValidationError("columnId columnName must be non-empty")

    try:
        col_order = int(col_order_raw)
    except (TypeError, ValueError):
        raise PipelineValidationError("columnId columnOrder must be an integer")

    if col_order < 0:
        raise PipelineValidationError("columnId columnOrder must be >= 0")

    return ColumnId(sheet_name=sheet_name, column_order=col_order, column_name=col_name)


def resolve_column_id(
    ws: Worksheet,
    *,
    selected_sheet: str,
    header_row_idx: int,
    column_id: str,
) -> int:
    """
    Resolve a columnId to a 1-based column index on the current worksheet state.

    Strict:
    - sheetName in columnId must match selected_sheet exactly
    - header row cell at (columnOrder + 1) must equal columnName exactly
    """
    parsed = parse_column_id(column_id)

    if parsed.sheet_name != selected_sheet:
        raise PipelineValidationError(
            f"columnId sheetName '{parsed.sheet_name}' does not match selected sheet '{selected_sheet}'"
        )

    excel_col = parsed.column_order + 1
    cell = ws.cell(row=header_row_idx, column=excel_col)
    if cell.value != parsed.column_name:
        raise PipelineValidationError(
            "columnId header mismatch. "
            f"Expected header[{parsed.column_order}] == '{parsed.column_name}', "
            f"got '{cell.value}'"
        )

    return excel_col

