from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from ..column_id import resolve_column_id
from ..errors import PipelineValidationError
from .common import (
    _collect_column_values,
    _determine_column_type,
    _filter_non_empty_values,
    _parse_date_cell_value,
    _require_non_empty_str,
)


_ALLOWED_DATE_OUTPUT_FORMATS: set[str] = {
    "YYYY/MM/DD",
    "DD/MM/YYYY",
    "YYYY.MM.DD",
    "DD.MM.YYYY",
    "YYYY-MM-DD",
    "DD-MM-YYYY",
}


_DATE_OUTPUT_FORMAT_TO_EXCEL_NUMBER_FORMAT: dict[str, str] = {
    "YYYY/MM/DD": "yyyy/mm/dd",
    "DD/MM/YYYY": "dd/mm/yyyy",
    "YYYY.MM.DD": "yyyy.mm.dd",
    "DD.MM.YYYY": "dd.mm.yyyy",
    "YYYY-MM-DD": "yyyy-mm-dd",
    "DD-MM-YYYY": "dd-mm-yyyy",
}


def _require_date_dtype_and_resolve(
    ws: Worksheet,
    *,
    header_row_idx: int,
    selected_sheet: str,
    column_id: str,
) -> int:
    col_idx = resolve_column_id(
        ws,
        selected_sheet=selected_sheet,
        header_row_idx=header_row_idx,
        column_id=column_id,
    )

    values = _collect_column_values(ws, header_row_idx=header_row_idx, col_idx=col_idx)
    non_empty = _filter_non_empty_values(values)
    if not non_empty:
        raise PipelineValidationError(
            f"ColumnId={column_id} must contain at least one non-empty value to parse as date"
        )

    inferred = _determine_column_type(values)
    if inferred != "date":
        raise PipelineValidationError(f"ColumnId={column_id} must be date. Inferred type={inferred}")

    return col_idx


def apply_parse_date(
    ws: Worksheet,
    *,
    header_row_idx: int,
    selected_sheet: str,
    column_ids: Any,
    output_format: Any,
) -> None:
    """
    Parse date-like columns and apply a consistent Excel date number format.

    Params:
    - column_ids: list of columnId strings
    - output_format: one of the supported 'YYYY/MM/DD' style strings
    """
    if not isinstance(column_ids, list):
        raise PipelineValidationError("columnIds must be a list")
    if not column_ids:
        raise PipelineValidationError("columnIds must be a non-empty list")
    if any(not isinstance(cid, str) or not cid.strip() for cid in column_ids):
        raise PipelineValidationError("columnIds must contain non-empty strings only")
    if len(set(column_ids)) != len(column_ids):
        raise PipelineValidationError("columnIds must not contain duplicates")

    output_format_str = _require_non_empty_str(output_format, field_name="outputFormat")
    if output_format_str not in _ALLOWED_DATE_OUTPUT_FORMATS:
        raise PipelineValidationError(
            f"outputFormat must be one of {sorted(_ALLOWED_DATE_OUTPUT_FORMATS)}"
        )
    excel_number_format = _DATE_OUTPUT_FORMAT_TO_EXCEL_NUMBER_FORMAT[output_format_str]

    for column_id in column_ids:
        column_id = column_id.strip()
        col_idx = _require_date_dtype_and_resolve(
            ws,
            header_row_idx=header_row_idx,
            selected_sheet=selected_sheet,
            column_id=column_id,
        )

        for row_idx in range(header_row_idx + 1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = cell.value
            if val is None or val == "":
                continue

            try:
                parsed = _parse_date_cell_value(val)
            except Exception as exc:
                raise PipelineValidationError(
                    f"Failed to parse date in ColumnId={column_id} at row={row_idx}: "
                    f"value={val!r} (type={type(val).__name__}). {exc}"
                )

            cell.value = parsed
            cell.number_format = excel_number_format

