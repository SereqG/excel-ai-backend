from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from .column_id import resolve_column_id
from .errors import PipelineValidationError


def _header_values(ws: Worksheet, *, header_row_idx: int) -> list[Any]:
    # openpyxl returns tuples of cells via iter_rows; we want current values.
    return [cell.value for cell in ws[header_row_idx]]


def apply_rename_column(
    ws: Worksheet,
    *,
    header_row_idx: int,
    selected_sheet: str,
    column_id: str,
    new_name: Any,
) -> None:
    """
    Rename a column header.

    - new_name must be a non-empty string
    - fail on name collision (new_name already exists in header row, excluding target column)
    - column_id must match selected_sheet and header (strict)
    """
    if not isinstance(new_name, str) or not new_name.strip():
        raise PipelineValidationError("newName must be a non-empty string")
    new_name = new_name.strip()

    col_idx = resolve_column_id(
        ws,
        selected_sheet=selected_sheet,
        header_row_idx=header_row_idx,
        column_id=column_id,
    )

    header_vals = _header_values(ws, header_row_idx=header_row_idx)
    for i, val in enumerate(header_vals, start=1):
        if i == col_idx:
            continue
        if val == new_name:
            raise PipelineValidationError(f"Rename collision: header already contains '{new_name}'")

    ws.cell(row=header_row_idx, column=col_idx).value = new_name


def apply_drop_column(
    ws: Worksheet,
    *,
    header_row_idx: int,
    selected_sheet: str,
    column_ids: Any,
) -> None:
    """
    Drop one or more columns.

    - column_ids must be a non-empty list of unique strings
    - resolve all ids on current worksheet state before deletion
    - delete from highest index to lowest to avoid shifting
    """
    if not isinstance(column_ids, list):
        raise PipelineValidationError("columnIds must be a list")
    if not column_ids:
        raise PipelineValidationError("columnIds must be a non-empty list")
    if any(not isinstance(cid, str) or not cid.strip() for cid in column_ids):
        raise PipelineValidationError("columnIds must contain non-empty strings only")
    if len(set(column_ids)) != len(column_ids):
        raise PipelineValidationError("columnIds must not contain duplicates")

    resolved: list[int] = []
    for cid in column_ids:
        resolved.append(
            resolve_column_id(
                ws,
                selected_sheet=selected_sheet,
                header_row_idx=header_row_idx,
                column_id=cid,
            )
        )

    for col_idx in sorted(set(resolved), reverse=True):
        ws.delete_cols(col_idx, 1)

