from __future__ import annotations

from copy import copy
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

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


def _require_non_empty_str(value: Any, *, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise PipelineValidationError(f"{field_name} must be a non-empty string")
    return value.strip()


def _parse_constant_source(source: Any) -> Any:
    if not isinstance(source, dict):
        raise PipelineValidationError("source must be an object")
    if set(source.keys()) != {"kind", "value"}:
        missing = {"kind", "value"} - set(source.keys())
        extra = set(source.keys()) - {"kind", "value"}
        raise PipelineValidationError(
            "source must have exactly keys: kind, value. "
            f"Missing={sorted(missing)} Extra={sorted(extra)}"
        )
    kind = source.get("kind")
    if kind != "constant":
        raise PipelineValidationError("source.kind must be 'constant'")
    return source.get("value")


def apply_add_column(
    ws: Worksheet,
    *,
    header_row_idx: int,
    column_name: Any,
    source: Any,
) -> None:
    """
    Append a new column as the last column.

    - column_name must be a non-empty string and must not collide with an existing header value
    - source must be an object with kind='constant' and a value to fill all data rows
    """
    column_name = _require_non_empty_str(column_name, field_name="columnName")
    fill_value = _parse_constant_source(source)

    header_vals = _header_values(ws, header_row_idx=header_row_idx)
    if column_name in header_vals:
        raise PipelineValidationError(f"Add column collision: header already contains '{column_name}'")

    template_col_idx = ws.max_column
    new_col_idx = template_col_idx + 1

    # Copy column dimension metadata from the current last column (n) to new last (n+1).
    _copy_column_dimensions(ws, from_col=template_col_idx, to_col=new_col_idx)

    # Copy styles from template cells, but overwrite values.
    for row_idx in range(1, ws.max_row + 1):
        src = ws.cell(row=row_idx, column=template_col_idx)
        dst = ws.cell(row=row_idx, column=new_col_idx)
        _copy_cell_style(src, dst)

    ws.cell(row=header_row_idx, column=new_col_idx).value = column_name
    for row_idx in range(header_row_idx + 1, ws.max_row + 1):
        ws.cell(row=row_idx, column=new_col_idx).value = fill_value


def _swap_cell_contents(a, b) -> None:
    # Values
    a_val, b_val = a.value, b.value

    # Core style attributes (aligned with workbook copy logic)
    a_font, b_font = copy(a.font), copy(b.font)
    a_border, b_border = copy(a.border), copy(b.border)
    a_fill, b_fill = copy(a.fill), copy(b.fill)
    a_alignment, b_alignment = copy(a.alignment), copy(b.alignment)
    a_protection, b_protection = copy(a.protection), copy(b.protection)
    a_number_format, b_number_format = a.number_format, b.number_format

    # Other commonly used cell metadata
    a_hyperlink, b_hyperlink = a.hyperlink, b.hyperlink
    a_comment, b_comment = a.comment, b.comment

    a.value, b.value = b_val, a_val

    a.font, b.font = b_font, a_font
    a.border, b.border = b_border, a_border
    a.fill, b.fill = b_fill, a_fill
    a.alignment, b.alignment = b_alignment, a_alignment
    a.protection, b.protection = b_protection, a_protection
    a.number_format, b.number_format = b_number_format, a_number_format

    a.hyperlink, b.hyperlink = b_hyperlink, a_hyperlink
    a.comment, b.comment = b_comment, a_comment


def _copy_cell_style(src, dst) -> None:
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)
    dst.number_format = src.number_format


_COLUMN_DIM_ATTRS: tuple[str, ...] = (
    "width",
    "hidden",
    "bestFit",
    "outlineLevel",
    "collapsed",
    "style",
)

def _set_dim_attr_if_possible(dim, attr: str, value: Any) -> None:
    """
    Some ColumnDimension attributes are read-only properties depending on openpyxl version.
    Skip any attribute that cannot be set.
    """
    try:
        setattr(dim, attr, value)
    except (AttributeError, TypeError):
        return


def _copy_column_dimensions(ws: Worksheet, *, from_col: int, to_col: int) -> None:
    letter_from = get_column_letter(from_col)
    letter_to = get_column_letter(to_col)
    dim_from = ws.column_dimensions[letter_from]
    dim_to = ws.column_dimensions[letter_to]

    for attr in _COLUMN_DIM_ATTRS:
        if hasattr(dim_from, attr) and hasattr(dim_to, attr):
            _set_dim_attr_if_possible(dim_to, attr, getattr(dim_from, attr))



def apply_reorder_columns(
    ws: Worksheet,
    *,
    header_row_idx: int,
    selected_sheet: str,
    column_ids: Any,
) -> None:
    """
    Reorder multiple columns (including header + all data rows).

    - column_ids must be a list of >= 2 unique non-empty strings
    - both columnIds must resolve against the current worksheet state
    - columns listed are reordered to match the given order while keeping all other columns in place
    """
    if not isinstance(column_ids, list):
        raise PipelineValidationError("columnIds must be a list")
    if len(column_ids) < 2:
        raise PipelineValidationError("columnIds must contain at least two items")
    if any(not isinstance(cid, str) or not cid.strip() for cid in column_ids):
        raise PipelineValidationError("columnIds must contain non-empty strings only")
    if len(set(column_ids)) != len(column_ids):
        raise PipelineValidationError("columnIds must not contain duplicates")

    resolved: list[int] = [
        resolve_column_id(
            ws,
            selected_sheet=selected_sheet,
            header_row_idx=header_row_idx,
            column_id=cid,
        )
        for cid in column_ids
    ]
    if len(set(resolved)) != len(resolved):
        raise PipelineValidationError("columnIds must refer to distinct columns")

    # We reorder within the set of positions currently occupied by these columns,
    # leaving other columns untouched.
    dest_cols = sorted(resolved)
    max_row = ws.max_row

    def snapshot_column(col_idx: int) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        cells: list[dict[str, Any]] = []
        for row_idx in range(1, max_row + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            cells.append(
                {
                    "value": c.value,
                    "font": copy(c.font),
                    "border": copy(c.border),
                    "fill": copy(c.fill),
                    "alignment": copy(c.alignment),
                    "protection": copy(c.protection),
                    "number_format": c.number_format,
                    "hyperlink": copy(c.hyperlink) if c.hyperlink else None,
                    "comment": copy(c.comment) if c.comment else None,
                }
            )

        letter = get_column_letter(col_idx)
        dim = ws.column_dimensions[letter]
        dim_snap: dict[str, Any] = {}
        for attr in _COLUMN_DIM_ATTRS:
            if hasattr(dim, attr):
                dim_snap[attr] = getattr(dim, attr)
        return cells, dim_snap

    def apply_column_snapshot(dest_col: int, *, cells: list[dict[str, Any]], dim_snap: dict[str, Any]) -> None:
        for row_idx, snap in enumerate(cells, start=1):
            c = ws.cell(row=row_idx, column=dest_col)
            c.value = snap["value"]
            c.font = snap["font"]
            c.border = snap["border"]
            c.fill = snap["fill"]
            c.alignment = snap["alignment"]
            c.protection = snap["protection"]
            c.number_format = snap["number_format"]
            c.hyperlink = snap["hyperlink"]
            c.comment = snap["comment"]

        letter = get_column_letter(dest_col)
        dim = ws.column_dimensions[letter]
        for attr, val in dim_snap.items():
            if hasattr(dim, attr):
                _set_dim_attr_if_possible(dim, attr, val)

    snapshots = [snapshot_column(col) for col in resolved]
    for dest_col, (cells, dim_snap) in zip(dest_cols, snapshots):
        apply_column_snapshot(dest_col, cells=cells, dim_snap=dim_snap)

