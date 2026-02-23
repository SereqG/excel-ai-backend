from __future__ import annotations

from datetime import date, datetime
from typing import Any, Optional

from openpyxl.worksheet.worksheet import Worksheet

from ..errors import PipelineValidationError


def _require_non_empty_str(value: Any, *, field_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise PipelineValidationError(f"{field_name} must be a non-empty string")
    return value.strip()


def _require_dict(value: Any, *, field_name: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise PipelineValidationError(f"{field_name} must be an object")
    return value


def _require_list(value: Any, *, field_name: str) -> list[Any]:
    if not isinstance(value, list):
        raise PipelineValidationError(f"{field_name} must be a list")
    return value


def _require_exact_keys(obj: dict[str, Any], *, expected_keys: set[str], field_name: str) -> None:
    actual_keys = set(obj.keys())
    if actual_keys == expected_keys:
        return
    missing = expected_keys - actual_keys
    extra = actual_keys - expected_keys
    raise PipelineValidationError(
        f"{field_name} must have exactly keys {sorted(expected_keys)}. "
        f"Missing={sorted(missing)} Extra={sorted(extra)}"
    )


def _set_dim_attr_if_possible(dim, attr: str, value: Any) -> None:
    """
    Some Row/ColumnDimension attributes are read-only depending on openpyxl version.
    Skip any attribute that cannot be set.
    """
    try:
        setattr(dim, attr, value)
    except (AttributeError, TypeError):
        return


def _filter_non_empty_values(values: list[Any]) -> list[Any]:
    return [v for v in values if v is not None and v != ""]


def _is_numeric_value(value: Any) -> bool:
    if isinstance(value, (int, float)):
        return True

    if isinstance(value, str):
        try:
            float(value.replace(",", ""))
            return True
        except (ValueError, AttributeError):
            return False

    return False


# Keep existing supported input parsing formats and extend with the requested ones.
_DATE_INPUT_STRPTIME_FORMATS: tuple[str, ...] = (
    "%Y/%m/%d",
    "%d/%m/%Y",
    "%Y.%m.%d",
    "%d.%m.%Y",
    "%Y-%m-%d",
    "%d-%m-%Y",
    # legacy / already supported
    "%m/%d/%Y",
)


def _is_date_value(value: Any) -> bool:
    if isinstance(value, (datetime, date)):
        return True

    if isinstance(value, str):
        for date_format in _DATE_INPUT_STRPTIME_FORMATS:
            try:
                datetime.strptime(value, date_format)
                return True
            except (ValueError, TypeError):
                continue

    return False


def _count_numeric_values(values: list[Any]) -> int:
    count = 0
    for value in values:
        if _is_numeric_value(value):
            count += 1
    return count


def _count_date_values(values: list[Any]) -> int:
    count = 0
    for value in values:
        if _is_date_value(value):
            count += 1
    return count


def _count_boolean_values(values: list[Any]) -> int:
    return sum(1 for v in values if isinstance(v, bool))


def _determine_type_from_counts(total: int, bool_count: int, date_count: int, number_count: int) -> str:
    if bool_count == total:
        return "boolean"

    if date_count / total >= 0.5:
        return "date"

    if number_count / total >= 0.8:
        return "number"

    return "string"


def _determine_column_type(values: list[Any]) -> Optional[str]:
    if not values:
        return None

    non_empty_values = _filter_non_empty_values(values)
    if not non_empty_values:
        return None

    bool_count = _count_boolean_values(non_empty_values)
    date_count = _count_date_values(non_empty_values)
    number_count = _count_numeric_values(non_empty_values)
    total = len(non_empty_values)

    return _determine_type_from_counts(total, bool_count, date_count, number_count)


def _is_empty_row(row: tuple[Any, ...]) -> bool:
    return all(cell is None or cell == "" for cell in row)


def _collect_column_values(ws: Worksheet, *, header_row_idx: int, col_idx: int) -> list[Any]:
    values: list[Any] = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if _is_empty_row(row):
            continue
        values.append(row[col_idx - 1] if (col_idx - 1) < len(row) else None)
    return values


def _parse_number_value(value: Any, *, field_name: str) -> float:
    if isinstance(value, bool):
        # bool is a subclass of int; treat it as invalid for numeric comparisons
        raise PipelineValidationError(f"{field_name} must be a number")
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if not s:
            raise PipelineValidationError(f"{field_name} must be a number")
        try:
            return float(s.replace(",", ""))
        except ValueError:
            raise PipelineValidationError(f"{field_name} must be a number")
    raise PipelineValidationError(f"{field_name} must be a number")


def _parse_cell_number(value: Any, *, column_id: str, row_idx: int) -> float:
    try:
        return _parse_number_value(value, field_name="cellValue")
    except PipelineValidationError as exc:
        raise PipelineValidationError(
            f"Failed to parse number in ColumnId={column_id} at row={row_idx}: "
            f"value={value!r} (type={type(value).__name__}). {exc}"
        )


def _parse_date_cell_value(value: Any) -> date:
    if isinstance(value, (datetime, date)):
        return value

    if isinstance(value, str):
        s = value.strip()
        if not s:
            raise ValueError("Empty string")
        for fmt in _DATE_INPUT_STRPTIME_FORMATS:
            try:
                return datetime.strptime(s, fmt)
            except (ValueError, TypeError):
                continue
        raise ValueError(f"Unsupported date string format: {s!r}")

    raise ValueError(f"Unsupported date cell type: {type(value).__name__}")


def _parse_date_value(value: Any, *, field_name: str) -> date:
    if value is None or value == "":
        raise PipelineValidationError(f"{field_name} must be a date")

    try:
        parsed = _parse_date_cell_value(value)
    except Exception as exc:
        raise PipelineValidationError(f"{field_name} must be a date. {exc}")

    if isinstance(parsed, datetime):
        return parsed.date()
    if isinstance(parsed, date):
        return parsed
    raise PipelineValidationError(f"{field_name} must be a date")


def _parse_cell_date(value: Any, *, column_id: str, row_idx: int) -> date:
    try:
        return _parse_date_value(value, field_name="cellValue")
    except PipelineValidationError as exc:
        raise PipelineValidationError(
            f"Failed to parse date in ColumnId={column_id} at row={row_idx}: "
            f"value={value!r} (type={type(value).__name__}). {exc}"
        )


def _infer_column_type_cached(
    ws: Worksheet,
    *,
    header_row_idx: int,
    col_idx: int,
    cache: dict[int, Optional[str]],
) -> Optional[str]:
    if col_idx in cache:
        return cache[col_idx]
    inferred = _determine_column_type(_collect_column_values(ws, header_row_idx=header_row_idx, col_idx=col_idx))
    cache[col_idx] = inferred
    return inferred

