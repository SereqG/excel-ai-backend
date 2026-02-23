"""
Domain-driven worksheet operations used by pipeline execution.

This package replaces the legacy `services/operations.py` module while keeping
the public API stable via re-exports of `apply_*` functions.
"""

from .columns import apply_add_column, apply_drop_column, apply_reorder_columns, apply_rename_column
from .dates import apply_parse_date
from .rows_filter import apply_filter_rows
from .rows_sort import apply_sort_rows
from .text import apply_normalize_case, apply_replace_text

__all__ = [
    "apply_add_column",
    "apply_drop_column",
    "apply_filter_rows",
    "apply_normalize_case",
    "apply_parse_date",
    "apply_rename_column",
    "apply_replace_text",
    "apply_reorder_columns",
    "apply_sort_rows",
]

