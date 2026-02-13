from __future__ import annotations

from typing import Any

from .errors import PipelineValidationError


SUPPORTED_OPERATION_IDS: set[str] = {"rename_column", "drop_column"}

# Strict param shapes (exact keys, no defaults, no extras).
ALLOWED_PARAMS_BY_OPERATION: dict[str, set[str]] = {
    "rename_column": {"columnId", "newName"},
    "drop_column": {"columnIds"},
}


def _require_dict(value: Any, *, field_name: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise PipelineValidationError(f"{field_name} must be an object")
    return value


def _require_list(value: Any, *, field_name: str) -> list[Any]:
    if not isinstance(value, list):
        raise PipelineValidationError(f"{field_name} must be a list")
    return value


def _require_str(value: Any, *, field_name: str) -> str:
    if not isinstance(value, str):
        raise PipelineValidationError(f"{field_name} must be a string")
    return value


def _validate_operation_shape(op: dict[str, Any], *, idx: int) -> None:
    required_keys = {"id", "operationId", "params"}
    actual_keys = set(op.keys())
    if actual_keys == required_keys:
        return

    missing = required_keys - actual_keys
    extra = actual_keys - required_keys
    raise PipelineValidationError(
        "Each operation must have exactly keys: id, operationId, params. "
        f"Missing={sorted(missing)} Extra={sorted(extra)}"
    )


def _validate_and_register_operation_id(*, op_id_raw: Any, idx: int, seen_ids: set[str]) -> str:
    op_id = _require_str(op_id_raw, field_name=f"pipeline_operations[{idx}].id").strip()
    if not op_id:
        raise PipelineValidationError(f"pipeline_operations[{idx}].id must be non-empty")
    if op_id in seen_ids:
        raise PipelineValidationError(f"Duplicate operation id: {op_id}")
    seen_ids.add(op_id)
    return op_id


def _validate_operation_id(*, operation_id_raw: Any, idx: int) -> str:
    operation_id = _require_str(
        operation_id_raw, field_name=f"pipeline_operations[{idx}].operationId"
    )
    if operation_id not in SUPPORTED_OPERATION_IDS:
        raise PipelineValidationError(f"Unsupported operationId: {operation_id}")
    return operation_id


def _validate_operation_params(*, operation_id: str, params_raw: Any, idx: int) -> dict[str, Any]:
    params = _require_dict(params_raw, field_name=f"pipeline_operations[{idx}].params")
    allowed_keys = ALLOWED_PARAMS_BY_OPERATION[operation_id]
    param_keys = set(params.keys())
    if param_keys != allowed_keys:
        missing = allowed_keys - param_keys
        extra = param_keys - allowed_keys
        raise PipelineValidationError(
            f"Invalid params for operationId={operation_id}. "
            f"Missing={sorted(missing)} Extra={sorted(extra)}"
        )
    return params


def _validate_and_normalize_operation(
    raw_op: Any, *, idx: int, seen_ids: set[str]
) -> dict[str, Any]:
    op = _require_dict(raw_op, field_name=f"pipeline_operations[{idx}]")
    _validate_operation_shape(op, idx=idx)

    op_id = _validate_and_register_operation_id(op_id_raw=op["id"], idx=idx, seen_ids=seen_ids)
    operation_id = _validate_operation_id(operation_id_raw=op["operationId"], idx=idx)
    params = _validate_operation_params(operation_id=operation_id, params_raw=op["params"], idx=idx)

    return {"id": op_id, "operationId": operation_id, "params": params}


def validate_pipeline_operations(pipeline_operations: Any) -> list[dict[str, Any]]:
    """
    Validate pipeline_operations strictly.

    Rules:
    - pipeline_operations must be a list (may be empty)
    - each item must have exactly keys: id, operationId, params
    - id must be string and unique
    - operationId must be one of SUPPORTED_OPERATION_IDS
    - params must be dict and contain exactly the allowed keys for that operationId
    """
    ops = _require_list(pipeline_operations, field_name="pipeline_operations")

    seen_ids: set[str] = set()
    return [
        _validate_and_normalize_operation(raw_op, idx=idx, seen_ids=seen_ids)
        for idx, raw_op in enumerate(ops)
    ]

