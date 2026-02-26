from __future__ import annotations

from typing import Any, Optional

from openpyxl.worksheet.worksheet import Worksheet

from ..column_id import resolve_column_id
from ..errors import PipelineValidationError
from .common import (
    _infer_column_type_cached,
    _parse_cell_date,
    _parse_cell_number,
    _parse_date_value,
    _parse_number_value,
    _require_dict,
    _require_exact_keys,
    _require_list,
)


_FILTER_ROWS_ACTIONS: set[str] = {"keep", "drop"}
_FILTER_ROWS_WHEN_OPS: set[str] = {"and", "or"}

_FILTER_ROWS_OPERATORS_BY_TYPE: dict[str, set[str]] = {
    "text": {
        "equals",
        "not_equals",
        "contains",
        "starts_with",
        "ends_with",
        "is_empty",
        "is_not_empty",
    },
    "number": {"eq", "ne", "gt", "gte", "lt", "lte", "between", "is_empty", "is_not_empty"},
    "date": {"on", "before", "after", "between", "is_empty", "is_not_empty"},
    "boolean": {"is_true", "is_false", "is_empty", "is_not_empty"},
}


def apply_filter_rows(
    ws: Worksheet,
    *,
    header_row_idx: int,
    selected_sheet: str,
    default_action: Any,
    rules: Any,
) -> None:
    """
    Filter worksheet rows (keep/drop) based on ordered rules.

    Semantics:
    - evaluate all rules in order (last matching rule wins)
    - if no rule matches, apply default_action
    - action is applied when the rule's `when` evaluates to True
    """

    def require_action(value: Any, *, field_name: str) -> str:
        if not isinstance(value, str):
            raise PipelineValidationError(f"{field_name} must be a string")
        s = value.strip()
        if s not in _FILTER_ROWS_ACTIONS:
            raise PipelineValidationError(f"{field_name} must be one of {sorted(_FILTER_ROWS_ACTIONS)}")
        return s

    default_action_str = require_action(default_action, field_name="defaultAction")

    rules_list = _require_list(rules, field_name="rules")
    if not rules_list:
        raise PipelineValidationError("rules must be a non-empty list")

    col_type_cache: dict[int, Optional[str]] = {}

    compiled_rules: list[dict[str, Any]] = []
    for rule_idx, raw_rule in enumerate(rules_list):
        rule = _require_dict(raw_rule, field_name=f"rules[{rule_idx}]")
        _require_exact_keys(rule, expected_keys={"action", "when"}, field_name=f"rules[{rule_idx}]")

        action_str = require_action(rule.get("action"), field_name=f"rules[{rule_idx}].action")
        when_raw = _require_dict(rule.get("when"), field_name=f"rules[{rule_idx}].when")
        _require_exact_keys(when_raw, expected_keys={"op", "conditions"}, field_name=f"rules[{rule_idx}].when")

        op_raw = when_raw.get("op")
        if not isinstance(op_raw, str):
            raise PipelineValidationError(f"rules[{rule_idx}].when.op must be a string")
        op_str = op_raw.strip()
        if op_str not in _FILTER_ROWS_WHEN_OPS:
            raise PipelineValidationError(
                f"rules[{rule_idx}].when.op must be one of {sorted(_FILTER_ROWS_WHEN_OPS)}"
            )

        conditions_list = _require_list(
            when_raw.get("conditions"), field_name=f"rules[{rule_idx}].when.conditions"
        )
        if not conditions_list:
            raise PipelineValidationError(f"rules[{rule_idx}].when.conditions must be a non-empty list")

        compiled_conditions: list[dict[str, Any]] = []
        for cond_idx, raw_cond in enumerate(conditions_list):
            cond = _require_dict(raw_cond, field_name=f"rules[{rule_idx}].when.conditions[{cond_idx}]")
            allowed_keys = {"type", "columnId", "operator", "value", "value2"}
            extra_keys = set(cond.keys()) - allowed_keys
            if extra_keys:
                raise PipelineValidationError(
                    f"rules[{rule_idx}].when.conditions[{cond_idx}] contains unknown keys: "
                    f"{sorted(extra_keys)}"
                )

            type_raw = cond.get("type")
            if not isinstance(type_raw, str):
                raise PipelineValidationError(
                    f"rules[{rule_idx}].when.conditions[{cond_idx}].type must be a string"
                )
            cond_type = type_raw.strip()
            if cond_type not in _FILTER_ROWS_OPERATORS_BY_TYPE:
                raise PipelineValidationError(
                    f"rules[{rule_idx}].when.conditions[{cond_idx}].type must be one of "
                    f"{sorted(_FILTER_ROWS_OPERATORS_BY_TYPE.keys())}"
                )

            column_id_raw = cond.get("columnId")
            if not isinstance(column_id_raw, str) or not column_id_raw.strip():
                raise PipelineValidationError(
                    f"rules[{rule_idx}].when.conditions[{cond_idx}].columnId must be a non-empty string"
                )
            column_id = column_id_raw.strip()
            col_idx = resolve_column_id(
                ws,
                selected_sheet=selected_sheet,
                header_row_idx=header_row_idx,
                column_id=column_id,
            )

            operator_raw = cond.get("operator")
            if not isinstance(operator_raw, str):
                raise PipelineValidationError(
                    f"rules[{rule_idx}].when.conditions[{cond_idx}].operator must be a string"
                )
            operator = operator_raw.strip()
            allowed_ops = _FILTER_ROWS_OPERATORS_BY_TYPE[cond_type]
            if operator not in allowed_ops:
                raise PipelineValidationError(
                    f"rules[{rule_idx}].when.conditions[{cond_idx}].operator must be one of "
                    f"{sorted(allowed_ops)} for type={cond_type}"
                )

            inferred = _infer_column_type_cached(
                ws, header_row_idx=header_row_idx, col_idx=col_idx, cache=col_type_cache
            )
            expected_inferred: Optional[str]
            if cond_type == "text":
                expected_inferred = "string"
            else:
                expected_inferred = cond_type

            if inferred is not None and inferred != expected_inferred:
                raise PipelineValidationError(
                    f"ColumnId={column_id} must be {cond_type}. Inferred type={inferred}"
                )

            compiled: dict[str, Any] = {
                "type": cond_type,
                "columnId": column_id,
                "colIdx": col_idx,
                "operator": operator,
            }

            # Validate and pre-parse constants
            if operator == "between":
                if "value" not in cond or "value2" not in cond:
                    raise PipelineValidationError(
                        f"rules[{rule_idx}].when.conditions[{cond_idx}] must include value and value2 for between"
                    )
                raw_min = cond.get("value")
                raw_max = cond.get("value2")
                if cond_type == "number":
                    min_v = _parse_number_value(raw_min, field_name="value")
                    max_v = _parse_number_value(raw_max, field_name="value2")
                    if min_v > max_v:
                        raise PipelineValidationError("between requires value <= value2")
                    compiled["value"] = min_v
                    compiled["value2"] = max_v
                elif cond_type == "date":
                    min_d = _parse_date_value(raw_min, field_name="value")
                    max_d = _parse_date_value(raw_max, field_name="value2")
                    if min_d > max_d:
                        raise PipelineValidationError("between requires value <= value2")
                    compiled["value"] = min_d
                    compiled["value2"] = max_d
                else:
                    raise PipelineValidationError("between operator is only supported for number/date")
            elif operator in {"is_empty", "is_not_empty", "is_true", "is_false"}:
                # no constants needed
                pass
            else:
                if "value" not in cond:
                    raise PipelineValidationError(
                        f"rules[{rule_idx}].when.conditions[{cond_idx}] must include value"
                    )
                raw_value = cond.get("value")
                if cond_type == "number":
                    compiled["value"] = _parse_number_value(raw_value, field_name="value")
                elif cond_type == "date":
                    compiled["value"] = _parse_date_value(raw_value, field_name="value")
                elif cond_type == "text":
                    if raw_value is None:
                        raise PipelineValidationError("value must be a string")
                    compiled["value"] = raw_value if isinstance(raw_value, str) else str(raw_value)
                elif cond_type == "boolean":
                    # boolean operators don't use value currently
                    pass

            compiled_conditions.append(compiled)

        compiled_rules.append({"action": action_str, "op": op_str, "conditions": compiled_conditions})

    def is_empty_cell(v: Any) -> bool:
        return v is None or v == ""

    def eval_condition(*, row_idx: int, compiled_condition: dict[str, Any]) -> bool:
        cond_type = compiled_condition["type"]
        operator = compiled_condition["operator"]
        col_idx = compiled_condition["colIdx"]
        column_id = compiled_condition["columnId"]

        cell_value = ws.cell(row=row_idx, column=col_idx).value

        if operator == "is_empty":
            return is_empty_cell(cell_value)
        if operator == "is_not_empty":
            return not is_empty_cell(cell_value)

        if is_empty_cell(cell_value):
            return False

        if cond_type == "text":
            s = cell_value if isinstance(cell_value, str) else str(cell_value)
            target = compiled_condition.get("value", "")
            if operator == "equals":
                return s == target
            if operator == "not_equals":
                return s != target
            if operator == "contains":
                return target in s
            if operator == "starts_with":
                return s.startswith(target)
            if operator == "ends_with":
                return s.endswith(target)
            raise PipelineValidationError(f"Unsupported text operator: {operator}")

        if cond_type == "number":
            num = _parse_cell_number(cell_value, column_id=column_id, row_idx=row_idx)
            if operator == "eq":
                return num == compiled_condition["value"]
            if operator == "ne":
                return num != compiled_condition["value"]
            if operator == "gt":
                return num > compiled_condition["value"]
            if operator == "gte":
                return num >= compiled_condition["value"]
            if operator == "lt":
                return num < compiled_condition["value"]
            if operator == "lte":
                return num <= compiled_condition["value"]
            if operator == "between":
                return compiled_condition["value"] <= num <= compiled_condition["value2"]
            raise PipelineValidationError(f"Unsupported number operator: {operator}")

        if cond_type == "date":
            d = _parse_cell_date(cell_value, column_id=column_id, row_idx=row_idx)
            if operator == "on":
                return d == compiled_condition["value"]
            if operator == "before":
                return d < compiled_condition["value"]
            if operator == "after":
                return d > compiled_condition["value"]
            if operator == "between":
                return compiled_condition["value"] <= d <= compiled_condition["value2"]
            raise PipelineValidationError(f"Unsupported date operator: {operator}")

        if cond_type == "boolean":
            if not isinstance(cell_value, bool):
                raise PipelineValidationError(
                    f"Failed to parse boolean in ColumnId={column_id} at row={row_idx}: "
                    f"value={cell_value!r} (type={type(cell_value).__name__})"
                )
            if operator == "is_true":
                return cell_value is True
            if operator == "is_false":
                return cell_value is False
            raise PipelineValidationError(f"Unsupported boolean operator: {operator}")

        raise PipelineValidationError(f"Unsupported condition type: {cond_type}")

    def eval_rule(*, row_idx: int, rule: dict[str, Any]) -> bool:
        op_str = rule["op"]
        conditions = rule["conditions"]
        if op_str == "and":
            for c in conditions:
                if not eval_condition(row_idx=row_idx, compiled_condition=c):
                    return False
            return True
        if op_str == "or":
            for c in conditions:
                if eval_condition(row_idx=row_idx, compiled_condition=c):
                    return True
            return False
        raise PipelineValidationError(f"Unsupported when.op: {op_str}")

    rows_to_drop: list[int] = []
    for row_idx in range(header_row_idx + 1, ws.max_row + 1):
        row_action: Optional[str] = None
        for rule in compiled_rules:
            if eval_rule(row_idx=row_idx, rule=rule):
                row_action = rule["action"]
        final_action = row_action or default_action_str
        if final_action == "drop":
            rows_to_drop.append(row_idx)

    for row_idx in sorted(rows_to_drop, reverse=True):
        ws.delete_rows(row_idx, 1)

