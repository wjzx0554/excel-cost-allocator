from typing import Dict, List, Sequence

from .allocator import ColumnInfo


TEMPLATE_VERSION = 1
TEMPLATE_TYPE = "allocation_schemes"
VALID_AMOUNT_MODES = {"target_total", "source_column", "manual"}
VALID_FILTER_LOGICS = {"OR", "AND"}
VALID_OPERATORS = {
    "equals",
    "not_equals",
    "contains",
    "not_contains",
    "regex",
    "blank",
    "not_blank",
}


def serialize_scheme_template(
    schemes: Sequence[Dict],
    headers: Sequence[ColumnInfo],
    sheet_name: str = "",
    header_row: int = 1,
) -> Dict:
    header_by_label = {item.label: item for item in headers}

    def column_ref(label: str) -> Dict:
        label = label or ""
        info = header_by_label.get(label)
        if info is None:
            return {
                "label": label,
                "header": _header_from_label(label),
                "letter": "",
                "index": None,
            }
        return {
            "label": info.label,
            "header": info.header,
            "letter": info.letter,
            "index": info.index,
        }

    template_schemes = []
    for scheme in schemes:
        rules = []
        for rule in scheme.get("filter_rules", []):
            column = rule.get("column", "")
            if not column:
                continue
            rules.append(
                {
                    "column": column_ref(column),
                    "operator": rule.get("operator", "equals"),
                    "value": rule.get("value", ""),
                }
            )

        template_schemes.append(
            {
                "name": scheme.get("name", ""),
                "amount_mode": scheme.get("amount_mode", "target_total"),
                "amount_column": column_ref(scheme.get("amount_column", "")),
                "manual_amount": scheme.get("manual_amount", ""),
                "allocation_column": column_ref(scheme.get("allocation_column", "")),
                "base_columns": [column_ref(label) for label in scheme.get("base_columns", [])],
                "filter_logic": scheme.get("filter_logic", "OR"),
                "filter_rules": rules,
            }
        )

    return {
        "version": TEMPLATE_VERSION,
        "template_type": TEMPLATE_TYPE,
        "app": "分摊工具",
        "sheet_name": sheet_name,
        "header_row": header_row,
        "headers": [column_ref(item.label) for item in headers],
        "schemes": template_schemes,
    }


def import_scheme_template(template: Dict, headers: Sequence[ColumnInfo]) -> List[Dict]:
    if not isinstance(template, dict):
        raise ValueError("模板文件格式不正确。")
    if template.get("template_type", TEMPLATE_TYPE) != TEMPLATE_TYPE:
        raise ValueError("模板类型不匹配，请选择分摊工具导出的方案模板。")

    schemes = template.get("schemes")
    if not isinstance(schemes, list) or not schemes:
        raise ValueError("模板里没有可导入的分摊方案。")

    labels = {item.label for item in headers}
    errors = []
    imported = []

    def resolve(reference, field_name: str, required: bool = True) -> str:
        if not reference:
            if required:
                errors.append(f"{field_name} 缺少列。")
            return ""

        if isinstance(reference, dict):
            label = reference.get("label", "") or ""
            header = reference.get("header", "") or ""
        else:
            label = str(reference)
            header = _header_from_label(label)

        if label in labels:
            return label

        if header:
            matches = [item.label for item in headers if item.header == header]
            if len(matches) == 1:
                return matches[0]
            if len(matches) > 1:
                errors.append(f"{field_name} 列名重复，无法自动匹配：{header}")
                return ""

        if required:
            errors.append(f"{field_name} 当前表缺少列：{header or label or '空标题'}")
        return ""

    for index, scheme in enumerate(schemes, start=1):
        if not isinstance(scheme, dict):
            errors.append(f"第 {index} 个方案格式不正确。")
            continue

        name = (scheme.get("name") or f"方案{index}").strip() or f"方案{index}"
        amount_mode = scheme.get("amount_mode", "target_total")
        if amount_mode not in VALID_AMOUNT_MODES:
            errors.append(f"{name} 的金额来源类型无效：{amount_mode}")
            amount_mode = "target_total"

        filter_logic = (scheme.get("filter_logic", "OR") or "OR").upper()
        if filter_logic not in VALID_FILTER_LOGICS:
            errors.append(f"{name} 的过滤规则关系无效：{filter_logic}")
            filter_logic = "OR"

        allocation_column = resolve(scheme.get("allocation_column"), f"{name} 的分摊结果列")
        base_columns = [
            label
            for label in (
                resolve(item, f"{name} 的占比计算列") for item in scheme.get("base_columns", [])
            )
            if label
        ]

        amount_column = ""
        if amount_mode == "source_column":
            amount_column = resolve(scheme.get("amount_column"), f"{name} 的金额来源列")
        elif scheme.get("amount_column"):
            amount_column = resolve(
                scheme.get("amount_column"),
                f"{name} 的备用金额来源列",
                required=False,
            )

        rules = []
        for rule_index, rule in enumerate(scheme.get("filter_rules", []), start=1):
            if not isinstance(rule, dict):
                errors.append(f"{name} 的第 {rule_index} 条过滤条件格式不正确。")
                continue
            operator = rule.get("operator", "equals")
            if operator not in VALID_OPERATORS:
                errors.append(f"{name} 的第 {rule_index} 条过滤条件类型无效：{operator}")
                operator = "equals"
            column = resolve(rule.get("column"), f"{name} 的第 {rule_index} 条过滤条件字段")
            if column:
                rules.append(
                    {
                        "column": column,
                        "operator": operator,
                        "value": rule.get("value", ""),
                    }
                )

        imported.append(
            {
                "name": name,
                "amount_mode": amount_mode,
                "amount_column": amount_column,
                "manual_amount": scheme.get("manual_amount", ""),
                "allocation_column": allocation_column,
                "base_columns": base_columns,
                "filter_logic": filter_logic,
                "filter_rules": rules,
            }
        )

    if errors:
        preview = "\n".join(f"- {item}" for item in errors[:12])
        if len(errors) > 12:
            preview += f"\n- 还有 {len(errors) - 12} 个问题未显示"
        raise ValueError(f"模板中的列无法匹配当前表头：\n{preview}")

    return imported


def _header_from_label(label: str) -> str:
    if "列 - " in label:
        return label.split("列 - ", 1)[1]
    return label
