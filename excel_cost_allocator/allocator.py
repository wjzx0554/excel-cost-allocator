from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


MONEY_QUANT = Decimal("0.01")


@dataclass
class ColumnInfo:
    index: int
    letter: str
    header: str

    @property
    def label(self) -> str:
        title = self.header or "空标题"
        return f"{self.letter}列 - {title}"


@dataclass
class AllocationConfig:
    input_path: str
    output_path: str
    sheet_name: str
    header_row: int
    base_columns: Sequence[int]
    allocation_columns: Sequence[int]
    filter_column: Optional[int] = None
    excluded_values: Optional[Set[str]] = None
    detail_sheet_name: str = "分摊明细"


@dataclass
class RowAllocation:
    row_number: int
    participates: bool
    reason: str
    filter_value: str
    base_values: List[Decimal]
    base_total: Decimal
    effective_base: Decimal
    ratio: Decimal
    allocations: Dict[int, Decimal]


@dataclass
class AllocationResult:
    output_path: str
    total_rows: int
    participating_rows: int
    excluded_rows: int
    base_total: Decimal
    allocation_totals: Dict[int, Decimal]


def normalize_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def display_value(value: str) -> str:
    return "<空白>" if value == "" else value


def parse_number(value) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, bool):
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        return Decimal(str(value))
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return Decimal("0")
        text = text.replace(",", "").replace("，", "").replace("￥", "").replace("¥", "")
        try:
            return Decimal(text)
        except InvalidOperation:
            return Decimal("0")
    return Decimal("0")


def round_money(value: Decimal) -> Decimal:
    return value.quantize(MONEY_QUANT, rounding=ROUND_HALF_UP)


def get_sheet_names(file_path: str) -> List[str]:
    wb = load_workbook(file_path, read_only=False, data_only=False)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def guess_header_row(file_path: str, sheet_name: str, max_scan_rows: int = 20) -> int:
    wb = load_workbook(file_path, read_only=False, data_only=True)
    try:
        ws = wb[sheet_name]
        best_row = 1
        best_count = -1
        scan_to = min(ws.max_row, max_scan_rows)
        for row in range(1, scan_to + 1):
            count = _row_value_count(ws, row)
            if count > best_count:
                best_count = count
                best_row = row
        return best_row
    finally:
        wb.close()


def get_headers(file_path: str, sheet_name: str, header_row: int) -> List[ColumnInfo]:
    wb = load_workbook(file_path, read_only=False, data_only=True)
    try:
        ws = wb[sheet_name]
        max_column = _detect_used_max_column(ws, header_row)
        headers: List[ColumnInfo] = []
        for col in range(1, max_column + 1):
            value = normalize_value(ws.cell(header_row, col).value)
            headers.append(ColumnInfo(col, get_column_letter(col), value))
        return headers
    finally:
        wb.close()


def _row_value_count(ws, row_number: int) -> int:
    count = 0
    for (_row, _col), cell in ws._cells.items():
        if _row == row_number and normalize_value(cell.value):
            count += 1
    return count


def _detect_used_max_column(ws, header_row: int, scan_rows: int = 2000) -> int:
    last_col = 0
    max_scan_row = min(ws.max_row or header_row, header_row + scan_rows)

    for (_row, _col), cell in ws._cells.items():
        if header_row <= _row <= max_scan_row and normalize_value(cell.value):
            last_col = max(last_col, _col)

    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= header_row <= merged_range.max_row:
            last_col = max(last_col, merged_range.max_col)

    return max(last_col, 1)


def get_unique_values(
    file_path: str,
    sheet_name: str,
    header_row: int,
    column_index: int,
    max_values: int = 5000,
) -> List[str]:
    wb = load_workbook(file_path, read_only=False, data_only=True)
    try:
        ws = wb[sheet_name]
        seen = set()
        values: List[str] = []
        for row in range(header_row + 1, ws.max_row + 1):
            text = normalize_value(ws.cell(row, column_index).value)
            if text not in seen:
                seen.add(text)
                values.append(text)
                if len(values) >= max_values:
                    break
        return values
    finally:
        wb.close()


def allocate_workbook(config: AllocationConfig) -> AllocationResult:
    _validate_config(config)

    input_path = Path(config.input_path)
    keep_vba = input_path.suffix.lower() == ".xlsm"
    formulas_wb = load_workbook(config.input_path, data_only=False, keep_vba=keep_vba)
    values_wb = load_workbook(config.input_path, data_only=True, keep_vba=keep_vba)

    try:
        ws = formulas_wb[config.sheet_name]
        values_ws = values_wb[config.sheet_name]
        data_start = config.header_row + 1
        max_row = values_ws.max_row
        excluded_values = {normalize_value(v) for v in (config.excluded_values or set())}

        rows: List[RowAllocation] = []
        base_total = Decimal("0")

        for row_number in range(data_start, max_row + 1):
            base_values = [
                parse_number(values_ws.cell(row_number, col).value)
                for col in config.base_columns
            ]
            row_base_total = sum(base_values, Decimal("0"))
            filter_value = (
                normalize_value(values_ws.cell(row_number, config.filter_column).value)
                if config.filter_column
                else ""
            )

            filtered_out = bool(config.filter_column and filter_value in excluded_values)
            if filtered_out:
                participates = False
                reason = "过滤排除"
            elif row_base_total <= 0:
                participates = False
                reason = "基数小于等于0"
            else:
                participates = True
                reason = ""

            effective_base = row_base_total if participates else Decimal("0")
            base_total += effective_base
            rows.append(
                RowAllocation(
                    row_number=row_number,
                    participates=participates,
                    reason=reason,
                    filter_value=filter_value,
                    base_values=base_values,
                    base_total=row_base_total,
                    effective_base=effective_base,
                    ratio=Decimal("0"),
                    allocations={},
                )
            )

        allocation_totals: Dict[int, Decimal] = {}
        for target_col in config.allocation_columns:
            target_total = round_money(
                sum(
                    parse_number(values_ws.cell(row, target_col).value)
                    for row in range(data_start, max_row + 1)
                )
            )
            allocation_totals[target_col] = target_total
            if target_total != 0 and base_total <= 0:
                raise ValueError("没有可参与分摊的行，无法分配非零费用。")

            rounded_sum = Decimal("0")
            for item in rows:
                item.ratio = item.effective_base / base_total if base_total > 0 else Decimal("0")
                amount = round_money(target_total * item.ratio) if item.participates else Decimal("0")
                item.allocations[target_col] = amount
                rounded_sum += amount

            residual = target_total - rounded_sum
            residual_row = _find_residual_row(rows)
            if residual_row is not None:
                residual_row.allocations[target_col] = round_money(
                    residual_row.allocations[target_col] + residual
                )

            for item in rows:
                cell = ws.cell(item.row_number, target_col)
                cell.value = float(item.allocations[target_col])
                cell.number_format = '#,##0.00'

        _write_detail_sheet(
            formulas_wb,
            config,
            rows,
            base_total,
            allocation_totals,
        )

        formulas_wb.save(config.output_path)
    finally:
        formulas_wb.close()
        values_wb.close()

    participating_rows = sum(1 for item in rows if item.participates)
    return AllocationResult(
        output_path=config.output_path,
        total_rows=len(rows),
        participating_rows=participating_rows,
        excluded_rows=len(rows) - participating_rows,
        base_total=round_money(base_total),
        allocation_totals=allocation_totals,
    )


def _validate_config(config: AllocationConfig) -> None:
    if not config.input_path:
        raise ValueError("请选择 Excel 文件。")
    if not config.output_path:
        raise ValueError("请选择输出文件。")
    if config.header_row < 1:
        raise ValueError("表头行必须大于等于 1。")
    if not config.base_columns:
        raise ValueError("请至少选择一个参与占比计算的列。")
    if not config.allocation_columns:
        raise ValueError("请至少选择一个需要分配的费用列。")
    overlap = set(config.base_columns) & set(config.allocation_columns)
    if overlap:
        letters = ", ".join(get_column_letter(col) for col in sorted(overlap))
        raise ValueError(f"参与占比列和需要分配列不能重复：{letters}")


def _find_residual_row(rows: Iterable[RowAllocation]) -> Optional[RowAllocation]:
    participating = [item for item in rows if item.participates and item.effective_base > 0]
    if not participating:
        return None
    return max(participating, key=lambda item: (item.effective_base, -item.row_number))


def _header(ws, row: int, col: int, value: str) -> None:
    cell = ws.cell(row, col, value)
    cell.fill = PatternFill("solid", fgColor="305496")
    cell.font = Font(bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )


def _write_detail_sheet(
    wb,
    config: AllocationConfig,
    rows: Sequence[RowAllocation],
    base_total: Decimal,
    allocation_totals: Dict[int, Decimal],
) -> None:
    if config.detail_sheet_name in wb.sheetnames:
        del wb[config.detail_sheet_name]

    detail = wb.create_sheet(config.detail_sheet_name)
    detail.sheet_properties.tabColor = "5B9BD5"
    detail["A1"] = "费用分摊明细"
    detail["A1"].font = Font(size=15, bold=True, color="1F4E78")

    participating = sum(1 for item in rows if item.participates)
    detail["A2"] = "原工作表"
    detail["B2"] = config.sheet_name
    detail["C2"] = "表头行"
    detail["D2"] = config.header_row
    detail["A3"] = "数据行数"
    detail["B3"] = len(rows)
    detail["C3"] = "参与行数"
    detail["D3"] = participating
    detail["E3"] = "不参与行数"
    detail["F3"] = len(rows) - participating
    detail["A4"] = "有效分摊基数合计"
    detail["B4"] = float(round_money(base_total))
    detail["B4"].number_format = '#,##0.00'

    source_ws = wb[config.sheet_name]
    source_headers = {
        col: normalize_value(source_ws.cell(config.header_row, col).value) or get_column_letter(col)
        for col in set(config.base_columns) | set(config.allocation_columns)
    }

    row_cursor = 6
    summary_headers = ["费用列", "原始合计", "分摊后合计"]
    for col, title in enumerate(summary_headers, start=1):
        _header(detail, row_cursor, col, title)
    for offset, col_index in enumerate(config.allocation_columns, start=1):
        row = row_cursor + offset
        total = allocation_totals[col_index]
        distributed = sum(item.allocations[col_index] for item in rows)
        detail.cell(row, 1, f"{get_column_letter(col_index)}列 - {source_headers[col_index]}")
        detail.cell(row, 2, float(total))
        detail.cell(row, 3, float(distributed))
        detail.cell(row, 2).number_format = '#,##0.00'
        detail.cell(row, 3).number_format = '#,##0.00'

    table_row = row_cursor + len(config.allocation_columns) + 3
    headers = ["源行号", "是否参与", "不参与原因", "过滤列值"]
    headers.extend(
        f"基数列 {get_column_letter(col)} - {source_headers[col]}"
        for col in config.base_columns
    )
    headers.extend(["分摊基数", "占比"])
    headers.extend(
        f"分摊结果 {get_column_letter(col)} - {source_headers[col]}"
        for col in config.allocation_columns
    )

    for col, title in enumerate(headers, start=1):
        _header(detail, table_row, col, title)

    for row_offset, item in enumerate(rows, start=1):
        row = table_row + row_offset
        col = 1
        detail.cell(row, col, item.row_number)
        col += 1
        detail.cell(row, col, "是" if item.participates else "否")
        col += 1
        detail.cell(row, col, item.reason)
        col += 1
        detail.cell(row, col, item.filter_value)
        col += 1
        for value in item.base_values:
            detail.cell(row, col, float(value))
            detail.cell(row, col).number_format = '#,##0.00'
            col += 1
        detail.cell(row, col, float(item.effective_base))
        detail.cell(row, col).number_format = '#,##0.00'
        col += 1
        detail.cell(row, col, float(item.ratio))
        detail.cell(row, col).number_format = '0.0000%'
        col += 1
        for target_col in config.allocation_columns:
            detail.cell(row, col, float(item.allocations[target_col]))
            detail.cell(row, col).number_format = '#,##0.00'
            col += 1

    detail.freeze_panes = detail.cell(table_row + 1, 1).coordinate
    detail.auto_filter.ref = f"A{table_row}:{get_column_letter(len(headers))}{table_row + len(rows)}"

    for col in range(1, len(headers) + 1):
        detail.column_dimensions[get_column_letter(col)].width = 16
    detail.column_dimensions["A"].width = 10
    detail.column_dimensions["C"].width = 16
    detail.column_dimensions["D"].width = 18
