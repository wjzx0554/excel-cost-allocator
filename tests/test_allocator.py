from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_cost_allocator.allocator import (
    AllocationConfig,
    allocate_workbook,
    get_headers,
    get_unique_values,
)


def test_allocate_with_filter_and_multiple_targets(tmp_path: Path):
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws.append(["说明", None, None, None, None])
    ws.append(["车间", "材料", "人工", "共耗料", "水电费"])
    ws.append(["生产车间", 100, 0, 10, 20])
    ws.append(["销售配货部", 100, 100, 30, 40])
    ws.append(["生产车间", 300, 100, 60, 40])
    wb.save(input_path)

    result = allocate_workbook(
        AllocationConfig(
            input_path=str(input_path),
            output_path=str(output_path),
            sheet_name="sheet1",
            header_row=2,
            base_columns=[2, 3],
            allocation_columns=[4, 5],
            filter_column=1,
            excluded_values={"销售配货部"},
        )
    )

    assert result.total_rows == 3
    assert result.participating_rows == 2
    assert result.excluded_rows == 1

    out = load_workbook(output_path, data_only=True)
    ws = out["sheet1"]

    assert ws["D3"].value == 20
    assert ws["D4"].value == 0
    assert ws["D5"].value == 80
    assert ws["E3"].value == 20
    assert ws["E4"].value == 0
    assert ws["E5"].value == 80

    assert "分摊明细" in out.sheetnames
    detail = out["分摊明细"]
    assert detail["B3"].value == 3
    assert detail["D3"].value == 2
    assert detail["F3"].value == 1


def test_headers_keep_trailing_blank_title_columns(tmp_path: Path):
    path = tmp_path / "headers.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws["A1"] = "表头1"
    ws["C1"] = "表头3"
    ws["D2"] = 123
    ws["F5"] = 456
    wb.save(path)

    headers = get_headers(str(path), "sheet1", 1)
    assert len(headers) >= 6
    assert headers[1].header == ""
    assert headers[2].header == "表头3"
    assert headers[5].header == ""


def test_unique_values_reads_filter_column(tmp_path: Path):
    path = tmp_path / "values.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"
    ws.append(["车间", "金额"])
    ws.append(["生产车间", 10])
    ws.append(["销售配货部", 20])
    ws.append(["生产车间", 30])
    ws.append([None, 40])
    wb.save(path)

    assert get_unique_values(str(path), "sheet1", 1, 1) == ["生产车间", "销售配货部", ""]
