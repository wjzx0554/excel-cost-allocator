from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_cost_allocator.allocator import AllocationConfig, allocate_workbook


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

