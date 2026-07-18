"""Modern Flet desktop UI for the Excel cost allocator.

Install the Python 3.8 compatible desktop runtime with
``python -m pip install "flet[desktop]==0.25.2"`` and run this file directly.
"""

import json
import sys
import threading
import traceback
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import flet as ft

from excel_cost_allocator.allocator import (
    AllocationScheme,
    BatchAllocationConfig,
    ColumnInfo,
    FilterRule,
    allocate_workbook_batch,
    create_sample_workbook,
    display_value,
    get_headers,
    get_sheet_names,
    get_unique_values,
    guess_header_row,
    preview_filter_matches,
    preview_workbook_batch,
)
from excel_cost_allocator.templates import (
    import_scheme_template,
    serialize_scheme_template,
)


OPERATORS = [
    ("equals", "等于"),
    ("not_equals", "不等于"),
    ("contains", "包含"),
    ("not_contains", "不包含"),
    ("regex", "正则"),
    ("blank", "为空"),
    ("not_blank", "非空"),
]

AMOUNT_MODES = [
    ("target_total", "取分摊列原始合计"),
    ("source_column", "取指定金额列合计"),
    ("manual", "手工输入金额"),
]

HELP_TEXT = """分摊工具使用说明

一、基础设置
1. Excel 文件：选择需要处理的 .xlsx 或 .xlsm 文件，WPS 另存出来的 Excel 文件也可以使用。
2. 工作表：选择需要分摊的 sheet。
3. 表头行：填写列名所在的行号。工具会从这一行读取所有列名。
4. 输出路径：分摊结果会另存为新文件，不会覆盖原文件。

二、分摊方案
一个方案代表一笔费用的一种分摊逻辑。需要分摊多笔费用时，新增多个方案即可。

方案名称：用于在预览、“分摊汇总”和各方案明细页里区分不同费用。
分摊结果列：分摊后的金额写入哪一列。建议一个方案只写一个结果列，方便财务核对。

方案模板：
保存模板：把当前所有方案保存为 .json 文件，适合月度重复使用。
导入模板：读取以前保存的 .json 模板，按当前表头匹配列名并恢复方案。
模板不绑定具体 Excel 文件，只保存规则；如果当前表缺少模板里的列，工具会提示缺少哪些列。

三、分摊金额来源
1. 取分摊列原始合计：把“分摊结果列”原来的合计金额拿来重新分摊。
   例：本期共耗料 consumptioncost 列原来合计 100000，就按条件重新分摊 100000。
2. 取指定金额列合计：金额来自另一列，结果写入分摊结果列。
   例：金额来源列是“待分摊运费”，结果写入“运费分摊”。
3. 手工输入金额：财务直接输入一笔金额，再按占比列和过滤条件分摊。
   例：手工金额 50000，按材料成本 + 人工费占比分摊到“运费分摊”列。

四、占比计算列
可以多选。工具会把所选列的数值相加作为每行分摊基数。
单行占比 = 当前行占比列合计 / 所有参与行占比列合计
单行分摊金额 = 待分摊金额 × 单行占比

五、过滤条件
过滤条件表示“不参与分摊”的行。
例如：车间 等于 销售配货部，则销售配货部不参与占比，也不会被写入分摊金额。
OR：任一条件满足就不参与分摊。
AND：所有条件同时满足才不参与分摊。
条件类型：等于、不等于、包含、不包含、正则、为空、非空。

正则示例：
^销售配货部$                       只匹配销售配货部
.*维修.*                           包含“维修”
^(销售配货部|售后服务部)$           匹配两个部门
^$                                  匹配空白内容

六、预览执行
执行前先点“生成预览”，核对：
1. 每个方案的待分摊金额。
2. 参与行数和不参与行数。
3. 占比基数合计。
4. 分摊结果列是否正确。

七、导出结果
分摊汇总：每个方案一行，显示金额来源、待分摊金额、参与行数、不参与行数、基数合计、分摊后合计和校验差额。
明细_01_方案名称、明细_02_方案名称：每个方案一个独立明细页，显示源行号、是否参与、不参与原因、过滤命中说明、基数列、占比和分摊结果。

八、测试数据样式
日期 | 车间 | 产品名称 | 产品类别 | 完工入库材料成本 wg_materialcost | 本期人工费 laborcost | 本期共耗料 consumptioncost | 运费分摊
2026-06 | 一车间 | A产品 | 成品 | 10000 | 3000 | 0 | 0
2026-06 | 二车间 | B产品 | 成品 | 20000 | 5000 | 0 | 0
2026-06 | 销售配货部 | 配货费用 | 内部 | 5000 | 1000 | 0 | 0
2026-06 | 三车间 | C产品 | 半成品 | 15000 | 4000 | 0 | 0

推荐测试方案：
方案名称：共耗料分摊
金额来源：手工输入金额 100000
占比计算列：完工入库材料成本 wg_materialcost、本期人工费 laborcost
分摊结果列：本期共耗料 consumptioncost
过滤条件：车间 等于 销售配货部
"""


class FilterRuleEditor:
    def __init__(
        self,
        app: "FletAllocatorApp",
        default_column: str = "",
        default_operator: str = "equals",
        default_value: str = "",
    ) -> None:
        self.app = app
        self.column = ft.Dropdown(
            label="字段",
            value=default_column or None,
            options=self.app.header_options(),
            border_radius=12,
            dense=True,
            col={"sm": 12, "md": 4},
        )
        self.operator = ft.Dropdown(
            label="条件",
            value=self.operator_code(default_operator),
            options=[ft.dropdown.Option(code, label) for code, label in OPERATORS],
            border_radius=12,
            dense=True,
            on_change=self._operator_changed,
            col={"sm": 6, "md": 2},
        )
        self.value = ft.TextField(
            label="值",
            value=display_value(default_value) if default_value == "" else default_value,
            hint_text="输入或选择已有值",
            border_radius=12,
            dense=True,
            col={"sm": 10, "md": 4},
        )
        self.value_button = ft.IconButton(
            icon=ft.icons.LIST_ALT,
            tooltip="从字段现有值中选择",
            on_click=lambda _event: self.app.open_unique_value_picker(self),
        )
        self.delete_button = ft.IconButton(
            icon=ft.icons.DELETE_OUTLINE,
            icon_color=ft.colors.RED_300,
            tooltip="删除规则",
            on_click=lambda _event: self.app.remove_filter_rule(self),
        )
        actions = ft.Row(
            [self.value_button, self.delete_button],
            alignment=ft.MainAxisAlignment.END,
            spacing=0,
        )
        self.control = ft.Container(
            content=ft.ResponsiveRow(
                [
                    self.column,
                    self.operator,
                    self.value,
                    ft.Container(content=actions, col={"sm": 2, "md": 2}),
                ],
                spacing=10,
                run_spacing=8,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=12,
            border_radius=14,
            bgcolor="#0F1A2E",
            border=ft.border.all(1, "#24324A"),
        )
        self._update_value_state()

    @staticmethod
    def operator_code(label_or_code: str) -> str:
        for code, label in OPERATORS:
            if label_or_code in {code, label}:
                return code
        return label_or_code or "equals"

    def _operator_changed(self, _event: ft.ControlEvent) -> None:
        self._update_value_state()
        self.app.page.update()

    def _update_value_state(self) -> None:
        disabled = self.operator_code(self.operator.value or "") in {"blank", "not_blank"}
        self.value.disabled = disabled
        self.value_button.disabled = disabled
        if disabled:
            self.value.value = ""

    def to_draft(self) -> Dict[str, str]:
        value = (self.value.value or "").strip()
        if value == "<空白>":
            value = ""
        return {
            "column": self.column.value or "",
            "operator": self.operator_code(self.operator.value or "equals"),
            "value": value,
        }

    def to_rule(self) -> FilterRule:
        draft = self.to_draft()
        column_index = self.app.label_to_column_index(draft["column"])
        if column_index is None:
            raise ValueError("过滤条件中的字段选择无效。")
        if draft["operator"] not in {"blank", "not_blank"} and not draft["value"]:
            raise ValueError("过滤条件中的值不能为空；不需要过滤时请删除这条规则。")
        return FilterRule(
            column=column_index,
            operator=draft["operator"],
            value=draft["value"],
        )


class FletAllocatorApp:
    def __init__(self, page: ft.Page) -> None:
        self.page = page
        self.headers: List[ColumnInfo] = []
        self.schemes: List[Dict[str, Any]] = []
        self.active_scheme_index = -1
        self.filter_rule_editors: List[FilterRuleEditor] = []
        self.unique_value_cache: Dict[int, List[str]] = {}
        self._loading_scheme = False
        self._busy = False
        self._job_lock = threading.Lock()
        self._job_buttons: List[ft.Control] = []
        self._busy_controls: List[ft.Control] = []
        self._configure_page()
        self._build_controls()
        self._build_file_pickers()
        self.page.add(self._build_layout())
        self.page.update()

    def _configure_page(self) -> None:
        self.page.title = "分摊工具 · Modern Excel Cost Allocator"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.theme = ft.Theme(
            color_scheme_seed=ft.colors.BLUE,
            font_family="Microsoft YaHei UI",
        )
        self.page.padding = 0
        self.page.spacing = 0
        self.page.bgcolor = "#080F1F"
        self.page.horizontal_alignment = ft.CrossAxisAlignment.STRETCH
        self.page.window.width = 1120
        self.page.window.height = 700
        self.page.window.min_width = 980
        self.page.window.min_height = 640
        icon_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "assets" / "app.ico"
        if icon_path.exists():
            self.page.window.icon = str(icon_path)

    def _build_controls(self) -> None:
        self.status_text = ft.Text(
            "请选择 Excel 文件。",
            color=ft.colors.BLUE_100,
            expand=True,
            max_lines=2,
        )
        self.progress = ft.ProgressBar(
            visible=False,
            color=ft.colors.BLUE_400,
            bgcolor="#1E293B",
        )
        self.file_path = ft.TextField(
            label="Excel 文件",
            hint_text="选择 .xlsx 或 .xlsm 文件",
            border_radius=12,
            prefix_icon=ft.icons.DESCRIPTION_OUTLINED,
            col={"sm": 12, "md": 10},
        )
        self.output_path = ft.TextField(
            label="输出路径",
            hint_text="结果将另存为新文件",
            border_radius=12,
            prefix_icon=ft.icons.SAVE_OUTLINED,
            col={"sm": 12, "md": 10},
        )
        self.sheet_name = ft.Dropdown(
            label="工作表",
            hint_text="请先选择 Excel 文件",
            border_radius=12,
            dense=True,
            on_change=self._sheet_changed,
            col={"sm": 12, "md": 6},
        )
        self.header_row = ft.TextField(
            label="表头行",
            value="1",
            hint_text="1-999",
            border_radius=12,
            dense=True,
            keyboard_type=ft.KeyboardType.NUMBER,
            on_submit=self._read_headers_clicked,
            col={"sm": 6, "md": 2},
        )
        self.header_summary = ft.Text("尚未读取表头。", color=ft.colors.BLUE_GREY_200)
        self.header_list = ft.ListView(height=270, spacing=6, padding=0)
        self.scheme_list = ft.ListView(height=280, spacing=8, padding=0)
        self.scheme_name = ft.TextField(
            label="方案名称",
            hint_text="例如：共耗料分摊",
            border_radius=12,
            on_blur=self._scheme_name_blurred,
            col={"sm": 12, "md": 6},
        )
        self.allocation_column = ft.Dropdown(
            label="分摊结果列",
            hint_text="选择写入金额的列",
            border_radius=12,
            dense=True,
            col={"sm": 12, "md": 6},
        )
        self.amount_mode = ft.Dropdown(
            label="金额来源",
            value="target_total",
            options=[ft.dropdown.Option(code, label) for code, label in AMOUNT_MODES],
            border_radius=12,
            dense=True,
            on_change=self._amount_mode_changed,
            col={"sm": 12, "md": 6},
        )
        self.amount_column = ft.Dropdown(
            label="金额来源列",
            hint_text="选择金额合计来源列",
            border_radius=12,
            dense=True,
            disabled=True,
            col={"sm": 12, "md": 6},
        )
        self.manual_amount = ft.TextField(
            label="手工金额",
            hint_text="例如：100000.00",
            border_radius=12,
            dense=True,
            keyboard_type=ft.KeyboardType.NUMBER,
            disabled=True,
            col={"sm": 12, "md": 6},
        )
        self.base_column_list = ft.ListView(height=225, spacing=2, padding=0)
        self.filter_logic = ft.Dropdown(
            label="规则关系",
            value="OR",
            options=[ft.dropdown.Option("OR"), ft.dropdown.Option("AND")],
            border_radius=12,
            dense=True,
            width=145,
        )
        self.filter_rule_column = ft.Column(spacing=8)
        self.filter_preview_summary = ft.Text("尚未测试规则。", color=ft.colors.BLUE_GREY_200)
        self.filter_preview_list = ft.ListView(height=180, spacing=6, padding=0)
        self.preview_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("方案名称")),
                ft.DataColumn(ft.Text("金额来源")),
                ft.DataColumn(ft.Text("分摊列")),
                ft.DataColumn(ft.Text("待分摊金额"), numeric=True),
                ft.DataColumn(ft.Text("参与行"), numeric=True),
                ft.DataColumn(ft.Text("不参与行"), numeric=True),
                ft.DataColumn(ft.Text("基数合计"), numeric=True),
                ft.DataColumn(ft.Text("分摊后合计"), numeric=True),
            ],
            rows=[],
            border=ft.border.all(1, "#25334A"),
            border_radius=12,
            heading_row_color="#16233B",
            data_row_color={ft.ControlState.HOVERED: "#132038"},
            column_spacing=28,
            horizontal_margin=18,
        )

    def _build_file_pickers(self) -> None:
        self.input_picker = ft.FilePicker(on_result=self._input_file_selected)
        self.output_picker = ft.FilePicker(on_result=self._output_file_selected)
        self.sample_picker = ft.FilePicker(on_result=self._sample_path_selected)
        self.template_save_picker = ft.FilePicker(on_result=self._template_save_path_selected)
        self.template_open_picker = ft.FilePicker(on_result=self._template_file_selected)
        self.page.overlay.extend(
            [
                self.input_picker,
                self.output_picker,
                self.sample_picker,
                self.template_save_picker,
                self.template_open_picker,
            ]
        )

    def _build_layout(self) -> ft.Control:
        self.views = ft.Column(
            [
                self._build_base_tab(),
                self._build_scheme_tab(),
                self._build_preview_tab(),
                self._build_help_tab(),
            ],
            expand=True,
            spacing=0,
        )
        for index, view in enumerate(self.views.controls):
            view.visible = index == 0
        self.navigation = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=104,
            bgcolor="#0D1729",
            indicator_color="#214A80",
            destinations=[
                ft.NavigationRailDestination(
                    icon=ft.icons.SETTINGS_OUTLINED,
                    selected_icon=ft.icons.SETTINGS,
                    label="基础设置",
                ),
                ft.NavigationRailDestination(
                    icon=ft.icons.ACCOUNT_TREE_OUTLINED,
                    selected_icon=ft.icons.ACCOUNT_TREE,
                    label="分摊方案",
                ),
                ft.NavigationRailDestination(
                    icon=ft.icons.FACT_CHECK_OUTLINED,
                    selected_icon=ft.icons.FACT_CHECK,
                    label="预览执行",
                ),
                ft.NavigationRailDestination(
                    icon=ft.icons.HELP_OUTLINE,
                    selected_icon=ft.icons.HELP,
                    label="使用说明",
                ),
            ],
            on_change=self._navigation_changed,
        )
        self.workspace = ft.Row(
            [
                self.navigation,
                ft.VerticalDivider(width=1, color="#26344A"),
                self.views,
            ],
            spacing=0,
            expand=True,
            vertical_alignment=ft.CrossAxisAlignment.STRETCH,
        )
        self._busy_controls.extend(
            [
                self.navigation,
                self.file_path,
                self.output_path,
                self.sheet_name,
                self.header_row,
                self.scheme_list,
                self.scheme_name,
                self.allocation_column,
                self.amount_mode,
                self.amount_column,
                self.manual_amount,
                self.filter_logic,
            ]
        )
        return ft.Column(
            [self._build_hero(), self.workspace, self._build_footer()],
            expand=True,
            spacing=0,
        )

    def _navigation_changed(self, event: ft.ControlEvent) -> None:
        selected = int(event.control.selected_index or 0)
        self._show_view(selected)

    def _show_view(self, index: int) -> None:
        self.navigation.selected_index = index
        for view_index, view in enumerate(self.views.controls):
            view.visible = view_index == index
        self.page.update()

    def _build_hero(self) -> ft.Control:
        return ft.Container(
            content=ft.Row(
                [
                    ft.Container(
                        content=ft.Icon(ft.icons.AUTO_GRAPH, size=28, color=ft.colors.BLUE_200),
                        width=52,
                        height=52,
                        alignment=ft.alignment.center,
                        border_radius=16,
                        bgcolor="#18345F",
                    ),
                    ft.Column(
                        [
                            ft.Text("分摊工具", size=24, weight=ft.FontWeight.BOLD, color=ft.colors.WHITE),
                            ft.Text(
                                "按方案配置费用来源、占比基数和过滤条件，生成可核对的 Excel/WPS 分摊结果",
                                color=ft.colors.BLUE_100,
                            ),
                        ],
                        spacing=2,
                        expand=True,
                    ),
                ],
                spacing=14,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            padding=ft.padding.symmetric(horizontal=24, vertical=16),
            gradient=ft.LinearGradient(
                begin=ft.alignment.center_left,
                end=ft.alignment.center_right,
                colors=["#102A56", "#173F78", "#102A56"],
            ),
        )

    def _build_base_tab(self) -> ft.Control:
        browse_button = ft.FilledButton(
            text="浏览文件",
            icon=ft.icons.FOLDER_OPEN,
            height=48,
            on_click=self._open_input_picker,
        )
        output_button = ft.ElevatedButton(
            text="另存为",
            icon=ft.icons.SAVE_AS,
            height=48,
            on_click=self._open_output_picker,
        )
        decrease_header_button = ft.IconButton(
            icon=ft.icons.REMOVE,
            tooltip="表头行减 1",
            on_click=lambda _event: self._step_header_row(-1),
        )
        increase_header_button = ft.IconButton(
            icon=ft.icons.ADD,
            tooltip="表头行加 1",
            on_click=lambda _event: self._step_header_row(1),
        )
        read_button = ft.ElevatedButton(
            text="读取表头",
            icon=ft.icons.TABLE_VIEW,
            height=48,
            on_click=self._read_headers_clicked,
        )
        sample_button = ft.ElevatedButton(
            text="生成测试数据",
            icon=ft.icons.SCIENCE_OUTLINED,
            on_click=self._open_sample_picker,
        )
        self._job_buttons.extend(
            [
                browse_button,
                output_button,
                decrease_header_button,
                increase_header_button,
                read_button,
                sample_button,
            ]
        )
        file_card = self._card(
            "文件与表头",
            ft.Column(
                [
                    ft.ResponsiveRow(
                        [
                            self.file_path,
                            ft.Container(content=browse_button, col={"sm": 12, "md": 2}),
                        ],
                        spacing=12,
                        run_spacing=10,
                    ),
                    ft.ResponsiveRow(
                        [
                            self.sheet_name,
                            ft.Container(
                                content=ft.Row(
                                    [decrease_header_button, self.header_row, increase_header_button],
                                    spacing=4,
                                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                                ),
                                col={"sm": 6, "md": 4},
                            ),
                            ft.Container(content=read_button, col={"sm": 6, "md": 2}),
                        ],
                        spacing=12,
                        run_spacing=10,
                    ),
                    ft.ResponsiveRow(
                        [
                            self.output_path,
                            ft.Container(content=output_button, col={"sm": 12, "md": 2}),
                        ],
                        spacing=12,
                        run_spacing=10,
                    ),
                ],
                spacing=14,
            ),
            ft.icons.FOLDER_COPY_OUTLINED,
        )
        headers_card = self._card(
            "读取结果",
            ft.Column(
                [
                    self.header_summary,
                    ft.Divider(height=1, color="#26344A"),
                    self.header_list,
                ],
                spacing=10,
            ),
            ft.icons.VIEW_LIST_OUTLINED,
        )
        sample_card = self._card(
            "测试数据",
            ft.Row(
                [
                    ft.Text(
                        "需要先试流程时，可生成示例 Excel 验证表头读取、过滤、预览和导出。",
                        color=ft.colors.BLUE_GREY_100,
                        expand=True,
                    ),
                    sample_button,
                ],
                spacing=12,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            ft.icons.SCIENCE_OUTLINED,
        )
        return ft.Container(
            content=ft.Column([file_card, headers_card, sample_card], spacing=14, scroll=ft.ScrollMode.AUTO),
            padding=18,
            expand=True,
        )

    def _build_scheme_tab(self) -> ft.Control:
        add_button = ft.FilledButton(text="新增", icon=ft.icons.ADD, on_click=self._add_scheme)
        copy_button = ft.ElevatedButton(text="复制", icon=ft.icons.CONTENT_COPY, on_click=self._copy_scheme)
        delete_button = ft.ElevatedButton(
            text="删除",
            icon=ft.icons.DELETE_OUTLINE,
            color=ft.colors.RED_200,
            on_click=self._delete_scheme,
        )
        import_button = ft.ElevatedButton(
            text="导入模板",
            icon=ft.icons.UPLOAD_FILE,
            on_click=self._open_template_picker,
        )
        save_button = ft.ElevatedButton(
            text="保存模板",
            icon=ft.icons.DOWNLOAD,
            on_click=self._open_template_save_picker,
        )
        self._job_buttons.extend([import_button, save_button])
        self._job_buttons.extend([add_button, copy_button, delete_button])
        scheme_panel = self._card(
            "方案列表",
            ft.Column(
                [
                    self.scheme_list,
                    ft.Row([add_button, copy_button], wrap=True, spacing=8, run_spacing=8),
                    ft.Row([delete_button], wrap=True),
                    ft.Divider(height=1, color="#26344A"),
                    ft.Column([import_button, save_button], spacing=8, horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
                ],
                spacing=10,
            ),
            ft.icons.FORMAT_LIST_NUMBERED,
        )
        scheme_panel.col = {"sm": 12, "lg": 3}
        editor = ft.Column(
            [
                self._card(
                    "方案基础信息",
                    ft.ResponsiveRow([self.scheme_name, self.allocation_column], spacing=12, run_spacing=10),
                    ft.icons.EDIT_NOTE,
                ),
                self._build_amount_card(),
                self._build_base_columns_card(),
                self._build_filter_card(),
            ],
            spacing=14,
        )
        editor_panel = ft.Container(content=editor, col={"sm": 12, "lg": 9})
        return ft.Container(
            content=ft.Column(
                [
                    ft.ResponsiveRow(
                        [scheme_panel, editor_panel],
                        spacing=14,
                        run_spacing=14,
                        vertical_alignment=ft.CrossAxisAlignment.START,
                    )
                ],
                scroll=ft.ScrollMode.AUTO,
            ),
            padding=18,
            expand=True,
        )

    def _build_amount_card(self) -> ft.Control:
        return self._card(
            "分摊金额来源",
            ft.Column(
                [
                    ft.ResponsiveRow([self.amount_mode, self.amount_column], spacing=12, run_spacing=10),
                    ft.ResponsiveRow(
                        [
                            self.manual_amount,
                            ft.Container(
                                content=ft.Text(
                                    "手工金额适合财务单独给出的一笔费用；建议一个方案只写入一个分摊结果列。",
                                    color=ft.colors.BLUE_GREY_300,
                                ),
                                padding=ft.padding.only(top=8),
                                col={"sm": 12, "md": 6},
                            ),
                        ],
                        spacing=12,
                        run_spacing=8,
                    ),
                ],
                spacing=10,
            ),
            ft.icons.PAID_OUTLINED,
        )

    def _build_base_columns_card(self) -> ft.Control:
        smart_button = ft.ElevatedButton(
            text="智能选择",
            icon=ft.icons.AUTO_FIX_HIGH,
            on_click=self._auto_select_base_columns,
        )
        clear_button = ft.ElevatedButton(
            text="清空选择",
            icon=ft.icons.CLEAR_ALL,
            on_click=self._clear_base_columns,
        )
        self._job_buttons.extend([smart_button, clear_button])
        return self._card(
            "占比计算列",
            ft.Column(
                [
                    ft.Text(
                        "可多选，例如：完工入库材料成本 + 本期人工费。",
                        color=ft.colors.BLUE_GREY_300,
                    ),
                    ft.Container(
                        content=self.base_column_list,
                        padding=10,
                        border_radius=12,
                        bgcolor="#0A1426",
                        border=ft.border.all(1, "#24324A"),
                    ),
                    ft.Row([smart_button, clear_button], wrap=True, spacing=8, run_spacing=8),
                ],
                spacing=10,
            ),
            ft.icons.CALCULATE_OUTLINED,
        )

    def _build_filter_card(self) -> ft.Control:
        add_rule_button = ft.ElevatedButton(
            text="添加规则",
            icon=ft.icons.ADD,
            on_click=self._add_filter_rule_clicked,
        )
        test_button = ft.FilledButton(
            text="测试命中",
            icon=ft.icons.FILTER_ALT,
            on_click=self._test_filter_rules,
        )
        self._job_buttons.append(test_button)
        self._job_buttons.append(add_rule_button)
        rule_panel = ft.Container(
            content=ft.Column(
                [
                    self.filter_rule_column,
                    ft.Text(
                        "正则示例：^销售配货部$ 精确匹配；.*维修.* 包含维修；^(销售配货部|售后服务部)$ 多值匹配；^$ 匹配空白。",
                        size=12,
                        color=ft.colors.BLUE_GREY_300,
                    ),
                ],
                spacing=10,
            ),
            col={"sm": 12, "xl": 8},
        )
        preview_panel = ft.Container(
            content=ft.Column(
                [
                    ft.Text("过滤命中预览", weight=ft.FontWeight.BOLD),
                    self.filter_preview_summary,
                    ft.Divider(height=1, color="#26344A"),
                    self.filter_preview_list,
                ],
                spacing=8,
            ),
            padding=14,
            border_radius=14,
            bgcolor="#0A1426",
            border=ft.border.all(1, "#24324A"),
            col={"sm": 12, "xl": 4},
        )
        return self._card(
            "不参与分摊的过滤条件",
            ft.Column(
                [
                    ft.Row(
                        [
                            self.filter_logic,
                            add_rule_button,
                            test_button,
                            ft.Text(
                                "OR=任一条件命中即排除；AND=全部条件满足才排除",
                                color=ft.colors.BLUE_GREY_300,
                            ),
                        ],
                        wrap=True,
                        spacing=10,
                        run_spacing=8,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    ft.ResponsiveRow(
                        [rule_panel, preview_panel],
                        spacing=12,
                        run_spacing=12,
                        vertical_alignment=ft.CrossAxisAlignment.START,
                    ),
                ],
                spacing=12,
            ),
            ft.icons.FILTER_ALT_OUTLINED,
        )

    def _build_preview_tab(self) -> ft.Control:
        preview_button = ft.ElevatedButton(
            text="生成预览",
            icon=ft.icons.PREVIEW,
            on_click=self._preview_allocation,
        )
        run_button = ft.FilledButton(
            text="开始分摊",
            icon=ft.icons.PLAY_ARROW,
            on_click=self._run_allocation,
        )
        self.preview_buttons = [preview_button]
        self.run_buttons = [run_button]
        self._job_buttons.extend([preview_button, run_button])
        toolbar = self._card(
            "执行前核对",
            ft.Row(
                [
                    ft.Text(
                        "先生成预览，确认每个方案的金额、参与行数、基数合计和分摊列，再执行导出。",
                        color=ft.colors.BLUE_GREY_100,
                        expand=True,
                    ),
                    preview_button,
                    run_button,
                ],
                spacing=10,
            ),
            ft.icons.VERIFIED_OUTLINED,
        )
        table_card = self._card(
            "方案预览",
            ft.Column(
                [
                    ft.Row([self.preview_table], scroll=ft.ScrollMode.AUTO),
                    ft.Container(
                        content=ft.Text(
                            "预览会检查：是否有重复分摊列、手工金额是否填写、占比基数是否大于 0、过滤条件是否可用。\n"
                            "如果某个方案待分摊金额不为 0，但参与行基数合计为 0，执行时会阻止导出。",
                            selectable=True,
                            color=ft.colors.BLUE_GREY_200,
                        ),
                        padding=14,
                        border_radius=12,
                        bgcolor="#0A1426",
                        border=ft.border.all(1, "#24324A"),
                    ),
                ],
                spacing=14,
            ),
            ft.icons.TABLE_CHART_OUTLINED,
        )
        return ft.Container(
            content=ft.Column([toolbar, table_card], spacing=14, scroll=ft.ScrollMode.AUTO),
            padding=18,
            expand=True,
        )

    def _build_help_tab(self) -> ft.Control:
        support = ft.Container(
            content=ft.Column(
                [
                    ft.Text("技术支持", size=20, weight=ft.FontWeight.BOLD, color=ft.colors.BLUE_200),
                    ft.Text("A0金蝶软件王朝", size=16),
                    ft.Text("电话：15939121371（微信同号）"),
                    ft.Text("河南 焦作"),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=4,
            ),
            padding=18,
            border_radius=16,
            bgcolor="#102344",
            border=ft.border.all(1, "#28518D"),
            alignment=ft.alignment.center,
        )
        help_card = self._card(
            "使用说明",
            ft.Column([support, ft.Divider(color="#26344A"), ft.Text(HELP_TEXT, selectable=True, size=14)], spacing=18),
            ft.icons.MENU_BOOK_OUTLINED,
        )
        return ft.Container(
            content=ft.Column([help_card], scroll=ft.ScrollMode.AUTO),
            padding=18,
            expand=True,
        )

    def _build_footer(self) -> ft.Control:
        preview_button = ft.ElevatedButton(
            text="生成预览",
            icon=ft.icons.PREVIEW,
            on_click=self._preview_allocation,
        )
        run_button = ft.FilledButton(
            text="开始分摊",
            icon=ft.icons.PLAY_ARROW,
            on_click=self._run_allocation,
        )
        self.preview_buttons.append(preview_button)
        self.run_buttons.append(run_button)
        self._job_buttons.extend([preview_button, run_button])
        return ft.Container(
            content=ft.Column(
                [
                    self.progress,
                    ft.Row(
                        [
                            ft.Icon(ft.icons.INFO_OUTLINE, color=ft.colors.BLUE_200),
                            self.status_text,
                            preview_button,
                            run_button,
                        ],
                        spacing=10,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                ],
                spacing=8,
            ),
            padding=ft.padding.symmetric(horizontal=20, vertical=12),
            bgcolor="#101C31",
            border=ft.border.only(top=ft.BorderSide(1, "#26344A")),
        )

    def _card(self, title: str, content: ft.Control, icon: str) -> ft.Container:
        return ft.Container(
            content=ft.Column(
                [
                    ft.Row(
                        [
                            ft.Icon(icon, size=20, color=ft.colors.BLUE_300),
                            ft.Text(title, size=16, weight=ft.FontWeight.BOLD),
                        ],
                        spacing=8,
                    ),
                    ft.Divider(height=1, color="#26344A"),
                    content,
                ],
                spacing=12,
            ),
            padding=18,
            border_radius=16,
            bgcolor="#111B2D",
            border=ft.border.all(1, "#22314A"),
        )

    def _open_input_picker(self, _event: ft.ControlEvent) -> None:
        if self._job_in_progress():
            return
        self.input_picker.pick_files(
            dialog_title="选择 Excel 文件",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xlsm"],
            allow_multiple=False,
        )

    def _input_file_selected(self, event: ft.FilePickerResultEvent) -> None:
        if not event.files or not event.files[0].path:
            return
        path = event.files[0].path

        def task() -> Tuple[List[str], str, int, List[ColumnInfo]]:
            sheets = get_sheet_names(path)
            if not sheets:
                raise ValueError("Excel 文件中没有可用的工作表。")
            sheet = sheets[0]
            row_number = guess_header_row(path, sheet)
            headers = get_headers(path, sheet, row_number)
            return sheets, sheet, row_number, headers

        def success(result: Tuple[List[str], str, int, List[ColumnInfo]]) -> None:
            sheets, sheet, row_number, headers = result
            self.file_path.value = path
            self._set_default_output(path)
            self.sheet_name.options = [ft.dropdown.Option(item) for item in sheets]
            self.sheet_name.value = sheet
            self.header_row.value = str(row_number)
            self._apply_headers(headers)
            self.status_text.value = "文件读取完成，请到“分摊方案”页设置方案。"

        self._start_job("正在读取 Excel 文件...", task, success, "读取失败")

    def _set_default_output(self, path: str) -> None:
        input_path = Path(path)
        suffix = ".xlsm" if input_path.suffix.lower() == ".xlsm" else ".xlsx"
        self.output_path.value = str(input_path.with_name("{}_分摊结果{}".format(input_path.stem, suffix)))

    def _open_output_picker(self, _event: ft.ControlEvent) -> None:
        if self._job_in_progress():
            return
        initial = self.output_path.value or self.file_path.value or ""
        self.output_picker.save_file(
            dialog_title="保存分摊结果",
            file_name=Path(initial).name if initial else "分摊结果.xlsx",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xlsm"],
        )

    def _output_file_selected(self, event: ft.FilePickerResultEvent) -> None:
        if not event.path or self._job_in_progress():
            return
        initial = self.output_path.value or self.file_path.value or ""
        suffix = ".xlsm" if initial.lower().endswith(".xlsm") else ".xlsx"
        self.output_path.value = self._ensure_extension(event.path, (".xlsx", ".xlsm"), suffix)
        self.page.update()

    def _step_header_row(self, delta: int) -> None:
        try:
            current = self._parse_header_row()
        except ValueError:
            current = 1
        self.header_row.value = str(min(999, max(1, current + delta)))
        self.page.update()
        if self.file_path.value and self.sheet_name.value:
            self._read_headers_clicked(None)

    def _sheet_changed(self, _event: ft.ControlEvent) -> None:
        path = self.file_path.value or ""
        sheet = self.sheet_name.value or ""
        if not path or not sheet:
            return

        def task() -> Tuple[int, List[ColumnInfo]]:
            row_number = guess_header_row(path, sheet)
            return row_number, get_headers(path, sheet, row_number)

        def success(result: Tuple[int, List[ColumnInfo]]) -> None:
            row_number, headers = result
            self.header_row.value = str(row_number)
            self._apply_headers(headers)
            self.status_text.value = "工作表读取完成，请检查分摊方案。"

        self._start_job("正在读取工作表...", task, success, "读取失败")

    def _read_headers_clicked(self, _event: Optional[ft.ControlEvent]) -> None:
        try:
            path = (self.file_path.value or "").strip()
            sheet = self.sheet_name.value or ""
            row_number = self._parse_header_row()
            if not path:
                raise ValueError("请选择 Excel 文件。")
            if not sheet:
                raise ValueError("请选择工作表。")
        except Exception as exc:
            self._show_error("配置不完整", exc, traceback.format_exc())
            return

        def task() -> List[ColumnInfo]:
            return get_headers(path, sheet, row_number)

        def success(headers: List[ColumnInfo]) -> None:
            self._apply_headers(headers)
            self.status_text.value = "已读取 {} 个表头。".format(len(headers))

        self._start_job("正在读取表头...", task, success, "读取失败")

    def _apply_headers(self, headers: List[ColumnInfo]) -> None:
        self.headers = headers
        self.unique_value_cache.clear()
        self.header_summary.value = "已读取 {} 个表头。".format(len(headers))
        self.header_list.controls = [
            ft.Container(
                content=ft.Text(item.label, selectable=True),
                padding=ft.padding.symmetric(horizontal=12, vertical=8),
                border_radius=10,
                bgcolor="#0A1426" if index % 2 == 0 else "#101C31",
            )
            for index, item in enumerate(headers)
        ]
        options = self.header_options()
        self.allocation_column.options = options
        self.amount_column.options = self.header_options()
        self.base_column_list.controls = [ft.Checkbox(label=item.label, value=False) for item in headers]
        self._reset_schemes_for_headers()

    def _open_sample_picker(self, _event: ft.ControlEvent) -> None:
        if self._job_in_progress():
            return
        self.sample_picker.save_file(
            dialog_title="保存测试数据",
            file_name="分摊工具测试数据.xlsx",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
        )

    def _sample_path_selected(self, event: ft.FilePickerResultEvent) -> None:
        if not event.path or self._job_in_progress():
            return
        path = self._ensure_extension(event.path, (".xlsx",), ".xlsx")

        def task() -> str:
            create_sample_workbook(path)
            return path

        def success(created_path: str) -> None:
            self.status_text.value = "测试数据已生成：{}".format(created_path)
            self._show_message("已生成", "测试数据已生成：\n{}".format(created_path))

        self._start_job("正在生成测试数据...", task, success, "生成失败")

    def header_options(self) -> List[ft.dropdown.Option]:
        return [ft.dropdown.Option(item.label) for item in self.headers]

    def header_labels(self) -> List[str]:
        return [item.label for item in self.headers]

    def label_to_column_index(self, label: str) -> Optional[int]:
        for item in self.headers:
            if item.label == label:
                return item.index
        return None

    def column_label_short(self, index: int) -> str:
        for item in self.headers:
            if item.index == index:
                return item.label
        return "{}列".format(index)

    def _reset_schemes_for_headers(self) -> None:
        self.schemes = []
        self.active_scheme_index = -1
        if self.headers:
            self.schemes.append(self._default_scheme("共耗料分摊"))
            self._load_scheme(0, update=False)
        self._refresh_scheme_list()

    def _default_scheme(self, name: str) -> Dict[str, Any]:
        target = self.suggest_target_column()
        filter_column = self.suggest_filter_column()
        rules = []
        if filter_column:
            rules.append({"column": filter_column, "operator": "equals", "value": "销售配货部"})
        return {
            "name": name,
            "amount_mode": "target_total",
            "amount_column": target,
            "manual_amount": "",
            "allocation_column": target,
            "base_columns": self.suggest_base_columns(),
            "filter_logic": "OR",
            "filter_rules": rules,
        }

    def _refresh_scheme_list(self) -> None:
        controls = []
        for index, scheme in enumerate(self.schemes):
            active = index == self.active_scheme_index
            controls.append(
                ft.Container(
                    content=ft.Row(
                        [
                            ft.Icon(
                                ft.icons.RADIO_BUTTON_CHECKED if active else ft.icons.RADIO_BUTTON_UNCHECKED,
                                size=18,
                                color=ft.colors.BLUE_300 if active else ft.colors.BLUE_GREY_400,
                            ),
                            ft.Text(
                                "{}. {}".format(index + 1, scheme.get("name") or "未命名方案"),
                                weight=ft.FontWeight.BOLD if active else ft.FontWeight.NORMAL,
                                expand=True,
                            ),
                        ],
                        spacing=8,
                    ),
                    padding=12,
                    border_radius=12,
                    bgcolor="#18365F" if active else "#0A1426",
                    border=ft.border.all(1, "#2E65AA" if active else "#24324A"),
                    ink=True,
                    on_click=lambda _event, item_index=index: self._select_scheme(item_index),
                )
            )
        self.scheme_list.controls = controls

    def _select_scheme(self, index: int) -> None:
        if self._job_in_progress():
            return
        if index == self.active_scheme_index:
            return
        self._save_current_scheme()
        self._load_scheme(index)

    def _load_scheme(self, index: int, update: bool = True) -> None:
        if index < 0 or index >= len(self.schemes):
            return
        self._loading_scheme = True
        try:
            self.active_scheme_index = index
            scheme = self.schemes[index]
            self.scheme_name.value = scheme.get("name", "")
            self.amount_mode.value = scheme.get("amount_mode", "target_total")
            self.amount_column.value = scheme.get("amount_column") or None
            self.manual_amount.value = scheme.get("manual_amount", "")
            self.allocation_column.value = scheme.get("allocation_column") or None
            self.filter_logic.value = scheme.get("filter_logic", "OR")
            selected = set(scheme.get("base_columns", []))
            for checkbox in self.base_column_list.controls:
                checkbox.value = checkbox.label in selected
            self._clear_filter_rules(update=False)
            for rule in scheme.get("filter_rules", []):
                self._add_filter_rule(
                    default_column=rule.get("column") or self.suggest_filter_column(),
                    default_operator=rule.get("operator", "equals"),
                    default_value=rule.get("value", ""),
                    update=False,
                )
            self.filter_preview_summary.value = "尚未测试规则。"
            self.filter_preview_list.controls = []
            self._update_amount_mode_state()
            self._refresh_scheme_list()
        finally:
            self._loading_scheme = False
        if update:
            self.page.update()

    def _save_current_scheme(self) -> None:
        if self._loading_scheme or not (0 <= self.active_scheme_index < len(self.schemes)):
            return
        scheme = self.schemes[self.active_scheme_index]
        scheme["name"] = (self.scheme_name.value or "").strip() or "方案{}".format(self.active_scheme_index + 1)
        scheme["amount_mode"] = self.amount_mode.value or "target_total"
        scheme["amount_column"] = self.amount_column.value or ""
        scheme["manual_amount"] = (self.manual_amount.value or "").strip()
        scheme["allocation_column"] = self.allocation_column.value or ""
        scheme["base_columns"] = [
            checkbox.label for checkbox in self.base_column_list.controls if checkbox.value
        ]
        scheme["filter_logic"] = self.filter_logic.value or "OR"
        scheme["filter_rules"] = [editor.to_draft() for editor in self.filter_rule_editors]

    def _scheme_name_blurred(self, _event: ft.ControlEvent) -> None:
        self._save_current_scheme()
        self._refresh_scheme_list()
        self.page.update()

    def _add_scheme(self, _event: ft.ControlEvent) -> None:
        if not self.headers:
            self._show_message("缺少表头", "请先在“基础设置”页读取表头。")
            return
        self._save_current_scheme()
        self.schemes.append(self._default_scheme("方案{}".format(len(self.schemes) + 1)))
        self._load_scheme(len(self.schemes) - 1)

    def _copy_scheme(self, _event: ft.ControlEvent) -> None:
        if self.active_scheme_index < 0:
            return
        self._save_current_scheme()
        source = self.schemes[self.active_scheme_index]
        copied = dict(source)
        copied["filter_rules"] = [dict(rule) for rule in source.get("filter_rules", [])]
        copied["base_columns"] = list(source.get("base_columns", []))
        copied["name"] = "{} 副本".format(source.get("name", "方案"))
        self.schemes.append(copied)
        self._load_scheme(len(self.schemes) - 1)

    def _delete_scheme(self, _event: ft.ControlEvent) -> None:
        if self.active_scheme_index < 0:
            return
        if len(self.schemes) <= 1:
            self._show_message("提示", "至少保留一个分摊方案。")
            return
        index = self.active_scheme_index
        del self.schemes[index]
        self.active_scheme_index = -1
        self._load_scheme(min(index, len(self.schemes) - 1))

    def _amount_mode_changed(self, _event: ft.ControlEvent) -> None:
        self._update_amount_mode_state()
        self.page.update()

    def _update_amount_mode_state(self) -> None:
        mode = self.amount_mode.value or "target_total"
        self.amount_column.disabled = mode != "source_column"
        self.manual_amount.disabled = mode != "manual"

    def suggest_base_columns(self) -> List[str]:
        keywords = ("完工入库材料成本", "本期人工费", "材料成本", "人工费")
        selected = [
            header.label
            for header in self.headers
            if any(keyword in header.header for keyword in keywords)
        ]
        return selected or ([self.headers[0].label] if self.headers else [])

    def suggest_target_column(self) -> str:
        keywords = ("共耗料", "水电费", "维修费", "折旧", "租赁费", "运费")
        for header in self.headers:
            if any(keyword in header.header for keyword in keywords):
                return header.label
        return self.headers[-1].label if self.headers else ""

    def suggest_filter_column(self) -> str:
        keywords = ("生产车间", "车间", "部门")
        for header in self.headers:
            if any(keyword in header.header for keyword in keywords):
                return header.label
        return self.headers[0].label if self.headers else ""

    def _auto_select_base_columns(self, _event: ft.ControlEvent) -> None:
        selected = set(self.suggest_base_columns())
        for checkbox in self.base_column_list.controls:
            checkbox.value = checkbox.label in selected
        self.page.update()

    def _clear_base_columns(self, _event: ft.ControlEvent) -> None:
        for checkbox in self.base_column_list.controls:
            checkbox.value = False
        self.page.update()

    def _add_filter_rule_clicked(self, _event: ft.ControlEvent) -> None:
        if not self.headers:
            self._show_message("缺少表头", "请先读取表头。")
            return
        self._add_filter_rule(default_column=self.suggest_filter_column())

    def _add_filter_rule(
        self,
        default_column: str = "",
        default_operator: str = "equals",
        default_value: str = "",
        update: bool = True,
    ) -> None:
        editor = FilterRuleEditor(self, default_column, default_operator, default_value)
        self.filter_rule_editors.append(editor)
        self.filter_rule_column.controls.append(editor.control)
        if update:
            self.page.update()

    def remove_filter_rule(self, editor: FilterRuleEditor) -> None:
        if editor in self.filter_rule_editors:
            index = self.filter_rule_editors.index(editor)
            del self.filter_rule_editors[index]
            del self.filter_rule_column.controls[index]
            self.page.update()

    def _clear_filter_rules(self, update: bool = True) -> None:
        self.filter_rule_editors = []
        self.filter_rule_column.controls = []
        if update:
            self.page.update()

    def open_unique_value_picker(self, editor: FilterRuleEditor) -> None:
        column_index = self.label_to_column_index(editor.column.value or "")
        if column_index is None:
            self._show_message("字段未选择", "请先选择过滤字段。")
            return
        if column_index in self.unique_value_cache:
            self._show_unique_values(editor, self.unique_value_cache[column_index])
            return
        try:
            path = (self.file_path.value or "").strip()
            sheet = self.sheet_name.value or ""
            row_number = self._parse_header_row()
            if not path or not sheet:
                raise ValueError("请先选择 Excel 文件和工作表。")
        except Exception as exc:
            self._show_error("读取失败", exc, traceback.format_exc())
            return

        def task() -> List[str]:
            return get_unique_values(path, sheet, row_number, column_index)

        def success(values: List[str]) -> None:
            self.unique_value_cache[column_index] = values
            self.status_text.value = "已读取 {} 个字段值。".format(len(values))
            self._show_unique_values(editor, values)

        self._start_job("正在读取字段现有值...", task, success, "读取失败")

    def _show_unique_values(self, editor: FilterRuleEditor, values: Sequence[str]) -> None:
        result_list = ft.ListView(expand=True, spacing=4)
        summary = ft.Text(color=ft.colors.BLUE_GREY_300)
        search = ft.TextField(
            label="搜索",
            hint_text="输入关键字筛选",
            prefix_icon=ft.icons.SEARCH,
            border_radius=12,
        )
        dialog: ft.AlertDialog

        def close_dialog(_event: Optional[ft.ControlEvent] = None) -> None:
            dialog.open = False
            self.page.update()

        def choose_value(raw_value: str) -> None:
            editor.value.value = display_value(raw_value)
            close_dialog()

        def refresh(_event: Optional[ft.ControlEvent] = None) -> None:
            query = (search.value or "").strip().lower()
            filtered = [item for item in values if query in display_value(item).lower()]
            shown = filtered[:200]
            result_list.controls = [
                ft.TextButton(
                    text=display_value(item),
                    on_click=lambda _click, raw=item: choose_value(raw),
                )
                for item in shown
            ]
            summary.value = "匹配 {} 项，当前显示前 {} 项。".format(len(filtered), len(shown))
            self.page.update()

        search.on_change = refresh
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("选择字段值"),
            content=ft.Container(
                width=560,
                height=460,
                content=ft.Column([search, summary, result_list], spacing=10),
            ),
            actions=[ft.TextButton(text="取消", on_click=close_dialog)],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        refresh()
        self.page.open(dialog)
        self.page.update()

    def _test_filter_rules(self, _event: ft.ControlEvent) -> None:
        try:
            rules = [editor.to_rule() for editor in self.filter_rule_editors]
            path = (self.file_path.value or "").strip()
            sheet = self.sheet_name.value or ""
            row_number = self._parse_header_row()
            logic = self.filter_logic.value or "OR"
            if not path or not sheet:
                raise ValueError("请先选择 Excel 文件和工作表。")
        except Exception as exc:
            self._show_error("测试失败", exc, traceback.format_exc())
            return

        def task() -> Tuple[int, List[Tuple[int, str]]]:
            return preview_filter_matches(path, sheet, row_number, rules, logic)

        def success(result: Tuple[int, List[Tuple[int, str]]]) -> None:
            count, samples = result
            self.filter_preview_summary.value = "命中 {} 行，显示前 {} 条。".format(count, len(samples))
            self.filter_preview_list.controls = [
                ft.Container(
                    content=ft.Text("第 {} 行 | {}".format(row_no, reason), selectable=True),
                    padding=8,
                    border_radius=8,
                    bgcolor="#101C31",
                )
                for row_no, reason in samples
            ]
            self.status_text.value = "过滤条件测试完成。"

        self._start_job("正在测试过滤规则...", task, success, "测试失败")

    def _open_template_save_picker(self, _event: ft.ControlEvent) -> None:
        if self._job_in_progress():
            return
        if not self.headers:
            self._show_message("缺少表头", "请先读取表头，再保存方案模板。")
            return
        self._save_current_scheme()
        self.template_save_picker.save_file(
            dialog_title="保存方案模板",
            file_name="分摊方案模板.json",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["json"],
        )

    def _template_save_path_selected(self, event: ft.FilePickerResultEvent) -> None:
        if not event.path or self._job_in_progress():
            return
        try:
            path = self._ensure_extension(event.path, (".json",), ".json")
            schemes = [
                dict(
                    scheme,
                    filter_rules=[dict(rule) for rule in scheme.get("filter_rules", [])],
                    base_columns=list(scheme.get("base_columns", [])),
                )
                for scheme in self.schemes
            ]
            headers = list(self.headers)
            sheet = self.sheet_name.value or ""
            row_number = self._parse_header_row()
        except Exception as exc:
            self._show_error("保存失败", exc, traceback.format_exc())
            return

        def task() -> str:
            template = serialize_scheme_template(schemes, headers, sheet_name=sheet, header_row=row_number)
            with open(path, "w", encoding="utf-8") as file_obj:
                json.dump(template, file_obj, ensure_ascii=False, indent=2)
            return path

        def success(saved_path: str) -> None:
            self.status_text.value = "方案模板已保存：{}".format(saved_path)
            self._show_message("已保存", "方案模板已保存：\n{}".format(saved_path))

        self._start_job("正在保存方案模板...", task, success, "保存失败")

    def _open_template_picker(self, _event: ft.ControlEvent) -> None:
        if self._job_in_progress():
            return
        if not self.headers:
            self._show_message("缺少表头", "请先读取当前表头，再导入模板。")
            return
        self.template_open_picker.pick_files(
            dialog_title="导入方案模板",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["json"],
            allow_multiple=False,
        )

    def _template_file_selected(self, event: ft.FilePickerResultEvent) -> None:
        if self._job_in_progress() or not event.files or not event.files[0].path:
            return
        path = event.files[0].path
        headers = list(self.headers)

        def task() -> List[Dict[str, Any]]:
            with open(path, "r", encoding="utf-8") as file_obj:
                template = json.load(file_obj)
            return import_scheme_template(template, headers)

        def success(imported: List[Dict[str, Any]]) -> None:
            self._prompt_template_merge(imported)

        self._start_job("正在导入方案模板...", task, success, "导入失败")

    def _prompt_template_merge(self, imported: List[Dict[str, Any]]) -> None:
        if not self.schemes:
            self._apply_imported_schemes(imported, True)
            return
        dialog: ft.AlertDialog

        def apply(replace: bool) -> None:
            dialog.open = False
            self._apply_imported_schemes(imported, replace)

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("导入方案模板"),
            content=ft.Text("是否用模板方案替换当前方案？选择“追加”会保留当前方案。"),
            actions=[
                ft.TextButton(text="取消", on_click=lambda _event: self._close_dialog(dialog)),
                ft.ElevatedButton(text="追加", on_click=lambda _event: apply(False)),
                ft.FilledButton(text="替换", on_click=lambda _event: apply(True)),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.open(dialog)
        self.page.update()

    def _apply_imported_schemes(self, imported: List[Dict[str, Any]], replace: bool) -> None:
        if replace:
            self.schemes = imported
        else:
            self._save_current_scheme()
            self.schemes.extend(imported)
        self.active_scheme_index = -1
        self._load_scheme(0, update=False)
        self._refresh_scheme_list()
        self.status_text.value = "已导入 {} 个方案模板。".format(len(imported))
        self.page.update()
        self._show_message("导入完成", "已导入 {} 个方案。".format(len(imported)))

    def _preview_allocation(self, _event: ft.ControlEvent) -> None:
        try:
            config = self._build_batch_config()
        except Exception as exc:
            self._show_error("配置不完整", exc, traceback.format_exc())
            return

        def success(result: Any) -> None:
            self._fill_preview_table(result)
            self.status_text.value = "预览完成，请核对后执行分摊。"
            self._show_view(2)

        self._start_job(
            "正在生成预览...",
            lambda: preview_workbook_batch(config),
            success,
            "预览失败",
        )

    def _run_allocation(self, _event: ft.ControlEvent) -> None:
        try:
            config = self._build_batch_config()
        except Exception as exc:
            self._show_error("配置不完整", exc, traceback.format_exc())
            return

        def success(result: Any) -> None:
            self._fill_preview_table(result)
            self.status_text.value = "完成：{}".format(result.output_path)
            self._show_view(2)
            summary = "\n".join(
                "{}：金额 {}，参与 {} 行，不参与 {} 行".format(
                    item.name,
                    item.target_total,
                    item.participating_rows,
                    item.excluded_rows,
                )
                for item in result.scheme_results
            )
            self._show_message(
                "分摊完成",
                "分摊完成。\n\n输出文件：{}\n\n{}".format(result.output_path, summary),
            )

        self._start_job(
            "正在分摊，请稍候...",
            lambda: allocate_workbook_batch(config),
            success,
            "分摊失败",
        )

    def _fill_preview_table(self, result: Any) -> None:
        self.preview_table.rows = [
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(item.name)),
                    ft.DataCell(ft.Text(self.result_source_label(item.amount_source, item.amount_column))),
                    ft.DataCell(ft.Text(self.column_label_short(item.allocation_column))),
                    ft.DataCell(ft.Text(str(item.target_total))),
                    ft.DataCell(ft.Text(str(item.participating_rows))),
                    ft.DataCell(ft.Text(str(item.excluded_rows))),
                    ft.DataCell(ft.Text(str(item.base_total))),
                    ft.DataCell(ft.Text(str(item.distributed_total))),
                ]
            )
            for item in result.scheme_results
        ]

    def _build_batch_config(self) -> BatchAllocationConfig:
        self._save_current_scheme()
        scheme_configs = []
        for index, scheme in enumerate(self.schemes, start=1):
            name = (scheme.get("name") or "").strip() or "方案{}".format(index)
            allocation_column = self.label_to_column_index(scheme.get("allocation_column", ""))
            if allocation_column is None:
                raise ValueError("{}：请选择分摊结果列。".format(name))
            base_columns = [self.label_to_column_index(label) for label in scheme.get("base_columns", [])]
            if any(column is None for column in base_columns):
                raise ValueError("{}：占比计算列选择无效，请重新选择。".format(name))
            amount_mode = scheme.get("amount_mode", "target_total")
            amount_source = "manual" if amount_mode == "manual" else "column_total"
            amount_column = None
            manual_amount = None
            if amount_mode == "source_column":
                amount_column = self.label_to_column_index(scheme.get("amount_column", ""))
                if amount_column is None:
                    raise ValueError("{}：请选择金额来源列。".format(name))
            elif amount_mode == "manual":
                manual_amount = self._parse_manual_amount(scheme.get("manual_amount", ""), name)
            filter_rules = self._build_rules_from_drafts(name, scheme.get("filter_rules", []))
            scheme_configs.append(
                AllocationScheme(
                    name=name,
                    amount_source=amount_source,
                    amount_column=amount_column,
                    manual_amount=manual_amount,
                    allocation_column=allocation_column,
                    base_columns=[column for column in base_columns if column is not None],
                    filter_rules=filter_rules,
                    filter_logic=scheme.get("filter_logic", "OR"),
                )
            )
        return BatchAllocationConfig(
            input_path=(self.file_path.value or "").strip(),
            output_path=(self.output_path.value or "").strip(),
            sheet_name=self.sheet_name.value or "",
            header_row=self._parse_header_row(),
            schemes=scheme_configs,
        )

    def _build_rules_from_drafts(
        self,
        scheme_name: str,
        drafts: Sequence[Dict[str, str]],
    ) -> List[FilterRule]:
        rules = []
        for draft in drafts:
            column_label = draft.get("column", "")
            if not column_label:
                continue
            column = self.label_to_column_index(column_label)
            if column is None:
                raise ValueError("{}：过滤条件字段无效，请重新选择。".format(scheme_name))
            operator = FilterRuleEditor.operator_code(draft.get("operator", "equals"))
            value = (draft.get("value") or "").strip()
            if operator not in {"blank", "not_blank"} and not value:
                raise ValueError("{}：过滤条件的值不能为空；不需要过滤时请删除该规则。".format(scheme_name))
            rules.append(FilterRule(column=column, operator=operator, value=value))
        return rules

    def _parse_manual_amount(self, value: str, scheme_name: str) -> Decimal:
        text = (value or "").strip().replace(",", "").replace("，", "").replace("￥", "").replace("¥", "")
        if not text:
            raise ValueError("{}：请输入手工金额。".format(scheme_name))
        try:
            return Decimal(text)
        except InvalidOperation as exc:
            raise ValueError("{}：手工金额格式不正确。".format(scheme_name)) from exc

    def _parse_header_row(self) -> int:
        try:
            value = int((self.header_row.value or "").strip())
        except (TypeError, ValueError) as exc:
            raise ValueError("表头行必须是 1 到 999 之间的整数。") from exc
        if value < 1 or value > 999:
            raise ValueError("表头行必须是 1 到 999 之间的整数。")
        return value

    @staticmethod
    def result_source_label(amount_source: str, amount_column: Optional[int]) -> str:
        if amount_source == "manual":
            return "手工输入"
        if amount_column:
            return "指定列合计"
        return "分摊列合计"

    def _start_job(
        self,
        status: str,
        task: Callable[[], Any],
        on_success: Callable[[Any], None],
        error_title: str,
    ) -> None:
        with self._job_lock:
            if self._busy:
                started = False
            else:
                self._busy = True
                started = True
        if not started:
            self._show_message("任务进行中", "请等待当前任务完成后再试。")
            return
        self.progress.visible = True
        self.status_text.value = status
        self._set_job_buttons_disabled(True)
        self.page.update()

        def worker() -> None:
            try:
                result = task()
            except Exception as exc:
                self.status_text.value = "{}。".format(error_title)
                self._show_error(error_title, exc, traceback.format_exc())
            else:
                try:
                    on_success(result)
                except Exception as exc:
                    self.status_text.value = "处理已完成，但界面刷新失败。"
                    self._show_error("界面刷新失败", exc, traceback.format_exc())
            finally:
                with self._job_lock:
                    self._busy = False
                self.progress.visible = False
                self._set_job_buttons_disabled(False)
                self.page.update()

        self.page.run_thread(worker)

    def _job_in_progress(self) -> bool:
        with self._job_lock:
            return self._busy

    def _set_job_buttons_disabled(self, disabled: bool) -> None:
        for button in self._job_buttons:
            button.disabled = disabled
        for control in self._busy_controls:
            control.disabled = disabled
        if not disabled:
            self._update_amount_mode_state()
        for checkbox in self.base_column_list.controls:
            checkbox.disabled = disabled
        for editor in self.filter_rule_editors:
            editor.column.disabled = disabled
            editor.operator.disabled = disabled
            editor.value.disabled = disabled or editor.operator.value in {"blank", "not_blank"}
            editor.value_button.disabled = disabled or editor.operator.value in {"blank", "not_blank"}
            editor.delete_button.disabled = disabled

    def _show_message(self, title: str, message: str) -> None:
        dialog: ft.AlertDialog

        def close(_event: ft.ControlEvent) -> None:
            self._close_dialog(dialog)

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text(title),
            content=ft.Container(
                content=ft.Text(message, selectable=True),
                width=560,
            ),
            actions=[ft.FilledButton(text="确定", on_click=close)],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.open(dialog)
        self.page.update()

    def _show_error(self, title: str, error: Exception, details: str) -> None:
        dialog: ft.AlertDialog
        error_text = "{}: {}\n\n详细信息：\n{}".format(type(error).__name__, error, details)

        def close(_event: ft.ControlEvent) -> None:
            self._close_dialog(dialog)

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                [
                    ft.Icon(ft.icons.ERROR_OUTLINE, color=ft.colors.RED_300),
                    ft.Text(title),
                ]
            ),
            content=ft.Container(
                width=680,
                height=360,
                content=ft.TextField(
                    value=error_text,
                    read_only=True,
                    multiline=True,
                    min_lines=10,
                    max_lines=16,
                    border_radius=12,
                ),
            ),
            actions=[ft.FilledButton(text="关闭", on_click=close)],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.open(dialog)
        self.page.update()

    def _close_dialog(self, dialog: ft.AlertDialog) -> None:
        dialog.open = False
        self.page.update()

    @staticmethod
    def _ensure_extension(path: str, allowed: Sequence[str], default: str) -> str:
        if Path(path).suffix.lower() not in allowed:
            return "{}{}".format(path, default)
        return path


def main(page: ft.Page) -> None:
    FletAllocatorApp(page)


if __name__ == "__main__":
    ft.app(target=main)
