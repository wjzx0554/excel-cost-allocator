import sys
import json
import threading
import traceback
import tkinter as tk
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .allocator import (
    AllocationScheme,
    BatchAllocationConfig,
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
from .templates import import_scheme_template, serialize_scheme_template


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


class FilterRuleRow:
    def __init__(self, app, parent, default_column="", default_operator="equals", default_value=""):
        self.app = app
        self.frame = ttk.Frame(parent, style="Rule.TFrame", padding=(8, 6))
        self.frame.columnconfigure(1, weight=1)
        self.frame.columnconfigure(5, weight=1)

        self.column_var = tk.StringVar(value=default_column)
        self.operator_var = tk.StringVar(value=self.operator_label(default_operator))
        self.value_var = tk.StringVar(value=display_value(default_value) if default_value == "" else default_value)

        ttk.Label(self.frame, text="字段", style="Rule.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        self.column_combo = ttk.Combobox(
            self.frame,
            textvariable=self.column_var,
            state="readonly",
            values=self.app.header_labels(),
        )
        self.column_combo.grid(row=0, column=1, sticky=tk.EW, padx=(0, 8))

        ttk.Label(self.frame, text="条件", style="Rule.TLabel").grid(row=0, column=2, sticky=tk.W, padx=(0, 6))
        self.operator_combo = ttk.Combobox(
            self.frame,
            textvariable=self.operator_var,
            state="readonly",
            width=10,
            values=[label for _code, label in OPERATORS],
        )
        self.operator_combo.grid(row=0, column=3, sticky=tk.W, padx=(0, 8))
        self.operator_combo.bind("<<ComboboxSelected>>", self._operator_selected)

        ttk.Label(self.frame, text="值", style="Rule.TLabel").grid(row=0, column=4, sticky=tk.W, padx=(0, 6))
        self.value_combo = ttk.Combobox(self.frame, textvariable=self.value_var)
        self.value_combo.grid(row=0, column=5, sticky=tk.EW, padx=(0, 8))
        self.column_combo.bind("<<ComboboxSelected>>", self._column_selected)
        self.value_combo.bind("<FocusIn>", self._load_values)

        ttk.Button(
            self.frame,
            text="删除",
            command=lambda: self.app.remove_filter_rule_row(self),
            style="Danger.TButton",
        ).grid(row=0, column=6, sticky=tk.E)

        self.refresh_headers()

    def pack(self):
        self.frame.pack(fill=tk.X, pady=4)

    def destroy(self):
        self.frame.destroy()

    def to_draft(self):
        value = self.value_var.get().strip()
        if value == "<空白>":
            value = ""
        return {
            "column": self.column_var.get(),
            "operator": self.operator_code(self.operator_var.get()),
            "value": value,
        }

    def to_rule(self):
        draft = self.to_draft()
        column_index = self.app.label_to_column_index(draft["column"])
        if column_index is None:
            raise ValueError("过滤条件中的字段选择无效。")

        operator = draft["operator"]
        value = draft["value"]
        if operator not in {"blank", "not_blank"} and value == "":
            raise ValueError("过滤条件中的值不能为空；不需要过滤时请删除这条规则。")

        return FilterRule(column=column_index, operator=operator, value=value)

    def refresh_headers(self):
        labels = self.app.header_labels()
        self.column_combo.configure(values=labels)
        if labels and self.column_var.get() not in labels:
            self.column_var.set(self.app.suggest_filter_column())
        self.refresh_column_values()

    def refresh_column_values(self):
        self._load_values()
        self._update_value_state()

    def _column_selected(self, _event=None):
        self.refresh_column_values()

    def _operator_selected(self, _event=None):
        self._update_value_state()

    def _load_values(self, _event=None):
        column_index = self.app.label_to_column_index(self.column_var.get())
        if column_index is None:
            self.value_combo.configure(values=[])
            return
        values = self.app.cached_unique_values(column_index)
        self.value_combo.configure(values=[display_value(value) for value in values])

    def _update_value_state(self):
        operator = self.operator_code(self.operator_var.get())
        if operator in {"blank", "not_blank"}:
            self.value_combo.configure(state="disabled")
            self.value_var.set("")
        else:
            self.value_combo.configure(state="normal")

    @staticmethod
    def operator_code(label_or_code):
        for code, label in OPERATORS:
            if label_or_code in {code, label}:
                return code
        return label_or_code or "equals"

    @staticmethod
    def operator_label(code):
        for item_code, label in OPERATORS:
            if item_code == code:
                return label
        return code


class AllocatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("分摊工具")
        self.geometry("1180x800")
        self.minsize(1060, 720)
        self.configure(bg="#eef3f8")
        self._set_window_icon()
        self._setup_style()

        self.file_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.header_row = tk.IntVar(value=1)
        self.status_text = tk.StringVar(value="请选择 Excel 文件。")

        self.scheme_name = tk.StringVar()
        self.amount_mode = tk.StringVar(value="取分摊列原始合计")
        self.amount_column = tk.StringVar()
        self.manual_amount = tk.StringVar()
        self.allocation_column = tk.StringVar()
        self.filter_logic = tk.StringVar(value="OR")

        self.headers = []
        self.schemes = []
        self.active_scheme_index = -1
        self.filter_rule_rows = []
        self.unique_value_cache = {}
        self._loading_scheme = False

        self._build_ui()

    def _asset_path(self, name):
        base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
        return base_dir / "assets" / name

    def _set_window_icon(self):
        icon_path = self._asset_path("app.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except tk.TclError:
                pass

    def _setup_style(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.default_font = ("Microsoft YaHei UI", 9)
        self.title_font = ("Microsoft YaHei UI", 18, "bold")
        self.subtitle_font = ("Microsoft YaHei UI", 9)
        self.section_font = ("Microsoft YaHei UI", 10, "bold")

        self.option_add("*Font", self.default_font)
        self.style.configure("App.TFrame", background="#eef3f8")
        self.style.configure("Hero.TFrame", background="#164078")
        self.style.configure("Panel.TFrame", background="#ffffff")
        self.style.configure("Rule.TFrame", background="#ffffff")
        self.style.configure("TNotebook", background="#eef3f8", borderwidth=0)
        self.style.configure("TNotebook.Tab", padding=(16, 8), font=("Microsoft YaHei UI", 9, "bold"))
        self.style.configure("Card.TLabelframe", background="#ffffff", bordercolor="#d8e0ee", relief="solid")
        self.style.configure("Card.TLabelframe.Label", background="#ffffff", foreground="#0f172a", font=self.section_font)
        self.style.configure("TLabel", background="#eef3f8", foreground="#334155")
        self.style.configure("Card.TLabel", background="#ffffff", foreground="#334155")
        self.style.configure("Rule.TLabel", background="#ffffff", foreground="#334155")
        self.style.configure("Hint.TLabel", background="#ffffff", foreground="#64748b")
        self.style.configure("HeroTitle.TLabel", background="#164078", foreground="#ffffff", font=self.title_font)
        self.style.configure("HeroSub.TLabel", background="#164078", foreground="#dbeafe", font=self.subtitle_font)
        self.style.configure("Status.TLabel", background="#e8f1ff", foreground="#1e3a8a", padding=(10, 6))
        self.style.configure("TButton", padding=(12, 6), background="#e2e8f0", foreground="#0f172a")
        self.style.map("TButton", background=[("active", "#cbd5e1")])
        self.style.configure("Primary.TButton", padding=(16, 8), background="#1d4ed8", foreground="#ffffff", font=("Microsoft YaHei UI", 10, "bold"))
        self.style.map("Primary.TButton", background=[("active", "#1e40af"), ("disabled", "#93c5fd")])
        self.style.configure("Danger.TButton", padding=(10, 5), background="#fee2e2", foreground="#991b1b")
        self.style.map("Danger.TButton", background=[("active", "#fecaca")])

    def _build_ui(self):
        root = ttk.Frame(self, padding=12, style="App.TFrame")
        root.pack(fill=tk.BOTH, expand=True)

        hero = ttk.Frame(root, padding=(18, 14), style="Hero.TFrame")
        hero.pack(fill=tk.X)
        ttk.Label(hero, text="分摊工具", style="HeroTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(
            hero,
            text="按方案配置费用来源、占比基数和过滤条件，生成可核对的 Excel/WPS 分摊结果",
            style="HeroSub.TLabel",
        ).pack(anchor=tk.W, pady=(4, 0))

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.base_tab = ttk.Frame(self.notebook, padding=12, style="App.TFrame")
        self.scheme_tab = ttk.Frame(self.notebook, padding=12, style="App.TFrame")
        self.preview_tab = ttk.Frame(self.notebook, padding=12, style="App.TFrame")
        self.help_tab = ttk.Frame(self.notebook, padding=12, style="App.TFrame")
        self.notebook.add(self.base_tab, text="1 基础设置")
        self.notebook.add(self.scheme_tab, text="2 分摊方案")
        self.notebook.add(self.preview_tab, text="3 预览执行")
        self.notebook.add(self.help_tab, text="4 使用说明")

        self._build_base_tab()
        self._build_scheme_tab()
        self._build_preview_tab()
        self._build_help_tab()

        action_frame = ttk.Frame(root, style="App.TFrame")
        action_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(action_frame, textvariable=self.status_text, style="Status.TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(action_frame, text="生成预览", command=self.preview_allocation).pack(side=tk.RIGHT, padx=(8, 0))
        self.run_button = ttk.Button(action_frame, text="开始分摊", command=self.run_allocation, style="Primary.TButton")
        self.run_button.pack(side=tk.RIGHT)

    def _build_base_tab(self):
        self.base_tab.columnconfigure(0, weight=1)

        file_frame = ttk.LabelFrame(self.base_tab, text="文件与表头", padding=12, style="Card.TLabelframe")
        file_frame.grid(row=0, column=0, sticky=tk.EW)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Excel 文件", style="Card.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(file_frame, textvariable=self.file_path).grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(file_frame, text="浏览", command=self.choose_file).grid(row=0, column=2, padx=(8, 0))

        ttk.Label(file_frame, text="工作表", style="Card.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=(10, 0))
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="readonly", width=28)
        self.sheet_combo.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_changed)

        ttk.Label(file_frame, text="表头行", style="Card.TLabel").grid(row=1, column=2, sticky=tk.E, padx=(8, 4), pady=(10, 0))
        ttk.Spinbox(file_frame, from_=1, to=999, width=8, textvariable=self.header_row, command=self.load_headers).grid(row=1, column=3, sticky=tk.W, pady=(10, 0))
        ttk.Button(file_frame, text="读取表头", command=self.load_headers).grid(row=1, column=4, padx=(8, 0), pady=(10, 0))

        ttk.Label(file_frame, text="输出路径", style="Card.TLabel").grid(row=2, column=0, sticky=tk.W, padx=(0, 8), pady=(10, 0))
        ttk.Entry(file_frame, textvariable=self.output_path).grid(row=2, column=1, columnspan=3, sticky=tk.EW, pady=(10, 0))
        ttk.Button(file_frame, text="另存为", command=self.choose_output).grid(row=2, column=4, padx=(8, 0), pady=(10, 0))

        quick_frame = ttk.LabelFrame(self.base_tab, text="读取结果", padding=12, style="Card.TLabelframe")
        quick_frame.grid(row=1, column=0, sticky=tk.NSEW, pady=(10, 0))
        quick_frame.columnconfigure(0, weight=1)
        quick_frame.rowconfigure(1, weight=1)
        self.header_summary = ttk.Label(quick_frame, text="尚未读取表头。", style="Card.TLabel")
        self.header_summary.grid(row=0, column=0, sticky=tk.W)
        self.header_list = self._create_listbox(quick_frame, height=12)
        self.header_list.grid(row=1, column=0, sticky=tk.NSEW, pady=(8, 0))
        header_scroll = ttk.Scrollbar(quick_frame, orient=tk.VERTICAL, command=self.header_list.yview)
        header_scroll.grid(row=1, column=1, sticky=tk.NS, pady=(8, 0))
        self.header_list.configure(yscrollcommand=header_scroll.set)

        sample_frame = ttk.LabelFrame(self.base_tab, text="测试数据", padding=12, style="Card.TLabelframe")
        sample_frame.grid(row=2, column=0, sticky=tk.EW, pady=(10, 0))
        ttk.Label(
            sample_frame,
            text="需要先试流程时，可以生成一份示例 Excel，再用它验证表头读取、过滤、预览和导出。",
            style="Card.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Button(sample_frame, text="生成测试数据", command=self.save_sample_workbook).pack(side=tk.RIGHT)

    def _build_scheme_tab(self):
        self.scheme_tab.columnconfigure(1, weight=1)
        self.scheme_tab.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(self.scheme_tab, text="方案列表", padding=10, style="Card.TLabelframe")
        left.grid(row=0, column=0, sticky=tk.NS, padx=(0, 10))
        left.rowconfigure(0, weight=1)
        self.scheme_list = self._create_listbox(left, height=18)
        self.scheme_list.grid(row=0, column=0, columnspan=3, sticky=tk.NS)
        self.scheme_list.bind("<<ListboxSelect>>", self.on_scheme_selected)
        ttk.Button(left, text="新增", command=self.add_scheme).grid(row=1, column=0, sticky=tk.EW, pady=(8, 0), padx=(0, 4))
        ttk.Button(left, text="复制", command=self.copy_scheme).grid(row=1, column=1, sticky=tk.EW, pady=(8, 0), padx=4)
        ttk.Button(left, text="删除", command=self.delete_scheme, style="Danger.TButton").grid(row=1, column=2, sticky=tk.EW, pady=(8, 0), padx=(4, 0))
        ttk.Button(left, text="导入模板", command=self.import_scheme_template).grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=(10, 0))
        ttk.Button(left, text="保存模板", command=self.save_scheme_template).grid(row=3, column=0, columnspan=3, sticky=tk.EW, pady=(6, 0))

        right = ttk.Frame(self.scheme_tab, style="App.TFrame")
        right.grid(row=0, column=1, sticky=tk.NSEW)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self.editor_canvas = tk.Canvas(right, bg="#eef3f8", highlightthickness=0, borderwidth=0)
        self.editor_canvas.grid(row=0, column=0, sticky=tk.NSEW)
        editor_scroll = ttk.Scrollbar(right, orient=tk.VERTICAL, command=self.editor_canvas.yview)
        editor_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.editor_canvas.configure(yscrollcommand=editor_scroll.set)
        self.editor_inner = ttk.Frame(self.editor_canvas, style="App.TFrame")
        self.editor_window = self.editor_canvas.create_window((0, 0), window=self.editor_inner, anchor="nw")
        self.editor_inner.bind("<Configure>", self._sync_editor_scrollregion)
        self.editor_canvas.bind("<Configure>", self._sync_editor_canvas_width)

        self._build_scheme_editor(self.editor_inner)

    def _build_scheme_editor(self, parent):
        parent.columnconfigure(0, weight=1)

        info = ttk.LabelFrame(parent, text="方案基础信息", padding=12, style="Card.TLabelframe")
        info.grid(row=0, column=0, sticky=tk.EW)
        info.columnconfigure(1, weight=1)
        info.columnconfigure(3, weight=1)

        ttk.Label(info, text="方案名称", style="Card.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(info, textvariable=self.scheme_name).grid(row=0, column=1, sticky=tk.EW, padx=(0, 12))
        ttk.Label(info, text="分摊结果列", style="Card.TLabel").grid(row=0, column=2, sticky=tk.W, padx=(0, 8))
        self.allocation_column_combo = ttk.Combobox(info, textvariable=self.allocation_column, state="readonly")
        self.allocation_column_combo.grid(row=0, column=3, sticky=tk.EW)

        amount = ttk.LabelFrame(parent, text="分摊金额来源", padding=12, style="Card.TLabelframe")
        amount.grid(row=1, column=0, sticky=tk.EW, pady=(10, 0))
        amount.columnconfigure(1, weight=1)
        amount.columnconfigure(3, weight=1)

        ttk.Label(amount, text="金额来源", style="Card.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        self.amount_mode_combo = ttk.Combobox(
            amount,
            textvariable=self.amount_mode,
            state="readonly",
            values=[label for _code, label in AMOUNT_MODES],
        )
        self.amount_mode_combo.grid(row=0, column=1, sticky=tk.EW, padx=(0, 12))
        self.amount_mode_combo.bind("<<ComboboxSelected>>", self.on_amount_mode_changed)
        ttk.Label(amount, text="金额来源列", style="Card.TLabel").grid(row=0, column=2, sticky=tk.W, padx=(0, 8))
        self.amount_column_combo = ttk.Combobox(amount, textvariable=self.amount_column, state="readonly")
        self.amount_column_combo.grid(row=0, column=3, sticky=tk.EW)

        ttk.Label(amount, text="手工金额", style="Card.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=(10, 0))
        self.manual_amount_entry = ttk.Entry(amount, textvariable=self.manual_amount)
        self.manual_amount_entry.grid(row=1, column=1, sticky=tk.EW, pady=(10, 0))
        ttk.Label(
            amount,
            text="手工金额适合财务单独给出的一笔费用；建议一个方案只写入一个分摊结果列，便于查账。",
            style="Hint.TLabel",
        ).grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=(10, 0))

        base = ttk.LabelFrame(parent, text="占比计算列", padding=12, style="Card.TLabelframe")
        base.grid(row=2, column=0, sticky=tk.NSEW, pady=(10, 0))
        base.columnconfigure(0, weight=1)
        base.rowconfigure(1, weight=1)
        ttk.Label(base, text="可以多选，例如：完工入库材料成本 + 本期人工费。", style="Hint.TLabel").grid(row=0, column=0, sticky=tk.W)
        self.base_list = self._create_listbox(base, height=8)
        self.base_list.grid(row=1, column=0, sticky=tk.NSEW, pady=(8, 0))
        base_scroll = ttk.Scrollbar(base, orient=tk.VERTICAL, command=self.base_list.yview)
        base_scroll.grid(row=1, column=1, sticky=tk.NS, pady=(8, 0))
        self.base_list.configure(yscrollcommand=base_scroll.set)
        base_toolbar = ttk.Frame(base, style="Panel.TFrame")
        base_toolbar.grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Button(base_toolbar, text="智能选择", command=self.auto_select_base_columns).pack(side=tk.LEFT)
        ttk.Button(base_toolbar, text="清空选择", command=lambda: self.base_list.selection_clear(0, tk.END)).pack(side=tk.LEFT, padx=(8, 0))

        filter_frame = ttk.LabelFrame(parent, text="不参与分摊的过滤条件", padding=12, style="Card.TLabelframe")
        filter_frame.grid(row=3, column=0, sticky=tk.NSEW, pady=(10, 0))
        filter_frame.columnconfigure(0, weight=2)
        filter_frame.columnconfigure(1, weight=1)
        filter_frame.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(filter_frame, style="Rule.TFrame")
        toolbar.grid(row=0, column=0, columnspan=2, sticky=tk.EW)
        ttk.Label(toolbar, text="规则关系", style="Rule.TLabel").pack(side=tk.LEFT)
        ttk.Combobox(toolbar, textvariable=self.filter_logic, state="readonly", width=8, values=["OR", "AND"]).pack(side=tk.LEFT, padx=(8, 12))
        ttk.Button(toolbar, text="+ 添加规则", command=self.add_filter_rule_row).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="测试命中", command=self.test_filter_rules).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(toolbar, text="OR=任一条件命中即排除；AND=全部条件满足才排除", style="Hint.TLabel").pack(side=tk.RIGHT)

        rules_panel = ttk.Frame(filter_frame, style="App.TFrame")
        rules_panel.grid(row=1, column=0, sticky=tk.NSEW, padx=(0, 10), pady=(8, 0))
        rules_panel.columnconfigure(0, weight=1)
        rules_panel.rowconfigure(0, weight=1)
        self.rule_canvas = tk.Canvas(rules_panel, bg="#eef3f8", highlightthickness=0, borderwidth=0, height=150)
        self.rule_canvas.grid(row=0, column=0, sticky=tk.NSEW)
        rule_scroll = ttk.Scrollbar(rules_panel, orient=tk.VERTICAL, command=self.rule_canvas.yview)
        rule_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.rule_canvas.configure(yscrollcommand=rule_scroll.set)
        self.rule_inner = ttk.Frame(self.rule_canvas, style="App.TFrame")
        self.rule_window = self.rule_canvas.create_window((0, 0), window=self.rule_inner, anchor="nw")
        self.rule_inner.bind("<Configure>", self._sync_rule_scrollregion)
        self.rule_canvas.bind("<Configure>", self._sync_rule_canvas_width)

        regex_help = (
            "正则示例：^销售配货部$ 精确匹配；.*维修.* 包含维修；"
            "^(销售配货部|售后服务部)$ 多值匹配；^$ 匹配空白。"
        )
        ttk.Label(rules_panel, text=regex_help, style="Hint.TLabel", wraplength=650).grid(row=1, column=0, sticky=tk.W, pady=(6, 0))

        preview_panel = ttk.LabelFrame(filter_frame, text="过滤命中预览", padding=10, style="Card.TLabelframe")
        preview_panel.grid(row=1, column=1, sticky=tk.NSEW, pady=(8, 0))
        preview_panel.columnconfigure(0, weight=1)
        preview_panel.rowconfigure(1, weight=1)
        self.filter_preview_summary = ttk.Label(preview_panel, text="尚未测试规则。", style="Card.TLabel")
        self.filter_preview_summary.grid(row=0, column=0, sticky=tk.W)
        self.filter_preview_list = self._create_listbox(preview_panel, height=7)
        self.filter_preview_list.grid(row=1, column=0, sticky=tk.NSEW, pady=(8, 0))
        preview_scroll = ttk.Scrollbar(preview_panel, orient=tk.VERTICAL, command=self.filter_preview_list.yview)
        preview_scroll.grid(row=1, column=1, sticky=tk.NS, pady=(8, 0))
        self.filter_preview_list.configure(yscrollcommand=preview_scroll.set)

    def _build_preview_tab(self):
        self.preview_tab.columnconfigure(0, weight=1)
        self.preview_tab.rowconfigure(1, weight=1)

        toolbar = ttk.LabelFrame(self.preview_tab, text="执行前核对", padding=12, style="Card.TLabelframe")
        toolbar.grid(row=0, column=0, sticky=tk.EW)
        ttk.Label(
            toolbar,
            text="先生成预览，确认每个方案的金额、参与行数、基数合计和分摊列，再执行导出。",
            style="Card.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="生成预览", command=self.preview_allocation).pack(side=tk.RIGHT)

        table_frame = ttk.LabelFrame(self.preview_tab, text="方案预览", padding=10, style="Card.TLabelframe")
        table_frame.grid(row=1, column=0, sticky=tk.NSEW, pady=(10, 0))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        columns = ("name", "source", "target", "amount", "participating", "excluded", "base", "distributed")
        self.preview_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=12)
        headings = {
            "name": "方案名称",
            "source": "金额来源",
            "target": "分摊列",
            "amount": "待分摊金额",
            "participating": "参与行",
            "excluded": "不参与行",
            "base": "基数合计",
            "distributed": "分摊后合计",
        }
        widths = {
            "name": 170,
            "source": 130,
            "target": 130,
            "amount": 120,
            "participating": 80,
            "excluded": 80,
            "base": 120,
            "distributed": 120,
        }
        for col in columns:
            self.preview_tree.heading(col, text=headings[col])
            self.preview_tree.column(col, width=widths[col], anchor=tk.CENTER)
        self.preview_tree.grid(row=0, column=0, sticky=tk.NSEW)
        tree_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.preview_tree.configure(yscrollcommand=tree_scroll.set)

        self.preview_note = tk.Text(table_frame, height=7, wrap=tk.WORD, bg="#ffffff", fg="#334155", relief=tk.SOLID, bd=1)
        self.preview_note.grid(row=1, column=0, sticky=tk.EW, pady=(10, 0))
        self.preview_note.insert(
            tk.END,
            "预览会检查：是否有重复分摊列、手工金额是否填写、占比基数是否大于 0、过滤条件是否可用。\n"
            "如果某个方案待分摊金额不为 0，但参与行基数合计为 0，执行时会阻止导出，避免生成错误结果。",
        )
        self.preview_note.configure(state=tk.DISABLED)

    def _build_help_tab(self):
        self.help_tab.columnconfigure(0, weight=1)
        self.help_tab.rowconfigure(0, weight=1)
        text = tk.Text(
            self.help_tab,
            wrap=tk.WORD,
            bg="#ffffff",
            fg="#1f2937",
            relief=tk.SOLID,
            bd=1,
            padx=18,
            pady=16,
            spacing1=3,
            spacing3=8,
        )
        text.grid(row=0, column=0, sticky=tk.NSEW)
        help_scroll = ttk.Scrollbar(self.help_tab, orient=tk.VERTICAL, command=text.yview)
        help_scroll.grid(row=0, column=1, sticky=tk.NS)
        text.configure(yscrollcommand=help_scroll.set)
        text.tag_configure("center", justify=tk.CENTER)
        text.tag_configure("support_title", justify=tk.CENTER, font=("Microsoft YaHei UI", 13, "bold"), foreground="#164078")
        text.tag_configure("support_line", justify=tk.CENTER, font=("Microsoft YaHei UI", 11), foreground="#334155")
        text.tag_configure("intro_gap", spacing3=14)
        self._insert_help_support_block(text)
        text.insert(tk.END, HELP_TEXT)
        text.configure(state=tk.DISABLED)

    def _insert_help_support_block(self, text):
        text.insert(tk.END, "技术支持\n", "support_title")
        text.insert(tk.END, "A0金蝶软件王朝\n", "support_line")
        text.insert(tk.END, "电话：15939121371（微信同号）\n", "support_line")
        text.insert(tk.END, "河南 焦作\n\n", ("support_line", "intro_gap"))

    def _create_listbox(self, parent, height=10):
        return tk.Listbox(
            parent,
            selectmode=tk.EXTENDED,
            exportselection=False,
            height=height,
            font=self.default_font,
            bg="#ffffff",
            fg="#0f172a",
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightbackground="#d8e0ee",
            highlightcolor="#1d4ed8",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            activestyle="none",
        )

    def _sync_editor_scrollregion(self, _event=None):
        self.editor_canvas.configure(scrollregion=self.editor_canvas.bbox("all"))

    def _sync_editor_canvas_width(self, event):
        self.editor_canvas.itemconfigure(self.editor_window, width=event.width)

    def _sync_rule_scrollregion(self, _event=None):
        self.rule_canvas.configure(scrollregion=self.rule_canvas.bbox("all"))

    def _sync_rule_canvas_width(self, event):
        self.rule_canvas.itemconfigure(self.rule_window, width=event.width)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.file_path.set(path)
        self._set_default_output(path)
        try:
            sheets = get_sheet_names(path)
            self.sheet_combo["values"] = sheets
            if sheets:
                self.sheet_name.set(sheets[0])
                self.header_row.set(guess_header_row(path, sheets[0]))
            self.load_headers()
            self.status_text.set("文件读取完成，请到“分摊方案”页设置方案。")
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))

    def _set_default_output(self, path):
        input_path = Path(path)
        suffix = ".xlsm" if input_path.suffix.lower() == ".xlsm" else ".xlsx"
        self.output_path.set(str(input_path.with_name(f"{input_path.stem}_分摊结果{suffix}")))

    def choose_output(self):
        initial = self.output_path.get() or self.file_path.get()
        suffix = ".xlsm" if initial.lower().endswith(".xlsm") else ".xlsx"
        path = filedialog.asksaveasfilename(
            title="保存分摊结果",
            defaultextension=suffix,
            initialfile=Path(initial).name if initial else "",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def save_sample_workbook(self):
        path = filedialog.asksaveasfilename(
            title="保存测试数据",
            defaultextension=".xlsx",
            initialfile="分摊工具测试数据.xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            create_sample_workbook(path)
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))
            return
        messagebox.showinfo("已生成", f"测试数据已生成：\n{path}")

    def on_sheet_changed(self, _event=None):
        try:
            if self.file_path.get() and self.sheet_name.get():
                self.header_row.set(guess_header_row(self.file_path.get(), self.sheet_name.get()))
            self.load_headers()
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))

    def load_headers(self):
        if not self.file_path.get() or not self.sheet_name.get():
            return
        self.headers = get_headers(self.file_path.get(), self.sheet_name.get(), int(self.header_row.get()))
        self.unique_value_cache.clear()
        labels = self.header_labels()

        self.header_list.delete(0, tk.END)
        for label in labels:
            self.header_list.insert(tk.END, label)
        self.header_summary.configure(text=f"已读取 {len(labels)} 个表头。")

        self._refresh_editor_header_values()
        self._reset_schemes_for_headers()
        self.status_text.set(f"已读取 {len(labels)} 个表头。")

    def _refresh_editor_header_values(self):
        labels = self.header_labels()
        for combo in (self.allocation_column_combo, self.amount_column_combo):
            combo.configure(values=labels)
        self.base_list.delete(0, tk.END)
        for label in labels:
            self.base_list.insert(tk.END, label)
        for row in self.filter_rule_rows:
            row.refresh_headers()

    def _reset_schemes_for_headers(self):
        self.schemes = []
        self.active_scheme_index = -1
        if self.headers:
            self.schemes.append(self._default_scheme("共耗料分摊"))
            self._refresh_scheme_list()
            self._load_scheme(0)
        else:
            self._refresh_scheme_list()

    def _default_scheme(self, name):
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

    def add_scheme(self):
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先在“基础设置”页读取表头。")
            return
        self._save_current_scheme()
        self.schemes.append(self._default_scheme(f"方案{len(self.schemes) + 1}"))
        self._refresh_scheme_list()
        self._load_scheme(len(self.schemes) - 1)

    def copy_scheme(self):
        if self.active_scheme_index < 0:
            return
        self._save_current_scheme()
        current = dict(self.schemes[self.active_scheme_index])
        current["filter_rules"] = [dict(rule) for rule in current.get("filter_rules", [])]
        current["base_columns"] = list(current.get("base_columns", []))
        current["name"] = f"{current.get('name', '方案')} 副本"
        self.schemes.append(current)
        self._refresh_scheme_list()
        self._load_scheme(len(self.schemes) - 1)

    def delete_scheme(self):
        if self.active_scheme_index < 0:
            return
        if len(self.schemes) <= 1:
            messagebox.showwarning("提示", "至少保留一个分摊方案。")
            return
        index = self.active_scheme_index
        del self.schemes[index]
        self._refresh_scheme_list()
        self._load_scheme(min(index, len(self.schemes) - 1))

    def save_scheme_template(self):
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先读取表头，再保存方案模板。")
            return
        self._save_current_scheme()
        path = filedialog.asksaveasfilename(
            title="保存方案模板",
            defaultextension=".json",
            initialfile="分摊方案模板.json",
            filetypes=[("方案模板", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            template = serialize_scheme_template(
                self.schemes,
                self.headers,
                sheet_name=self.sheet_name.get(),
                header_row=int(self.header_row.get()),
            )
            with open(path, "w", encoding="utf-8") as file_obj:
                json.dump(template, file_obj, ensure_ascii=False, indent=2)
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))
            return
        self.status_text.set(f"方案模板已保存：{path}")
        messagebox.showinfo("已保存", f"方案模板已保存：\n{path}")

    def import_scheme_template(self):
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先在“基础设置”页读取当前表头，再导入模板。")
            return
        path = filedialog.askopenfilename(
            title="导入方案模板",
            filetypes=[("方案模板", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as file_obj:
                template = json.load(file_obj)
            imported = import_scheme_template(template, self.headers)
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            return

        if self.schemes:
            replace = messagebox.askyesno(
                "导入方案模板",
                "是否用模板方案替换当前方案？\n\n选择“否”会追加到当前方案列表。",
            )
            if replace:
                self.schemes = imported
            else:
                self._save_current_scheme()
                self.schemes.extend(imported)
        else:
            self.schemes = imported

        self._refresh_scheme_list()
        self._load_scheme(0)
        self.status_text.set(f"已导入 {len(imported)} 个方案模板。")
        messagebox.showinfo("导入完成", f"已导入 {len(imported)} 个方案。")

    def on_scheme_selected(self, _event=None):
        selection = self.scheme_list.curselection()
        if not selection:
            return
        index = selection[0]
        if index == self.active_scheme_index:
            return
        self._save_current_scheme()
        self._load_scheme(index)

    def _refresh_scheme_list(self):
        self.scheme_list.delete(0, tk.END)
        for index, scheme in enumerate(self.schemes, start=1):
            self.scheme_list.insert(tk.END, f"{index}. {scheme.get('name') or '未命名方案'}")

    def _load_scheme(self, index):
        if index < 0 or index >= len(self.schemes):
            return
        self._loading_scheme = True
        self.active_scheme_index = index
        scheme = self.schemes[index]
        self.scheme_list.selection_clear(0, tk.END)
        self.scheme_list.selection_set(index)
        self.scheme_list.activate(index)

        self.scheme_name.set(scheme.get("name", ""))
        self.amount_mode.set(self.amount_mode_label(scheme.get("amount_mode", "target_total")))
        self.amount_column.set(scheme.get("amount_column", ""))
        self.manual_amount.set(scheme.get("manual_amount", ""))
        self.allocation_column.set(scheme.get("allocation_column", ""))
        self.filter_logic.set(scheme.get("filter_logic", "OR"))

        self.base_list.selection_clear(0, tk.END)
        selected = set(scheme.get("base_columns", []))
        for idx, label in enumerate(self.header_labels()):
            if label in selected:
                self.base_list.selection_set(idx)

        self._clear_filter_rows()
        for rule in scheme.get("filter_rules", []):
            self.add_filter_rule_row(
                default_column=rule.get("column") or self.suggest_filter_column(),
                default_operator=rule.get("operator", "equals"),
                default_value=rule.get("value", ""),
            )
        self.filter_preview_list.delete(0, tk.END)
        self.filter_preview_summary.configure(text="尚未测试规则。")
        self.on_amount_mode_changed()
        self._loading_scheme = False

    def _save_current_scheme(self):
        if self._loading_scheme or self.active_scheme_index < 0 or self.active_scheme_index >= len(self.schemes):
            return
        scheme = self.schemes[self.active_scheme_index]
        scheme["name"] = self.scheme_name.get().strip() or f"方案{self.active_scheme_index + 1}"
        scheme["amount_mode"] = self.amount_mode_code(self.amount_mode.get())
        scheme["amount_column"] = self.amount_column.get()
        scheme["manual_amount"] = self.manual_amount.get().strip()
        scheme["allocation_column"] = self.allocation_column.get()
        scheme["base_columns"] = [self.header_labels()[index] for index in self.base_list.curselection()]
        scheme["filter_logic"] = self.filter_logic.get() or "OR"
        scheme["filter_rules"] = [row.to_draft() for row in self.filter_rule_rows]
        self._refresh_scheme_list()
        self.scheme_list.selection_set(self.active_scheme_index)

    def on_amount_mode_changed(self, _event=None):
        mode = self.amount_mode_code(self.amount_mode.get())
        if mode == "manual":
            self.amount_column_combo.configure(state="disabled")
            self.manual_amount_entry.configure(state="normal")
        elif mode == "source_column":
            self.amount_column_combo.configure(state="readonly")
            self.manual_amount_entry.configure(state="disabled")
        else:
            self.amount_column_combo.configure(state="disabled")
            self.manual_amount_entry.configure(state="disabled")

    def auto_select_base_columns(self):
        self.base_list.selection_clear(0, tk.END)
        selected = set(self.suggest_base_columns())
        for idx, label in enumerate(self.header_labels()):
            if label in selected:
                self.base_list.selection_set(idx)

    def add_filter_rule_row(self, default_column=None, default_operator="equals", default_value=""):
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先读取表头。")
            return
        row = FilterRuleRow(
            self,
            self.rule_inner,
            default_column=default_column or self.suggest_filter_column(),
            default_operator=default_operator,
            default_value=default_value,
        )
        row.pack()
        self.filter_rule_rows.append(row)
        self._sync_rule_scrollregion()

    def remove_filter_rule_row(self, row):
        if row in self.filter_rule_rows:
            self.filter_rule_rows.remove(row)
            row.destroy()
            self._sync_rule_scrollregion()

    def _clear_filter_rows(self):
        for row in list(self.filter_rule_rows):
            row.destroy()
        self.filter_rule_rows = []
        self._sync_rule_scrollregion()

    def test_filter_rules(self):
        try:
            rules = [row.to_rule() for row in self.filter_rule_rows]
            count, samples = preview_filter_matches(
                self.file_path.get(),
                self.sheet_name.get(),
                int(self.header_row.get()),
                rules,
                self.filter_logic.get(),
            )
        except Exception as exc:
            messagebox.showerror("测试失败", str(exc))
            return

        self.filter_preview_list.delete(0, tk.END)
        for row_no, reason in samples:
            self.filter_preview_list.insert(tk.END, f"第 {row_no} 行 | {reason}")
        self.filter_preview_summary.configure(text=f"命中 {count} 行，显示前 {len(samples)} 条。")

    def preview_allocation(self):
        try:
            config = self._build_batch_config()
        except Exception as exc:
            messagebox.showwarning("配置不完整", str(exc))
            return

        self.status_text.set("正在生成预览...")
        self.run_button.configure(state=tk.DISABLED)

        def worker():
            try:
                result = preview_workbook_batch(config)
            except Exception as exc:
                error = f"{exc}\n\n{traceback.format_exc()}"
                self.after(0, self._preview_failed, error)
                return
            self.after(0, self._preview_finished, result)

        threading.Thread(target=worker, daemon=True).start()

    def _preview_finished(self, result):
        self.run_button.configure(state=tk.NORMAL)
        self._fill_preview_tree(result)
        self.notebook.select(self.preview_tab)
        self.status_text.set("预览完成，请核对后执行分摊。")

    def _preview_failed(self, error):
        self.run_button.configure(state=tk.NORMAL)
        self.status_text.set("预览失败。")
        messagebox.showerror("预览失败", error)

    def run_allocation(self):
        try:
            config = self._build_batch_config()
        except Exception as exc:
            messagebox.showwarning("配置不完整", str(exc))
            return

        self.run_button.configure(state=tk.DISABLED)
        self.status_text.set("正在分摊，请稍候...")

        def worker():
            try:
                result = allocate_workbook_batch(config)
            except Exception as exc:
                error = f"{exc}\n\n{traceback.format_exc()}"
                self.after(0, self._allocation_failed, error)
                return
            self.after(0, self._allocation_finished, result)

        threading.Thread(target=worker, daemon=True).start()

    def _allocation_finished(self, result):
        self.run_button.configure(state=tk.NORMAL)
        self._fill_preview_tree(result)
        self.notebook.select(self.preview_tab)
        summary = "\n".join(
            f"{item.name}：金额 {item.target_total}，参与 {item.participating_rows} 行，不参与 {item.excluded_rows} 行"
            for item in result.scheme_results
        )
        self.status_text.set(f"完成：{result.output_path}")
        messagebox.showinfo(
            "分摊完成",
            "分摊完成。\n\n"
            f"输出文件：{result.output_path}\n\n"
            f"{summary}",
        )

    def _allocation_failed(self, error):
        self.run_button.configure(state=tk.NORMAL)
        self.status_text.set("分摊失败。")
        messagebox.showerror("分摊失败", error)

    def _fill_preview_tree(self, result):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        for item in result.scheme_results:
            self.preview_tree.insert(
                "",
                tk.END,
                values=(
                    item.name,
                    self.result_source_label(item.amount_source, item.amount_column),
                    self.column_label_short(item.allocation_column),
                    f"{item.target_total}",
                    item.participating_rows,
                    item.excluded_rows,
                    f"{item.base_total}",
                    f"{item.distributed_total}",
                ),
            )

    def _build_batch_config(self):
        self._save_current_scheme()
        scheme_configs = []
        for index, scheme in enumerate(self.schemes, start=1):
            name = scheme.get("name", "").strip() or f"方案{index}"
            allocation_col = self.label_to_column_index(scheme.get("allocation_column", ""))
            if allocation_col is None:
                raise ValueError(f"{name}：请选择分摊结果列。")

            base_columns = [self.label_to_column_index(label) for label in scheme.get("base_columns", [])]
            if any(col is None for col in base_columns):
                raise ValueError(f"{name}：占比计算列选择无效，请重新选择。")

            amount_mode = scheme.get("amount_mode", "target_total")
            amount_source = "manual" if amount_mode == "manual" else "column_total"
            amount_column = None
            manual_amount = None
            if amount_mode == "source_column":
                amount_column = self.label_to_column_index(scheme.get("amount_column", ""))
                if amount_column is None:
                    raise ValueError(f"{name}：请选择金额来源列。")
            elif amount_mode == "manual":
                manual_amount = self._parse_manual_amount(scheme.get("manual_amount", ""), name)

            filter_rules = self._build_rules_from_draft(name, scheme.get("filter_rules", []))
            scheme_configs.append(
                AllocationScheme(
                    name=name,
                    amount_source=amount_source,
                    amount_column=amount_column,
                    manual_amount=manual_amount,
                    allocation_column=allocation_col,
                    base_columns=[col for col in base_columns if col is not None],
                    filter_rules=filter_rules,
                    filter_logic=scheme.get("filter_logic", "OR"),
                )
            )

        return BatchAllocationConfig(
            input_path=self.file_path.get(),
            output_path=self.output_path.get(),
            sheet_name=self.sheet_name.get(),
            header_row=int(self.header_row.get()),
            schemes=scheme_configs,
        )

    def _build_rules_from_draft(self, scheme_name, drafts):
        rules = []
        for draft in drafts:
            column_label = draft.get("column", "")
            if not column_label:
                continue
            column = self.label_to_column_index(column_label)
            if column is None:
                raise ValueError(f"{scheme_name}：过滤条件字段无效，请重新选择。")
            operator = FilterRuleRow.operator_code(draft.get("operator", "equals"))
            value = (draft.get("value") or "").strip()
            if operator not in {"blank", "not_blank"} and value == "":
                raise ValueError(f"{scheme_name}：过滤条件的值不能为空；不需要过滤时请删除该规则。")
            rules.append(FilterRule(column=column, operator=operator, value=value))
        return rules

    def _parse_manual_amount(self, value, scheme_name):
        text = (value or "").strip().replace(",", "").replace("，", "").replace("￥", "").replace("¥", "")
        if not text:
            raise ValueError(f"{scheme_name}：请输入手工金额。")
        try:
            return Decimal(text)
        except InvalidOperation:
            raise ValueError(f"{scheme_name}：手工金额格式不正确。")

    def header_labels(self):
        return [item.label for item in self.headers]

    def label_to_column_index(self, label):
        for item in self.headers:
            if item.label == label:
                return item.index
        return None

    def column_label_short(self, index):
        for item in self.headers:
            if item.index == index:
                return item.label
        return f"{index}列"

    def suggest_base_columns(self):
        base_keywords = ("完工入库材料成本", "本期人工费", "材料成本", "人工费")
        selected = [
            header.label
            for header in self.headers
            if any(keyword in header.header for keyword in base_keywords)
        ]
        if selected:
            return selected
        return [self.headers[0].label] if self.headers else []

    def suggest_target_column(self):
        target_keywords = ("共耗料", "水电费", "维修费", "折旧", "租赁费", "运费")
        for header in self.headers:
            if any(keyword in header.header for keyword in target_keywords):
                return header.label
        return self.headers[-1].label if self.headers else ""

    def suggest_filter_column(self):
        filter_keywords = ("生产车间", "车间", "部门")
        for header in self.headers:
            if any(keyword in header.header for keyword in filter_keywords):
                return header.label
        return self.headers[0].label if self.headers else ""

    def cached_unique_values(self, column_index):
        if column_index not in self.unique_value_cache:
            self.unique_value_cache[column_index] = get_unique_values(
                self.file_path.get(),
                self.sheet_name.get(),
                int(self.header_row.get()),
                column_index,
            )
        return self.unique_value_cache[column_index]

    @staticmethod
    def amount_mode_code(label_or_code):
        for code, label in AMOUNT_MODES:
            if label_or_code in {code, label}:
                return code
        return label_or_code or "target_total"

    @staticmethod
    def amount_mode_label(code):
        for item_code, label in AMOUNT_MODES:
            if item_code == code:
                return label
        return code

    @staticmethod
    def result_source_label(amount_source, amount_column):
        if amount_source == "manual":
            return "手工输入"
        if amount_column:
            return "指定列合计"
        return "分摊列合计"


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
公式：
单行占比 = 当前行占比列合计 / 所有参与行占比列合计
单行分摊金额 = 待分摊金额 × 单行占比

五、过滤条件
过滤条件表示“不参与分摊”的行。
例如：车间 等于 销售配货部，则销售配货部不参与占比，也不会被写入分摊金额。

规则关系：
OR：任一条件满足就不参与分摊。
AND：所有条件同时满足才不参与分摊。

条件类型：
等于、不等于、包含、不包含、正则、为空、非空。

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
执行后会生成：
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


def main():
    app = AllocatorApp()
    app.mainloop()
