import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .allocator import (
    AllocationConfig,
    FilterRule,
    allocate_workbook,
    display_value,
    get_headers,
    get_sheet_names,
    guess_header_row,
    preview_filter_matches,
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


class FilterRuleRow:
    def __init__(self, app, parent, default_column="", default_operator="equals", default_value=""):
        self.app = app
        self.frame = ttk.Frame(parent, style="Rule.TFrame", padding=(8, 6))
        self.frame.columnconfigure(1, weight=1)
        self.frame.columnconfigure(5, weight=1)

        self.column_var = tk.StringVar(value=default_column)
        self.operator_var = tk.StringVar(value=default_operator)
        self.value_var = tk.StringVar(value=default_value)

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
        self.operator_combo.set(self.operator_label(default_operator))
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

        self.refresh_column_values()

    def pack(self):
        self.frame.pack(fill=tk.X, pady=4)

    def destroy(self):
        self.frame.destroy()

    def to_rule(self):
        column_index = self.app.label_to_column_index(self.column_var.get())
        if column_index is None:
            raise ValueError("过滤条件中的字段选择无效。")

        operator = self.operator_code(self.operator_combo.get())
        value = self.value_var.get().strip()
        if value == "<空白>":
            value = ""
        if operator not in {"blank", "not_blank"} and not value:
            raise ValueError("过滤条件中的值不能为空。")

        return FilterRule(column=column_index, operator=operator, value=value)

    def refresh_headers(self):
        labels = self.app.header_labels()
        self.column_combo.configure(values=labels)
        if self.column_var.get() not in labels:
            self.column_var.set(self.app.suggest_filter_column())
        self.refresh_column_values()

    def refresh_column_values(self):
        self._load_values()
        self._update_value_state()

    def _column_selected(self, _event=None):
        self.refresh_column_values()

    def _operator_selected(self, _event=None):
        self.operator_var.set(self.operator_code(self.operator_combo.get()))
        self._update_value_state()

    def _load_values(self, _event=None):
        column_index = self.app.label_to_column_index(self.column_var.get())
        if column_index is None:
            self.value_combo.configure(values=[])
            return
        values = self.app.cached_unique_values(column_index)
        self.value_combo.configure(values=[display_value(value) for value in values])

    def _update_value_state(self):
        operator = self.operator_code(self.operator_combo.get())
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
        self.geometry("1120x760")
        self.minsize(1020, 700)
        self.configure(bg="#f3f6fb")
        self._set_window_icon()
        self._setup_style()

        self.file_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.header_row = tk.IntVar(value=1)
        self.filter_logic = tk.StringVar(value="OR")
        self.status_text = tk.StringVar(value="请选择 Excel 文件。")

        self.headers = []
        self.filter_rule_rows = []
        self.unique_value_cache = {}

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
        self.style.configure("App.TFrame", background="#f3f6fb")
        self.style.configure("Hero.TFrame", background="#173b78")
        self.style.configure("Card.TLabelframe", background="#ffffff", bordercolor="#d8e0ee", relief="solid")
        self.style.configure("Card.TLabelframe.Label", background="#ffffff", foreground="#0f172a", font=self.section_font)
        self.style.configure("Rule.TFrame", background="#ffffff")
        self.style.configure("TLabel", background="#f3f6fb", foreground="#334155")
        self.style.configure("Card.TLabel", background="#ffffff", foreground="#334155")
        self.style.configure("Rule.TLabel", background="#ffffff", foreground="#334155")
        self.style.configure("Hint.TLabel", background="#ffffff", foreground="#64748b")
        self.style.configure("HeroTitle.TLabel", background="#173b78", foreground="#ffffff", font=self.title_font)
        self.style.configure("HeroSub.TLabel", background="#173b78", foreground="#dbeafe", font=self.subtitle_font)
        self.style.configure("Status.TLabel", background="#eef4ff", foreground="#1e3a8a", padding=(10, 6))
        self.style.configure("TButton", padding=(12, 6), background="#e2e8f0", foreground="#0f172a")
        self.style.map("TButton", background=[("active", "#cbd5e1")])
        self.style.configure("Primary.TButton", padding=(16, 8), background="#1d4ed8", foreground="#ffffff", font=("Microsoft YaHei UI", 10, "bold"))
        self.style.map("Primary.TButton", background=[("active", "#1e40af"), ("disabled", "#93c5fd")])
        self.style.configure("Danger.TButton", padding=(10, 5), background="#fee2e2", foreground="#991b1b")
        self.style.map("Danger.TButton", background=[("active", "#fecaca")])

    def _build_ui(self):
        root = ttk.Frame(self, padding=12, style="App.TFrame")
        root.pack(fill=tk.BOTH, expand=True)

        hero = ttk.Frame(root, padding=(18, 16), style="Hero.TFrame")
        hero.pack(fill=tk.X)
        ttk.Label(hero, text="分摊工具", style="HeroTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(
            hero,
            text="Excel/WPS 费用分摊、规则过滤、明细核对，一次配置后快速生成结果文件",
            style="HeroSub.TLabel",
        ).pack(anchor=tk.W, pady=(4, 0))

        file_frame = ttk.LabelFrame(root, text="1. 文件与表头", padding=10, style="Card.TLabelframe")
        file_frame.pack(fill=tk.X, pady=(10, 0))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Excel 文件", style="Card.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(file_frame, textvariable=self.file_path).grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(file_frame, text="浏览", command=self.choose_file).grid(row=0, column=2, padx=(8, 0))

        ttk.Label(file_frame, text="工作表", style="Card.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=(8, 0))
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="readonly")
        self.sheet_combo.grid(row=1, column=1, sticky=tk.W, pady=(8, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_changed)

        ttk.Label(file_frame, text="表头行", style="Card.TLabel").grid(row=1, column=2, sticky=tk.E, padx=(8, 4), pady=(8, 0))
        ttk.Spinbox(file_frame, from_=1, to=999, width=8, textvariable=self.header_row, command=self.load_headers).grid(row=1, column=3, sticky=tk.W, pady=(8, 0))
        ttk.Button(file_frame, text="读取表头", command=self.load_headers).grid(row=1, column=4, padx=(8, 0), pady=(8, 0))

        select_frame = ttk.Frame(root, style="App.TFrame")
        select_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        select_frame.columnconfigure(0, weight=1)
        select_frame.columnconfigure(1, weight=1)

        base_frame = ttk.LabelFrame(select_frame, text="2. 参与占比计算的列", padding=10, style="Card.TLabelframe")
        base_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 5))
        base_frame.rowconfigure(0, weight=1)
        base_frame.columnconfigure(0, weight=1)
        self.base_list = self._create_listbox(base_frame)
        self.base_list.grid(row=0, column=0, sticky=tk.NSEW)
        base_scroll = ttk.Scrollbar(base_frame, orient=tk.VERTICAL, command=self.base_list.yview)
        base_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.base_list.configure(yscrollcommand=base_scroll.set)

        target_frame = ttk.LabelFrame(select_frame, text="3. 需要分配的费用列", padding=10, style="Card.TLabelframe")
        target_frame.grid(row=0, column=1, sticky=tk.NSEW, padx=(5, 0))
        target_frame.rowconfigure(0, weight=1)
        target_frame.columnconfigure(0, weight=1)
        self.target_list = self._create_listbox(target_frame)
        self.target_list.grid(row=0, column=0, sticky=tk.NSEW)
        target_scroll = ttk.Scrollbar(target_frame, orient=tk.VERTICAL, command=self.target_list.yview)
        target_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.target_list.configure(yscrollcommand=target_scroll.set)

        filter_frame = ttk.LabelFrame(root, text="4. 不参与计算的过滤规则", padding=10, style="Card.TLabelframe")
        filter_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
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
        self.rule_canvas = tk.Canvas(rules_panel, bg="#f3f6fb", highlightthickness=0, borderwidth=0)
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
        ttk.Label(rules_panel, text=regex_help, style="Hint.TLabel", wraplength=660).grid(row=1, column=0, sticky=tk.W, pady=(6, 0))

        preview_panel = ttk.LabelFrame(filter_frame, text="命中预览", padding=10, style="Card.TLabelframe")
        preview_panel.grid(row=1, column=1, sticky=tk.NSEW, pady=(8, 0))
        preview_panel.columnconfigure(0, weight=1)
        preview_panel.rowconfigure(1, weight=1)
        self.preview_summary = ttk.Label(preview_panel, text="尚未测试规则。", style="Card.TLabel")
        self.preview_summary.grid(row=0, column=0, sticky=tk.W)
        self.preview_list = self._create_listbox(preview_panel, height=8)
        self.preview_list.grid(row=1, column=0, sticky=tk.NSEW, pady=(8, 0))
        preview_scroll = ttk.Scrollbar(preview_panel, orient=tk.VERTICAL, command=self.preview_list.yview)
        preview_scroll.grid(row=1, column=1, sticky=tk.NS, pady=(8, 0))
        self.preview_list.configure(yscrollcommand=preview_scroll.set)

        output_frame = ttk.LabelFrame(root, text="5. 输出文件", padding=10, style="Card.TLabelframe")
        output_frame.pack(fill=tk.X, pady=(10, 0))
        output_frame.columnconfigure(1, weight=1)
        ttk.Label(output_frame, text="输出路径", style="Card.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(output_frame, textvariable=self.output_path).grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(output_frame, text="另存为", command=self.choose_output).grid(row=0, column=2, padx=(8, 0))

        action_frame = ttk.Frame(root, style="App.TFrame")
        action_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(action_frame, textvariable=self.status_text, style="Status.TLabel").pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.run_button = ttk.Button(action_frame, text="开始分摊", command=self.run_allocation, style="Primary.TButton")
        self.run_button.pack(side=tk.RIGHT)

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
            self.status_text.set("文件读取完成，请选择列和过滤规则。")
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
        labels = self.header_labels()
        self.unique_value_cache.clear()

        self.base_list.delete(0, tk.END)
        self.target_list.delete(0, tk.END)
        for label in labels:
            self.base_list.insert(tk.END, label)
            self.target_list.insert(tk.END, label)

        self._auto_select_columns()
        self._reset_rule_rows()
        self.status_text.set(f"已读取 {len(labels)} 个表头。")

    def _auto_select_columns(self):
        base_keywords = ("完工入库材料成本", "本期人工费", "材料成本", "人工费")
        target_keywords = ("共耗料", "水电费", "维修费", "折旧", "租赁费")
        for idx, header in enumerate(self.headers):
            text = header.header
            if any(keyword in text for keyword in base_keywords):
                self.base_list.selection_set(idx)
            if any(keyword in text for keyword in target_keywords):
                self.target_list.selection_set(idx)

    def _reset_rule_rows(self):
        for row in list(self.filter_rule_rows):
            row.destroy()
        self.filter_rule_rows = []
        if self.headers:
            self.add_filter_rule_row(default_value="销售配货部")
        self.preview_list.delete(0, tk.END)
        self.preview_summary.configure(text="尚未测试规则。")

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
        if len(self.filter_rule_rows) <= 1:
            messagebox.showwarning("提示", "至少保留一条规则。")
            return
        self.filter_rule_rows.remove(row)
        row.destroy()
        self._sync_rule_scrollregion()

    def test_filter_rules(self):
        try:
            rules = self._build_filter_rules()
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

        self.preview_list.delete(0, tk.END)
        for row_no, reason in samples:
            self.preview_list.insert(tk.END, f"第 {row_no} 行 | {reason}")
        self.preview_summary.configure(text=f"命中 {count} 行，显示前 {len(samples)} 条。")

    def run_allocation(self):
        try:
            config = self._build_config()
        except Exception as exc:
            messagebox.showwarning("配置不完整", str(exc))
            return

        self.run_button.configure(state=tk.DISABLED)
        self.status_text.set("正在分摊，请稍候...")

        def worker():
            try:
                result = allocate_workbook(config)
            except Exception as exc:
                error = f"{exc}\n\n{traceback.format_exc()}"
                self.after(0, self._allocation_failed, error)
                return
            self.after(0, self._allocation_finished, result)

        threading.Thread(target=worker, daemon=True).start()

    def _allocation_finished(self, result):
        self.run_button.configure(state=tk.NORMAL)
        self.status_text.set(f"完成：{result.output_path}")
        messagebox.showinfo(
            "分摊完成",
            "分摊完成。\n\n"
            f"输出文件：{result.output_path}\n"
            f"数据行数：{result.total_rows}\n"
            f"参与行数：{result.participating_rows}\n"
            f"不参与行数：{result.excluded_rows}\n"
            f"分摊基数合计：{result.base_total}",
        )

    def _allocation_failed(self, error):
        self.run_button.configure(state=tk.NORMAL)
        self.status_text.set("分摊失败。")
        messagebox.showerror("分摊失败", error)

    def _build_config(self):
        return AllocationConfig(
            input_path=self.file_path.get(),
            output_path=self.output_path.get(),
            sheet_name=self.sheet_name.get(),
            header_row=int(self.header_row.get()),
            base_columns=self._selected_columns(self.base_list),
            allocation_columns=self._selected_columns(self.target_list),
            filter_rules=self._build_filter_rules(),
            filter_logic=self.filter_logic.get(),
        )

    def _build_filter_rules(self):
        if not self.filter_rule_rows:
            return []
        return [row.to_rule() for row in self.filter_rule_rows]

    def _selected_columns(self, listbox):
        return [self.headers[index].index for index in listbox.curselection()]

    def header_labels(self):
        return [item.label for item in self.headers]

    def label_to_column_index(self, label):
        for item in self.headers:
            if item.label == label:
                return item.index
        return None

    def suggest_filter_column(self):
        filter_keywords = ("生产车间", "车间", "部门")
        for header in self.headers:
            if any(keyword in header.header for keyword in filter_keywords):
                return header.label
        return self.headers[0].label if self.headers else ""

    def cached_unique_values(self, column_index):
        if column_index not in self.unique_value_cache:
            from .allocator import get_unique_values

            self.unique_value_cache[column_index] = get_unique_values(
                self.file_path.get(),
                self.sheet_name.get(),
                int(self.header_row.get()),
                column_index,
            )
        return self.unique_value_cache[column_index]


def main():
    app = AllocatorApp()
    app.mainloop()
