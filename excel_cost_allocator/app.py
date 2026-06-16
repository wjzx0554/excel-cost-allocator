import os
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .allocator import (
    AllocationConfig,
    allocate_workbook,
    display_value,
    get_headers,
    get_sheet_names,
    get_unique_values,
    guess_header_row,
)


class AllocatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 费用自动分摊工具")
        self.geometry("980x680")
        self.minsize(900, 620)

        self.file_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.header_row = tk.IntVar(value=1)
        self.filter_column = tk.StringVar()
        self.status_text = tk.StringVar(value="请选择 Excel 文件。")

        self.headers = []
        self.filter_values = []

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(root, text="1. 选择文件和工作表", padding=8)
        file_frame.pack(fill=tk.X)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Excel 文件").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(file_frame, textvariable=self.file_path).grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(file_frame, text="浏览", command=self.choose_file).grid(row=0, column=2, padx=(8, 0))

        ttk.Label(file_frame, text="工作表").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=(8, 0))
        self.sheet_combo = ttk.Combobox(file_frame, textvariable=self.sheet_name, state="readonly")
        self.sheet_combo.grid(row=1, column=1, sticky=tk.W, pady=(8, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_changed)

        ttk.Label(file_frame, text="表头行").grid(row=1, column=2, sticky=tk.E, padx=(8, 4), pady=(8, 0))
        ttk.Spinbox(
            file_frame,
            from_=1,
            to=999,
            width=8,
            textvariable=self.header_row,
            command=self.load_headers,
        ).grid(row=1, column=3, sticky=tk.W, pady=(8, 0))
        ttk.Button(file_frame, text="读取表头", command=self.load_headers).grid(row=1, column=4, padx=(8, 0), pady=(8, 0))

        select_frame = ttk.Frame(root)
        select_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        select_frame.columnconfigure(0, weight=1)
        select_frame.columnconfigure(1, weight=1)

        base_frame = ttk.LabelFrame(select_frame, text="2. 选择参与占比计算的列", padding=8)
        base_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 5))
        base_frame.rowconfigure(0, weight=1)
        base_frame.columnconfigure(0, weight=1)
        self.base_list = tk.Listbox(base_frame, selectmode=tk.EXTENDED, exportselection=False)
        self.base_list.grid(row=0, column=0, sticky=tk.NSEW)
        base_scroll = ttk.Scrollbar(base_frame, orient=tk.VERTICAL, command=self.base_list.yview)
        base_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.base_list.configure(yscrollcommand=base_scroll.set)

        target_frame = ttk.LabelFrame(select_frame, text="3. 选择需要分配的费用列", padding=8)
        target_frame.grid(row=0, column=1, sticky=tk.NSEW, padx=(5, 0))
        target_frame.rowconfigure(0, weight=1)
        target_frame.columnconfigure(0, weight=1)
        self.target_list = tk.Listbox(target_frame, selectmode=tk.EXTENDED, exportselection=False)
        self.target_list.grid(row=0, column=0, sticky=tk.NSEW)
        target_scroll = ttk.Scrollbar(target_frame, orient=tk.VERTICAL, command=self.target_list.yview)
        target_scroll.grid(row=0, column=1, sticky=tk.NS)
        self.target_list.configure(yscrollcommand=target_scroll.set)

        filter_frame = ttk.LabelFrame(root, text="4. 设置不参与计算的过滤条件", padding=8)
        filter_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        filter_frame.columnconfigure(1, weight=1)
        filter_frame.columnconfigure(3, weight=1)
        filter_frame.rowconfigure(1, weight=1)

        ttk.Label(filter_frame, text="过滤列").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        self.filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_column, state="readonly")
        self.filter_combo.grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(filter_frame, text="读取过滤值", command=self.load_filter_values).grid(row=0, column=2, padx=(8, 12))
        ttk.Label(filter_frame, text="手工补充值，一行一个").grid(row=0, column=3, sticky=tk.W)

        self.filter_list = tk.Listbox(filter_frame, selectmode=tk.EXTENDED, exportselection=False, height=7)
        self.filter_list.grid(row=1, column=0, columnspan=3, sticky=tk.NSEW, pady=(8, 0), padx=(0, 12))
        filter_scroll = ttk.Scrollbar(filter_frame, orient=tk.VERTICAL, command=self.filter_list.yview)
        filter_scroll.grid(row=1, column=2, sticky="nse", pady=(8, 0))
        self.filter_list.configure(yscrollcommand=filter_scroll.set)

        self.manual_values = tk.Text(filter_frame, height=7, width=32)
        self.manual_values.grid(row=1, column=3, sticky=tk.NSEW, pady=(8, 0))
        self.manual_values.insert("1.0", "销售配货部")

        output_frame = ttk.LabelFrame(root, text="5. 输出文件", padding=8)
        output_frame.pack(fill=tk.X, pady=(10, 0))
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="输出路径").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(output_frame, textvariable=self.output_path).grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(output_frame, text="另存为", command=self.choose_output).grid(row=0, column=2, padx=(8, 0))

        action_frame = ttk.Frame(root)
        action_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(action_frame, textvariable=self.status_text).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.run_button = ttk.Button(action_frame, text="开始分摊", command=self.run_allocation)
        self.run_button.pack(side=tk.RIGHT)

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
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
            self.status_text.set("文件读取完成，请选择列和过滤条件。")
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))

    def _set_default_output(self, path):
        input_path = Path(path)
        suffix = ".xlsm" if input_path.suffix.lower() == ".xlsm" else ".xlsx"
        output = input_path.with_name(f"{input_path.stem}_分摊结果{suffix}")
        self.output_path.set(str(output))

    def choose_output(self):
        initial = self.output_path.get() or self.file_path.get()
        suffix = ".xlsm" if initial.lower().endswith(".xlsm") else ".xlsx"
        path = filedialog.asksaveasfilename(
            title="保存分摊结果",
            defaultextension=suffix,
            initialfile=Path(initial).name if initial else "",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
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
        labels = [item.label for item in self.headers]

        self.base_list.delete(0, tk.END)
        self.target_list.delete(0, tk.END)
        self.filter_combo["values"] = labels
        self.filter_column.set("")
        self.filter_list.delete(0, tk.END)
        self.filter_values = []

        for label in labels:
            self.base_list.insert(tk.END, label)
            self.target_list.insert(tk.END, label)

        self._auto_select_columns()
        self.status_text.set(f"已读取 {len(labels)} 个表头。")

    def _auto_select_columns(self):
        base_keywords = ("完工入库材料成本", "本期人工费", "材料成本", "人工费")
        target_keywords = ("共耗料", "水电费", "维修费", "折旧", "租赁费")
        filter_keywords = ("生产车间", "车间", "部门")
        for idx, header in enumerate(self.headers):
            text = header.header
            if any(keyword in text for keyword in base_keywords):
                self.base_list.selection_set(idx)
            if any(keyword in text for keyword in target_keywords):
                self.target_list.selection_set(idx)
            if not self.filter_column.get() and any(keyword in text for keyword in filter_keywords):
                self.filter_column.set(header.label)

    def load_filter_values(self):
        column_index = self._selected_filter_column()
        if column_index is None:
            messagebox.showwarning("缺少过滤列", "请先选择过滤列。")
            return
        try:
            values = get_unique_values(
                self.file_path.get(),
                self.sheet_name.get(),
                int(self.header_row.get()),
                column_index,
            )
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return

        self.filter_values = values
        self.filter_list.delete(0, tk.END)
        for value in values:
            self.filter_list.insert(tk.END, display_value(value))
        for idx, value in enumerate(values):
            if value == "销售配货部":
                self.filter_list.selection_set(idx)
        self.status_text.set(f"已读取 {len(values)} 个过滤值。")

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
        base_columns = self._selected_columns(self.base_list)
        target_columns = self._selected_columns(self.target_list)
        excluded_values = set(self._selected_filter_values())
        excluded_values.update(self._manual_filter_values())

        return AllocationConfig(
            input_path=self.file_path.get(),
            output_path=self.output_path.get(),
            sheet_name=self.sheet_name.get(),
            header_row=int(self.header_row.get()),
            base_columns=base_columns,
            allocation_columns=target_columns,
            filter_column=self._selected_filter_column(),
            excluded_values=excluded_values,
        )

    def _selected_columns(self, listbox):
        return [self.headers[index].index for index in listbox.curselection()]

    def _selected_filter_column(self):
        label = self.filter_column.get()
        if not label:
            return None
        for item in self.headers:
            if item.label == label:
                return item.index
        return None

    def _selected_filter_values(self):
        result = []
        for index in self.filter_list.curselection():
            if 0 <= index < len(self.filter_values):
                result.append(self.filter_values[index])
        return result

    def _manual_filter_values(self):
        text = self.manual_values.get("1.0", tk.END)
        parts = []
        for line in text.replace("，", "\n").replace(",", "\n").replace(";", "\n").splitlines():
            value = line.strip()
            if value:
                parts.append(value)
        return parts


def main():
    app = AllocatorApp()
    app.mainloop()
