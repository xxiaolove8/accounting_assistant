"""Chinese Tkinter interface for the Excel filtering and accounting tool."""

from __future__ import annotations

import queue
import threading
import os
import re
import tkinter as tk
from dataclasses import replace
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Callable

from accounting import AccountingConfig
from analysis_core import AnalysisMetric, AnalysisOptions, CalculatedVariable, analyze_to_workbook
from filter_core import (
    BatchJob,
    CancelledError,
    ExportOptions,
    FilterRule,
    export_batch,
    export_filtered,
    export_named_batch,
    inspect_headers,
    list_sheets,
    scan_values,
)
from merge_core import MergeOptions, merge_workbooks


_EMPTY_LABEL = "（空白）"
_SCAN_LIMIT = 10_000


class ExcelFilterApp(tk.Tk):
    """Single-window desktop workflow; workbook operations run off the UI thread."""

    def __init__(self) -> None:
        super().__init__()
        self.title("财务 Excel 筛选、合并与分析工具")
        self.geometry("1160x800")
        self.minsize(920, 680)

        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.header_start_var = tk.StringVar(value="1")
        self.header_end_var = tk.StringVar(value="1")
        self.data_start_var = tk.StringVar()
        self.data_end_var = tk.StringVar()
        self.stop_summary_var = tk.BooleanVar(value=True)
        self.preserve_styles_var = tk.BooleanVar(value=False)
        self.column_var = tk.StringVar()
        self.operator_var = tk.StringVar(value="完全匹配")
        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择源 Excel 文件。")
        self.values_info_var = tk.StringVar(value="读取表头后，可扫描所选字段的值。")
        self.accounting_info_var = tk.StringVar(value="会计辅助：未启用")
        self.batch_tasks_info_var = tk.StringVar(value="已保存 0 套批量任务")
        self.batch_parent_var = tk.StringVar()

        self.headers: list[str] = []
        self.available_values: list[str] = []
        self.visible_values: list[str] = []
        self.selected_values: set[str] = set()
        self.rules: list[FilterRule] = []
        self.batch_jobs: list[BatchJob] = []
        self.accounting_config = AccountingConfig()
        self._last_batch_output_dir: Path | None = None
        self._active_operation = ""
        self._header_context: tuple[Path, str, int, int] | None = None

        self._events: queue.Queue[tuple] = queue.Queue()
        self._busy = False
        self._cancel_event: threading.Event | None = None
        self._success_handler: Callable[[object], None] | None = None
        self._closing_requested = False
        self._search_after: str | None = None
        self._progress_indeterminate = False
        self._controls: list[tuple[tk.Widget, str]] = []

        self._configure_style()
        self._build_ui()
        self.search_var.trace_add("write", self._schedule_value_search)
        self.after(100, self._drain_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 16, "bold"))
        style.configure("Hint.TLabel", foreground="#566575")
        style.configure("Status.TLabel", foreground="#24476B")
        style.configure("TButton", padding=(9, 5))

    def _track_control(self, widget: tk.Widget, normal_state: str = "normal") -> tk.Widget:
        self._controls.append((widget, normal_state))
        return widget

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)

        title = ttk.Frame(root)
        title.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(title, text="Excel 数据筛选与分析", style="Title.TLabel").pack(side="left")
        ttk.Label(
            title,
            text="支持筛选、拆分、合并、计算变量及图表分析",
            style="Hint.TLabel",
        ).pack(side="left", padx=(18, 0), pady=(5, 0))

        self._build_source_section(root)
        self._build_filter_section(root)
        self._build_footer(root)

    def _build_source_section(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(parent, text="1  文件与数据范围", padding=(12, 10))
        box.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="源文件").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        source = self._track_control(ttk.Entry(box, textvariable=self.source_var))
        source.grid(row=0, column=1, columnspan=7, sticky="ew", pady=3)
        source.bind("<Return>", lambda _event: self._load_sheets())
        source_browse = self._track_control(ttk.Button(box, text="浏览…", command=self._browse_source))
        source_browse.grid(row=0, column=8, padx=(8, 0), pady=3)
        load = self._track_control(ttk.Button(box, text="读取工作表", command=self._load_sheets))
        load.grid(row=0, column=9, padx=(8, 0), pady=3)

        ttk.Label(box, text="工作表").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        self.sheet_combo = self._track_control(
            ttk.Combobox(box, textvariable=self.sheet_var, state="readonly", width=24), "readonly"
        )
        self.sheet_combo.grid(row=1, column=1, sticky="ew", pady=3)
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _event: self._read_headers())
        ranges = ttk.Frame(box)
        ranges.grid(row=1, column=2, columnspan=7, sticky="e", padx=(12, 0))
        ttk.Label(ranges, text="表头行").pack(side="left", padx=(0, 5))
        self._number_pair(ranges, self.header_start_var, self.header_end_var)
        ttk.Label(ranges, text="明细行").pack(side="left", padx=(14, 5))
        self._number_pair(ranges, self.data_start_var, self.data_end_var)
        read = self._track_control(ttk.Button(box, text="读取表头", command=self._read_headers))
        read.grid(row=1, column=9, padx=(8, 0), pady=3)

        ttk.Label(box, text="输出文件").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=3)
        output = self._track_control(ttk.Entry(box, textvariable=self.output_var))
        output.grid(row=2, column=1, columnspan=7, sticky="ew", pady=3)
        output_browse = self._track_control(ttk.Button(box, text="另存为…", command=self._browse_output))
        output_browse.grid(row=2, column=8, columnspan=2, sticky="ew", padx=(8, 0), pady=3)

        checks = ttk.Frame(box)
        checks.grid(row=3, column=1, columnspan=9, sticky="w", pady=(5, 0))
        self._track_control(ttk.Checkbutton(
            checks, text="遇到合计 / 汇总行时停止", variable=self.stop_summary_var
        )).pack(side="left")
        self._track_control(ttk.Checkbutton(
            checks, text="尽量保留单元格格式（导出较慢）", variable=self.preserve_styles_var
        )).pack(side="left", padx=(20, 0))
        ttk.Label(box, text="留空的明细起始行为表头后第一行；结束行留空表示到工作表末尾。", style="Hint.TLabel").grid(
            row=4, column=1, columnspan=9, sticky="w", pady=(5, 0)
        )

    def _number_pair(
        self,
        parent: ttk.Frame,
        start_var: tk.StringVar,
        end_var: tk.StringVar,
    ) -> None:
        first = self._track_control(ttk.Entry(parent, textvariable=start_var, width=6))
        first.pack(side="left")
        ttk.Label(parent, text="至").pack(side="left", padx=4)
        last = self._track_control(ttk.Entry(parent, textvariable=end_var, width=6))
        last.pack(side="left")

    def _build_filter_section(self, parent: ttk.Frame) -> None:
        section = ttk.LabelFrame(parent, text="2  筛选条件", padding=(12, 10))
        section.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        section.columnconfigure(0, weight=1)
        section.columnconfigure(1, weight=1)
        section.rowconfigure(1, weight=1)

        ttk.Label(
            section,
            text="不同字段之间为“且”；同一字段选中的多个值为“或”。重新添加同一字段会替换原条件。",
            style="Hint.TLabel",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        selector = ttk.Frame(section)
        selector.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        selector.columnconfigure(0, weight=1)
        selector.rowconfigure(3, weight=1)

        column_row = ttk.Frame(selector)
        column_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        column_row.columnconfigure(1, weight=1)
        ttk.Label(column_row, text="字段").grid(row=0, column=0, padx=(0, 8))
        self.column_combo = self._track_control(
            ttk.Combobox(column_row, textvariable=self.column_var, state="readonly"), "readonly"
        )
        self.column_combo.grid(row=0, column=1, sticky="ew")
        self.column_combo.bind("<<ComboboxSelected>>", self._column_changed)
        self.scan_button = self._track_control(ttk.Button(column_row, text="扫描可选值", command=self._scan_values))
        self.scan_button.grid(row=0, column=2, padx=(8, 0))

        search_row = ttk.Frame(selector)
        search_row.grid(row=1, column=0, sticky="ew", pady=(0, 5))
        search_row.columnconfigure(1, weight=1)
        ttk.Label(search_row, text="查找值").grid(row=0, column=0, padx=(0, 8))
        self._track_control(ttk.Entry(search_row, textvariable=self.search_var)).grid(row=0, column=1, sticky="ew")
        ttk.Label(selector, textvariable=self.values_info_var, style="Hint.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, 4)
        )

        list_frame = ttk.Frame(selector)
        list_frame.grid(row=3, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.values_list = tk.Listbox(
            list_frame, selectmode=tk.MULTIPLE, exportselection=False, height=10,
            font=("Microsoft YaHei UI", 10), activestyle="none",
        )
        self.values_list.grid(row=0, column=0, sticky="nsew")
        self.values_list.bind("<<ListboxSelect>>", self._value_selection_changed)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.values_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.values_list.configure(yscrollcommand=scrollbar.set)
        self.split_current_button = self._track_control(ttk.Button(
            selector,
            text="按此字段的每个不同值分别导出 Excel…",
            command=self._export_each_value,
        ))
        self.split_current_button.grid(row=4, column=0, sticky="ew", pady=(9, 3))
        ttk.Label(
            selector,
            text="例如 7 个不同值就生成 7 个文件，无需勾选。超过 200 类可在“批量导出…”中调整上限。",
            style="Hint.TLabel",
            wraplength=520,
        ).grid(row=5, column=0, sticky="w")

        right = ttk.Frame(section)
        right.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(3, weight=1)
        mode_row = ttk.Frame(right)
        mode_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(mode_row, text="匹配方式").pack(side="left", padx=(0, 8))
        self.operator_combo = self._track_control(ttk.Combobox(
            mode_row, textvariable=self.operator_var, state="readonly",
            values=("完全匹配", "包含文字"), width=12,
        ), "readonly")
        self.operator_combo.pack(side="left")

        ttk.Label(right, text="也可手动输入筛选值，每行一个；与上方选中值合并。", style="Hint.TLabel").grid(
            row=1, column=0, sticky="w", pady=(0, 4)
        )
        self.manual_values = self._track_control(tk.Text(right, height=3, undo=True, font=("Microsoft YaHei UI", 10)))
        self.manual_values.grid(row=2, column=0, sticky="ew", pady=(0, 7))

        rule_frame = ttk.Frame(right)
        rule_frame.grid(row=3, column=0, sticky="nsew")
        rule_frame.columnconfigure(0, weight=1)
        rule_frame.rowconfigure(1, weight=1)
        row = ttk.Frame(rule_frame)
        row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        self._track_control(ttk.Button(row, text="添加 / 更新条件", command=self._add_rule)).pack(side="left")
        self._track_control(ttk.Button(row, text="删除选中条件", command=self._remove_rule)).pack(side="left", padx=(8, 0))
        self._track_control(ttk.Button(
            row, text="保存为批量任务…", command=self._save_batch_job
        )).pack(side="left", padx=(8, 0))
        self.rules_tree = ttk.Treeview(
            rule_frame, columns=("field", "mode", "values"), show="headings", height=7,
            selectmode="extended",
        )
        self.rules_tree.heading("field", text="字段")
        self.rules_tree.heading("mode", text="方式")
        self.rules_tree.heading("values", text="筛选值")
        self.rules_tree.column("field", width=130, minwidth=90)
        self.rules_tree.column("mode", width=75, minwidth=65, stretch=False)
        self.rules_tree.column("values", width=260, minwidth=100)
        self.rules_tree.grid(row=1, column=0, sticky="nsew")
        rule_scroll = ttk.Scrollbar(rule_frame, orient="vertical", command=self.rules_tree.yview)
        rule_scroll.grid(row=1, column=1, sticky="ns")
        self.rules_tree.configure(yscrollcommand=rule_scroll.set)

    def _build_footer(self, parent: ttk.Frame) -> None:
        footer = ttk.Frame(parent)
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(1, weight=1)

        accounting = self._track_control(ttk.Button(footer, text="会计辅助设置…", command=self._show_accounting_dialog))
        accounting.grid(row=0, column=0, sticky="w")
        ttk.Label(footer, textvariable=self.accounting_info_var, style="Hint.TLabel").grid(
            row=0, column=1, sticky="w", padx=(12, 6)
        )
        self.export_button = self._track_control(ttk.Button(footer, text="生成筛选结果", command=self._export))
        self.export_button.grid(row=0, column=2, padx=(8, 0))
        self.batch_button = self._track_control(ttk.Button(footer, text="批量导出…", command=self._show_batch_dialog))
        self.batch_button.grid(row=0, column=3, padx=(8, 0))
        self.merge_button = self._track_control(ttk.Button(footer, text="合并总表…", command=self._show_merge_dialog))
        self.merge_button.grid(row=0, column=4, padx=(8, 0))
        self.analysis_button = self._track_control(ttk.Button(
            footer, text="数据分析与图表…", command=self._show_analysis_dialog,
        ))
        self.analysis_button.grid(row=0, column=5, padx=(8, 0))
        self.cancel_button = ttk.Button(footer, text="取消操作", command=self._cancel, state="disabled")
        self.cancel_button.grid(row=0, column=6, padx=(8, 0))
        self.open_batch_folder_button = ttk.Button(
            footer, text="打开上次批量结果文件夹", command=self._open_batch_folder, state="disabled"
        )
        self.open_batch_folder_button.grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Label(footer, textvariable=self.batch_tasks_info_var, style="Hint.TLabel").grid(
            row=1, column=1, columnspan=6, sticky="w", padx=(12, 0), pady=(8, 0)
        )
        self.progress = ttk.Progressbar(footer, mode="determinate", maximum=100)
        self.progress.grid(row=2, column=0, columnspan=7, sticky="ew", pady=(10, 4))
        ttk.Label(footer, textvariable=self.status_var, style="Status.TLabel").grid(
            row=3, column=0, columnspan=7, sticky="w"
        )

    def _browse_source(self) -> None:
        path = filedialog.askopenfilename(
            parent=self, title="选择源 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.source_var.set(path)
        self.output_var.set(str(self._suggest_output(Path(path))))
        self.batch_parent_var.set(str(Path(path).parent))
        self._reset_workbook_state()
        self._load_sheets()

    def _suggest_output(self, source: Path) -> Path:
        base = source.with_name(source.stem + "_筛选结果.xlsx")
        if not base.exists():
            return base
        for suffix in range(2, 10000):
            candidate = source.with_name(f"{source.stem}_筛选结果_{suffix}.xlsx")
            if not candidate.exists():
                return candidate
        return base

    def _browse_output(self) -> None:
        suggested = Path(self.output_var.get().strip() or "筛选结果.xlsx")
        path = filedialog.asksaveasfilename(
            parent=self, title="保存筛选结果", defaultextension=".xlsx",
            initialdir=str(suggested.parent) if suggested.parent != Path(".") else None,
            initialfile=suggested.name,
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _reset_workbook_state(self) -> None:
        self.sheet_var.set("")
        self.sheet_combo.configure(values=())
        self._reset_headers()

    def _reset_headers(self) -> None:
        self.headers = []
        self.column_var.set("")
        self.column_combo.configure(values=())
        self.available_values = []
        self.visible_values = []
        self.selected_values.clear()
        self.values_list.delete(0, "end")
        self.values_info_var.set("读取表头后，可扫描所选字段的值。")
        self.rules.clear()
        self._refresh_rules()
        self.batch_jobs.clear()
        self._refresh_batch_tasks_info()
        self.accounting_config = AccountingConfig()
        self._header_context = None
        self._refresh_accounting_info()

    def _load_sheets(self) -> None:
        raw = self.source_var.get().strip()
        if not raw:
            messagebox.showwarning("缺少源文件", "请先选择源 Excel 文件。", parent=self)
            return
        source = Path(raw).expanduser()
        self._reset_workbook_state()
        self._run_operation("读取工作表", lambda _progress, _cancel: list_sheets(source), self._sheets_loaded)

    def _sheets_loaded(self, result: object) -> None:
        sheets = list(result)
        if not sheets:
            messagebox.showwarning("没有工作表", "源文件中没有可读取的工作表。", parent=self)
            return
        self.sheet_combo.configure(values=sheets)
        self.sheet_var.set(sheets[0])
        if not self.output_var.get().strip():
            self.output_var.set(str(self._suggest_output(Path(self.source_var.get().strip()))))
        self.status_var.set(f"找到 {len(sheets)} 个工作表，正在读取表头…")
        self._read_headers()

    @staticmethod
    def _positive_int(text: str, label: str, *, optional: bool = False) -> int | None:
        raw = text.strip()
        if optional and not raw:
            return None
        try:
            value = int(raw)
        except ValueError as exc:
            raise ValueError(f"{label}必须是正整数。") from exc
        if value < 1:
            raise ValueError(f"{label}必须是正整数。")
        return value

    def _layout(self) -> tuple[int, int, int | None, int | None]:
        header_start = self._positive_int(self.header_start_var.get(), "表头起始行")
        header_end = self._positive_int(self.header_end_var.get(), "表头结束行")
        data_start = self._positive_int(self.data_start_var.get(), "明细起始行", optional=True)
        data_end = self._positive_int(self.data_end_var.get(), "明细结束行", optional=True)
        assert header_start is not None and header_end is not None
        if header_end < header_start:
            raise ValueError("表头结束行不能早于起始行。")
        if data_start is not None and data_start <= header_end:
            raise ValueError("明细起始行必须在表头结束行之后。")
        if data_end is not None and data_end < (data_start or header_end + 1):
            raise ValueError("明细结束行不能早于明细起始行。")
        return header_start, header_end, data_start, data_end

    def _read_headers(self) -> None:
        if not self.source_var.get().strip() or not self.sheet_var.get().strip():
            messagebox.showwarning("缺少工作表", "请先读取源文件并选择工作表。", parent=self)
            return
        try:
            start, end, _, _ = self._layout()
        except ValueError as exc:
            messagebox.showerror("行号无效", str(exc), parent=self)
            return
        source = Path(self.source_var.get().strip()).expanduser()
        sheet = self.sheet_var.get()
        self._reset_headers()
        self._run_operation(
            "读取表头",
            lambda _progress, _cancel: inspect_headers(source, sheet, start, end),
            self._headers_loaded,
        )

    def _headers_loaded(self, result: object) -> None:
        self.headers = list(result)
        start, end, _, _ = self._layout()
        self._header_context = (
            Path(self.source_var.get().strip()).expanduser().resolve(), self.sheet_var.get(), start, end
        )
        self.column_combo.configure(values=self.headers)
        if self.headers:
            self.column_combo.current(0)
        self.values_info_var.set("点击“扫描可选值”；也可以直接手动输入。")
        self.status_var.set(f"已读取 {len(self.headers)} 个字段。")

    def _build_options(self, *, require_output: bool = True) -> ExportOptions:
        source_text = self.source_var.get().strip()
        output_text = self.output_var.get().strip()
        sheet = self.sheet_var.get().strip()
        if not source_text:
            raise ValueError("请选择源 Excel 文件。")
        if not sheet:
            raise ValueError("请选择工作表。")
        if require_output and not output_text:
            raise ValueError("请选择输出文件。")
        start, end, data_start, data_end = self._layout()
        source = Path(source_text).expanduser()
        return ExportOptions(
            source=source, sheet=sheet,
            output=Path(output_text).expanduser() if output_text else self._suggest_output(source),
            header_start=start, header_end=end,
            data_start=data_start, data_end=data_end,
            stop_at_summary=self.stop_summary_var.get(),
            preserve_styles=self.preserve_styles_var.get(),
            accounting=self.accounting_config,
        )

    def _column_changed(self, _event=None) -> None:
        self.available_values = []
        self.visible_values = []
        self.selected_values.clear()
        self.search_var.set("")
        self.values_list.delete(0, "end")
        self.values_info_var.set("点击“扫描可选值”；也可以直接手动输入。")

    def _scan_values(self) -> None:
        column = self.column_combo.current() + 1
        if column < 1:
            messagebox.showwarning("缺少字段", "请先读取表头并选择字段。", parent=self)
            return
        try:
            options = self._build_options()
        except ValueError as exc:
            messagebox.showerror("设置无效", str(exc), parent=self)
            return
        if not self._headers_match_options(options):
            messagebox.showwarning("表头已变化", "源文件、工作表或表头行已变化，请重新读取表头。", parent=self)
            return
        self.available_values = []
        self.visible_values = []
        self.selected_values.clear()
        self.values_list.delete(0, "end")
        self._run_operation(
            "扫描字段值",
            lambda progress, cancel: scan_values(
                options, column, limit=_SCAN_LIMIT, progress=progress, cancel_event=cancel
            ),
            self._values_scanned,
        )

    def _values_scanned(self, result: object) -> None:
        values, truncated, scanned_rows = result
        self.available_values = list(values)
        self._refresh_value_list()
        note = f"扫描 {scanned_rows:,} 行，发现 {len(values):,} 个不同值。"
        if truncated:
            note += " 已达到显示上限；仍可手动输入其他值。"
        self.values_info_var.set(note)
        self.status_var.set("字段值扫描完成。")

    def _export_each_value(self) -> None:
        """Create one output workbook for every distinct value of this field."""
        column = self.column_combo.current() + 1
        if column < 1 or column > len(self.headers):
            messagebox.showwarning("缺少字段", "请先读取表头并选择一个字段。", parent=self)
            return
        try:
            options = self._build_options(require_output=False)
        except ValueError as exc:
            messagebox.showerror("设置无效", str(exc), parent=self)
            return
        if not self._headers_match_options(options):
            messagebox.showwarning("表头已变化", "源文件、工作表或表头行已变化，请重新读取表头。", parent=self)
            return
        initial = Path(self.batch_parent_var.get().strip()).expanduser()
        if not initial.is_dir():
            initial = options.source.resolve().parent
        selected = filedialog.askdirectory(
            parent=self,
            title="选择按字段分别导出的父文件夹",
            initialdir=str(initial),
        )
        if not selected:
            return
        output_dir = Path(selected)
        self.batch_parent_var.set(str(output_dir))
        self._run_operation(
            "按当前字段批量导出",
            lambda progress, cancel: export_batch(
                options, (), (column,), output_dir,
                progress=progress, cancel_event=cancel,
            ),
            self._batch_export_complete,
        )

    def _schedule_value_search(self, *_args) -> None:
        if self._search_after is not None:
            self.after_cancel(self._search_after)
        self._search_after = self.after(180, self._refresh_value_list)

    def _value_selection_changed(self, _event=None) -> None:
        selected_indices = set(self.values_list.curselection())
        self.selected_values.difference_update(self.visible_values)
        self.selected_values.update(
            self.visible_values[index] for index in selected_indices if index < len(self.visible_values)
        )

    def _refresh_value_list(self) -> None:
        self._search_after = None
        self._value_selection_changed()
        needle = self.search_var.get().strip().casefold()
        if needle:
            self.visible_values = [
                value for value in self.available_values
                if needle in (_EMPTY_LABEL if value == "" else value).casefold()
            ]
        else:
            self.visible_values = list(self.available_values)
        self.values_list.delete(0, "end")
        for index, value in enumerate(self.visible_values):
            self.values_list.insert("end", _EMPTY_LABEL if value == "" else value)
            if value in self.selected_values:
                self.values_list.selection_set(index)

    def _add_rule(self) -> None:
        column = self.column_combo.current() + 1
        if column < 1 or column > len(self.headers):
            messagebox.showwarning("缺少字段", "请先读取表头并选择字段。", parent=self)
            return
        self._value_selection_changed()
        chosen = [value for value in self.available_values if value in self.selected_values]
        manually_entered = [line.strip() for line in self.manual_values.get("1.0", "end").splitlines()]
        values = tuple(dict.fromkeys([*chosen, *(value for value in manually_entered if value)]))
        if not values:
            messagebox.showwarning("缺少筛选值", "请从列表选择值，或在右侧每行输入一个值。", parent=self)
            return
        operator = "equals" if self.operator_var.get() == "完全匹配" else "contains"
        self.rules = [rule for rule in self.rules if rule.column != column]
        self.rules.append(FilterRule(column=column, operator=operator, values=values))
        self.rules.sort(key=lambda rule: rule.column)
        self._refresh_rules()
        self.manual_values.delete("1.0", "end")
        self.status_var.set(f"已添加 {len(self.rules)} 个字段条件。")

    def _remove_rule(self) -> None:
        selected = {int(item) for item in self.rules_tree.selection()}
        if not selected:
            return
        self.rules = [rule for index, rule in enumerate(self.rules) if index not in selected]
        self._refresh_rules()
        self.status_var.set(f"已保留 {len(self.rules)} 个字段条件。")

    def _refresh_rules(self) -> None:
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)
        for index, rule in enumerate(self.rules):
            field = self.headers[rule.column - 1] if rule.column <= len(self.headers) else f"第 {rule.column} 列"
            mode = "完全匹配" if rule.operator == "equals" else "包含文字"
            display_values = "、".join(_EMPTY_LABEL if value == "" else value for value in rule.values)
            if len(display_values) > 110:
                display_values = display_values[:107] + "…"
            self.rules_tree.insert("", "end", iid=str(index), values=(field, mode, display_values))

    def _refresh_batch_tasks_info(self) -> None:
        count = len(self.batch_jobs)
        self.batch_tasks_info_var.set(
            f"已保存 {count} 套批量任务；可继续修改条件并保存下一套。"
            if count else "已保存 0 套批量任务。按字段拆分无需预先保存任务。"
        )

    def _save_batch_job(self) -> None:
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先读取表头。", parent=self)
            return
        if not self.rules:
            messagebox.showwarning("缺少条件", "请先添加至少一个筛选条件，再保存为批量任务。", parent=self)
            return
        try:
            options = self._build_options(require_output=False)
        except ValueError as exc:
            messagebox.showerror("设置无效", str(exc), parent=self)
            return
        if not self._headers_match_options(options):
            messagebox.showwarning("表头已变化", "源文件、工作表或表头行已变化，请重新读取表头。", parent=self)
            return
        name = simpledialog.askstring(
            "保存批量任务", "给当前这套筛选条件起一个名称（也用于输出文件名）：",
            parent=self,
        )
        if name is None:
            return
        name = name.strip()
        if not name:
            messagebox.showwarning("任务名为空", "请输入任务名称。", parent=self)
            return
        existing = next((index for index, job in enumerate(self.batch_jobs) if job.name == name), None)
        job = BatchJob(name=name, rules=tuple(self.rules))
        if existing is not None:
            if not messagebox.askyesno("更新批量任务", f"“{name}”已存在。用当前条件更新它吗？", parent=self):
                return
            self.batch_jobs[existing] = job
        else:
            self.batch_jobs.append(job)
        self._refresh_batch_tasks_info()
        self.status_var.set(f"已保存批量任务“{name}”，共 {len(self.batch_jobs)} 套。")

    def _refresh_accounting_info(self) -> None:
        config = self.accounting_config
        if not any((config.amount_columns, config.group_column, config.required_columns, config.duplicate_columns)):
            self.accounting_info_var.set("会计辅助：未启用")
            return
        parts = []
        if config.amount_columns:
            parts.append(f"金额 {len(config.amount_columns)} 列")
        if config.group_column is not None:
            parts.append("分组 1 列")
        if config.required_columns:
            parts.append(f"必填 {len(config.required_columns)} 列")
        if config.duplicate_columns:
            parts.append(f"重复键 {len(config.duplicate_columns)} 列")
        self.accounting_info_var.set("会计辅助：" + "，".join(parts))

    def _show_accounting_dialog(self) -> None:
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先读取表头，再设置会计辅助功能。", parent=self)
            return
        dialog = tk.Toplevel(self)
        dialog.title("会计辅助设置")
        dialog.geometry("830x620")
        dialog.minsize(700, 500)
        dialog.transient(self)
        dialog.grab_set()
        outer = ttk.Frame(dialog, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(3, weight=1)

        ttk.Label(
            outer,
            text="这些检查只针对筛选后的明细行。留空的项目不会生成对应的辅助结果。",
            style="Hint.TLabel", wraplength=780,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        group_frame = ttk.LabelFrame(outer, text="分组字段（可选）", padding=10)
        group_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        group_frame.columnconfigure(0, weight=1)
        group_var = tk.StringVar()
        group_combo = ttk.Combobox(
            group_frame, textvariable=group_var, state="readonly",
            values=["不分组", *self.headers],
        )
        group_combo.grid(row=0, column=0, sticky="ew")
        group_combo.current(self.accounting_config.group_column or 0)
        ttk.Label(group_frame, text="分组后显示每组行数；勾选金额列后也显示每组金额合计。", style="Hint.TLabel").grid(
            row=1, column=0, sticky="w", pady=(5, 0)
        )

        def make_list(row: int, column: int, title: str, initial: tuple[int, ...], explanation: str) -> tk.Listbox:
            frame = ttk.LabelFrame(outer, text=title, padding=8)
            frame.grid(row=row, column=column, sticky="nsew", padx=(0, 6) if column == 0 else (6, 0), pady=(0, 8))
            frame.columnconfigure(0, weight=1)
            frame.rowconfigure(0, weight=1)
            widget = tk.Listbox(frame, selectmode=tk.MULTIPLE, exportselection=False, height=7)
            widget.grid(row=0, column=0, sticky="nsew")
            bar = ttk.Scrollbar(frame, orient="vertical", command=widget.yview)
            bar.grid(row=0, column=1, sticky="ns")
            widget.configure(yscrollcommand=bar.set)
            for index, label in enumerate(self.headers):
                widget.insert("end", label)
                if index + 1 in initial:
                    widget.selection_set(index)
            ttk.Label(frame, text=explanation, style="Hint.TLabel", wraplength=350).grid(
                row=1, column=0, columnspan=2, sticky="w", pady=(5, 0)
            )
            return widget

        amount_list = make_list(2, 0, "金额列（可多选）", self.accounting_config.amount_columns, "输出合计；无效或空白金额记入检查结果。")
        required_list = make_list(2, 1, "必填检查列（可多选）", self.accounting_config.required_columns, "检查筛选后明细中的空白值。")
        duplicate_list = make_list(3, 0, "重复键列（可多选）", self.accounting_config.duplicate_columns, "所选列组合成一个键，用于查找重复明细。")
        ttk.Label(
            outer, text="提示：按住 Ctrl 或逐项点击可以选取多列。重复键的多列会作为一组联合检查。",
            style="Hint.TLabel", wraplength=350,
        ).grid(row=3, column=1, sticky="nw", padx=(14, 0), pady=(12, 0))

        buttons = ttk.Frame(outer)
        buttons.grid(row=4, column=0, columnspan=2, sticky="e", pady=(8, 0))

        def clear() -> None:
            group_combo.current(0)
            for widget in (amount_list, required_list, duplicate_list):
                widget.selection_clear(0, "end")

        def save() -> None:
            try:
                self.accounting_config = AccountingConfig(
                    amount_columns=tuple(index + 1 for index in amount_list.curselection()),
                    group_column=group_combo.current() or None,
                    required_columns=tuple(index + 1 for index in required_list.curselection()),
                    duplicate_columns=tuple(index + 1 for index in duplicate_list.curselection()),
                )
            except ValueError as exc:
                messagebox.showerror("会计设置无效", str(exc), parent=dialog)
                return
            self._refresh_accounting_info()
            dialog.destroy()

        ttk.Button(buttons, text="清空设置", command=clear).pack(side="left", padx=(0, 14))
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="保存设置", command=save).pack(side="left")
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.focus_set()

    def _export(self) -> None:
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先读取表头。", parent=self)
            return
        if not self.rules:
            messagebox.showwarning("缺少条件", "请先添加至少一个筛选条件。", parent=self)
            return
        try:
            options = self._build_options()
        except ValueError as exc:
            messagebox.showerror("设置无效", str(exc), parent=self)
            return
        if not self._headers_match_options(options):
            messagebox.showwarning("表头已变化", "源文件、工作表或表头行已变化，请重新读取表头。", parent=self)
            return
        if options.output.exists() and not messagebox.askyesno(
            "覆盖已有文件", f"输出文件已存在：\n{options.output}\n\n确定覆盖吗？", parent=self
        ):
            return
        rules = tuple(self.rules)
        self._run_operation(
            "筛选并导出",
            lambda progress, cancel: export_filtered(options, rules, progress=progress, cancel_event=cancel),
            self._export_complete,
        )

    def _show_batch_dialog(self) -> None:
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先读取表头。", parent=self)
            return
        try:
            options = self._build_options(require_output=False)
        except ValueError as exc:
            messagebox.showerror("设置无效", str(exc), parent=self)
            return
        if not self._headers_match_options(options):
            messagebox.showwarning("表头已变化", "源文件、工作表或表头行已变化，请重新读取表头。", parent=self)
            return

        dialog = tk.Toplevel(self)
        dialog.title("批量导出 Excel 文件")
        dialog.geometry("760x610")
        dialog.minsize(660, 510)
        dialog.transient(self)
        dialog.grab_set()

        outer = ttk.Frame(dialog, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)
        ttk.Label(
            outer,
            text="选择一种批量方式。每个结果都会保存为独立的 Excel 文件，统一放进新建的结果文件夹。",
            style="Hint.TLabel", wraplength=710,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        notebook = ttk.Notebook(outer)
        notebook.grid(row=1, column=0, sticky="nsew")

        split_tab = ttk.Frame(notebook, padding=12)
        split_tab.columnconfigure(0, weight=1)
        split_tab.rowconfigure(1, weight=1)
        notebook.add(split_tab, text="按字段自动拆分")
        ttk.Label(
            split_tab,
            text="选一个或多个字段，程序会按每种值或值的组合各生成一个文件。"
                 "主界面已有的筛选条件会先作为公共范围；没有条件则处理全部明细。",
            style="Hint.TLabel", wraplength=685,
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))
        split_list_frame = ttk.Frame(split_tab)
        split_list_frame.grid(row=1, column=0, sticky="nsew")
        split_list_frame.columnconfigure(0, weight=1)
        split_list_frame.rowconfigure(0, weight=1)
        split_list = tk.Listbox(
            split_list_frame, selectmode=tk.MULTIPLE, exportselection=False,
            font=("Microsoft YaHei UI", 10), activestyle="none",
        )
        split_list.grid(row=0, column=0, sticky="nsew")
        for header in self.headers:
            split_list.insert("end", header)
        split_scroll = ttk.Scrollbar(split_list_frame, orient="vertical", command=split_list.yview)
        split_scroll.grid(row=0, column=1, sticky="ns")
        split_list.configure(yscrollcommand=split_scroll.set)
        ttk.Label(
            split_tab, text="例如同时选“部门”和“月份”，每个部门 × 月份组合各得一个文件。点击多项即可多选。",
            style="Hint.TLabel", wraplength=685,
        ).grid(row=2, column=0, sticky="w", pady=(7, 5))
        prefix_row = ttk.Frame(split_tab)
        prefix_row.grid(row=3, column=0, sticky="ew")
        prefix_row.columnconfigure(1, weight=1)
        ttk.Label(prefix_row, text="文件名前缀（可选）").grid(row=0, column=0, sticky="w", padx=(0, 8))
        prefix_var = tk.StringVar()
        ttk.Entry(prefix_row, textvariable=prefix_var).grid(row=0, column=1, sticky="ew")

        jobs_tab = ttk.Frame(notebook, padding=12)
        jobs_tab.columnconfigure(0, weight=1)
        jobs_tab.rowconfigure(1, weight=1)
        notebook.add(jobs_tab, text="多套条件批量导出")
        ttk.Label(
            jobs_tab,
            text="在主界面设置条件后点“保存为批量任务…”，再修改条件并保存下一套。"
                 "此处一次导出全部已保存任务，每套条件生成一个文件。",
            style="Hint.TLabel", wraplength=685,
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))
        jobs_frame = ttk.Frame(jobs_tab)
        jobs_frame.grid(row=1, column=0, sticky="nsew")
        jobs_frame.columnconfigure(0, weight=1)
        jobs_frame.rowconfigure(0, weight=1)
        jobs_tree = ttk.Treeview(
            jobs_frame, columns=("name", "rules"), show="headings", selectmode="extended",
        )
        jobs_tree.heading("name", text="任务名 / 文件名")
        jobs_tree.heading("rules", text="保存时的筛选条件")
        jobs_tree.column("name", width=170, minwidth=120, stretch=False)
        jobs_tree.column("rules", width=470, minwidth=230)
        jobs_tree.grid(row=0, column=0, sticky="nsew")
        jobs_scroll = ttk.Scrollbar(jobs_frame, orient="vertical", command=jobs_tree.yview)
        jobs_scroll.grid(row=0, column=1, sticky="ns")
        jobs_tree.configure(yscrollcommand=jobs_scroll.set)

        def refresh_jobs() -> None:
            jobs_tree.delete(*jobs_tree.get_children())
            for index, job in enumerate(self.batch_jobs):
                parts = []
                for rule in job.rules:
                    field = self.headers[rule.column - 1] if rule.column <= len(self.headers) else f"第 {rule.column} 列"
                    values = " / ".join(_EMPTY_LABEL if value == "" else value for value in rule.values)
                    parts.append(f"{field} {'=' if rule.operator == 'equals' else '包含'} {values}")
                jobs_tree.insert("", "end", iid=str(index), values=(job.name, "；".join(parts)))

        refresh_jobs()

        def delete_jobs() -> None:
            selected = {int(item) for item in jobs_tree.selection()}
            if not selected:
                return
            self.batch_jobs = [job for index, job in enumerate(self.batch_jobs) if index not in selected]
            self._refresh_batch_tasks_info()
            refresh_jobs()

        ttk.Button(jobs_tab, text="删除选中任务", command=delete_jobs).grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )

        output_box = ttk.LabelFrame(outer, text="输出设置（两种方式共用）", padding=(10, 8))
        output_box.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        output_box.columnconfigure(1, weight=1)
        ttk.Label(output_box, text="输出父文件夹").grid(row=0, column=0, sticky="w", padx=(0, 8))
        folder_var = tk.StringVar(value=self.batch_parent_var.get().strip() or str(options.source.resolve().parent))
        ttk.Entry(output_box, textvariable=folder_var).grid(row=0, column=1, sticky="ew")

        def browse_folder() -> None:
            initial = Path(folder_var.get().strip()).expanduser()
            path = filedialog.askdirectory(
                parent=dialog, title="选择批量输出的父文件夹",
                initialdir=str(initial) if initial.is_dir() else str(options.source.resolve().parent),
            )
            if path:
                folder_var.set(path)

        ttk.Button(output_box, text="选择…", command=browse_folder).grid(row=0, column=2, padx=(8, 0))
        ttk.Label(output_box, text="最多生成文件数").grid(row=1, column=0, sticky="w", pady=(8, 0))
        max_files_var = tk.StringVar(value="200")
        ttk.Entry(output_box, textvariable=max_files_var, width=10).grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Label(
            output_box, text="程序会新建唯一的结果子文件夹，已有文件不会被覆盖。",
            style="Hint.TLabel",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        buttons = ttk.Frame(outer)
        buttons.grid(row=3, column=0, sticky="e", pady=(12, 0))

        def start_batch() -> None:
            raw_folder = folder_var.get().strip()
            if not raw_folder:
                messagebox.showwarning("缺少文件夹", "请选择输出父文件夹。", parent=dialog)
                return
            output_dir = Path(raw_folder).expanduser()
            if not output_dir.is_dir():
                messagebox.showerror("文件夹无效", "输出父文件夹不存在，请选择已有文件夹。", parent=dialog)
                return
            try:
                max_files = self._positive_int(max_files_var.get(), "最多生成文件数")
            except ValueError as exc:
                messagebox.showerror("文件数量无效", str(exc), parent=dialog)
                return
            assert max_files is not None
            self.batch_parent_var.set(str(output_dir))
            if notebook.index(notebook.select()) == 0:
                split_columns = tuple(index + 1 for index in split_list.curselection())
                if not split_columns:
                    messagebox.showwarning("缺少拆分字段", "请选择至少一个拆分字段。", parent=dialog)
                    return
                rules = tuple(self.rules)
                prefix = prefix_var.get().strip() or None
                dialog.destroy()
                self._run_operation(
                    "批量拆分导出",
                    lambda progress, cancel: export_batch(
                        options, rules, split_columns, output_dir, prefix=prefix,
                        max_files=max_files, progress=progress, cancel_event=cancel,
                    ),
                    self._batch_export_complete,
                )
            else:
                jobs = tuple(self.batch_jobs)
                if not jobs:
                    messagebox.showwarning("缺少批量任务", "请先在主界面保存至少一套筛选条件。", parent=dialog)
                    return
                dialog.destroy()
                self._run_operation(
                    "多套条件批量导出",
                    lambda progress, cancel: export_named_batch(
                        options, jobs, output_dir, max_files=max_files,
                        progress=progress, cancel_event=cancel,
                    ),
                    self._batch_export_complete,
                )

        ttk.Button(buttons, text="关闭", command=dialog.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="开始批量导出", command=start_batch).pack(side="left")
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.focus_set()

    def _show_analysis_dialog(self) -> None:
        """Define calculated fields, grouped metrics and an Excel chart."""
        if self._busy:
            return
        if not self.headers:
            messagebox.showwarning("缺少表头", "请先选择源文件并读取表头。", parent=self)
            return
        try:
            source_options = self._build_options(require_output=False)
        except ValueError as exc:
            messagebox.showerror("设置无效", str(exc), parent=self)
            return
        if not self._headers_match_options(source_options):
            messagebox.showwarning("表头已变化", "源文件、工作表或表头行已变化，请重新读取表头。", parent=self)
            return

        dialog = tk.Toplevel(self)
        dialog.title("数据分析与图表")
        initial_height = min(790, max(690, dialog.winfo_screenheight() - 90))
        dialog.geometry(f"1040x{initial_height}")
        dialog.minsize(850, 690)
        dialog.transient(self)
        dialog.grab_set()

        outer = ttk.Frame(dialog, padding=15)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        scope = ttk.LabelFrame(outer, text="数据来源与筛选范围", padding=(10, 8))
        scope.grid(row=0, column=0, sticky="ew", pady=(0, 9))
        scope.columnconfigure(0, weight=1)
        ttk.Label(
            scope,
            text=f"源文件：{source_options.source.name}    工作表：{source_options.sheet}    "
                 f"字段数：{len(self.headers)}",
            style="Hint.TLabel", wraplength=850,
        ).grid(row=0, column=0, sticky="w")
        apply_rules_var = tk.BooleanVar(value=bool(self.rules))
        if self.rules:
            rule_parts = []
            for rule in self.rules:
                field = self.headers[rule.column - 1]
                values = "、".join(_EMPTY_LABEL if value == "" else value for value in rule.values)
                rule_parts.append(f"{field} {'=' if rule.operator == 'equals' else '包含'} {values}")
            rules_text = "；".join(rule_parts)
            if len(rules_text) > 230:
                rules_text = rules_text[:227] + "…"
            ttk.Checkbutton(
                scope, text="应用主界面当前筛选条件", variable=apply_rules_var,
            ).grid(row=1, column=0, sticky="w", pady=(6, 0))
            ttk.Label(scope, text=rules_text, style="Hint.TLabel", wraplength=850).grid(
                row=2, column=0, sticky="w", pady=(3, 0)
            )
        else:
            ttk.Label(scope, text="当前无筛选条件，将分析全部明细。要分析某位业务员，请先在主界面添加筛选条件。",
                      style="Hint.TLabel", wraplength=850).grid(row=1, column=0, sticky="w", pady=(5, 0))

        notebook = ttk.Notebook(outer)
        notebook.grid(row=1, column=0, sticky="nsew")
        metrics_tab = ttk.Frame(notebook, padding=10)
        metrics_tab.columnconfigure(0, weight=1)
        metrics_tab.rowconfigure(2, weight=1)
        notebook.add(metrics_tab, text="1  分组与指标")
        variables_tab = ttk.Frame(notebook, padding=12)
        variables_tab.columnconfigure(0, weight=1)
        variables_tab.rowconfigure(1, weight=1)
        notebook.add(variables_tab, text="2  定义新变量")
        chart_tab = ttk.Frame(notebook, padding=12)
        chart_tab.columnconfigure(1, weight=1)
        notebook.add(chart_tab, text="3  图表设置")

        variables: list[CalculatedVariable] = []
        ttk.Label(
            variables_tab,
            text="例如 回款比例 = C5 / C6（C5、C6 只是示意，请按本表字段选择）。"
                 "先用下方按钮插入字段或已有变量，再写 +、-、*、/ 和括号。"
                 "变量按列表顺序计算；V1 表示第 1 个新变量。",
            style="Hint.TLabel", wraplength=850,
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        variable_frame = ttk.Frame(variables_tab)
        variable_frame.grid(row=1, column=0, sticky="nsew")
        variable_frame.columnconfigure(0, weight=1)
        variable_frame.rowconfigure(0, weight=1)
        variable_tree = ttk.Treeview(
            variable_frame, columns=("reference", "name", "expression", "percent"),
            show="headings", selectmode="browse", height=7,
        )
        for key, title, width in (
            ("reference", "引用", 60), ("name", "新变量名称", 190),
            ("expression", "计算公式", 380), ("percent", "显示", 110),
        ):
            variable_tree.heading(key, text=title)
            variable_tree.column(key, width=width, minwidth=55, stretch=key == "expression")
        variable_tree.grid(row=0, column=0, sticky="nsew")
        variable_scroll = ttk.Scrollbar(variable_frame, orient="vertical", command=variable_tree.yview)
        variable_scroll.grid(row=0, column=1, sticky="ns")
        variable_tree.configure(yscrollcommand=variable_scroll.set)

        editor = ttk.LabelFrame(variables_tab, text="新增或编辑变量", padding=(10, 9))
        editor.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        editor.columnconfigure(1, weight=1)
        name_var = tk.StringVar()
        expression_var = tk.StringVar()
        variable_percent_var = tk.BooleanVar(value=False)
        ttk.Label(editor, text="名称").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(editor, textvariable=name_var).grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Checkbutton(editor, text="按百分比显示", variable=variable_percent_var).grid(
            row=0, column=2, sticky="w", padx=(12, 0)
        )
        ttk.Label(editor, text="公式").grid(row=1, column=0, sticky="w", padx=(0, 8))
        expression_entry = ttk.Entry(editor, textvariable=expression_var)
        expression_entry.grid(row=1, column=1, columnspan=2, sticky="ew", pady=3)

        source_choices = [f"C{index}｜{header}" for index, header in enumerate(self.headers, 1)]
        source_choice_var = tk.StringVar(value=source_choices[0])
        variable_choice_var = tk.StringVar()
        insert_row = ttk.Frame(editor)
        insert_row.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(5, 0))
        insert_row.columnconfigure(0, weight=1)
        insert_row.columnconfigure(2, weight=1)
        ttk.Combobox(
            insert_row, textvariable=source_choice_var, values=source_choices, state="readonly",
        ).grid(row=0, column=0, sticky="ew")

        def insert_reference(choice: str) -> None:
            if not choice:
                return
            expression_entry.insert("insert", choice.split("｜", 1)[0])
            expression_entry.focus_set()

        ttk.Button(
            insert_row, text="插入字段", command=lambda: insert_reference(source_choice_var.get()),
        ).grid(row=0, column=1, padx=(5, 12))
        variable_insert_combo = ttk.Combobox(
            insert_row, textvariable=variable_choice_var, state="disabled",
        )
        variable_insert_combo.grid(row=0, column=2, sticky="ew")
        ttk.Button(
            insert_row, text="插入变量", command=lambda: insert_reference(variable_choice_var.get()),
        ).grid(row=0, column=3, padx=(5, 0))
        ttk.Label(
            editor,
            text="百分比只改变 Excel 显示格式，不会把计算结果乘以 100；例如 0.25 显示为 25%。",
            style="Hint.TLabel", wraplength=800,
        ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(5, 0))

        group_var = tk.StringVar()
        metric_name_var = tk.StringVar()
        metric_mode_var = tk.StringVar(value="求和")
        metric_value_var = tk.StringVar()
        numerator_var = tk.StringVar()
        denominator_var = tk.StringVar()
        metric_percent_var = tk.BooleanVar(value=False)
        metric_unit_var = tk.StringVar()
        sort_metric_var = tk.StringVar()
        chart_type_var = tk.StringVar(value="柱状图")
        sort_var = tk.StringVar(value="按指标从高到低")
        chart_limit_var = tk.StringVar(value="20")
        include_details_var = tk.BooleanVar(value=True)
        metric_modes = {
            "求和": "sum", "平均": "average", "计数": "count",
            "先求和，再相除": "ratio_of_sums", "逐行相除，再平均": "average_of_ratios",
        }
        chart_types = {"柱状图": "bar", "饼图": "pie", "折线图": "line"}
        sort_modes = {
            "按指标从高到低": "value_desc", "按指标从低到高": "value_asc", "按分组名称": "group",
        }

        def suggest_amount_column(keywords: tuple[str, ...]) -> str | None:
            candidates: list[tuple[int, int]] = []
            for index, header in enumerate(self.headers, 1):
                label = header.split("·", 1)[-1].strip().casefold()
                if any(word in label for word in ("比例", "比率", "百分比", "%", "％")):
                    continue
                score = max((len(word) for word in keywords if word in label), default=0)
                if score:
                    candidates.append((score, index))
            if not candidates:
                return None
            return f"C{max(candidates, key=lambda item: (item[0], -item[1]))[1]}"

        suggested_numerator = suggest_amount_column((
            "预计回款金额", "期望回款金额", "计划回款金额", "预计回款", "期望回款",
            "计划回款", "实际回款", "实收金额", "回款金额", "收款金额", "回款", "实收",
        ))
        suggested_denominator = suggest_amount_column((
            "应收账款金额", "应收款金额", "应收金额", "合同总金额", "销售总金额",
            "应收账款", "应收款", "应收", "合同金额", "销售金额", "销售额",
        ))
        suggested_actual = suggest_amount_column(("实际回款金额", "实际回款", "实收金额", "实收"))
        suggested_planned = suggest_amount_column(("计划回款金额", "计划回款"))
        choices_initialized = False
        metric_records: list[dict] = []
        editor_metric_id: str | None = None
        editor_snapshot: tuple = ()
        metric_selection_guard = False
        variable_selection_guard = False
        variable_editor_index: int | None = None
        variable_snapshot: tuple = ("", "", False)
        custom_metric_number = 0

        def unique_metric_name(name: str, reference: str = "") -> str:
            used = {record["metric"].name.casefold() for record in metric_records}
            if name.casefold() not in used:
                return name
            base = f"{name}（{reference}）" if reference else name
            candidate = base
            suffix = 2
            while candidate.casefold() in used:
                candidate = f"{base} {suffix}"
                suffix += 1
            return candidate

        for index, header in enumerate(self.headers, 1):
            label = header.split("·", 1)[-1].strip()
            as_percent = any(word in label for word in ("比例", "比率", "完成率", "回款率", "%", "％"))
            reference = f"C{index}"
            metric_records.append({
                "id": reference, "origin": reference, "checked": False,
                "metric": AnalysisMetric(
                    name=unique_metric_name(label, reference),
                    mode="average" if as_percent else "sum", value_ref=reference,
                    as_percent=as_percent,
                ),
            })

        def refresh_metric_choices() -> None:
            nonlocal choices_initialized
            choices = [*source_choices, *(
                f"V{index}｜{variable.name}" for index, variable in enumerate(variables, 1)
            )]
            for combo, selected in (
                (value_combo, metric_value_var),
                (numerator_combo, numerator_var),
                (denominator_combo, denominator_var),
            ):
                previous = selected.get().split("｜", 1)[0]
                combo.configure(values=choices)
                selected.set(next((choice for choice in choices if choice.split("｜", 1)[0] == previous), ""))
            if not choices_initialized:
                if suggested_numerator:
                    numerator_var.set(next(choice for choice in choices if choice.startswith(suggested_numerator + "｜")))
                if suggested_denominator and suggested_denominator != suggested_numerator:
                    denominator_var.set(next(choice for choice in choices if choice.startswith(suggested_denominator + "｜")))
                choices_initialized = True

        def refresh_variable_choices() -> None:
            editable_index = variable_editor_index if variable_editor_index is not None else len(variables)
            choices = [f"V{index}｜{variable.name}" for index, variable in enumerate(
                variables[:editable_index], 1
            )]
            variable_insert_combo.configure(values=choices, state="readonly" if choices else "disabled")
            variable_choice_var.set(choices[0] if choices else "")

        def refresh_variables(select_index: int | None = None) -> None:
            nonlocal variable_selection_guard
            variable_selection_guard = True
            variable_tree.delete(*variable_tree.get_children())
            for index, variable in enumerate(variables):
                variable_tree.insert(
                    "", "end", iid=str(index),
                    values=(f"V{index + 1}", variable.name, variable.expression,
                            "百分比" if variable.as_percent else "数值"),
                )
            if select_index is not None:
                variable_tree.selection_set(str(select_index))
                variable_tree.see(str(select_index))
            variable_selection_guard = False
            refresh_variable_choices()
            refresh_metric_choices()
            sync_variable_metrics()

        def select_variable(_event=None) -> None:
            nonlocal variable_editor_index, variable_snapshot, variable_selection_guard
            if variable_selection_guard:
                return
            selected = variable_tree.selection()
            index = int(selected[0]) if selected else None
            if index == variable_editor_index:
                return
            if variable_form() != variable_snapshot and not save_variable():
                variable_selection_guard = True
                if variable_editor_index is not None:
                    variable_tree.selection_set(str(variable_editor_index))
                else:
                    variable_tree.selection_remove(*variable_tree.selection())
                variable_selection_guard = False
                return
            variable_editor_index = index
            if index is not None:
                variable_tree.selection_set(str(index))
                variable = variables[index]
                name_var.set(variable.name)
                expression_var.set(variable.expression)
                variable_percent_var.set(variable.as_percent)
            variable_snapshot = variable_form()
            refresh_variable_choices()

        variable_tree.bind("<<TreeviewSelect>>", select_variable)

        def new_variable() -> None:
            nonlocal variable_editor_index, variable_snapshot, variable_selection_guard
            if variable_form() != variable_snapshot and not save_variable():
                return
            variable_selection_guard = True
            variable_tree.selection_remove(*variable_tree.selection())
            variable_editor_index = None
            name_var.set("")
            expression_var.set("")
            variable_percent_var.set(False)
            variable_snapshot = variable_form()
            variable_selection_guard = False
            refresh_variable_choices()
            expression_entry.focus_set()

        def variable_form() -> tuple:
            return (name_var.get().strip(), expression_var.get().strip(), variable_percent_var.get())

        def save_variable() -> bool:
            nonlocal variable_editor_index, variable_snapshot
            name = name_var.get().strip()
            expression = expression_var.get().strip()
            if not name and not expression and variable_editor_index is None:
                return True
            if not name or not expression:
                messagebox.showwarning("变量未完成", "请填写新变量名称和计算公式。", parent=dialog)
                notebook.select(variables_tab)
                return False
            target_index = variable_editor_index if variable_editor_index is not None else len(variables)
            if any(index != target_index and item.name.casefold() == name.casefold()
                   for index, item in enumerate(variables)):
                messagebox.showerror("名称重复", "新变量名称不能重复。", parent=dialog)
                notebook.select(variables_tab)
                return False
            for column in re.findall(r"\bC(\d+)\b", expression):
                if not 1 <= int(column) <= len(self.headers):
                    messagebox.showerror("字段引用无效", f"公式中的 C{column} 不在源表字段范围内。", parent=dialog)
                    notebook.select(variables_tab)
                    return False
            for variable_index in re.findall(r"\bV(\d+)\b", expression):
                if not 1 <= int(variable_index) <= target_index:
                    messagebox.showerror("变量引用无效", "公式只能引用列表中排在自己前面的变量。", parent=dialog)
                    notebook.select(variables_tab)
                    return False
            variable = CalculatedVariable(name=name, expression=expression, as_percent=variable_percent_var.get())
            if variable_editor_index is not None:
                old_variable = variables[target_index]
                default_record = next((record for record in metric_records
                                       if record["origin"] == f"V{target_index + 1}"), None)
                if default_record:
                    if editor_metric_id == default_record["id"] and not save_metric_editor():
                        return False
                    metric = default_record["metric"]
                    updates = {}
                    if metric.name == old_variable.name and not any(
                        record is not default_record and record["metric"].name.casefold() == name.casefold()
                        for record in metric_records
                    ):
                        updates["name"] = name
                    if (metric.value_ref == f"V{target_index + 1}"
                            and metric.as_percent == old_variable.as_percent
                            and metric.mode == ("average" if old_variable.as_percent else "sum")):
                        updates.update(as_percent=variable.as_percent,
                                       mode="average" if variable.as_percent else "sum")
                    if updates:
                        default_record["metric"] = replace(metric, **updates)
                    if editor_metric_id == default_record["id"]:
                        load_metric(editor_metric_id)
                variables[target_index] = variable
                variable_snapshot = variable_form()
                refresh_variables(target_index)
            else:
                variables.append(variable)
                variable_editor_index = target_index
                variable_snapshot = variable_form()
                refresh_variables(target_index)
            return True

        def remove_variable() -> None:
            nonlocal variable_editor_index, variable_snapshot
            selected = variable_tree.selection()
            if not selected:
                return
            index = int(selected[0])
            if index != len(variables) - 1:
                messagebox.showwarning(
                    "请从末尾删除", "变量可能被后面的公式引用。请先删除列表中排在它后面的变量。",
                    parent=dialog,
                )
                return
            reference = f"V{index + 1}"
            if not save_metric_editor():
                return
            dependent = [record["metric"].name for record in metric_records
                         if record["origin"] != reference and reference in (
                             record["metric"].value_ref, record["metric"].numerator_ref,
                             record["metric"].denominator_ref,
                         )]
            if dependent:
                messagebox.showwarning(
                    "变量仍被指标使用", "请先修改或删除引用此变量的指标：\n" + "、".join(dependent), parent=dialog,
                )
                return
            variables.pop()
            variable_editor_index = None
            name_var.set("")
            expression_var.set("")
            variable_percent_var.set(False)
            variable_snapshot = variable_form()
            refresh_variables()

        variable_actions = ttk.Frame(variables_tab)
        variable_actions.grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Button(variable_actions, text="新增变量", command=new_variable).pack(side="left")
        ttk.Button(variable_actions, text="添加 / 更新变量", command=save_variable).pack(side="left", padx=(7, 0))
        ttk.Button(variable_actions, text="删除选中变量", command=remove_variable).pack(side="left", padx=(7, 0))

        group_row = ttk.Frame(metrics_tab)
        group_row.grid(row=0, column=0, sticky="ew")
        group_row.columnconfigure(1, weight=1)
        ttk.Label(group_row, text="先确定分组字段（横轴）").grid(row=0, column=0, sticky="w", padx=(0, 10))
        person_columns = [
            index for index, header in enumerate(self.headers)
            if any(word in header for word in ("业务员", "销售员", "姓名", "经办人", "负责人"))
        ]
        company_column = next((index for index, header in enumerate(self.headers)
                               if "公司" in header or "客户" in header), None)
        person_already_filtered = any(
            rule.column == index + 1 for rule in self.rules for index in person_columns
        )
        default_group = (
            person_columns[0] if person_columns and not person_already_filtered else
            company_column if company_column is not None else
            person_columns[0] if person_columns else 0
        )
        group_var.set(source_choices[default_group])
        ttk.Combobox(group_row, textvariable=group_var, values=source_choices, state="readonly").grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Label(
            metrics_tab,
            text="再勾选需要展示的指标：点击左侧 □ / ☑ 或按空格切换；点击某一行可修改口径。"
                 "人名请选择“业务员 / 姓名”作为分组；指标名称显示在图例中。",
            style="Hint.TLabel", wraplength=940,
        ).grid(row=1, column=0, sticky="w", pady=(6, 6))
        metric_frame = ttk.Frame(metrics_tab)
        metric_frame.grid(row=2, column=0, sticky="nsew")
        metric_frame.columnconfigure(0, weight=1)
        metric_frame.rowconfigure(0, weight=1)
        metric_tree = ttk.Treeview(
            metric_frame, columns=("checked", "name", "mode", "reference", "format", "unit"),
            show="headings", selectmode="browse", height=5,
        )
        for key, title, width in (
            ("checked", "勾选", 50), ("name", "指标名称（图例）", 200),
            ("mode", "汇总口径", 145), ("reference", "字段 / 公式", 210),
            ("format", "格式", 70), ("unit", "单位", 70),
        ):
            metric_tree.heading(key, text=title)
            metric_tree.column(key, width=width, minwidth=45, stretch=key in {"name", "reference"})
        metric_tree.grid(row=0, column=0, sticky="nsew")
        metric_scroll = ttk.Scrollbar(metric_frame, orient="vertical", command=metric_tree.yview)
        metric_scroll.grid(row=0, column=1, sticky="ns")
        metric_hscroll = ttk.Scrollbar(metric_frame, orient="horizontal", command=metric_tree.xview)
        metric_hscroll.grid(row=1, column=0, sticky="ew")
        metric_tree.configure(yscrollcommand=metric_scroll.set, xscrollcommand=metric_hscroll.set)
        metric_tree.tag_configure("checked", foreground="#176342")
        metric_actions = ttk.Frame(metrics_tab)
        metric_actions.grid(row=3, column=0, sticky="ew", pady=(6, 5))
        metric_count_label = ttk.Label(metric_actions, style="Hint.TLabel")
        metric_count_label.pack(side="right")

        metric_editor = ttk.LabelFrame(metrics_tab, text="当前指标设置（切换行或导出时自动保存）", padding=(8, 5))
        metric_editor.grid(row=4, column=0, sticky="ew")
        metric_editor.columnconfigure(1, weight=1)
        metric_editor.columnconfigure(3, weight=1)
        for row, left, right in ((0, "指标名称", "汇总口径"), (1, "指标字段 / 变量", "单位（可选）"),
                                 (2, "分子", "分母")):
            ttk.Label(metric_editor, text=left).grid(row=row, column=0, sticky="w", padx=(0, 6), pady=3)
            ttk.Label(metric_editor, text=right).grid(row=row, column=2, sticky="w", padx=(12, 6), pady=3)
        ttk.Entry(metric_editor, textvariable=metric_name_var).grid(row=0, column=1, sticky="ew", pady=3)
        ttk.Combobox(metric_editor, textvariable=metric_mode_var, values=tuple(metric_modes), state="readonly").grid(
            row=0, column=3, sticky="ew", pady=3,
        )
        value_combo = ttk.Combobox(metric_editor, textvariable=metric_value_var, state="readonly")
        value_combo.grid(row=1, column=1, sticky="ew", pady=3)
        ttk.Entry(metric_editor, textvariable=metric_unit_var, width=14).grid(row=1, column=3, sticky="ew", pady=3)
        numerator_combo = ttk.Combobox(metric_editor, textvariable=numerator_var, state="readonly")
        numerator_combo.grid(row=2, column=1, sticky="ew", pady=3)
        denominator_combo = ttk.Combobox(metric_editor, textvariable=denominator_var, state="readonly")
        denominator_combo.grid(row=2, column=3, sticky="ew", pady=3)
        editor_actions = ttk.Frame(metric_editor)
        editor_actions.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(4, 0))
        metric_percent_check = ttk.Checkbutton(editor_actions, text="百分比显示（0.25 → 25%）", variable=metric_percent_var)
        metric_percent_check.pack(side="left")
        mouth_note = ttk.Label(metric_editor, style="Hint.TLabel", wraplength=910)
        mouth_note.grid(row=4, column=0, columnspan=4, sticky="w", pady=(4, 0))

        def update_metric_mode(*_args, reset_format: bool = True) -> None:
            mode = metric_modes[metric_mode_var.get()]
            if reset_format:
                metric_percent_var.set(mode in {"ratio_of_sums", "average_of_ratios"})
            if mode == "count":
                metric_percent_var.set(False)
            metric_percent_check.configure(state="disabled" if mode == "count" else "normal")
            value_combo.configure(state="readonly" if mode in {"sum", "average"} else "disabled")
            numerator_combo.configure(state="readonly" if mode in {"ratio_of_sums", "average_of_ratios"} else "disabled")
            denominator_combo.configure(state="readonly" if mode in {"ratio_of_sums", "average_of_ratios"} else "disabled")
            if mode == "ratio_of_sums":
                mouth_note.configure(text="仅统计分子与分母都有效的成对数据，在每组内分别合计后相除；缺失或非数值行成对跳过。")
            elif mode == "average_of_ratios":
                mouth_note.configure(text="每行先计算分子 ÷ 分母，再按组取算术平均；分母为零或缺失的行会跳过。")
            elif mode == "count":
                mouth_note.configure(text="统计每个分组的明细行数，无需选择指标字段。")
            else:
                mouth_note.configure(text="按所选字段或新变量，在每个分组内求和或平均。非数值单元格会跳过。")

        metric_mode_var.trace_add("write", update_metric_mode)
        mode_labels = {value: label for label, value in metric_modes.items()}

        def token(text: str) -> str | None:
            return text.split("｜", 1)[0] if text else None

        def metric_form() -> tuple:
            return (metric_name_var.get().strip(), metric_mode_var.get(), token(metric_value_var.get()),
                    token(numerator_var.get()), token(denominator_var.get()), metric_percent_var.get(),
                    metric_unit_var.get().strip())

        def find_metric(record_id: str | None) -> dict | None:
            return next((record for record in metric_records if record["id"] == record_id), None)

        def checked_metrics() -> list[AnalysisMetric]:
            return [record["metric"] for record in metric_records if record["checked"]]

        def refresh_sort_metrics() -> None:
            names = [metric.name for metric in checked_metrics()]
            previous = sort_metric_var.get()
            sort_metric_combo.configure(values=names)
            sort_metric_var.set(previous if previous in names else names[0] if names else "")
            metric_count_label.configure(text=f"已勾选 {len(names)} 项")

        def refresh_metric_tree(select_id: str | None = None) -> None:
            nonlocal metric_selection_guard
            metric_selection_guard = True
            previous_view = metric_tree.yview()
            metric_tree.delete(*metric_tree.get_children())
            choices = {choice.split("｜", 1)[0]: choice for choice in (
                *source_choices, *(f"V{i}｜{variable.name}" for i, variable in enumerate(variables, 1))
            )}
            for record in metric_records:
                metric = record["metric"]
                if metric.mode == "count":
                    reference = "明细行数"
                elif metric.mode in {"ratio_of_sums", "average_of_ratios"}:
                    reference = f"{metric.numerator_ref} ÷ {metric.denominator_ref}"
                else:
                    reference = choices.get(metric.value_ref, metric.value_ref or "")
                metric_tree.insert("", "end", iid=record["id"], values=(
                    "☑" if record["checked"] else "□", metric.name, mode_labels[metric.mode], reference,
                    "百分比" if metric.as_percent else "数值", metric.unit or "—",
                ), tags=("checked",) if record["checked"] else ())
            target = select_id if select_id is not None else editor_metric_id
            if target and metric_tree.exists(target):
                metric_tree.selection_set(target)
                metric_tree.focus(target)
            if previous_view:
                metric_tree.yview_moveto(previous_view[0])
            metric_selection_guard = False
            refresh_sort_metrics()

        def load_metric(record_id: str | None) -> None:
            nonlocal editor_metric_id, editor_snapshot
            record = find_metric(record_id)
            editor_metric_id = record_id if record else None
            metric = record["metric"] if record else AnalysisMetric(name="", mode="sum")
            metric_name_var.set(metric.name)
            metric_mode_var.set(mode_labels[metric.mode])
            choices = [*source_choices, *(f"V{i}｜{variable.name}" for i, variable in enumerate(variables, 1))]
            for variable, reference in ((metric_value_var, metric.value_ref), (numerator_var, metric.numerator_ref),
                                        (denominator_var, metric.denominator_ref)):
                variable.set(next((choice for choice in choices if token(choice) == reference), ""))
            metric_percent_var.set(metric.as_percent)
            metric_unit_var.set(metric.unit)
            update_metric_mode(reset_format=False)
            editor_snapshot = metric_form()

        def save_metric_editor(check: bool = False) -> bool:
            nonlocal editor_metric_id, editor_snapshot, custom_metric_number
            record = find_metric(editor_metric_id)
            form = metric_form()
            if form == editor_snapshot and not check:
                return True
            if not form[0]:
                if record is None and form == ("", "求和", None, None, None, False, ""):
                    return True
                messagebox.showwarning("缺少指标名称", "请填写当前指标的名称。", parent=dialog)
                notebook.select(metrics_tab)
                return False
            if any(other is not record and other["metric"].name.casefold() == form[0].casefold()
                   for other in metric_records):
                messagebox.showwarning("指标名称重复", "各指标名称必须不同，请为当前指标换一个名称。", parent=dialog)
                notebook.select(metrics_tab)
                return False
            mode = metric_modes[form[1]]
            if not form[5] and form[6] in {"%", "％", "百分比"}:
                messagebox.showwarning(
                    "单位与显示格式不一致", "百分比单位请同时勾选“百分比显示”；普通数值请填写元、件等单位，或留空。",
                    parent=dialog,
                )
                notebook.select(metrics_tab)
                return False
            available = {f"C{i}" for i in range(1, len(self.headers) + 1)} | {f"V{i}" for i in range(1, len(variables) + 1)}
            references = (form[2],) if mode in {"sum", "average"} else (form[3], form[4]) if mode != "count" else ()
            if any(reference not in available for reference in references):
                messagebox.showwarning("指标字段未完成", "请选择有效的指标字段；比值需要同时选择分子与分母。", parent=dialog)
                notebook.select(metrics_tab)
                return False
            metric = AnalysisMetric(
                name=form[0], mode=mode, value_ref=form[2] if mode in {"sum", "average"} else None,
                numerator_ref=form[3] if mode in {"ratio_of_sums", "average_of_ratios"} else None,
                denominator_ref=form[4] if mode in {"ratio_of_sums", "average_of_ratios"} else None,
                as_percent=bool(form[5]) if mode != "count" else False,
                unit=form[6],
            )
            if record is None:
                custom_metric_number += 1
                editor_metric_id = f"M{custom_metric_number}"
                metric_records.append({"id": editor_metric_id, "origin": None, "checked": True, "metric": metric})
            else:
                record["metric"] = metric
                if check:
                    record["checked"] = True
            editor_snapshot = form
            refresh_metric_tree()
            return True

        def select_metric(_event=None) -> None:
            nonlocal metric_selection_guard
            if metric_selection_guard:
                return
            selection = metric_tree.selection()
            target = selection[0] if selection else None
            if target == editor_metric_id:
                return
            if not save_metric_editor():
                metric_selection_guard = True
                if editor_metric_id and metric_tree.exists(editor_metric_id):
                    metric_tree.selection_set(editor_metric_id)
                else:
                    metric_tree.selection_remove(*metric_tree.selection())
                metric_selection_guard = False
                return
            if target:
                metric_tree.selection_set(target)
                load_metric(target)

        def toggle_metric(event=None) -> str | None:
            if event is not None and getattr(event, "num", None) == 1:
                if metric_tree.identify_region(event.x, event.y) != "cell" or metric_tree.identify_column(event.x) != "#1":
                    return None
                target = metric_tree.identify_row(event.y)
            else:
                selection = metric_tree.selection()
                target = selection[0] if selection else ""
            if not target or not save_metric_editor():
                return "break"
            record = find_metric(target)
            record["checked"] = not record["checked"]
            load_metric(target)
            refresh_metric_tree(target)
            return "break"

        def clear_metrics() -> None:
            if not save_metric_editor():
                return
            for record in metric_records:
                record["checked"] = False
            refresh_metric_tree()

        def quick_amount_metrics() -> None:
            if not save_metric_editor():
                return
            references = {suggest_amount_column(words) for words in (
                ("计划回款金额", "预计回款金额", "期望回款金额", "计划回款", "预计回款", "期望回款"),
                ("实际回款金额", "实际回款", "实收金额", "实收"),
            )}
            references.discard(None)
            if not references:
                messagebox.showinfo("未找到回款字段", "本表未发现计划 / 实际回款字段，请直接勾选所需金额字段。", parent=dialog)
                return
            for record in metric_records:
                if record["origin"] in references:
                    record["checked"] = True
                    record["metric"] = replace(record["metric"], mode="sum", value_ref=record["origin"],
                                               numerator_ref=None, denominator_ref=None, as_percent=False)
            if editor_metric_id:
                load_metric(editor_metric_id)
            refresh_metric_tree()

        def new_metric() -> None:
            nonlocal editor_snapshot, metric_selection_guard
            if not save_metric_editor():
                return
            metric_selection_guard = True
            metric_tree.selection_remove(*metric_tree.selection())
            metric_selection_guard = False
            load_metric(None)
            is_completion = bool(suggested_actual and suggested_planned and suggested_actual != suggested_planned)
            metric_name_var.set(unique_metric_name("回款完成率" if is_completion else "自定义比值"))
            metric_mode_var.set("先求和，再相除")
            for selected, reference in (
                (numerator_var, suggested_actual if is_completion else suggested_numerator),
                (denominator_var, suggested_planned if is_completion else suggested_denominator),
            ):
                selected.set(next((choice for choice in source_choices if token(choice) == reference), ""))
            notebook.select(metrics_tab)

        def remove_metric() -> None:
            nonlocal editor_metric_id
            record = find_metric(editor_metric_id)
            if record is None:
                load_metric(None)
                return
            if record["origin"] is None:
                metric_records.remove(record)
                load_metric(None)
            else:
                if not save_metric_editor():
                    return
                record["checked"] = False
            refresh_metric_tree()

        def move_metric(offset: int) -> None:
            if not save_metric_editor():
                return
            record = find_metric(editor_metric_id)
            if record is None:
                return
            index = metric_records.index(record)
            target = max(0, min(len(metric_records) - 1, index + offset))
            metric_records.insert(target, metric_records.pop(index))
            refresh_metric_tree()
            metric_tree.see(record["id"])

        def sync_variable_metrics() -> None:
            valid_origins = {f"V{i}" for i in range(1, len(variables) + 1)}
            metric_records[:] = [record for record in metric_records
                                 if not (record["origin"] and record["origin"].startswith("V")
                                         and record["origin"] not in valid_origins)]
            for index, variable in enumerate(variables, 1):
                origin = f"V{index}"
                record = next((item for item in metric_records if item["origin"] == origin), None)
                if record is None:
                    metric_records.append({
                        "id": origin, "origin": origin, "checked": False,
                        "metric": AnalysisMetric(name=unique_metric_name(variable.name, origin),
                                                 mode="average" if variable.as_percent else "sum",
                                                 value_ref=origin, as_percent=variable.as_percent),
                    })
            if editor_metric_id and find_metric(editor_metric_id) is None:
                load_metric(None)
            refresh_metric_tree()

        metric_tree.bind("<<TreeviewSelect>>", select_metric)
        metric_tree.bind("<Button-1>", toggle_metric)
        metric_tree.bind("<space>", toggle_metric)
        ttk.Button(metric_actions, text="勾选计划 / 实际回款", command=quick_amount_metrics).pack(side="left")
        ttk.Button(metric_actions, text="全部取消", command=clear_metrics).pack(side="left", padx=(6, 0))
        ttk.Button(metric_actions, text="新增比值 / 计数", command=new_metric).pack(side="left", padx=(6, 0))
        ttk.Button(metric_actions, text="↑", width=3, command=lambda: move_metric(-1)).pack(side="left", padx=(6, 0))
        ttk.Button(metric_actions, text="↓", width=3, command=lambda: move_metric(1)).pack(side="left", padx=(3, 0))
        ttk.Button(metric_actions, text="移除 / 取消勾选", command=remove_metric).pack(side="left", padx=(6, 0))
        ttk.Button(editor_actions, text="应用指标设置", command=save_metric_editor).pack(side="right")
        ttk.Button(editor_actions, text="保存并勾选", command=lambda: save_metric_editor(True)).pack(side="right", padx=(0, 6))

        ttk.Label(chart_tab, text="图表类型").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Combobox(chart_tab, textvariable=chart_type_var, values=tuple(chart_types), state="readonly").grid(
            row=0, column=1, sticky="ew", pady=4
        )
        ttk.Label(chart_tab, text="排序方式").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Combobox(chart_tab, textvariable=sort_var, values=tuple(sort_modes), state="readonly").grid(
            row=1, column=1, sticky="ew", pady=4
        )
        ttk.Label(chart_tab, text="排序依据指标").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=4)
        sort_metric_combo = ttk.Combobox(chart_tab, textvariable=sort_metric_var, state="readonly")
        sort_metric_combo.grid(row=2, column=1, sticky="ew", pady=4)

        def sort_changed(*_args) -> None:
            sort_metric_combo.configure(state="disabled" if sort_var.get() == "按分组名称" else "readonly")

        sort_var.trace_add("write", sort_changed)

        def chart_changed(*_args) -> None:
            if chart_type_var.get() == "折线图":
                sort_var.set("按分组名称")

        chart_type_var.trace_add("write", chart_changed)
        limit_row = ttk.Frame(chart_tab)
        limit_row.grid(row=3, column=1, sticky="w", pady=4)
        ttk.Entry(limit_row, textvariable=chart_limit_var, width=8).pack(side="left")
        ttk.Label(limit_row, text="个分组；默认前 20 个，填 0 显示全部；汇总表始终完整", style="Hint.TLabel").pack(
            side="left", padx=(7, 0)
        )
        ttk.Label(chart_tab, text="图中最多显示").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Checkbutton(chart_tab, text="同时导出匹配的明细与新变量", variable=include_details_var).grid(
            row=4, column=1, sticky="w", pady=(6, 0)
        )
        ttk.Label(
            chart_tab,
            text="柱状图：同一分组的多个指标并排显示；折线图：每个指标一条线。\n\n"
                 "同单位指标放在同一张图中，图例标明各指标名称。百分比和普通数值、不同单位自动分图。"
                 "如金额与件数同时选择，请分别填写“元”“件”；单位只用于区分图表，不会换算数值。\n\n"
                 "饼图：每个勾选指标生成一张图，展示该指标在各分组中的占比；"
                 "仅支持非负的求和 / 计数指标，不接受百分比指标。\n\n"
                 "柱状图 / 折线图共用所选排序指标和分组；饼图按各自指标独立排序和选取分组。"
                 "每张图遵循分组上限，汇总表始终保留全部分组和全部勾选指标。",
            style="Hint.TLabel", wraplength=710,
        ).grid(row=5, column=1, sticky="w", pady=(14, 0))
        refresh_metric_choices()
        load_metric(None)
        refresh_metric_tree()

        output_box = ttk.LabelFrame(outer, text="输出 Excel（含汇总表与图表）", padding=(10, 8))
        output_box.grid(row=2, column=0, sticky="ew", pady=(9, 0))
        output_box.columnconfigure(0, weight=1)
        source_path = source_options.source.resolve()
        output_path = source_path.with_name(f"{source_path.stem}_分析图表.xlsx")
        output_var = tk.StringVar(value=str(output_path))
        ttk.Entry(output_box, textvariable=output_var).grid(row=0, column=0, sticky="ew")

        def browse_output() -> None:
            current = Path(output_var.get().strip() or output_path).expanduser()
            selected = filedialog.asksaveasfilename(
                parent=dialog, title="保存分析结果", defaultextension=".xlsx",
                initialdir=str(current.parent) if current.parent.is_dir() else str(source_path.parent),
                initialfile=current.name, filetypes=[("Excel 工作簿", "*.xlsx")],
            )
            if selected:
                output_var.set(selected)

        ttk.Button(output_box, text="另存为…", command=browse_output).grid(row=0, column=1, padx=(8, 0))

        actions = ttk.Frame(outer)
        actions.grid(row=3, column=0, sticky="e", pady=(9, 0))

        def start_analysis() -> None:
            if not save_variable():
                return
            if not save_metric_editor():
                return
            metrics = checked_metrics()
            if not metrics:
                messagebox.showwarning("尚未勾选指标", "请先在“分组与指标”页勾选至少一个需要展示的指标。", parent=dialog)
                notebook.select(metrics_tab)
                return
            if len(metrics) > 255:
                messagebox.showwarning("勾选指标过多", "一次最多选择 255 个指标，请取消部分指标后再导出。", parent=dialog)
                notebook.select(metrics_tab)
                return
            chart_type = chart_types[chart_type_var.get()]
            unsuitable = [metric.name for metric in metrics if metric.mode not in {"sum", "count"} or metric.as_percent]
            if chart_type == "pie" and unsuitable:
                messagebox.showwarning(
                    "饼图口径不合适", "饼图仅支持非百分比的“求和”或“计数”。请修改或取消这些指标，"
                    "也可以选择柱状图 / 折线图：\n" + "、".join(unsuitable),
                    parent=dialog,
                )
                notebook.select(chart_tab)
                return
            raw_output = output_var.get().strip()
            if not raw_output:
                messagebox.showwarning("缺少输出文件", "请选择分析结果的保存位置。", parent=dialog)
                return
            output = Path(raw_output).expanduser().resolve()
            if output.suffix.lower() != ".xlsx":
                messagebox.showerror("文件格式无效", "分析结果必须保存为 .xlsx 文件。", parent=dialog)
                return
            if output == source_path:
                messagebox.showerror("路径冲突", "分析结果不能覆盖源文件。", parent=dialog)
                return
            if not output.parent.is_dir():
                messagebox.showerror("文件夹不存在", "请选择已有的输出文件夹。", parent=dialog)
                return
            try:
                chart_limit_number = int(chart_limit_var.get().strip())
                if chart_limit_number < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("分组数无效", "图中分组数请填写 0 或正整数。", parent=dialog)
                return
            chart_limit = chart_limit_number or None
            equal_ratio_metrics = [metric.name for metric in metrics
                                   if metric.mode in {"ratio_of_sums", "average_of_ratios"}
                                   and metric.numerator_ref == metric.denominator_ref]
            if equal_ratio_metrics:
                if not messagebox.askyesno(
                    "分子与分母相同", "这些指标的分子与分母相同，结果通常为 100%：\n"
                    + "、".join(equal_ratio_metrics) + "\n仍要继续吗？",
                    parent=dialog,
                ):
                    return
            if output.exists() and not messagebox.askyesno(
                "覆盖已有文件", f"输出文件已存在：\n{output}\n\n确定覆盖吗？", parent=dialog,
            ):
                return
            options = AnalysisOptions(
                source=source_path, sheet=source_options.sheet, output=output,
                group_column=int(group_var.get().split("｜", 1)[0][1:]),
                metrics=tuple(metrics), chart_type=chart_type, variables=tuple(variables),
                rules=tuple(self.rules) if apply_rules_var.get() else (),
                header_start=source_options.header_start, header_end=source_options.header_end,
                data_start=source_options.data_start, data_end=source_options.data_end,
                stop_at_summary=source_options.stop_at_summary,
                chart_limit=chart_limit, sort_by=sort_modes[sort_var.get()],
                sort_metric_index=next((index for index, metric in enumerate(metrics)
                                        if metric.name == sort_metric_var.get()), 0),
                include_details=include_details_var.get(),
            )
            dialog.destroy()
            self._run_operation(
                "分析并生成图表",
                lambda progress, cancel: analyze_to_workbook(options, progress=progress, cancel_event=cancel),
                self._analysis_complete,
            )

        ttk.Button(actions, text="取消", command=dialog.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="生成分析 Excel", command=start_analysis).pack(side="left")
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.focus_set()

    def _show_merge_dialog(self) -> None:
        """Collect several workbooks and combine their data by header name."""
        if self._busy:
            return

        current = Path(self.source_var.get().strip()).expanduser() if self.source_var.get().strip() else None
        sources: list[Path] = [current.resolve()] if current is not None and current.is_file() else []

        def suggest_output(folder: Path) -> Path:
            base = folder / "合并总表.xlsx"
            if not base.exists():
                return base
            for suffix in range(2, 10000):
                candidate = folder / f"合并总表_{suffix}.xlsx"
                if not candidate.exists():
                    return candidate
            return base

        suggested = str(suggest_output(sources[0].parent if sources else Path.cwd()))
        dialog = tk.Toplevel(self)
        dialog.title("合并多个 Excel 文件为总表")
        dialog.geometry("820x760")
        dialog.minsize(700, 670)
        dialog.transient(self)
        dialog.grab_set()

        outer = ttk.Frame(dialog, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)
        ttk.Label(
            outer,
            text="选择至少两个工作簿。程序根据字段表头对齐数据，按列表顺序将明细合并到一个总表。",
            style="Hint.TLabel", wraplength=780,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        source_box = ttk.LabelFrame(outer, text="1  待合并文件", padding=(10, 8))
        source_box.grid(row=1, column=0, sticky="nsew", pady=(0, 9))
        source_box.columnconfigure(0, weight=1)
        source_box.rowconfigure(0, weight=1)
        list_frame = ttk.Frame(source_box)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        source_list = tk.Listbox(list_frame, selectmode=tk.EXTENDED, exportselection=False, height=7)
        source_list.grid(row=0, column=0, sticky="nsew")
        source_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=source_list.yview)
        source_scroll.grid(row=0, column=1, sticky="ns")
        source_xscroll = ttk.Scrollbar(list_frame, orient="horizontal", command=source_list.xview)
        source_xscroll.grid(row=1, column=0, sticky="ew")
        source_list.configure(yscrollcommand=source_scroll.set, xscrollcommand=source_xscroll.set)

        count_var = tk.StringVar()

        def refresh_sources(selected: int | None = None) -> None:
            source_list.delete(0, "end")
            for path in sources:
                source_list.insert("end", str(path))
            count_var.set(f"已选 {len(sources)} 个文件")
            if selected is not None and 0 <= selected < len(sources):
                source_list.selection_set(selected)
                source_list.see(selected)

        def add_sources() -> None:
            nonlocal suggested
            folder = sources[0].parent if sources else Path.cwd()
            paths = filedialog.askopenfilenames(
                parent=dialog, title="选择要合并的 Excel 文件（可多选）",
                initialdir=str(folder),
                filetypes=[("Excel 工作簿", "*.xlsx *.xlsm")],
            )
            if not paths:
                return
            existing = {os.path.normcase(str(path)) for path in sources}
            had_sources = bool(sources)
            for raw in paths:
                path = Path(raw).resolve()
                key = os.path.normcase(str(path))
                if key not in existing:
                    sources.append(path)
                    existing.add(key)
            if not had_sources and sources and output_var.get().strip() == suggested:
                suggested = str(suggest_output(sources[0].parent))
                output_var.set(suggested)
            refresh_sources()

        def remove_sources() -> None:
            for index in reversed(source_list.curselection()):
                del sources[index]
            refresh_sources()

        def move_source(direction: int) -> None:
            selected = source_list.curselection()
            if len(selected) != 1:
                return
            old = selected[0]
            new = old + direction
            if 0 <= new < len(sources):
                sources[old], sources[new] = sources[new], sources[old]
                refresh_sources(new)

        source_actions = ttk.Frame(source_box)
        source_actions.grid(row=0, column=1, sticky="ns", padx=(9, 0))
        ttk.Button(source_actions, text="添加文件…", command=add_sources).pack(fill="x", pady=(0, 5))
        ttk.Button(source_actions, text="移除选中", command=remove_sources).pack(fill="x", pady=(0, 5))
        ttk.Button(source_actions, text="上移", command=lambda: move_source(-1)).pack(fill="x", pady=(0, 5))
        ttk.Button(source_actions, text="下移", command=lambda: move_source(1)).pack(fill="x")
        ttk.Label(source_box, textvariable=count_var, style="Hint.TLabel").grid(
            row=1, column=0, sticky="w", pady=(5, 0)
        )
        refresh_sources()

        settings = ttk.LabelFrame(outer, text="2  读取和合并方式", padding=(10, 8))
        settings.grid(row=2, column=0, sticky="ew", pady=(0, 9))
        settings.columnconfigure(1, weight=1)
        sheet_mode = tk.StringVar(value="first")
        sheet_name = tk.StringVar(value=self.sheet_var.get().strip())
        ttk.Label(settings, text="工作表").grid(row=0, column=0, sticky="w", padx=(0, 10))
        sheet_row = ttk.Frame(settings)
        sheet_row.grid(row=0, column=1, sticky="ew")
        ttk.Radiobutton(sheet_row, text="每个文件的第一个工作表", variable=sheet_mode, value="first").pack(side="left")
        ttk.Radiobutton(sheet_row, text="每个文件中同名工作表：", variable=sheet_mode, value="named").pack(
            side="left", padx=(14, 4)
        )
        sheet_entry = ttk.Entry(sheet_row, textvariable=sheet_name, width=18)
        sheet_entry.pack(side="left", fill="x", expand=True)

        def update_sheet_state(*_args) -> None:
            sheet_entry.configure(state="normal" if sheet_mode.get() == "named" else "disabled")

        sheet_mode.trace_add("write", update_sheet_state)
        update_sheet_state()

        header_start = tk.StringVar(value=self.header_start_var.get().strip() or "1")
        header_end = tk.StringVar(value=self.header_end_var.get().strip() or "1")
        data_start = tk.StringVar(value=self.data_start_var.get().strip())
        data_end = tk.StringVar(value=self.data_end_var.get().strip())
        range_row = ttk.Frame(settings)
        range_row.grid(row=1, column=1, sticky="w", pady=(9, 0))
        ttk.Label(settings, text="数据范围").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(9, 0))
        for label, variable in (
            ("表头起始行", header_start), ("至", header_end),
            ("明细起始行", data_start), ("至", data_end),
        ):
            ttk.Label(range_row, text=label).pack(side="left", padx=(0, 5) if label != "至" else (3, 5))
            ttk.Entry(range_row, textvariable=variable, width=6).pack(side="left", padx=(0, 12) if label == "至" else 0)
        ttk.Label(
            settings, text="明细起始行留空表示表头后一行；结束行留空表示读取至末尾。",
            style="Hint.TLabel",
        ).grid(row=2, column=1, sticky="w", pady=(5, 0))

        match_mode = tk.StringVar(value="union")
        mode_row = ttk.Frame(settings)
        mode_row.grid(row=3, column=1, sticky="w", pady=(9, 0))
        ttk.Label(settings, text="字段匹配").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=(9, 0))
        ttk.Radiobutton(
            mode_row, text="字段并集（缺少的字段留空）", variable=match_mode, value="union"
        ).pack(side="left")
        ttk.Radiobutton(
            mode_row, text="要求字段一致（顺序可不同）", variable=match_mode, value="strict"
        ).pack(side="left", padx=(15, 0))

        stop_summary = tk.BooleanVar(value=self.stop_summary_var.get())
        add_source = tk.BooleanVar(value=False)
        deduplicate = tk.BooleanVar(value=False)
        check_row = ttk.Frame(settings)
        check_row.grid(row=4, column=1, sticky="w", pady=(9, 1))
        ttk.Checkbutton(check_row, text="遇到合计 / 汇总行时停止", variable=stop_summary).pack(side="left")
        ttk.Checkbutton(check_row, text="添加来源文件列", variable=add_source).pack(side="left", padx=(15, 0))
        ttk.Checkbutton(check_row, text="按整行数据去重", variable=deduplicate).pack(side="left", padx=(15, 0))
        ttk.Label(
            settings, text="去重以全部业务字段为准，来源文件列不参与比较；重复行保留首次出现的记录。",
            style="Hint.TLabel", wraplength=745,
        ).grid(row=5, column=1, sticky="w", pady=(4, 0))

        output_box = ttk.LabelFrame(outer, text="3  输出总表", padding=(10, 8))
        output_box.grid(row=3, column=0, sticky="ew", pady=(0, 9))
        output_box.columnconfigure(1, weight=1)
        output_var = tk.StringVar(value=suggested)
        ttk.Label(output_box, text="输出文件").grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Entry(output_box, textvariable=output_var).grid(row=0, column=1, sticky="ew")

        def browse_output() -> None:
            current_output = Path(output_var.get().strip() or suggested).expanduser()
            path = filedialog.asksaveasfilename(
                parent=dialog, title="保存合并总表", defaultextension=".xlsx",
                initialdir=str(current_output.parent) if current_output.parent.is_dir() else str(Path.cwd()),
                initialfile=current_output.name,
                filetypes=[("Excel 工作簿", "*.xlsx")],
            )
            if path:
                output_var.set(path)

        ttk.Button(output_box, text="另存为…", command=browse_output).grid(row=0, column=2, padx=(8, 0))
        ttk.Label(
            output_box, text="按表头名称对齐后，所有明细写入同一总表；超出 Excel 单表行数时自动分到后续工作表。",
            style="Hint.TLabel", wraplength=745,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(5, 0))

        actions = ttk.Frame(outer)
        actions.grid(row=4, column=0, sticky="e")

        def start_merge() -> None:
            if len(sources) < 2:
                messagebox.showwarning("源文件不足", "请至少添加两个 Excel 文件。", parent=dialog)
                return
            raw_output = output_var.get().strip()
            if not raw_output:
                messagebox.showwarning("缺少输出文件", "请选择总表的保存位置。", parent=dialog)
                return
            output = Path(raw_output).expanduser().resolve()
            if output.suffix.lower() != ".xlsx":
                messagebox.showerror("文件格式无效", "总表必须保存为 .xlsx 文件。", parent=dialog)
                return
            if not output.parent.is_dir():
                messagebox.showerror("文件夹不存在", "输出文件夹不存在，请选择已有文件夹。", parent=dialog)
                return
            if any(output == source for source in sources):
                messagebox.showerror("文件路径冲突", "输出文件不能与任一源文件相同。", parent=dialog)
                return
            if sheet_mode.get() == "named" and not sheet_name.get().strip():
                messagebox.showwarning("缺少工作表名", "请输入每个文件中共同的工作表名称。", parent=dialog)
                return
            try:
                h_start = self._positive_int(header_start.get(), "表头起始行")
                h_end = self._positive_int(header_end.get(), "表头结束行")
                d_start = self._positive_int(data_start.get(), "明细起始行", optional=True)
                d_end = self._positive_int(data_end.get(), "明细结束行", optional=True)
                assert h_start is not None and h_end is not None
                if h_end < h_start:
                    raise ValueError("表头结束行不能早于起始行。")
                if d_start is not None and d_start <= h_end:
                    raise ValueError("明细起始行必须在表头结束行之后。")
                if d_end is not None and d_end < (d_start or h_end + 1):
                    raise ValueError("明细结束行不能早于明细起始行。")
            except ValueError as exc:
                messagebox.showerror("行号无效", str(exc), parent=dialog)
                return
            missing = [path for path in sources if not path.is_file()]
            if missing:
                messagebox.showerror("找不到源文件", f"找不到源文件：\n{missing[0]}", parent=dialog)
                return
            if output.exists() and not messagebox.askyesno(
                "覆盖已有文件", f"输出文件已存在：\n{output}\n\n确定覆盖吗？", parent=dialog
            ):
                return
            options = MergeOptions(
                sources=tuple(sources), output=output,
                sheet_name=sheet_name.get().strip() if sheet_mode.get() == "named" else None,
                header_start=h_start, header_end=h_end,
                data_start=d_start, data_end=d_end,
                stop_at_summary=stop_summary.get(),
                add_source_column=add_source.get(),
                deduplicate=deduplicate.get(),
                match_mode=match_mode.get(),
            )
            dialog.destroy()
            self._run_operation(
                "合并总表",
                lambda progress, cancel: merge_workbooks(options, progress=progress, cancel_event=cancel),
                self._merge_complete,
            )

        ttk.Button(actions, text="取消", command=dialog.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="开始合并", command=start_merge).pack(side="left")
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.focus_set()

    def _headers_match_options(self, options: ExportOptions) -> bool:
        return self._header_context == (
            options.source.resolve(), options.sheet, options.header_start, options.header_end
        )

    def _export_complete(self, result: object) -> None:
        note = f"完成：从 {result.scanned_rows:,} 行明细中导出 {result.matched_rows:,} 行。"
        if result.stopped_at_row is not None:
            note += f" 在原表第 {result.stopped_at_row:,} 行汇总处停止。"
        self.status_var.set(note)
        messagebox.showinfo("导出完成", f"{note}\n\n文件已保存至：\n{result.output}", parent=self)

    def _batch_export_complete(self, result: object) -> None:
        note = (
            f"批量导出完成：扫描 {result.scanned_rows:,} 行，累计导出 {result.matched_rows:,} 行明细，"
            f"生成 {result.file_count:,} 个文件。"
        )
        if result.stopped_at_row is not None:
            note += f" 在原表第 {result.stopped_at_row:,} 行汇总处停止。"
        self._last_batch_output_dir = Path(result.output_dir)
        self.open_batch_folder_button.configure(state="normal")
        self.status_var.set(note + f" 结果文件夹：{result.output_dir}")
        messagebox.showinfo("批量导出完成", f"{note}\n\n结果文件夹：\n{result.output_dir}", parent=self)

    def _merge_complete(self, result: object) -> None:
        note = (
            f"合并完成：{result.files:,} 个源文件，扫描 {result.scanned_rows:,} 行明细，"
            f"写入 {result.merged_rows:,} 行，跳过重复 {result.duplicates_skipped:,} 行。"
        )
        if result.sheets > 1:
            note += f" 总表分为 {result.sheets:,} 个工作表。"
        self.status_var.set(note + f" 保存位置：{result.output}")
        messagebox.showinfo("合并完成", f"{note}\n\n总表已保存至：\n{result.output}", parent=self)

    def _analysis_complete(self, result: object) -> None:
        note = (
            f"分析完成：扫描 {result.scanned_rows:,} 行，匹配 {result.matched_rows:,} 行，"
            f"得到 {result.group_count:,} 个分组、{result.metric_count:,} 个指标，"
            f"生成 {result.chart_count:,} 张图，每图最多显示 {result.chart_points:,} 个分组。"
        )
        if result.invalid_metric_rows:
            note += f" 有 {result.invalid_metric_rows:,} 行至少一项指标无法计算（详见“指标校验”表）。"
        if result.zero_denominator_groups:
            note += f" {result.zero_denominator_groups:,} 个分组的合计分母为零，比例留空。"
        self.status_var.set(note + f" 保存位置：{result.output}")
        messagebox.showinfo("分析完成", f"{note}\n\n分析 Excel 已保存至：\n{result.output}", parent=self)

    def _open_batch_folder(self) -> None:
        folder = self._last_batch_output_dir
        if folder is None or not folder.is_dir():
            messagebox.showerror("找不到文件夹", "上次批量输出的文件夹不存在。", parent=self)
            return
        try:
            os.startfile(str(folder))
        except OSError as exc:
            messagebox.showerror("无法打开文件夹", str(exc), parent=self)

    def _run_operation(
        self,
        title: str,
        job: Callable[[Callable[[str, int, int | None, int], None], threading.Event], object],
        success_handler: Callable[[object], None],
    ) -> None:
        if self._busy:
            return
        self._busy = True
        self._active_operation = title
        self._cancel_event = threading.Event()
        self._success_handler = success_handler
        self._set_controls_busy(True)
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self._progress_indeterminate = True
        self.status_var.set(f"正在{title}…")

        def progress(phase: str, processed: int, total: int | None, matched: int) -> None:
            self._events.put(("progress", phase, processed, total, matched))

        def worker() -> None:
            try:
                result = job(progress, self._cancel_event)
            except CancelledError:
                self._events.put(("cancelled",))
            except Exception as exc:
                self._events.put(("error", str(exc) or type(exc).__name__))
            else:
                self._events.put(("success", result))

        threading.Thread(target=worker, daemon=True, name="excel-workbook-worker").start()

    def _set_controls_busy(self, busy: bool) -> None:
        for widget, normal_state in self._controls:
            widget.configure(state="disabled" if busy else normal_state)
        self.values_list.configure(state="disabled" if busy else "normal")
        self.rules_tree.configure(selectmode="none" if busy else "extended")
        self.cancel_button.configure(state="normal" if busy else "disabled")
        self.open_batch_folder_button.configure(
            state="disabled" if busy or self._last_batch_output_dir is None else "normal"
        )

    def _cancel(self) -> None:
        if self._cancel_event is not None:
            self._cancel_event.set()
            self.cancel_button.configure(state="disabled")
            self.status_var.set("正在取消，请稍候…")

    def _drain_events(self) -> None:
        try:
            while True:
                item = self._events.get_nowait()
                kind = item[0]
                if kind == "progress":
                    self._show_progress(*item[1:])
                else:
                    self._finish_operation()
                    if self._closing_requested:
                        self.destroy()
                        return
                    if kind == "success" and self._success_handler is not None:
                        self._success_handler(item[1])
                    elif kind == "cancelled":
                        self.status_var.set("操作已取消。")
                    elif kind == "error":
                        self.status_var.set("操作失败。")
                        messagebox.showerror("操作失败", item[1], parent=self)
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    def _show_progress(self, phase: str, processed: int, total: int | None, matched: int) -> None:
        if self._active_operation == "分析并生成图表":
            if phase == "done":
                self.progress.stop()
                self._progress_indeterminate = False
                self.progress.configure(mode="determinate", value=100)
                return
            if phase == "save":
                if not self._progress_indeterminate:
                    self.progress.configure(mode="indeterminate")
                    self.progress.start(12)
                    self._progress_indeterminate = True
                self.status_var.set(f"正在保存分析结果与图表，已匹配 {matched:,} 行…")
                return
            if total is not None and total > 0:
                if self._progress_indeterminate:
                    self.progress.stop()
                    self._progress_indeterminate = False
                self.progress.configure(mode="determinate", value=min(100, processed / total * 100))
                amount = f"{processed:,} / {total:,}"
            else:
                if not self._progress_indeterminate:
                    self.progress.configure(mode="indeterminate")
                    self.progress.start(12)
                    self._progress_indeterminate = True
                amount = f"{processed:,}"
            self.status_var.set(f"正在分析明细：已查看 {amount} 行，匹配 {matched:,} 行。")
            return
        if self._active_operation == "合并总表":
            if phase == "done":
                self.progress.stop()
                self._progress_indeterminate = False
                self.progress.configure(mode="determinate", value=100)
                return
            if phase == "save":
                if not self._progress_indeterminate:
                    self.progress.configure(mode="indeterminate")
                    self.progress.start(12)
                    self._progress_indeterminate = True
                self.status_var.set(f"正在保存合并总表，已写入 {matched:,} 行…")
                return
            if total is not None and total > 0:
                if self._progress_indeterminate:
                    self.progress.stop()
                    self._progress_indeterminate = False
                self.progress.configure(mode="determinate", value=min(100, processed / total * 100))
                amount = f"{processed:,} / {total:,}"
            else:
                if not self._progress_indeterminate:
                    self.progress.configure(mode="indeterminate")
                    self.progress.start(12)
                    self._progress_indeterminate = True
                amount = f"{processed:,}"
            if phase == "inspect":
                self.status_var.set(f"正在检查源文件：已检查 {amount} 个。")
            else:
                self.status_var.set(f"正在合并明细：已查看 {amount} 行，写入 {matched:,} 行。")
            return
        is_batch = "批量" in self._active_operation
        if phase == "save":
            if is_batch and total is not None and total > 0:
                if self._progress_indeterminate:
                    self.progress.stop()
                    self._progress_indeterminate = False
                self.progress.configure(mode="determinate", value=min(100, processed / total * 100))
                self.status_var.set(f"正在生成批量 Excel 文件：已写入 {processed:,} / {total:,} 行。")
                return
            if not self._progress_indeterminate:
                self.progress.configure(mode="indeterminate")
                self.progress.start(12)
                self._progress_indeterminate = True
            self.status_var.set(
                f"正在生成批量 Excel 文件，累计匹配 {matched:,} 行…"
                if is_batch else f"正在保存结果，已匹配 {matched:,} 行…"
            )
            return
        if phase == "done":
            self.progress.stop()
            self._progress_indeterminate = False
            self.progress.configure(mode="determinate", value=100)
            return
        if total is not None and total > 0:
            if self._progress_indeterminate:
                self.progress.stop()
                self._progress_indeterminate = False
            self.progress.configure(mode="determinate", value=min(100, processed / total * 100))
            amount = f"{processed:,} / {total:,}"
        else:
            if not self._progress_indeterminate:
                self.progress.configure(mode="indeterminate")
                self.progress.start(12)
                self._progress_indeterminate = True
            amount = f"{processed:,}"
        if phase == "scan":
            self.status_var.set(f"扫描字段值：已查看 {amount} 行，发现 {matched:,} 个不同值。")
        elif is_batch:
            self.status_var.set(f"批量导出：已查看 {amount} 行，累计匹配 {matched:,} 行。")
        else:
            self.status_var.set(f"筛选导出：已查看 {amount} 行，匹配 {matched:,} 行。")

    def _finish_operation(self) -> None:
        self._busy = False
        self._active_operation = ""
        self._cancel_event = None
        self._set_controls_busy(False)
        self.progress.stop()
        self._progress_indeterminate = False
        self.progress.configure(mode="determinate", value=0)

    def _on_close(self) -> None:
        if self._busy:
            self._closing_requested = True
            self._cancel()
            return
        self.destroy()


def run_app() -> None:
    ExcelFilterApp().mainloop()


if __name__ == "__main__":
    run_app()
