"""Streaming Excel row filtering used by the desktop application.

The source workbook is opened read-only and is never saved.  The result is
written to a temporary file first, then moved into place after a full save.
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
import os
import pickle
import re
import shutil
import sqlite3
import tempfile
import unicodedata
import uuid
from types import SimpleNamespace
from typing import Callable, Iterable, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from accounting import AccountingAccumulator, AccountingConfig


ProgressCallback = Callable[[str, int, int | None, int], None]
_SUMMARY_LABEL = re.compile(
    r"^\s*(?:合\s*计|小\s*计|总\s*计|汇\s*总|总\s*合\s*计|"
    r"grand\s+total|sub\s*total|total)\s*[:：]?\s*$",
    re.IGNORECASE,
)


class CancelledError(Exception):
    """Raised when the user cancels a workbook scan or export."""


@dataclass(frozen=True)
class FilterRule:
    column: int  # 1-based Excel column number
    operator: str  # "equals" or "contains"
    values: tuple[str, ...]  # alternatives within this column (OR)


@dataclass(frozen=True)
class ExportOptions:
    source: Path
    sheet: str
    output: Path
    header_start: int = 1
    header_end: int = 1
    data_start: int | None = None
    data_end: int | None = None
    stop_at_summary: bool = True
    preserve_styles: bool = False
    accounting: AccountingConfig = field(default_factory=AccountingConfig)


@dataclass(frozen=True)
class ExportResult:
    scanned_rows: int
    matched_rows: int
    stopped_at_row: int | None
    output: Path


@dataclass(frozen=True)
class BatchJob:
    """One independently filtered output in a named batch."""

    name: str
    rules: tuple[FilterRule, ...]


@dataclass(frozen=True)
class BatchResult:
    """Result of a batch export.

    matched_rows counts detail rows across all output files. A source row
    matching two named jobs is counted twice.
    """

    scanned_rows: int
    matched_rows: int
    file_count: int
    output_dir: Path
    files: tuple[Path, ...]
    stopped_at_row: int | None


def normalize_value(value: object) -> str:
    """Make cell values comparable with values typed in the user interface."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        if value.time() == time.min:
            return value.date().isoformat()
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, float):
        return format(value, ".15g")
    if isinstance(value, Decimal):
        return format(value, "f")
    return str(value).strip()


def _check_source(source: Path) -> Path:
    path = Path(source).expanduser()
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("请选择 .xlsx 或 .xlsm 文件；旧版 .xls 请先用 Excel 另存为 .xlsx。")
    if not path.is_file():
        raise FileNotFoundError(f"找不到源文件：{path}")
    return path


def _check_layout(header_start: int, header_end: int, data_start: int | None, data_end: int | None) -> int:
    if header_start < 1 or header_end < header_start:
        raise ValueError("表头起止行无效。")
    first_data = data_start if data_start is not None else header_end + 1
    if first_data <= header_end:
        raise ValueError("明细起始行必须在表头结束行之后。")
    if data_end is not None and data_end < first_data:
        raise ValueError("明细结束行不能早于明细起始行。")
    return first_data


def _open_sheet(source: Path, sheet: str):
    book = load_workbook(_check_source(source), read_only=True, data_only=True, keep_links=False)
    if sheet not in book.sheetnames:
        book.close()
        raise ValueError(f"工作表不存在：{sheet}")
    return book, book[sheet]


def list_sheets(source: Path) -> list[str]:
    """Return worksheet names without loading cell data into memory."""
    book = load_workbook(_check_source(source), read_only=True, data_only=True, keep_links=False)
    try:
        return list(book.sheetnames)
    finally:
        book.close()


def inspect_headers(source: Path, sheet: str, header_start: int, header_end: int) -> list[str]:
    """Return uniquely identified column labels, including multi-row headers."""
    _check_layout(header_start, header_end, None, None)
    book, ws = _open_sheet(source, sheet)
    try:
        header_rows = list(ws.iter_rows(min_row=header_start, max_row=header_end, values_only=True))
        width = max(
            (index for row in header_rows for index, value in enumerate(row, 1) if normalize_value(value)),
            default=0,
        )
        if width == 0:
            raise ValueError("指定表头行没有内容，请调整表头行号。")
        labels: list[str] = []
        for column in range(1, width + 1):
            parts: list[str] = []
            for row in header_rows:
                part = normalize_value(row[column - 1]) if column <= len(row) else ""
                if part and (not parts or part != parts[-1]):
                    parts.append(part)
            title = " / ".join(parts) if parts else "未命名列"
            labels.append(f"{get_column_letter(column)} · {title}")
        return labels
    finally:
        book.close()


def _is_summary_row(cells: Sequence[object]) -> bool:
    # A summary label is normally at the left of a report.  Restricting the
    # check to the first four columns avoids matching an ordinary note field.
    return any(
        _SUMMARY_LABEL.fullmatch(normalize_value(getattr(cell, "value", cell)))
        for cell in cells[:4]
    )


def _is_blank_row(cells: Sequence[object]) -> bool:
    return all(not normalize_value(getattr(cell, "value", cell)) for cell in cells)


def _progress(callback: ProgressCallback | None, phase: str, processed: int, total: int | None, matched: int) -> None:
    if callback is not None:
        callback(phase, processed, total, matched)


def _total_rows(ws, start: int, end: int | None) -> int | None:
    maximum = getattr(ws, "max_row", None)
    if maximum is None and end is None:
        return None
    final = min(maximum, end) if maximum is not None and end is not None else (end or maximum)
    return max(0, final - start + 1) if final is not None else None


def scan_values(
    options: ExportOptions,
    column: int,
    limit: int = 100000,
    progress: ProgressCallback | None = None,
    cancel_event=None,
) -> tuple[list[str], bool, int]:
    """Scan one column's distinct values, stopping if the distinct limit is hit."""
    first_data = _check_layout(options.header_start, options.header_end, options.data_start, options.data_end)
    if column < 1 or limit < 1:
        raise ValueError("列号或可选值上限无效。")
    # Use the same data width as exports. A row whose selected field is blank
    # may still be a real detail row when a later column contains data.
    width = len(inspect_headers(options.source, options.sheet, options.header_start, options.header_end))
    if column > width:
        raise ValueError("列号超出表头范围。")
    book, ws = _open_sheet(options.source, options.sheet)
    try:
        seen: set[str] = set()
        distinct: list[str] = []
        scanned = 0
        total = _total_rows(ws, first_data, options.data_end)
        for row_number, cells in enumerate(
            ws.iter_rows(min_row=first_data, max_row=options.data_end, max_col=width),
            start=first_data,
        ):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError()
            if options.stop_at_summary and _is_summary_row(cells):
                break
            if _is_blank_row(cells):
                continue
            scanned += 1
            value = normalize_value(cells[column - 1].value)
            if value not in seen:
                if len(seen) >= limit:
                    _progress(progress, "scan", row_number - first_data + 1, total, len(distinct))
                    return sorted(distinct, key=lambda item: (item != "", item.casefold())), True, scanned
                seen.add(value)
                distinct.append(value)
            if scanned % 2000 == 0:
                _progress(progress, "scan", row_number - first_data + 1, total, len(distinct))
        _progress(progress, "scan", scanned, total, len(distinct))
        return sorted(distinct, key=lambda item: (item != "", item.casefold())), False, scanned
    finally:
        book.close()


def _validate_rules(rules: Iterable[FilterRule], width: int) -> list[tuple[int, str, tuple[str, ...]]]:
    prepared = []
    for rule in rules:
        if rule.column < 1 or rule.column > width:
            raise ValueError(f"筛选列 {rule.column} 超出表头范围。")
        if rule.operator not in {"equals", "contains"}:
            raise ValueError(f"不支持的筛选方式：{rule.operator}")
        if not rule.values:
            raise ValueError("筛选条件至少需要一个值。")
        prepared.append((rule.column - 1, rule.operator, tuple(normalize_value(value).casefold() for value in rule.values)))
    return prepared


def _matches(cells: Sequence[object], rules: Sequence[tuple[int, str, tuple[str, ...]]]) -> bool:
    for index, operator, alternatives in rules:
        value = normalize_value(cells[index].value).casefold()
        if operator == "equals":
            passed = value in alternatives
        else:
            passed = any((candidate in value if candidate else value == "") for candidate in alternatives)
        if not passed:
            return False
    return True


def _display_width(value: str) -> int:
    return sum(2 if unicodedata.east_asian_width(char) in {"F", "W"} else 1 for char in value)


def _output_cell(ws_out, source_cell, *, header: bool, preserve_styles: bool):
    value = getattr(source_cell, "value", None)
    number_format = getattr(source_cell, "number_format", "General") or "General"
    has_style = bool(getattr(source_cell, "has_style", False))
    needs_cell = header or preserve_styles and has_style or number_format != "General" or isinstance(value, str) and value.startswith("=")
    if not needs_cell:
        return value
    target = WriteOnlyCell(ws_out, value=value)
    if isinstance(value, str) and value.startswith("="):
        target.data_type = "s"  # Keep a literal source string as text.
    if preserve_styles and has_style:
        for attr in ("font", "fill", "border", "alignment", "protection"):
            setattr(target, attr, copy(getattr(source_cell, attr)))
        target.number_format = number_format
    else:
        if number_format != "General":
            target.number_format = number_format
        if header:
            target.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
            target.fill = PatternFill(fill_type="solid", fgColor="24476B")
            target.alignment = Alignment(vertical="center", wrap_text=True)
    return target


def _append_accounting_row(sheet, values: Sequence[object], *, header: bool = False) -> None:
    cells = []
    for value in values:
        cell = WriteOnlyCell(sheet, value=value)
        if header:
            cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill(fill_type="solid", fgColor="24476B")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        elif isinstance(value, Decimal):
            cell.number_format = '#,##0.00;[Red](#,##0.00)'
        elif isinstance(value, int):
            cell.number_format = '#,##0'
        cells.append(cell)
    sheet.append(cells)


def _validate_accounting(config: AccountingConfig, width: int) -> None:
    columns = (*config.amount_columns, *config.required_columns, *config.duplicate_columns)
    if config.group_column is not None:
        columns = (*columns, config.group_column)
    if any(column > width for column in columns):
        raise ValueError("会计辅助功能选中的列超出表头范围。")


def export_filtered(
    options: ExportOptions,
    rules: Sequence[FilterRule],
    progress: ProgressCallback | None = None,
    cancel_event=None,
) -> ExportResult:
    """Export matching detail rows to an atomic .xlsx file."""
    first_data = _check_layout(options.header_start, options.header_end, options.data_start, options.data_end)
    source = _check_source(options.source)
    output = Path(options.output).expanduser()
    if output.suffix.lower() != ".xlsx":
        raise ValueError("输出文件必须是 .xlsx。")
    if source.resolve() == output.resolve():
        raise ValueError("输出路径不能与源文件相同。")
    if not output.parent.is_dir():
        raise FileNotFoundError(f"输出文件夹不存在：{output.parent}")

    labels = inspect_headers(source, options.sheet, options.header_start, options.header_end)
    width = len(labels)
    prepared_rules = _validate_rules(rules, width)
    _validate_accounting(options.accounting, width)
    accounting_enabled = bool(
        options.accounting.amount_columns
        or options.accounting.group_column is not None
        or options.accounting.required_columns
        or options.accounting.duplicate_columns
    )
    accountant = AccountingAccumulator(options.accounting) if accounting_enabled else None
    book, ws = _open_sheet(source, options.sheet)
    result_book = Workbook(write_only=True)
    result_sheet = result_book.create_sheet(options.sheet)
    result_sheet.freeze_panes = f"A{options.header_end - options.header_start + 2}"
    for index, label in enumerate(labels, 1):
        result_sheet.column_dimensions[get_column_letter(index)].width = min(max(_display_width(label) + 2, 12), 38)

    matched = 0
    scanned = 0
    stopped_at = None
    temp_path: Path | None = None
    try:
        for source_row in ws.iter_rows(
            min_row=options.header_start, max_row=options.header_end, max_col=width
        ):
            result_sheet.append([
                _output_cell(result_sheet, cell, header=True, preserve_styles=options.preserve_styles)
                for cell in source_row
            ])

        total = _total_rows(ws, first_data, options.data_end)
        for row_number, cells in enumerate(
            ws.iter_rows(min_row=first_data, max_row=options.data_end, max_col=width),
            start=first_data,
        ):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError()
            if options.stop_at_summary and _is_summary_row(cells):
                stopped_at = row_number
                break
            if _is_blank_row(cells):
                continue
            scanned += 1
            if _matches(cells, prepared_rules):
                result_sheet.append([
                    _output_cell(result_sheet, cell, header=False, preserve_styles=options.preserve_styles)
                    for cell in cells
                ])
                matched += 1
                if accountant is not None:
                    accountant.add_row([cell.value for cell in cells], row_number)
            if scanned % 2000 == 0:
                _progress(progress, "export", row_number - first_data + 1, total, matched)

        last_output_row = options.header_end - options.header_start + 1 + matched
        result_sheet.auto_filter.ref = f"A{options.header_end - options.header_start + 1}:{get_column_letter(width)}{last_output_row}"
        if accountant is not None:
            if options.accounting.amount_columns or options.accounting.group_column is not None:
                summary_sheet = result_book.create_sheet("金额汇总")
                summary_sheet.freeze_panes = "A2"
                for index, row in enumerate(accountant.iter_summary_rows(labels)):
                    _append_accounting_row(summary_sheet, row, header=index == 0)
                summary_sheet.column_dimensions["A"].width = 14
                summary_sheet.column_dimensions["B"].width = 26
                summary_sheet.column_dimensions["C"].width = 15
                for index in range(4, 4 + len(options.accounting.amount_columns)):
                    summary_sheet.column_dimensions[get_column_letter(index)].width = 24
            if (
                options.accounting.amount_columns
                or options.accounting.required_columns
                or options.accounting.duplicate_columns
            ):
                issues_sheet = result_book.create_sheet("数据检查")
                issues_sheet.freeze_panes = "A2"
                _append_accounting_row(issues_sheet, ["原始行号", "问题类型", "说明"], header=True)
                issues = accountant.issue_rows()
                if issues:
                    for issue in issues:
                        _append_accounting_row(issues_sheet, issue)
                    omitted = accountant.issue_count - len(issues)
                    if omitted:
                        _append_accounting_row(issues_sheet, ["", "记录上限", f"另外 {omitted} 条问题未列出"])
                else:
                    _append_accounting_row(issues_sheet, ["", "无问题", "所选检查未发现异常"])
                issues_sheet.column_dimensions["A"].width = 15
                issues_sheet.column_dimensions["B"].width = 16
                issues_sheet.column_dimensions["C"].width = 70
        _progress(progress, "save", scanned, total, matched)
        fd, temp_name = tempfile.mkstemp(prefix=".excel_filter_", suffix=".xlsx", dir=output.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        result_book.save(temp_path)
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        os.replace(temp_path, output)
        _progress(progress, "done", scanned, total, matched)
        return ExportResult(scanned, matched, stopped_at, output)
    finally:
        if accountant is not None:
            accountant.close()
        book.close()
        result_book.close()
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()


_INVALID_FILE_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{n}" for n in range(1, 10)), *(f"LPT{n}" for n in range(1, 10))}
_STYLE_ATTRIBUTES = ("font", "fill", "border", "alignment", "protection")


def _safe_file_part(value: object, fallback: str, limit: int = 50) -> str:
    """Make a readable component that also works on Windows filesystems."""
    part = unicodedata.normalize("NFKC", str(value or ""))
    part = _INVALID_FILE_CHARS.sub("_", part)
    part = re.sub(r"\s+", " ", part).strip(" .")
    part = part[:limit].rstrip(" .") or fallback
    if part.upper() in _WINDOWS_RESERVED:
        part += "_"
    return part


def _unique_xlsx_name(stem: str, used: set[str]) -> str:
    """Avoid collisions after sanitization, truncation and case folding."""
    stem = stem[:120].rstrip(" .") or "结果"
    suffix = 1
    while True:
        tail = "" if suffix == 1 else f" ({suffix})"
        name = f"{stem[:120 - len(tail)].rstrip(' .')}{tail}.xlsx"
        key = unicodedata.normalize("NFKC", name).casefold()
        if key not in used:
            used.add(key)
            return name
        suffix += 1


def _capture_row(cells: Sequence[object], preserve_styles: bool, styles: dict[int, tuple[object, ...]]) -> tuple[tuple[object, str, int | None], ...]:
    """Keep values and formats on disk while caching each distinct style once."""
    captured = []
    for cell in cells:
        style_id = None
        if preserve_styles and bool(getattr(cell, "has_style", False)):
            # ReadOnlyCell exposes _style_id in openpyxl 3.1; ordinary Cell
            # exposes style_id. A missing identifier gets a fresh cache slot.
            raw_style_id = getattr(cell, "_style_id", None)
            if raw_style_id is None:
                raw_style_id = getattr(cell, "style_id", None)
            style_id = int(raw_style_id) if raw_style_id is not None else min(styles, default=0) - 1
            if style_id not in styles:
                styles[style_id] = tuple(copy(getattr(cell, attr)) for attr in _STYLE_ATTRIBUTES)
        captured.append((
            getattr(cell, "value", None),
            getattr(cell, "number_format", "General") or "General",
            style_id,
        ))
    return tuple(captured)


def _restore_cell(record: tuple[object, str, int | None], styles: dict[int, tuple[object, ...]]):
    value, number_format, style_id = record
    attributes = {"value": value, "number_format": number_format, "has_style": style_id is not None}
    if style_id is not None:
        attributes.update(zip(_STYLE_ATTRIBUTES, styles[style_id]))
    return SimpleNamespace(**attributes)


def _add_accounting_sheets(book: Workbook, accountant: AccountingAccumulator, config: AccountingConfig, labels: list[str]) -> None:
    """Use the same per-file summaries and checks as a single export."""
    if config.amount_columns or config.group_column is not None:
        summary_sheet = book.create_sheet("金额汇总")
        summary_sheet.freeze_panes = "A2"
        for index, row in enumerate(accountant.iter_summary_rows(labels)):
            _append_accounting_row(summary_sheet, row, header=index == 0)
        summary_sheet.column_dimensions["A"].width = 14
        summary_sheet.column_dimensions["B"].width = 26
        summary_sheet.column_dimensions["C"].width = 15
        for index in range(4, 4 + len(config.amount_columns)):
            summary_sheet.column_dimensions[get_column_letter(index)].width = 24
    if config.amount_columns or config.required_columns or config.duplicate_columns:
        issues_sheet = book.create_sheet("数据检查")
        issues_sheet.freeze_panes = "A2"
        _append_accounting_row(issues_sheet, ["原始行号", "问题类型", "说明"], header=True)
        issues = accountant.issue_rows()
        if issues:
            for issue in issues:
                _append_accounting_row(issues_sheet, issue)
            omitted = accountant.issue_count - len(issues)
            if omitted:
                _append_accounting_row(issues_sheet, ["", "记录上限", f"另外 {omitted} 条问题未列出"])
        else:
            _append_accounting_row(issues_sheet, ["", "无问题", "所选检查未发现异常"])
        issues_sheet.column_dimensions["A"].width = 15
        issues_sheet.column_dimensions["B"].width = 16
        issues_sheet.column_dimensions["C"].width = 70


def _remove_owned_temp(root: Path, target: Path | None) -> None:
    if target is None or not target.exists():
        return
    # A recursive removal is only permitted for the temporary directory we
    # created immediately within the explicitly selected output directory.
    resolved_root = root.resolve()
    resolved_target = target.resolve()
    if resolved_target.parent != resolved_root or not target.name.startswith(".excel_filter_"):
        raise RuntimeError(f"拒绝清理非本次创建的临时目录：{target}")
    shutil.rmtree(target)


def _commit_batch_dir(root: Path, stage: Path, folder_stem: str) -> Path:
    for _ in range(20):
        destination = root / f"{folder_stem}_{uuid.uuid4().hex[:8]}"
        if destination.exists():
            continue
        try:
            os.rename(stage, destination)
            return destination
        except FileExistsError:
            continue
    raise FileExistsError("无法创建唯一的批量导出文件夹，请重试。")


def _close_batch_workbook(book: Workbook) -> None:
    """Release write-only XML files when saving is interrupted."""
    for sheet in book.worksheets:
        rows = getattr(sheet, "_rows", None)
        if rows is not None:
            try:
                rows.close()
            except Exception:
                pass
        writer = getattr(sheet, "_writer", None)
        if writer is not None and os.path.exists(writer.out):
            try:
                writer.close()
                writer.cleanup()
            except (OSError, ValueError):
                pass
    book.close()


def _write_batch_file(
    target: Path,
    sheet_name: str,
    labels: list[str],
    header_rows: Sequence[tuple[tuple[object, str, int | None], ...]],
    row_cursor,
    styles: dict[int, tuple[object, ...]],
    preserve_styles: bool,
    accounting: AccountingConfig,
    cancel_event,
    progress: ProgressCallback | None,
    written_before: int,
    total_written: int,
) -> int:
    book = Workbook(write_only=True)
    sheet = book.create_sheet(sheet_name)
    sheet.freeze_panes = f"A{len(header_rows) + 1}"
    for index, label in enumerate(labels, 1):
        sheet.column_dimensions[get_column_letter(index)].width = min(max(_display_width(label) + 2, 12), 38)

    accounting_enabled = bool(
        accounting.amount_columns or accounting.group_column is not None
        or accounting.required_columns or accounting.duplicate_columns
    )
    accountant = AccountingAccumulator(accounting) if accounting_enabled else None
    written = 0
    try:
        for row in header_rows:
            sheet.append([
                _output_cell(sheet, _restore_cell(cell, styles), header=True, preserve_styles=preserve_styles)
                for cell in row
            ])
        for source_row, blob in row_cursor:
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError()
            row = pickle.loads(blob)
            sheet.append([
                _output_cell(sheet, _restore_cell(cell, styles), header=False, preserve_styles=preserve_styles)
                for cell in row
            ])
            written += 1
            if accountant is not None:
                accountant.add_row([cell[0] for cell in row], source_row)
            if written % 2000 == 0:
                _progress(progress, "save", written_before + written, total_written, written_before + written)

        first_header_row = len(header_rows)
        sheet.auto_filter.ref = f"A{first_header_row}:{get_column_letter(len(labels))}{first_header_row + written}"
        if accountant is not None:
            _add_accounting_sheets(book, accountant, accounting, labels)
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        book.save(target)
        return written
    finally:
        if accountant is not None:
            accountant.close()
        _close_batch_workbook(book)


def _export_batch_impl(
    options: ExportOptions,
    output_dir: Path,
    *,
    common_rules: Sequence[FilterRule] = (),
    split_columns: Sequence[int] = (),
    jobs: Sequence[BatchJob] | None = None,
    prefix: str | None = None,
    max_files: int = 200,
    progress: ProgressCallback | None = None,
    cancel_event=None,
) -> BatchResult:
    first_data = _check_layout(options.header_start, options.header_end, options.data_start, options.data_end)
    source = _check_source(options.source)
    root = Path(output_dir).expanduser().resolve()
    if root.exists() and not root.is_dir():
        raise ValueError(f"输出位置不是文件夹：{root}")
    root.mkdir(parents=True, exist_ok=True)
    if not isinstance(max_files, int) or isinstance(max_files, bool) or max_files < 1:
        raise ValueError("最多导出文件数必须是正整数。")

    labels = inspect_headers(source, options.sheet, options.header_start, options.header_end)
    width = len(labels)
    _validate_accounting(options.accounting, width)
    split_mode = jobs is None
    if split_mode:
        columns = tuple(split_columns)
        if not columns or len(set(columns)) != len(columns) or any(
            not isinstance(column, int) or isinstance(column, bool) or column < 1 or column > width
            for column in columns
        ):
            raise ValueError("拆分列必须是表头范围内且不重复的列号。")
        prepared_common = _validate_rules(common_rules, width)
        groups: list[tuple[str, ...] | str] = []
        group_lookup: dict[tuple[str, ...], int] = {}
        prepared_jobs: list[list[tuple[int, str, tuple[str, ...]]]] = []
    else:
        columns = ()
        job_list = list(jobs)
        if not job_list:
            raise ValueError("请至少添加一套批量筛选条件。")
        if len(job_list) > max_files:
            raise ValueError(f"筛选任务有 {len(job_list)} 套，超过最多 {max_files} 个文件的限制。")
        groups = [job.name for job in job_list]
        prepared_jobs = [_validate_rules(job.rules, width) for job in job_list]
        prepared_common = []
        group_lookup = {}

    spool_root = Path(tempfile.gettempdir()).resolve()
    spool_dir: Path | None = None
    stage: Path | None = None
    db: sqlite3.Connection | None = None
    book = None
    try:
        spool_dir = Path(tempfile.mkdtemp(prefix=".excel_filter_spool_", dir=spool_root))
        db = sqlite3.connect(spool_dir / "rows.sqlite3")
        db.execute("CREATE TABLE rows (batch_id INTEGER NOT NULL, source_row INTEGER NOT NULL, payload BLOB NOT NULL)")
        styles: dict[int, tuple[object, ...]] = {}
        book, ws = _open_sheet(source, options.sheet)
        header_rows = tuple(
            _capture_row(row, options.preserve_styles, styles)
            for row in ws.iter_rows(min_row=options.header_start, max_row=options.header_end, max_col=width)
        )
        total = _total_rows(ws, first_data, options.data_end)
        scanned = 0
        matched = 0
        stopped_at = None
        for row_number, cells in enumerate(
            ws.iter_rows(min_row=first_data, max_row=options.data_end, max_col=width),
            start=first_data,
        ):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError()
            if options.stop_at_summary and _is_summary_row(cells):
                stopped_at = row_number
                break
            if _is_blank_row(cells):
                continue
            scanned += 1
            if split_mode:
                if not _matches(cells, prepared_common):
                    batch_ids: list[int] = []
                else:
                    key = tuple(normalize_value(cells[column - 1].value) for column in columns)
                    batch_id = group_lookup.get(key)
                    if batch_id is None:
                        if len(groups) >= max_files:
                            raise ValueError(
                                f"拆分组合超过 {max_files} 个文件的上限；请增加公共筛选条件或调整拆分列。"
                            )
                        batch_id = len(groups)
                        groups.append(key)
                        group_lookup[key] = batch_id
                    batch_ids = [batch_id]
            else:
                batch_ids = [index for index, rules in enumerate(prepared_jobs) if _matches(cells, rules)]
            if batch_ids:
                record = _capture_row(cells, options.preserve_styles, styles)
                blob = pickle.dumps(record, protocol=pickle.HIGHEST_PROTOCOL)
                db.executemany(
                    "INSERT INTO rows (batch_id, source_row, payload) VALUES (?, ?, ?)",
                    ((batch_id, row_number, blob) for batch_id in batch_ids),
                )
                matched += len(batch_ids)
            if scanned % 2000 == 0:
                db.commit()
                _progress(progress, "export", row_number - first_data + 1, total, matched)
        db.commit()
        _progress(progress, "export", scanned, total, matched)
        book.close()
        book = None

        if split_mode and not groups:
            raise ValueError("公共筛选条件没有命中明细行，未生成拆分文件。")
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        db.execute("CREATE INDEX rows_by_batch ON rows (batch_id)")
        db.commit()

        stage = Path(tempfile.mkdtemp(prefix=".excel_filter_stage_", dir=root))
        safe_prefix = _safe_file_part(prefix if prefix is not None else source.stem, "批量导出", 38)
        used_names: set[str] = set()
        filenames: list[str] = []
        for group in groups:
            if split_mode:
                assert isinstance(group, tuple)
                parts = []
                for column, value in zip(columns, group):
                    label = labels[column - 1].split(" · ", 1)[-1]
                    parts.append(f"{_safe_file_part(label, get_column_letter(column), 25)}={_safe_file_part(value, '（空白）', 35)}")
                name_stem = f"{safe_prefix}_{'_'.join(parts)}"
            else:
                name_stem = f"{safe_prefix}_{_safe_file_part(group, '条件', 70)}"
            filenames.append(_unique_xlsx_name(name_stem, used_names))

        written = 0
        _progress(progress, "save", 0, matched, 0)
        for batch_id, filename in enumerate(filenames):
            if cancel_event is not None and cancel_event.is_set():
                raise CancelledError()
            cursor = db.execute(
                "SELECT source_row, payload FROM rows WHERE batch_id = ? ORDER BY rowid",
                (batch_id,),
            )
            written += _write_batch_file(
                stage / filename, options.sheet, labels, header_rows, cursor, styles,
                options.preserve_styles, options.accounting, cancel_event, progress, written, matched,
            )
            _progress(progress, "save", written, matched, written)
        if cancel_event is not None and cancel_event.is_set():
            raise CancelledError()
        if written != matched:
            raise RuntimeError("批量导出行数校验失败，未保留本次结果。")
        db.close()
        db = None
        folder_stem = f"{safe_prefix}_批量导出_{datetime.now():%Y%m%d_%H%M%S}"
        destination = _commit_batch_dir(root, stage, folder_stem)
        stage = None
        files = tuple(destination / filename for filename in filenames)
        _progress(progress, "done", scanned, total, matched)
        return BatchResult(scanned, matched, len(files), destination, files, stopped_at)
    finally:
        if book is not None:
            book.close()
        if db is not None:
            db.close()
        _remove_owned_temp(root, stage)
        _remove_owned_temp(spool_root, spool_dir)


def export_batch(
    options: ExportOptions,
    rules: Sequence[FilterRule],
    split_columns: Sequence[int],
    output_dir: Path,
    prefix: str | None = None,
    max_files: int = 200,
    progress: ProgressCallback | None = None,
    cancel_event=None,
) -> BatchResult:
    """Export one file per actual split-column value combination.

    Rules first restrict the common source rows. Multiple values inside a
    FilterRule are alternatives; different rules are combined with AND.
    Source rows are scanned once and spooled to disk before writing one
    workbook at a time into a unique output subdirectory.
    """
    return _export_batch_impl(
        options, output_dir, common_rules=rules, split_columns=split_columns,
        prefix=prefix, max_files=max_files, progress=progress, cancel_event=cancel_event,
    )


def export_named_batch(
    options: ExportOptions,
    jobs: Sequence[BatchJob],
    output_dir: Path,
    max_files: int = 200,
    progress: ProgressCallback | None = None,
    cancel_event=None,
) -> BatchResult:
    """Export each named job's independent conditions to a separate file.

    Jobs with no matching rows still produce a header-only workbook. A source
    row may appear in more than one file when job rules overlap.
    """
    return _export_batch_impl(
        options, output_dir, jobs=jobs, max_files=max_files,
        progress=progress, cancel_event=cancel_event,
    )
