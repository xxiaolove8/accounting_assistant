"""Streaming, header-aligned consolidation of Excel detail worksheets.

Source workbooks are opened read-only.  The destination is saved to a file in
its own directory and replaced only after the workbook has been fully written.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path
import os
import pickle
import re
import sqlite3
import tempfile
from typing import Callable

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from filter_core import CancelledError, normalize_value


ProgressCallback = Callable[[str, int, int | None, int], None]
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
_SUMMARY_LABEL = re.compile(
    r"^\s*(?:合\s*计|小\s*计|总\s*计|汇\s*总|总\s*合\s*计|"
    r"grand\s+total|sub\s*total|total)\s*[:：]?\s*$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class MergeOptions:
    sources: tuple[Path, ...]
    output: Path
    sheet_name: str | None = None
    header_start: int = 1
    header_end: int = 1
    data_start: int | None = None
    data_end: int | None = None
    stop_at_summary: bool = True
    add_source_column: bool = False
    deduplicate: bool = False
    match_mode: str = "union"  # "union" or "strict"


@dataclass(frozen=True)
class MergeResult:
    files: int
    scanned_rows: int
    merged_rows: int
    duplicates_skipped: int
    output: Path
    headers: tuple[str, ...]
    sheets: int


@dataclass(frozen=True)
class _SourceInfo:
    path: Path
    sheet: str
    headers: tuple[str, ...]
    keys: tuple[str, ...]
    estimated_rows: int | None


def _progress(callback: ProgressCallback | None, phase: str, processed: int,
              total: int | None, merged: int) -> None:
    if callback is not None:
        callback(phase, processed, total, merged)


def _cancel_if_requested(cancel_event) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError()


def _header_key(label: str) -> str:
    # The first encountered spelling remains the output display name.
    return " ".join(label.split()).casefold()


def _validate_options(options: MergeOptions) -> tuple[tuple[Path, ...], Path, int]:
    sources = tuple(Path(path).expanduser().resolve() for path in options.sources)
    if not sources:
        raise ValueError("请至少选择一个源 Excel 文件。")
    if len(set(sources)) != len(sources):
        raise ValueError("源文件列表包含重复文件，请移除重复项。")
    for path in sources:
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError(f"源文件必须是 .xlsx 或 .xlsm：{path.name}")
        if not path.is_file():
            raise FileNotFoundError(f"找不到源文件：{path}")

    output = Path(options.output).expanduser().resolve()
    if output.suffix.lower() != ".xlsx":
        raise ValueError("输出文件必须是 .xlsx。")
    if output in sources:
        raise ValueError("输出路径不能与任何源文件相同。")
    if not output.parent.is_dir():
        raise FileNotFoundError(f"输出文件夹不存在：{output.parent}")
    if options.match_mode not in {"union", "strict"}:
        raise ValueError("表头匹配方式必须是 union 或 strict。")
    if not isinstance(options.header_start, int) or isinstance(options.header_start, bool) or options.header_start < 1:
        raise ValueError("表头起始行必须是正整数。")
    if not isinstance(options.header_end, int) or isinstance(options.header_end, bool) or options.header_end < options.header_start:
        raise ValueError("表头结束行不能早于表头起始行。")
    first_data = options.data_start if options.data_start is not None else options.header_end + 1
    if not isinstance(first_data, int) or isinstance(first_data, bool) or first_data <= options.header_end:
        raise ValueError("明细起始行必须在表头结束行之后。")
    if options.data_end is not None and (
        not isinstance(options.data_end, int) or isinstance(options.data_end, bool)
        or options.data_end < first_data
    ):
        raise ValueError("明细结束行不能早于明细起始行。")
    if options.sheet_name is not None and not str(options.sheet_name).strip():
        raise ValueError("工作表名称不能为空。")
    return sources, output, first_data


def _read_headers(ws, path: Path, start: int, end: int) -> tuple[tuple[str, ...], tuple[str, ...]]:
    rows = list(ws.iter_rows(min_row=start, max_row=end, values_only=True))
    width = max(
        (index for row in rows for index, value in enumerate(row, 1) if normalize_value(value)),
        default=0,
    )
    if width == 0:
        raise ValueError(f"{path.name} 的第 {start}–{end} 行没有表头，请调整表头行号。")
    if width > EXCEL_MAX_COLUMNS:
        raise ValueError(f"{path.name} 的表头超过 Excel 的 {EXCEL_MAX_COLUMNS} 列限制。")

    headers: list[str] = []
    keys: list[str] = []
    first_column_by_key: dict[str, int] = {}
    for column in range(1, width + 1):
        parts: list[str] = []
        for row in rows:
            part = normalize_value(row[column - 1]) if column <= len(row) else ""
            if part and (not parts or part != parts[-1]):
                parts.append(part)
        if not parts:
            raise ValueError(
                f"{path.name} 的表头第 {get_column_letter(column)} 列为空；"
                "请填写列名或调整表头范围，避免合并时遗漏数据。"
            )
        label = " / ".join(parts)
        key = _header_key(label)
        earlier = first_column_by_key.get(key)
        if earlier is not None:
            raise ValueError(
                f"{path.name} 的表头“{label}”在第 {get_column_letter(earlier)} 列和"
                f"第 {get_column_letter(column)} 列重复，请先区分这两列。"
            )
        first_column_by_key[key] = column
        headers.append(label)
        keys.append(key)
    return tuple(headers), tuple(keys)


def _inspect_source(path: Path, options: MergeOptions, first_data: int) -> _SourceInfo:
    book = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        if options.sheet_name is None:
            if not book.sheetnames:
                raise ValueError(f"{path.name} 没有可读取的工作表。")
            sheet_name = book.sheetnames[0]
        else:
            sheet_name = options.sheet_name
            if sheet_name not in book.sheetnames:
                raise ValueError(f"{path.name} 缺少指定工作表“{sheet_name}”。")
        ws = book[sheet_name]
        headers, keys = _read_headers(ws, path, options.header_start, options.header_end)
        maximum = ws.max_row
        if maximum is None:
            estimated = None
        else:
            last = min(maximum, options.data_end) if options.data_end is not None else maximum
            estimated = max(0, last - first_data + 1)
        return _SourceInfo(path, sheet_name, headers, keys, estimated)
    finally:
        book.close()


def _output_headers(infos: list[_SourceInfo], options: MergeOptions) -> tuple[tuple[str, ...], tuple[str, ...]]:
    first_keys = set(infos[0].keys)
    if options.match_mode == "strict":
        for info in infos[1:]:
            actual = set(info.keys)
            if actual != first_keys:
                missing = [label for label, key in zip(infos[0].headers, infos[0].keys) if key not in actual]
                extra = [label for label, key in zip(info.headers, info.keys) if key not in first_keys]
                details = []
                if missing:
                    details.append("缺少：" + "、".join(missing[:8]))
                if extra:
                    details.append("多出：" + "、".join(extra[:8]))
                raise ValueError(f"{info.path.name} 与首个文件的表头不一致（{'；'.join(details)}）。")

    labels: list[str] = []
    keys: list[str] = []
    seen: set[str] = set()
    for info in infos:
        for label, key in zip(info.headers, info.keys):
            if key not in seen:
                seen.add(key)
                labels.append(label)
                keys.append(key)

    if options.add_source_column:
        source_label = "来源文件"
        number = 2
        while _header_key(source_label) in seen:
            source_label = f"来源文件（合并{number}）"
            number += 1
        labels.append(source_label)

    if len(labels) > EXCEL_MAX_COLUMNS:
        raise ValueError(f"合并后共有 {len(labels)} 列，超过 Excel 的 {EXCEL_MAX_COLUMNS} 列限制。")
    return tuple(labels), tuple(keys)


def _is_blank_row(cells) -> bool:
    return all(not normalize_value(cell.value) for cell in cells)


def _is_summary_row(cells) -> bool:
    return any(_SUMMARY_LABEL.fullmatch(normalize_value(cell.value)) for cell in cells[:4])


def _header_cell(sheet, value: str):
    cell = WriteOnlyCell(sheet, value=value)
    cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    cell.fill = PatternFill(fill_type="solid", fgColor="24476B")
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    return cell


def _data_cell(sheet, source_cell):
    if source_cell is None:
        return None
    value = source_cell.value
    number_format = source_cell.number_format or "General"
    if number_format == "General" and not (isinstance(value, str) and value.startswith("=")):
        return value
    cell = WriteOnlyCell(sheet, value=value)
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"  # Cached formula text is data, not a new formula.
    cell.number_format = number_format
    return cell


def _new_sheet(book: Workbook, number: int, headers: tuple[str, ...]):
    sheet = book.create_sheet("总表" if number == 1 else f"总表{number}")
    sheet.freeze_panes = "A2"
    for index, label in enumerate(headers, 1):
        # Set dimensions before rows are streamed, as required by write-only mode.
        sheet.column_dimensions[get_column_letter(index)].width = min(max(len(label) + 2, 12), 38)
    sheet.append([_header_cell(sheet, label) for label in headers])
    return sheet


def _close_write_only_book(book: Workbook) -> None:
    for sheet in book.worksheets:
        rows = getattr(sheet, "_rows", None)
        if rows is not None:
            try:
                rows.close()
            except Exception:
                pass
        writer = getattr(sheet, "_writer", None)
        if writer is not None:
            try:
                if os.path.exists(writer.out):
                    writer.close()
                    writer.cleanup()
            except (OSError, ValueError):
                pass
    book.close()


def merge_workbooks(options: MergeOptions, progress: ProgressCallback | None = None,
                    cancel_event=None) -> MergeResult:
    """Merge source detail rows by heading into one atomic .xlsx workbook.

    ``union`` adds new columns as they occur; ``strict`` requires the same
    heading set in every source, regardless of column order.  Duplicate rows
    compare aligned business values and ignore the optional source-file column.
    Source formulas are read as their last saved cached values.
    """
    sources, output, first_data = _validate_options(options)
    infos: list[_SourceInfo] = []
    for index, path in enumerate(sources, 1):
        _cancel_if_requested(cancel_event)
        infos.append(_inspect_source(path, options, first_data))
        _progress(progress, "inspect", index, len(sources), 0)
    headers, keys = _output_headers(infos, options)
    total = None if any(info.estimated_rows is None for info in infos) else sum(
        info.estimated_rows or 0 for info in infos
    )

    book = Workbook(write_only=True)
    sheet_count = 1
    sheet = _new_sheet(book, sheet_count, headers)
    rows_on_sheet = 1
    scanned_rows = 0
    examined_rows = 0
    merged_rows = 0
    duplicates_skipped = 0
    temp_path: Path | None = None
    dedup_dir: tempfile.TemporaryDirectory[str] | None = None
    db: sqlite3.Connection | None = None
    try:
        if options.deduplicate:
            dedup_dir = tempfile.TemporaryDirectory(prefix=".excel_merge_dedup_")
            db = sqlite3.connect(Path(dedup_dir.name) / "seen.sqlite3")
            db.execute("CREATE TABLE seen (row_values BLOB PRIMARY KEY) WITHOUT ROWID")

        for info in infos:
            _cancel_if_requested(cancel_event)
            source_book = load_workbook(info.path, read_only=True, data_only=True, keep_links=False)
            try:
                ws = source_book[info.sheet]
                width = len(info.headers)
                by_key = {key: index for index, key in enumerate(info.keys)}
                positions = tuple(by_key.get(key) for key in keys)
                for cells in ws.iter_rows(min_row=first_data, max_row=options.data_end, max_col=width):
                    _cancel_if_requested(cancel_event)
                    examined_rows += 1
                    if options.stop_at_summary and _is_summary_row(cells):
                        break
                    if _is_blank_row(cells):
                        continue
                    scanned_rows += 1
                    aligned = tuple(cells[index] if index is not None else None for index in positions)
                    if db is not None:
                        values = tuple(cell.value if cell is not None else None for cell in aligned)
                        blob = pickle.dumps(values, protocol=pickle.HIGHEST_PROTOCOL)
                        cursor = db.execute("INSERT OR IGNORE INTO seen (row_values) VALUES (?)", (blob,))
                        if cursor.rowcount == 0:
                            duplicates_skipped += 1
                            if scanned_rows % 2000 == 0:
                                db.commit()
                                _progress(progress, "merge", examined_rows, total, merged_rows)
                            continue
                        if scanned_rows % 2000 == 0:
                            db.commit()

                    if rows_on_sheet >= EXCEL_MAX_ROWS:
                        sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{rows_on_sheet}"
                        sheet_count += 1
                        sheet = _new_sheet(book, sheet_count, headers)
                        rows_on_sheet = 1
                    row = [_data_cell(sheet, cell) for cell in aligned]
                    if options.add_source_column:
                        filename_cell = WriteOnlyCell(sheet, value=info.path.name)
                        filename_cell.data_type = "s"
                        row.append(filename_cell)
                    sheet.append(row)
                    rows_on_sheet += 1
                    merged_rows += 1
                    if scanned_rows % 2000 == 0:
                        _progress(progress, "merge", examined_rows, total, merged_rows)
            finally:
                source_book.close()
            _progress(progress, "merge", examined_rows, total, merged_rows)

        sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{rows_on_sheet}"
        if db is not None:
            db.commit()
        _cancel_if_requested(cancel_event)
        _progress(progress, "save", examined_rows, total, merged_rows)
        descriptor, temp_name = tempfile.mkstemp(prefix=".excel_merge_", suffix=".xlsx", dir=output.parent)
        os.close(descriptor)
        temp_path = Path(temp_name)
        book.save(temp_path)
        _cancel_if_requested(cancel_event)
        os.replace(temp_path, output)
        temp_path = None
        _progress(progress, "done", examined_rows, total, merged_rows)
        return MergeResult(len(sources), scanned_rows, merged_rows, duplicates_skipped,
                           output, headers, sheet_count)
    finally:
        if db is not None:
            db.close()
        if dedup_dir is not None:
            dedup_dir.cleanup()
        _close_write_only_book(book)
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except OSError:
                pass
