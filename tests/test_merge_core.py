"""Behavioral tests for merging independently saved Excel detail tables.

All inputs are temporary workbooks. The user's financial sample is not opened.
"""

from __future__ import annotations

from datetime import datetime
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event
import unittest

from openpyxl import Workbook, load_workbook

from filter_core import CancelledError, ExportOptions, export_batch
from merge_core import MergeOptions, merge_workbooks


class MergeCoreTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)

    def make_source(
        self,
        name: str,
        headers: list[object],
        rows: list[tuple[object, ...]] | None = None,
    ) -> Path:
        path = self.directory / name
        book = Workbook()
        sheet = book.active
        sheet.title = "明细"
        sheet.append(headers)
        for row in rows or []:
            sheet.append(row)
        book.save(path)
        book.close()
        return path

    def options(self, *sources: Path, name: str = "combined.xlsx", **changes) -> MergeOptions:
        return MergeOptions(sources=tuple(sources), output=self.directory / name, **changes)

    @staticmethod
    def output_rows(path: Path) -> list[tuple[object, ...]]:
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            # Write-only sheets can omit a dimension tag and read-only mode
            # then trims trailing empty cells. Pad each row to header width.
            sheet = book.active
            rows = [tuple(row) for row in sheet.iter_rows(values_only=True)]
            width = len(rows[0]) if rows else 0
            return [row + (None,) * (width - len(row)) for row in rows]
        finally:
            book.close()

    def test_matches_headers_even_when_source_columns_are_reordered(self) -> None:
        first = self.make_source(
            "first.xlsx",
            ["公司", "部门", "金额", "日期"],
            [("甲公司", "销售一部", 100, datetime(2026, 6, 1))],
        )
        second = self.make_source(
            "second.xlsx",
            ["金额", "日期", "部门", "公司"],
            [(250, datetime(2026, 6, 2), "销售二部", "乙公司")],
        )
        options = self.options(first, second)

        result = merge_workbooks(options)

        self.assertEqual(result.merged_rows, 2)
        self.assertEqual(result.scanned_rows, 2)
        self.assertEqual(result.headers, ("公司", "部门", "金额", "日期"))
        self.assertEqual(
            self.output_rows(options.output),
            [
                ("公司", "部门", "金额", "日期"),
                ("甲公司", "销售一部", 100, datetime(2026, 6, 1)),
                ("乙公司", "销售二部", 250, datetime(2026, 6, 2)),
            ],
        )

    def test_union_adds_missing_columns_and_leaves_missing_cells_blank(self) -> None:
        first = self.make_source("first.xlsx", ["公司", "金额"], [("甲公司", 100)])
        second = self.make_source("second.xlsx", ["备注", "公司"], [("新客户", "乙公司")])
        options = self.options(first, second, match_mode="union")

        result = merge_workbooks(options)

        self.assertEqual(result.headers, ("公司", "金额", "备注"))
        self.assertEqual(
            self.output_rows(options.output),
            [("公司", "金额", "备注"), ("甲公司", 100, None), ("乙公司", None, "新客户")],
        )

    def test_strict_mode_rejects_mismatched_field_sets_before_creating_output(self) -> None:
        first = self.make_source("first.xlsx", ["公司", "金额"], [("甲公司", 100)])
        second = self.make_source("second.xlsx", ["公司", "备注"], [("乙公司", "新客户")])
        options = self.options(first, second, match_mode="strict")

        with self.assertRaises(ValueError):
            merge_workbooks(options)

        self.assertFalse(options.output.exists())

    def test_strict_mode_accepts_the_same_fields_in_a_different_order(self) -> None:
        first = self.make_source("first.xlsx", ["公司", "金额"], [("甲公司", 100)])
        second = self.make_source("second.xlsx", ["金额", "公司"], [(200, "乙公司")])
        options = self.options(first, second, match_mode="strict")

        result = merge_workbooks(options)

        self.assertEqual(result.merged_rows, 2)
        self.assertEqual(self.output_rows(options.output), [("公司", "金额"), ("甲公司", 100), ("乙公司", 200)])

    def test_duplicate_or_empty_header_is_rejected_as_ambiguous(self) -> None:
        for name, headers in [
            ("duplicate.xlsx", ["公司", "金额", "金额"]),
            ("blank.xlsx", ["公司", None, "金额"]),
        ]:
            with self.subTest(name=name):
                source = self.make_source(name, headers, [("甲公司", 100, 200)])
                options = self.options(source, name=f"{Path(name).stem}-result.xlsx")

                with self.assertRaises(ValueError):
                    merge_workbooks(options)

                self.assertFalse(options.output.exists())

    def test_stops_at_summary_and_does_not_include_trailing_rows(self) -> None:
        first = self.make_source(
            "first.xlsx", ["公司", "金额"],
            [("甲公司", 100), ("合计", 100), ("尾部记录", 999)],
        )
        second = self.make_source("second.xlsx", ["公司", "金额"], [("乙公司", 200)])
        options = self.options(first, second)

        result = merge_workbooks(options)

        self.assertEqual(result.merged_rows, 2)
        self.assertEqual(self.output_rows(options.output), [("公司", "金额"), ("甲公司", 100), ("乙公司", 200)])

    def test_source_column_and_deduplication_across_files(self) -> None:
        first = self.make_source("first.xlsx", ["公司", "金额"], [("甲公司", 100), ("乙公司", 200)])
        second = self.make_source("second.xlsx", ["金额", "公司"], [(100, "甲公司"), (300, "丙公司")])
        options = self.options(first, second, add_source_column=True, deduplicate=True)

        result = merge_workbooks(options)

        self.assertEqual(result.merged_rows, 3)
        self.assertEqual(result.duplicates_skipped, 1)
        self.assertEqual(len(result.headers), 3)
        rows = self.output_rows(options.output)
        company = result.headers.index("公司")
        amount = result.headers.index("金额")
        source = next(index for index, header in enumerate(result.headers) if header not in {"公司", "金额"})
        self.assertEqual(rows[0], result.headers)
        self.assertEqual(
            [(row[company], row[amount], row[source]) for row in rows[1:]],
            [("甲公司", 100, "first.xlsx"), ("乙公司", 200, "first.xlsx"), ("丙公司", 300, "second.xlsx")],
        )

    def test_typed_values_number_formats_and_source_bytes_are_preserved(self) -> None:
        date = datetime(2026, 7, 15)
        first = self.make_source("first.xlsx", ["日期", "金额"], [(date, 1234.5)])
        second = self.make_source("second.xlsx", ["金额", "日期"], [(20.75, datetime(2026, 7, 16))])
        book = load_workbook(first)
        try:
            book.active["A2"].number_format = "yyyy-mm-dd"
            book.active["B2"].number_format = "#,##0.00"
            book.save(first)
        finally:
            book.close()
        before = {source: sha256(source.read_bytes()).digest() for source in (first, second)}
        options = self.options(first, second)

        merge_workbooks(options)

        self.assertEqual({source: sha256(source.read_bytes()).digest() for source in (first, second)}, before)
        output = load_workbook(options.output, read_only=False, data_only=True)
        try:
            sheet = output.active
            self.assertEqual(sheet["A2"].value, date)
            self.assertEqual(sheet["B2"].value, 1234.5)
            self.assertIsInstance(sheet["A2"].value, datetime)
            self.assertIsInstance(sheet["B2"].value, (int, float))
            self.assertEqual(sheet["A2"].number_format, "yyyy-mm-dd")
            self.assertEqual(sheet["B2"].number_format, "#,##0.00")
            self.assertEqual(sheet["A3"].value, datetime(2026, 7, 16))
            self.assertEqual(sheet["B3"].value, 20.75)
        finally:
            output.close()

    def test_header_only_file_contributes_fields_but_no_detail_rows(self) -> None:
        empty = self.make_source("empty.xlsx", ["公司", "金额", "备注"])
        populated = self.make_source("populated.xlsx", ["金额", "公司"], [(50, "甲公司")])
        options = self.options(empty, populated)

        result = merge_workbooks(options)

        self.assertEqual(result.merged_rows, 1)
        self.assertEqual(self.output_rows(options.output), [("公司", "金额", "备注"), ("甲公司", 50, None)])

    def test_seven_value_split_then_merge_restores_all_detail_rows(self) -> None:
        details = [
            ("A001", "销售一部", 100),
            ("B002", "销售二部", 200),
            ("C003", "销售三部", 300),
            ("D004", "销售四部", 400),
            ("E005", "财务部", 500),
            ("F006", "行政部", 600),
            ("G007", "研发部", 700),
            ("H008", "销售一部", 800),
        ]
        source = self.make_source(
            "original.xlsx", ["编号", "部门", "金额"],
            [*details, ("合计", None, 3600), ("AFTER", "销售一部", 9999)],
        )
        batch = export_batch(
            ExportOptions(source=source, sheet="明细", output=self.directory / "unused.xlsx"),
            [], (2,), self.directory / "by-department",
        )
        self.assertEqual(batch.file_count, 7)

        options = self.options(*batch.files, name="restored.xlsx")
        result = merge_workbooks(options)

        self.assertEqual(result.files, 7)
        self.assertEqual(result.merged_rows, 8)
        rows = self.output_rows(options.output)
        self.assertEqual(rows[0], ("编号", "部门", "金额"))
        self.assertCountEqual(rows[1:], details)
        self.assertEqual({row[0] for row in rows[1:]}, {row[0] for row in details})

    def test_output_path_must_not_overwrite_a_source(self) -> None:
        source = self.make_source("source.xlsx", ["公司", "金额"], [("甲公司", 100)])
        before = sha256(source.read_bytes()).digest()
        options = MergeOptions(sources=(source,), output=source)

        with self.assertRaises(ValueError):
            merge_workbooks(options)

        self.assertEqual(sha256(source.read_bytes()).digest(), before)

    def test_cancel_leaves_no_output_or_staging_files(self) -> None:
        source = self.make_source("source.xlsx", ["公司", "金额"], [("甲公司", 100)])
        options = self.options(source)
        before = set(self.directory.iterdir())
        cancel_event = Event()
        cancel_event.set()

        with self.assertRaises(CancelledError):
            merge_workbooks(options, cancel_event=cancel_event)

        self.assertFalse(options.output.exists())
        self.assertEqual(set(self.directory.iterdir()), before)

    def test_cancel_after_rows_are_staged_keeps_existing_output_untouched(self) -> None:
        source = self.make_source("source.xlsx", ["公司", "金额"], [("甲公司", 100)])
        options = self.options(source)
        options.output.write_bytes(b"previous result")
        before = set(self.directory.iterdir())
        cancel_event = Event()
        saw_merge_progress = False

        def cancel_after_merge(phase: str, processed: int, total: int | None, merged: int) -> None:
            nonlocal saw_merge_progress
            if phase == "merge" and processed >= 1:
                saw_merge_progress = True
                cancel_event.set()

        with self.assertRaises(CancelledError):
            merge_workbooks(options, progress=cancel_after_merge, cancel_event=cancel_event)

        self.assertTrue(saw_merge_progress)
        self.assertEqual(options.output.read_bytes(), b"previous result")
        self.assertEqual(set(self.directory.iterdir()), before)


if __name__ == "__main__":
    unittest.main()
