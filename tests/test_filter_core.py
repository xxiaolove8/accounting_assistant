"""Behavioral tests for the workbook filtering engine.

The fixtures are created in a temporary directory.  No test reads or changes the
user's financial workbook.
"""

from __future__ import annotations

from dataclasses import replace
from datetime import datetime
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from accounting import AccountingConfig
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


HEADERS = ["公司", "部门", "日期", "金额", "备注", "编号"]


class FilterCoreTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.source = self.directory / "source.xlsx"

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "明细"
        sheet.append(HEADERS)
        sheet.append(["甲公司", "销售一部", datetime(2026, 6, 1), 100, "重点客户", "A001"])
        sheet.append(["乙公司", "销售二部", datetime(2026, 6, 2), 200, "重点订单", "B002"])
        sheet.append(["丙公司", "销售一部", None, 0, None, "C003"])
        sheet.append(["丁公司", "财务部", datetime(2026, 6, 3), 300, "常规业务", "D004"])
        sheet["A6"] = "合计"
        sheet["D6"] = "=SUM(D2:D5)"
        # This row must never be exported when summary stopping is enabled.
        sheet.append(["尾部记录", "销售一部", datetime(2026, 6, 4), 999, "重点客户", "AFTER"])
        sheet["C2"].number_format = "yyyy-mm-dd"
        sheet["D2"].number_format = "#,##0.00"
        sheet["D2"].fill = PatternFill(fill_type="solid", fgColor="FFCC99")
        workbook.create_sheet("说明")
        workbook.save(self.source)

    def options(self, name: str, *, preserve_styles: bool = False) -> ExportOptions:
        return ExportOptions(
            source=self.source,
            sheet="明细",
            output=self.directory / name,
            preserve_styles=preserve_styles,
        )

    @staticmethod
    def nonempty_rows(path: Path) -> list[tuple]:
        workbook = load_workbook(path, read_only=True, data_only=False)
        try:
            return [
                tuple(row)
                for row in workbook.active.iter_rows(values_only=True)
                if any(value is not None for value in row)
            ]
        finally:
            workbook.close()

    def batch_detail_ids(self, files: tuple[Path, ...]) -> set[frozenset[str]]:
        """Identify output groups by source IDs, regardless of generated filenames."""
        return {
            frozenset(row[5] for row in self.nonempty_rows(Path(path))[1:])
            for path in files
        }

    def assert_safe_batch_files(self, result) -> None:
        self.assertEqual(result.file_count, len(result.files))
        self.assertEqual(len(set(result.files)), len(result.files))
        self.assertTrue(Path(result.output_dir).is_dir())
        for path in result.files:
            path = Path(path)
            self.assertTrue(path.is_file())
            self.assertEqual(path.suffix.lower(), ".xlsx")
            self.assertEqual(path.parent.resolve(), Path(result.output_dir).resolve())
            self.assertNotRegex(path.name, r'[<>:"/\\|?*]')
            self.assertFalse(path.stem.endswith((" ", ".")))

    def test_sheet_and_header_inspection(self) -> None:
        self.assertEqual(list_sheets(self.source), ["明细", "说明"])
        labels = inspect_headers(self.source, "明细", 1, 1)
        self.assertEqual(len(labels), len(HEADERS))
        for index, (label, header) in enumerate(zip(labels, HEADERS), start=1):
            self.assertTrue(label.startswith(get_column_letter(index)), label)
            self.assertTrue(label.endswith(header), label)

    def test_same_column_or_and_different_column_and(self) -> None:
        rules = [
            FilterRule(column=2, operator="equals", values=("销售一部", "销售二部")),
            FilterRule(column=5, operator="contains", values=("重点",)),
        ]
        options = self.options("combined.xlsx")
        result = export_filtered(options, rules)

        self.assertEqual(result.matched_rows, 2)
        self.assertEqual(result.stopped_at_row, 6)
        rows = self.nonempty_rows(options.output)
        self.assertEqual(rows[0], tuple(HEADERS))
        self.assertEqual([row[5] for row in rows[1:]], ["A001", "B002"])
        self.assertEqual(rows[1][3], 100)
        self.assertEqual(rows[2][3], 200)
        self.assertEqual(rows[1][2], datetime(2026, 6, 1))

    def test_scan_values_round_trip_for_date_number_and_blank(self) -> None:
        options = self.options("typed.xlsx")
        date_values, date_truncated, _ = scan_values(options, 3, limit=100)
        number_values, number_truncated, _ = scan_values(options, 4, limit=100)
        remark_values, remark_truncated, _ = scan_values(options, 5, limit=100)
        self.assertFalse(date_truncated or number_truncated or remark_truncated)
        self.assertTrue(all(isinstance(value, str) for value in date_values + number_values + remark_values))

        date_choice = next(value for value in date_values if "2026" in value and "01" in value)
        number_choice = next(value for value in number_values if value.replace(",", "").startswith("100"))
        blank_choices = set(remark_values) - {"重点客户", "重点订单", "常规业务"}
        self.assertEqual(len(blank_choices), 1)
        blank_choice = blank_choices.pop()

        for name, column, choice, expected_id in [
            ("date.xlsx", 3, date_choice, "A001"),
            ("number.xlsx", 4, number_choice, "A001"),
            ("blank.xlsx", 5, blank_choice, "C003"),
        ]:
            with self.subTest(column=column):
                target = self.options(name)
                result = export_filtered(
                    target, [FilterRule(column=column, operator="equals", values=(choice,))]
                )
                self.assertEqual(result.matched_rows, 1)
                self.assertEqual([row[5] for row in self.nonempty_rows(target.output)[1:]], [expected_id])

    def test_scan_values_includes_blank_field_when_later_column_has_data(self) -> None:
        source = self.directory / "late-column-data.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "明细"
        sheet.append(HEADERS[:5])
        sheet.append((None, None, None, None, "仅第五列有值"))
        sheet.append(("甲公司", "销售一部", None, None, "普通记录"))
        workbook.save(source)
        workbook.close()

        options = replace(self.options("unused.xlsx"), source=source)
        values, truncated, scanned = scan_values(options, column=2)
        result = export_batch(options, [], (2,), self.directory / "late-column-batch")

        self.assertFalse(truncated)
        self.assertEqual(scanned, 2)
        self.assertEqual(values, ["", "销售一部"])
        self.assert_safe_batch_files(result)
        self.assertEqual(result.file_count, 2)
        detail_by_department = {
            rows[1][1] or "": rows[1:]
            for path in result.files
            if (rows := self.nonempty_rows(Path(path)))
        }
        self.assertEqual(set(detail_by_department), set(values))
        self.assertEqual(detail_by_department[""], [(None, None, None, None, "仅第五列有值")])

    def test_summary_stops_scan_and_source_is_unchanged(self) -> None:
        before = sha256(self.source.read_bytes()).digest()
        options = self.options("department.xlsx", preserve_styles=True)
        result = export_filtered(
            options, [FilterRule(column=2, operator="equals", values=("销售一部",))]
        )

        self.assertEqual(sha256(self.source.read_bytes()).digest(), before)
        self.assertEqual(result.scanned_rows, 4)
        self.assertEqual(result.matched_rows, 2)
        self.assertEqual(result.stopped_at_row, 6)
        rows = self.nonempty_rows(options.output)
        self.assertEqual([row[5] for row in rows[1:]], ["A001", "C003"])
        self.assertNotIn("AFTER", [row[5] for row in rows])

        original = load_workbook(self.source, read_only=False)
        filtered = load_workbook(options.output, read_only=False)
        try:
            self.assertEqual(filtered.active["C2"].number_format, original["明细"]["C2"].number_format)
            self.assertEqual(filtered.active["D2"].number_format, original["明细"]["D2"].number_format)
            self.assertEqual(filtered.active["D2"].fill.fgColor.rgb, original["明细"]["D2"].fill.fgColor.rgb)
        finally:
            original.close()
            filtered.close()

    def test_zero_matches_still_creates_header_only_workbook(self) -> None:
        options = self.options("empty.xlsx")
        result = export_filtered(
            options, [FilterRule(column=2, operator="equals", values=("研发部",))]
        )

        self.assertEqual(result.matched_rows, 0)
        self.assertTrue(options.output.is_file())
        self.assertEqual(self.nonempty_rows(options.output), [tuple(HEADERS)])

    def test_accounting_totals_and_groups_use_only_matched_detail_rows(self) -> None:
        options = replace(
            self.options("accounting_totals.xlsx"),
            accounting=AccountingConfig(amount_columns=(4,), group_column=2),
        )
        result = export_filtered(
            options,
            [FilterRule(column=2, operator="equals", values=("销售一部", "销售二部"))],
        )

        self.assertEqual(result.matched_rows, 3)
        self.assertEqual(result.stopped_at_row, 6)
        output = load_workbook(options.output, read_only=True, data_only=True)
        try:
            self.assertEqual(output.sheetnames, ["明细", "金额汇总", "数据检查"])
            summary = list(output["金额汇总"].iter_rows(values_only=True))
            self.assertEqual(summary[0][:3], ("范围", "分组", "命中行数"))
            self.assertTrue(summary[0][3].endswith("金额合计"))
            self.assertEqual(summary[1][:4], ("总计", None, 3, 300))
            groups = {row[1]: (row[2], row[3]) for row in summary[2:]}
            self.assertEqual(groups, {"销售一部": (2, 100), "销售二部": (1, 200)})
            self.assertEqual(output["数据检查"]["B2"].value, "无问题")
        finally:
            output.close()

    def test_accounting_checks_required_composite_duplicate_and_invalid_amount(self) -> None:
        source_book = load_workbook(self.source)
        try:
            source_sheet = source_book["明细"]
            source_sheet["F3"] = "A001"  # Same ID, different department: not a duplicate key.
            source_sheet["F4"] = "A001"  # Both department and ID match source row 2.
            source_sheet["D4"] = "not-a-number"
            source_sheet["E4"] = None
            source_sheet["D5"] = "also-invalid"  # Unmatched row must not be checked.
            source_sheet["E5"] = None
            source_book.save(self.source)
        finally:
            source_book.close()

        options = replace(
            self.options("accounting_checks.xlsx"),
            accounting=AccountingConfig(
                amount_columns=(4,),
                group_column=2,
                required_columns=(5,),
                duplicate_columns=(2, 6),
            ),
        )
        result = export_filtered(
            options,
            [FilterRule(column=2, operator="equals", values=("销售一部", "销售二部"))],
        )

        self.assertEqual(result.matched_rows, 3)
        output = load_workbook(options.output, read_only=True, data_only=True)
        try:
            summary = list(output["金额汇总"].iter_rows(values_only=True))
            self.assertEqual(summary[1][:4], ("总计", None, 3, 300))
            groups = {row[1]: (row[2], row[3]) for row in summary[2:]}
            self.assertEqual(groups, {"销售一部": (2, 100), "销售二部": (1, 200)})

            issue_rows = list(output["数据检查"].iter_rows(min_row=2, values_only=True))
            self.assertEqual(
                {(row[0], row[1]) for row in issue_rows},
                {(4, "金额无效"), (4, "必填缺失"), (4, "重复键")},
            )
            duplicate_description = next(row[2] for row in issue_rows if row[1] == "重复键")
            self.assertIn("第 2 行", duplicate_description)
        finally:
            output.close()

    def test_batch_splits_departments_and_excludes_summary_tail(self) -> None:
        before = sha256(self.source.read_bytes()).digest()
        result = export_batch(
            self.options("unused.xlsx"), [], (2,), self.directory / "by-department"
        )

        self.assert_safe_batch_files(result)
        self.assertEqual(result.scanned_rows, 4)
        self.assertEqual(result.matched_rows, 4)
        self.assertEqual(result.stopped_at_row, 6)
        self.assertEqual(result.file_count, 3)
        self.assertEqual(
            self.batch_detail_ids(result.files),
            {frozenset({"A001", "C003"}), frozenset({"B002"}), frozenset({"D004"})},
        )
        for path in result.files:
            rows = self.nonempty_rows(Path(path))
            self.assertEqual(rows[0], tuple(HEADERS))
            self.assertNotIn("AFTER", {row[5] for row in rows[1:]})
        self.assertEqual(sha256(self.source.read_bytes()).digest(), before)

    def test_batch_single_field_with_seven_values_creates_seven_files(self) -> None:
        source = self.directory / "seven-departments.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "明细"
        sheet.append(HEADERS)
        detail_rows = [
            ("甲公司", "销售一部", datetime(2026, 6, 1), 100, "订单一", "A001"),
            ("乙公司", "销售二部", datetime(2026, 6, 2), 200, "订单二", "B002"),
            ("丙公司", "销售三部", datetime(2026, 6, 3), 300, "订单三", "C003"),
            ("丁公司", "销售四部", datetime(2026, 6, 4), 400, "订单四", "D004"),
            ("戊公司", "财务部", datetime(2026, 6, 5), 500, "订单五", "E005"),
            ("己公司", "行政部", datetime(2026, 6, 6), 600, "订单六", "F006"),
            ("庚公司", "研发部", datetime(2026, 6, 7), 700, "订单七", "G007"),
            ("辛公司", "销售一部", datetime(2026, 6, 8), 800, "重复部门", "H008"),
        ]
        for row in detail_rows:
            sheet.append(row)
        sheet.append(("合计", None, None, "=SUM(D2:D9)", None, None))
        sheet.append(("尾部记录", "销售一部", datetime(2026, 6, 9), 999, "不应导出", "AFTER"))
        workbook.save(source)
        workbook.close()

        options = replace(self.options("unused.xlsx"), source=source)
        result = export_batch(options, [], (2,), self.directory / "seven-value-batch")

        self.assert_safe_batch_files(result)
        self.assertEqual(result.scanned_rows, 8)
        self.assertEqual(result.matched_rows, 8)
        self.assertEqual(result.stopped_at_row, 10)
        self.assertEqual(result.file_count, 7)
        expected_by_department = {
            department: [row for row in detail_rows if row[1] == department]
            for department in {row[1] for row in detail_rows}
        }
        actual_by_department = {}
        for path in result.files:
            rows = self.nonempty_rows(Path(path))
            self.assertEqual(rows[0], tuple(HEADERS))
            self.assertTrue(rows[1:])
            department = rows[1][1]
            self.assertNotIn(department, actual_by_department)
            self.assertTrue(all(row[1] == department for row in rows[1:]))
            actual_by_department[department] = rows[1:]
        self.assertEqual(actual_by_department, expected_by_department)

    def test_batch_splits_only_actual_department_date_combinations(self) -> None:
        workbook = load_workbook(self.source)
        try:
            # Two rows share one department/date combination; the other two
            # combinations occur once each. No cartesian-product files exist.
            workbook["明细"]["C4"] = datetime(2026, 6, 1)
            workbook.save(self.source)
        finally:
            workbook.close()

        result = export_batch(
            self.options("unused.xlsx"), [], (2, 3), self.directory / "by-combination"
        )

        self.assert_safe_batch_files(result)
        self.assertEqual(result.file_count, 3)
        self.assertEqual(
            self.batch_detail_ids(result.files),
            {frozenset({"A001", "C003"}), frozenset({"B002"}), frozenset({"D004"})},
        )
        for path in result.files:
            detail = self.nonempty_rows(Path(path))[1:]
            self.assertEqual(len({(row[1], row[2]) for row in detail}), 1)

    def test_batch_applies_common_filters_before_splitting(self) -> None:
        result = export_batch(
            self.options("unused.xlsx"),
            [FilterRule(column=5, operator="contains", values=("重点",))],
            (2,),
            self.directory / "filtered-batches",
        )

        self.assert_safe_batch_files(result)
        self.assertEqual(result.matched_rows, 2)
        self.assertEqual(result.file_count, 2)
        self.assertEqual(
            self.batch_detail_ids(result.files),
            {frozenset({"A001"}), frozenset({"B002"})},
        )

    def test_batch_keeps_zero_and_blank_split_values_distinct(self) -> None:
        workbook = load_workbook(self.source)
        try:
            workbook["明细"]["B2"] = 0
            workbook["明细"]["B3"] = None
            workbook.save(self.source)
        finally:
            workbook.close()

        result = export_batch(
            self.options("unused.xlsx"), [], (2,), self.directory / "zero-and-blank"
        )

        self.assert_safe_batch_files(result)
        self.assertEqual(result.file_count, 4)
        rows_by_id = {
            rows[1][5]: rows[1][1]
            for path in result.files
            if (rows := self.nonempty_rows(Path(path)))[1][5] in {"A001", "B002"}
        }
        self.assertEqual(rows_by_id, {"A001": 0, "B002": None})
        self.assertEqual(
            self.batch_detail_ids(result.files),
            {frozenset({"A001"}), frozenset({"B002"}), frozenset({"C003"}), frozenset({"D004"})},
        )

    def test_batch_uses_safe_unique_names_for_colliding_values(self) -> None:
        workbook = load_workbook(self.source)
        try:
            workbook["明细"]["B2"] = "财/务"
            workbook["明细"]["B3"] = "财:务"
            workbook.save(self.source)
        finally:
            workbook.close()

        result = export_batch(
            self.options("unused.xlsx"),
            [FilterRule(column=5, operator="contains", values=("重点",))],
            (2,),
            self.directory / "safe-names",
            prefix="财/务:*?",
        )

        self.assert_safe_batch_files(result)
        self.assertEqual(result.file_count, 2)
        self.assertEqual(
            self.batch_detail_ids(result.files),
            {frozenset({"A001"}), frozenset({"B002"})},
        )

    def test_batch_max_files_failure_leaves_no_partial_results(self) -> None:
        parent = self.directory / "too-many"
        parent.mkdir()
        sentinel = parent / "keep.txt"
        sentinel.write_text("pre-existing", encoding="utf-8")

        with self.assertRaises(ValueError):
            export_batch(self.options("unused.xlsx"), [], (2,), parent, max_files=2)

        self.assertEqual(list(parent.iterdir()), [sentinel])
        self.assertEqual(sentinel.read_text(encoding="utf-8"), "pre-existing")

    def test_batch_cancel_leaves_no_partial_results(self) -> None:
        parent = self.directory / "cancelled"
        parent.mkdir()
        sentinel = parent / "keep.txt"
        sentinel.write_text("pre-existing", encoding="utf-8")
        cancel_event = Event()
        cancel_event.set()

        with self.assertRaises(CancelledError):
            export_batch(
                self.options("unused.xlsx"), [], (2,), parent, cancel_event=cancel_event
            )

        self.assertEqual(list(parent.iterdir()), [sentinel])
        self.assertEqual(sentinel.read_text(encoding="utf-8"), "pre-existing")

    def test_batch_cancel_after_first_staged_file_removes_whole_batch(self) -> None:
        parent = self.directory / "cancel-during-save"
        parent.mkdir()
        cancel_event = Event()
        saved_progress = []

        def cancel_after_first_file(phase: str, processed: int, total: int | None, matched: int) -> None:
            if phase == "save" and processed > 0:
                saved_progress.append(processed)
                cancel_event.set()

        with self.assertRaises(CancelledError):
            export_batch(
                self.options("unused.xlsx"), [], (2,), parent,
                progress=cancel_after_first_file, cancel_event=cancel_event,
            )

        self.assertTrue(saved_progress, "The test must cancel after a file is staged")
        self.assertEqual(list(parent.iterdir()), [])

    def test_batch_zero_common_matches_creates_no_files(self) -> None:
        parent = self.directory / "no-matches"
        parent.mkdir()

        with self.assertRaises(ValueError):
            export_batch(
                self.options("unused.xlsx"),
                [FilterRule(2, "equals", ("研发部",))],
                (2,), parent,
            )

        self.assertEqual(list(parent.iterdir()), [])

    def test_batch_preserves_selected_source_styles(self) -> None:
        result = export_batch(
            self.options("unused.xlsx", preserve_styles=True),
            [], (2,), self.directory / "styled-batches",
        )
        sales_one_file = next(
            Path(path) for path in result.files
            if any(row[5] == "A001" for row in self.nonempty_rows(Path(path))[1:])
        )
        original = load_workbook(self.source, read_only=False)
        filtered = load_workbook(sales_one_file, read_only=False)
        try:
            self.assertEqual(filtered.active["C2"].number_format, original["明细"]["C2"].number_format)
            self.assertEqual(filtered.active["D2"].number_format, original["明细"]["D2"].number_format)
            self.assertEqual(filtered.active["D2"].fill.fgColor.rgb, original["明细"]["D2"].fill.fgColor.rgb)
        finally:
            original.close()
            filtered.close()

    def test_batch_accounting_is_independent_for_each_file(self) -> None:
        workbook = load_workbook(self.source)
        try:
            # The same ID in different output files must not be a duplicate.
            workbook["明细"]["F3"] = "A001"
            workbook.save(self.source)
        finally:
            workbook.close()
        options = replace(
            self.options("unused.xlsx"),
            accounting=AccountingConfig(
                amount_columns=(4,),
                group_column=1,
                required_columns=(5,),
                duplicate_columns=(6,),
            ),
        )

        result = export_batch(options, [], (2,), self.directory / "accounting-batches")

        self.assert_safe_batch_files(result)
        expected = {
            "销售一部": (2, 100, {(4, "必填缺失")}),
            "销售二部": (1, 200, set()),
            "财务部": (1, 300, set()),
        }
        self.assertEqual(result.file_count, len(expected))
        for path in result.files:
            output = load_workbook(path, read_only=True, data_only=True)
            try:
                detail = list(output["明细"].iter_rows(min_row=2, values_only=True))
                department = detail[0][1]
                count, amount, issues_expected = expected[department]
                self.assertEqual(len(detail), count)
                self.assertEqual(sum(row[3] for row in detail), amount)
                summary = list(output["金额汇总"].iter_rows(values_only=True))
                self.assertEqual(summary[1][2:4], (count, amount))
                issues = {
                    (row[0], row[1])
                    for row in output["数据检查"].iter_rows(min_row=2, values_only=True)
                    if isinstance(row[0], int)
                }
                self.assertEqual(issues, issues_expected)
            finally:
                output.close()

    def test_named_batch_exports_independent_rule_sets_and_header_only_job(self) -> None:
        jobs = (
            BatchJob("销售/一部", (FilterRule(2, "equals", ("销售一部",)),)),
            BatchJob("销售:二部", (
                FilterRule(2, "equals", ("销售二部",)),
                FilterRule(5, "contains", ("重点",)),
            )),
            BatchJob("研发*无匹配", (FilterRule(2, "equals", ("研发部",)),)),
        )
        result = export_named_batch(
            self.options("unused.xlsx"), jobs, self.directory / "named-batches"
        )

        self.assert_safe_batch_files(result)
        self.assertEqual(result.file_count, 3)
        self.assertEqual(
            self.batch_detail_ids(result.files),
            {frozenset({"A001", "C003"}), frozenset({"B002"}), frozenset()},
        )
        header_only = [path for path in result.files if len(self.nonempty_rows(Path(path))) == 1]
        self.assertEqual(len(header_only), 1)
        self.assertEqual(self.nonempty_rows(Path(header_only[0])), [tuple(HEADERS)])


if __name__ == "__main__":
    unittest.main()
