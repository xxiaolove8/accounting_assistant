"""Behavior tests for calculated fields, grouped metrics and Excel charts.

All workbooks in this module are synthetic and live in a temporary directory.
The user's financial sample is never opened or modified.
"""

from __future__ import annotations

from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event
import unittest

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart

from analysis_core import (
    AnalysisMetric,
    AnalysisOptions,
    CalculatedVariable,
    analyze_to_workbook,
)
from filter_core import CancelledError, FilterRule


class AnalysisCoreTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.source = self.directory / "financial_records.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "明细"
        sheet.append(["业务员", "公司", "应收金额", "期望回款", "费用"])
        sheet.append(["张三", "甲公司", 100, 30, 5])
        sheet.append(["张三", "甲公司", 900, 450, 50])
        sheet.append(["张三", "乙公司", 200, 100, 20])
        sheet.append(["张三", "丙公司", 0, 10, 1])
        sheet.append(["李四", "甲公司", 1000, 1000, 100])
        sheet.append(["合计", None, 2200, 1590, 176])
        sheet.append(["张三", "尾部公司", 100, 100, 1])
        book.save(self.source)
        book.close()

    def options(self, name: str, metric: AnalysisMetric, **changes) -> AnalysisOptions:
        chart_type = changes.pop("chart_type", "bar")
        return AnalysisOptions(
            source=self.source,
            sheet="明细",
            output=self.directory / name,
            group_column=2,
            metric=metric,
            chart_type=chart_type,
            **changes,
        )

    @staticmethod
    def ratio_metric(mode: str) -> AnalysisMetric:
        return AnalysisMetric(
            name="期望回款比例",
            mode=mode,
            numerator_ref="C4",
            denominator_ref="C3",
            as_percent=True,
        )

    @staticmethod
    def salesperson_rule() -> tuple[FilterRule, ...]:
        return (FilterRule(column=1, operator="equals", values=("张三",)),)

    @staticmethod
    def summary(path: Path) -> dict[str, tuple[object, ...]]:
        book = load_workbook(path, data_only=True)
        try:
            sheet = book["分析汇总"]
            assert tuple(sheet.cell(7, col).value for col in range(1, 8)) == (
                "分组", "指标值", "明细行数", "有效行数", "无效行数", "分子合计", "分母合计"
            )
            result: dict[str, tuple[object, ...]] = {}
            for row in sheet.iter_rows(min_row=8, min_col=1, max_col=7, values_only=True):
                if row[0] is not None:
                    result[str(row[0])] = tuple(row[1:])
            return result
        finally:
            book.close()

    def test_salesperson_filter_and_two_optional_ratio_rules(self) -> None:
        rules = self.salesperson_rule()
        weighted = self.options(
            "weighted.xlsx", self.ratio_metric("ratio_of_sums"), rules=rules
        )
        simple_average = self.options(
            "simple_average.xlsx", self.ratio_metric("average_of_ratios"), rules=rules
        )

        analyze_to_workbook(weighted)
        analyze_to_workbook(simple_average)
        weighted_rows = self.summary(weighted.output)
        average_rows = self.summary(simple_average.output)

        self.assertEqual(set(weighted_rows), {"甲公司", "乙公司", "丙公司"})
        self.assertNotIn("尾部公司", weighted_rows)
        self.assertAlmostEqual(weighted_rows["甲公司"][0], 480 / 1000)
        self.assertAlmostEqual(average_rows["甲公司"][0], (30 / 100 + 450 / 900) / 2)
        self.assertNotEqual(weighted_rows["甲公司"][0], average_rows["甲公司"][0])
        self.assertEqual(weighted_rows["甲公司"][1], 2)
        self.assertAlmostEqual(weighted_rows["乙公司"][0], 0.5)
        book = load_workbook(weighted.output, read_only=True, data_only=True)
        try:
            self.assertIn("张三", book["分析汇总"]["B6"].value)
        finally:
            book.close()
        book = load_workbook(weighted.output)
        try:
            sheet = book["分析汇总"]
            company_row = next(
                row for row in range(8, sheet.max_row + 1)
                if sheet.cell(row, 1).value == "甲公司"
            )
            self.assertIn("%", sheet.cell(company_row, 2).number_format)
        finally:
            book.close()

    def test_person_names_are_cached_as_chart_categories_and_axes_are_visible(self) -> None:
        for chart_type in ("bar", "line", "pie"):
            with self.subTest(chart_type=chart_type):
                options = AnalysisOptions(
                    source=self.source,
                    sheet="明细",
                    output=self.directory / f"person_{chart_type}.xlsx",
                    group_column=1,
                    metric=AnalysisMetric("期望回款", "sum", value_ref="C4"),
                    chart_type=chart_type,
                )
                analyze_to_workbook(options)
                book = load_workbook(options.output)
                try:
                    chart = book["分析汇总"]._charts[0]
                    category = chart.series[0].cat
                    self.assertIsNone(category.numRef)
                    self.assertIn("$I$", category.strRef.f)
                    names = {point.v for point in category.strRef.strCache.pt}
                    self.assertEqual(names, {"张三", "李四"})
                    self.assertEqual(chart.series[0].val.numRef.numCache.ptCount, 2)
                    if chart_type == "pie":
                        self.assertTrue(chart.dataLabels.showCatName)
                        self.assertTrue(chart.dataLabels.showPercent)
                        self.assertFalse(chart.dataLabels.showSerName)
                        self.assertFalse(chart.dataLabels.showVal)
                    else:
                        self.assertEqual(chart.x_axis.axPos, "b")
                        self.assertEqual(chart.y_axis.axPos, "l")
                        self.assertEqual(chart.x_axis.tickLblPos, "nextTo")
                        self.assertEqual(chart.y_axis.tickLblPos, "nextTo")
                        self.assertEqual(chart.layout.manualLayout.layoutTarget, "inner")
                        self.assertLess(chart.layout.manualLayout.h, 0.7)
                        x_title = "".join(run.t for para in chart.x_axis.title.tx.rich.p for run in para.r)
                        y_title = "".join(run.t for para in chart.y_axis.title.tx.rich.p for run in para.r)
                        self.assertEqual(x_title, "业务员")
                        self.assertEqual(y_title, "期望回款")
                finally:
                    book.close()

    def test_calculated_variables_support_arithmetic_parentheses_and_chaining(self) -> None:
        variables = (
            CalculatedVariable(name="双倍净回款", expression="(C4 - C5) * 2"),
            CalculatedVariable(name="调整后比例", expression="(V1 + 10) / C3"),
        )
        options = self.options(
            "calculated.xlsx",
            AnalysisMetric(name="调整后比例均值", mode="average", value_ref="V2"),
            variables=variables,
            rules=self.salesperson_rule(),
        )

        analyze_to_workbook(options)
        rows = self.summary(options.output)

        expected_甲 = (((30 - 5) * 2 + 10) / 100 + ((450 - 50) * 2 + 10) / 900) / 2
        self.assertAlmostEqual(rows["甲公司"][0], expected_甲)
        self.assertAlmostEqual(rows["乙公司"][0], ((100 - 20) * 2 + 10) / 200)
        book = load_workbook(options.output, read_only=True, data_only=True)
        try:
            definitions = book["变量定义"]
            self.assertEqual(definitions["A4"].value, "V1")
            self.assertEqual(definitions["B4"].value, "双倍净回款")
            self.assertEqual(definitions["C4"].value, "(C4 - C5) * 2")
            self.assertEqual(definitions["A5"].value, "V2")
            self.assertEqual(definitions["C5"].value, "(V1 + 10) / C3")
        finally:
            book.close()

    def test_accounting_text_numbers_are_parsed_without_turning_bad_text_into_zero(self) -> None:
        source = self.directory / "text_numbers.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "明细"
        sheet.append(["公司", "金额"])
        sheet.append(["甲公司", "1,000.50"])
        sheet.append(["甲公司", "(200.25)"])
        sheet.append(["甲公司", "不详"])
        sheet.append(["乙公司", "25%"])
        book.save(source)
        book.close()
        options = AnalysisOptions(
            source=source, sheet="明细", output=self.directory / "text_analysis.xlsx",
            group_column=1, metric=AnalysisMetric("金额合计", "sum", value_ref="C2"),
            chart_type="bar",
        )

        result = analyze_to_workbook(options)
        rows = self.summary(options.output)

        self.assertAlmostEqual(rows["甲公司"][0], 800.25)
        self.assertEqual(rows["甲公司"][3], 1)
        self.assertAlmostEqual(rows["乙公司"][0], 0.25)
        self.assertEqual(result.invalid_metric_rows, 1)

    def test_three_native_chart_types_reference_an_excel_data_range(self) -> None:
        cases = (("bar", BarChart), ("pie", PieChart), ("line", LineChart))
        for chart_type, expected_type in cases:
            with self.subTest(chart_type=chart_type):
                options = self.options(
                    f"{chart_type}.xlsx",
                    AnalysisMetric(name="期望回款", mode="sum", value_ref="C4"),
                    chart_type=chart_type,
                    rules=self.salesperson_rule(),
                )
                analyze_to_workbook(options)
                book = load_workbook(options.output)
                try:
                    summary = book["分析汇总"]
                    self.assertEqual(len(summary._charts), 1)
                    chart = summary._charts[0]
                    self.assertIsInstance(chart, expected_type)
                    self.assertGreaterEqual(len(chart.series), 1)
                    series = chart.series[0]
                    value_ref = series.val.numRef.f if series.val and series.val.numRef else ""
                    category_ref = ""
                    if series.cat:
                        if series.cat.numRef:
                            category_ref = series.cat.numRef.f
                        elif series.cat.strRef:
                            category_ref = series.cat.strRef.f
                    self.assertIn("分析汇总", value_ref)
                    self.assertIn("$J$", value_ref)
                    self.assertIn("分析汇总", category_ref)
                    self.assertIn("$I$", category_ref)
                    self.assertEqual(summary["I7"].value, "图表分组")
                    self.assertEqual(summary["J7"].value, "图表数值")
                finally:
                    book.close()

    def test_chart_limit_does_not_truncate_full_group_summary(self) -> None:
        options = self.options(
            "limited_chart.xlsx",
            AnalysisMetric(name="期望回款", mode="sum", value_ref="C4"),
            chart_type="bar",
            chart_limit=2,
            sort_by="value_desc",
            rules=self.salesperson_rule(),
        )

        analyze_to_workbook(options)
        rows = self.summary(options.output)
        book = load_workbook(options.output, data_only=True)
        try:
            sheet = book["分析汇总"]
            projection = [
                (sheet.cell(row, 9).value, sheet.cell(row, 10).value)
                for row in range(8, 11)
            ]
        finally:
            book.close()

        self.assertEqual(set(rows), {"甲公司", "乙公司", "丙公司"})
        self.assertEqual(projection[:2], [("甲公司", 480), ("乙公司", 100)])
        self.assertEqual(projection[2], (None, None))

    def test_pie_chart_rolls_small_groups_into_other_without_losing_value(self) -> None:
        options = self.options(
            "pie_other.xlsx",
            AnalysisMetric(name="期望回款", mode="sum", value_ref="C4"),
            chart_type="pie",
            chart_limit=2,
            rules=self.salesperson_rule(),
        )

        analyze_to_workbook(options)
        full_rows = self.summary(options.output)
        book = load_workbook(options.output, data_only=True)
        try:
            sheet = book["分析汇总"]
            chart_values = [sheet.cell(row, 10).value for row in (8, 9)]
            self.assertEqual(sheet["I9"].value, "其他（其余分组）")
            self.assertEqual(sheet["I10"].value, None)
        finally:
            book.close()

        self.assertEqual(sum(chart_values), sum(row[0] for row in full_rows.values()))

    def test_missing_and_zero_denominator_are_reported_without_infinite_chart_values(self) -> None:
        book = load_workbook(self.source)
        sheet = book["明细"]
        sheet.insert_rows(6)
        sheet["A6"], sheet["B6"], sheet["C6"], sheet["D6"] = (
            "张三", "乙公司", None, 50
        )
        book.save(self.source)
        book.close()
        options = self.options(
            "invalid_denominators.xlsx",
            self.ratio_metric("average_of_ratios"),
            rules=self.salesperson_rule(),
        )

        analyze_to_workbook(options)
        rows = self.summary(options.output)

        self.assertAlmostEqual(rows["乙公司"][0], 0.5)
        self.assertGreaterEqual(rows["乙公司"][3], 1)  # invalid-row count
        self.assertIsNone(rows["丙公司"][0])
        self.assertGreaterEqual(rows["丙公司"][3], 1)

    def test_source_is_unchanged_and_output_contains_optional_detail_sheet(self) -> None:
        before = sha256(self.source.read_bytes()).digest()
        options = self.options(
            "analysis.xlsx",
            AnalysisMetric(name="期望回款", mode="sum", value_ref="C4"),
            rules=self.salesperson_rule(),
            include_details=True,
        )

        analyze_to_workbook(options)

        self.assertEqual(sha256(self.source.read_bytes()).digest(), before)
        book = load_workbook(options.output, read_only=True, data_only=True)
        try:
            self.assertIn("分析汇总", book.sheetnames)
            self.assertIn("分析明细", book.sheetnames)
        finally:
            book.close()

    def test_cancel_and_invalid_expression_leave_no_partial_output(self) -> None:
        cancelled = Event()
        cancelled.set()
        cancel_options = self.options(
            "cancelled.xlsx",
            AnalysisMetric(name="金额", mode="sum", value_ref="C4"),
        )
        with self.assertRaises(CancelledError):
            analyze_to_workbook(cancel_options, cancel_event=cancelled)
        self.assertFalse(cancel_options.output.exists())

        invalid_options = self.options(
            "invalid.xlsx",
            AnalysisMetric(name="恶意公式", mode="sum", value_ref="V1"),
            variables=(CalculatedVariable(name="恶意公式", expression="__import__('os')"),),
        )
        with self.assertRaises(ValueError):
            analyze_to_workbook(invalid_options)
        self.assertFalse(invalid_options.output.exists())

    def test_header_and_detail_ranges_are_honored(self) -> None:
        source = self.directory / "multirow_header.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "数据"
        sheet.append(["2026 财务数据", None, None])
        sheet.append(["人员", "客户", "金额"])
        sheet.append(["业务员", "公司", "期望回款"])
        sheet.append(["说明", "本行非明细", 9999])
        sheet.append(["张三", "甲公司", 10])
        sheet.append(["张三", "乙公司", 20])
        sheet.append(["张三", "丙公司", 30])
        book.save(source)
        book.close()
        options = AnalysisOptions(
            source=source,
            sheet="数据",
            output=self.directory / "ranged.xlsx",
            group_column=2,
            metric=AnalysisMetric(name="期望回款", mode="sum", value_ref="C3"),
            chart_type="line",
            header_start=2,
            header_end=3,
            data_start=5,
            data_end=6,
        )

        analyze_to_workbook(options)

        rows = self.summary(options.output)
        self.assertEqual(set(rows), {"甲公司", "乙公司"})
        self.assertEqual(rows["甲公司"][0], 10)
        self.assertEqual(rows["乙公司"][0], 20)


if __name__ == "__main__":
    unittest.main()
