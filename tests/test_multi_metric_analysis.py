"""Synthetic behavior checks for grouped analysis with several selected metrics."""

from __future__ import annotations

from dataclasses import replace
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event
import unittest
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart

from analysis_core import (
    AnalysisMetric,
    AnalysisOptions,
    CalculatedVariable,
    analyze_to_workbook,
)
from filter_core import CancelledError, FilterRule


class MultiMetricAnalysisTests(unittest.TestCase):
    def setUp(self) -> None:
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.source = self.directory / "synthetic_receipts.xlsx"
        self.write_source([
            ["张三", "甲公司", 100, 80, 2],
            ["张三", "乙公司", 200, 100, 3],
            ["李四", "甲公司", 500, 400, 6],
            ["王五", "丙公司", 50, None, 1],
            ["赵六", "丁公司", 150, 200, 4],
            ["合计", None, 1000, 780, 16],
            ["尾部人员", "不应纳入", 999, 999, 999],
        ])

    def write_source(self, rows: list[list[object]]) -> None:
        book = Workbook()
        sheet = book.active
        sheet.title = "明细"
        sheet.append(["业务员", "公司", "计划回款", "实际回款", "数量"])
        for row in rows:
            sheet.append(row)
        book.save(self.source)
        book.close()

    @staticmethod
    def amounts() -> tuple[AnalysisMetric, ...]:
        return (
            AnalysisMetric("计划回款", "sum", value_ref="C3", unit="元"),
            AnalysisMetric("实际回款", "sum", value_ref="C4", unit="元"),
        )

    @staticmethod
    def completion(mode: str = "ratio_of_sums") -> AnalysisMetric:
        return AnalysisMetric(
            "回款完成率", mode, numerator_ref="C4", denominator_ref="C3",
            as_percent=True,
        )

    def options(self, name: str = "analysis.xlsx", **changes) -> AnalysisOptions:
        kwargs = dict(
            source=self.source, sheet="明细", output=self.directory / name,
            group_column=1, metrics=self.amounts(), chart_limit=None,
        )
        kwargs.update(changes)
        return AnalysisOptions(**kwargs)

    def open_result(self, options: AnalysisOptions):
        result = analyze_to_workbook(options)
        book = load_workbook(options.output)
        self.addCleanup(book.close)
        return result, book

    @staticmethod
    def table(book, names: list[str]) -> dict[str, tuple[object, ...]]:
        sheet = book["分析汇总"]
        width = len(names) + 2
        headers = [sheet.cell(7, col).value for col in range(1, width + 1)]
        if headers != ["分组", *names, "明细行数"]:
            raise AssertionError(f"Unexpected summary headers: {headers!r}")
        return {
            str(row[0]): tuple(row[1:])
            for row in sheet.iter_rows(min_row=8, max_col=width, values_only=True)
            if row[0] is not None
        }

    @staticmethod
    def categories(series) -> list[str]:
        cache = series.cat.strRef.strCache
        points = {point.idx: point.v for point in cache.pt}
        return [points[index] for index in range(cache.ptCount)]

    @staticmethod
    def values(series) -> list[float | None]:
        cache = series.val.numRef.numCache
        points = {point.idx: float(point.v) for point in cache.pt}
        return [points.get(index) for index in range(cache.ptCount)]

    @staticmethod
    def series_name(series) -> str:
        if series.tx is None or series.tx.strRef is None:
            raise AssertionError("Native series must retain a worksheet title reference")
        return series.tx.strRef.strCache.pt[0].v

    @staticmethod
    def title_text(title) -> str:
        return "".join(run.t for paragraph in title.tx.rich.p for run in paragraph.r)

    def test_amounts_and_weighted_completion_aggregate_each_group_independently(self) -> None:
        metrics = (*self.amounts(), self.completion())
        result, book = self.open_result(self.options(metrics=metrics))
        rows = self.table(book, [metric.name for metric in metrics])

        self.assertEqual(set(rows), {"张三", "李四", "王五", "赵六"})
        self.assertEqual(rows["张三"][:2], (300, 180))
        self.assertAlmostEqual(rows["张三"][2], 180 / 300)
        self.assertEqual(rows["张三"][3], 2)
        self.assertEqual(rows["李四"][:2], (500, 400))
        self.assertAlmostEqual(rows["李四"][2], 0.8)
        self.assertEqual(rows["王五"][:3], (50, None, None))
        self.assertEqual(result.matched_rows, 5)
        self.assertEqual(result.group_count, 4)
        self.assertEqual(result.metric_count, 3)
        self.assertEqual(result.chart_count, 2)
        self.assertEqual(len(book["分析汇总"]._charts), 2)

    def test_bar_series_have_names_distinct_colors_person_categories_and_axes(self) -> None:
        _, book = self.open_result(self.options())
        charts = book["分析汇总"]._charts
        self.assertEqual(len(charts), 1)
        chart = charts[0]
        self.assertIsInstance(chart, BarChart)
        self.assertEqual(chart.grouping, "clustered")
        self.assertIsNotNone(chart.legend)
        self.assertEqual([self.series_name(s) for s in chart.series], ["计划回款", "实际回款"])
        self.assertEqual(self.categories(chart.series[0]), ["李四", "张三", "赵六", "王五"])
        self.assertEqual(self.categories(chart.series[1]), self.categories(chart.series[0]))
        self.assertEqual(self.values(chart.series[0]), [500, 300, 150, 50])
        self.assertEqual(self.values(chart.series[1]), [400, 180, 200, None])
        self.assertNotEqual(chart.series[0].graphicalProperties.solidFill,
                            chart.series[1].graphicalProperties.solidFill)
        self.assertEqual(self.title_text(chart.x_axis.title), "业务员")
        self.assertIn("元", self.title_text(chart.y_axis.title))
        self.assertEqual(chart.x_axis.axPos, "b")
        self.assertEqual(chart.y_axis.axPos, "l")
        self.assertFalse(chart.x_axis.delete)
        self.assertFalse(chart.y_axis.delete)
        self.assertEqual(chart.x_axis.tickLblSkip, 1)

    def test_sort_metric_drives_shared_top_n_without_truncating_summary(self) -> None:
        metrics = (*self.amounts(), self.completion())
        _, book = self.open_result(self.options(
            metrics=metrics, sort_metric_index=1, sort_by="value_desc", chart_limit=2,
        ))
        rows = self.table(book, [metric.name for metric in metrics])
        self.assertEqual(len(rows), 4)
        for chart in book["分析汇总"]._charts:
            for series in chart.series:
                self.assertEqual(self.categories(series), ["李四", "赵六"])
        numeric = next(chart for chart in book["分析汇总"]._charts if len(chart.series) == 2)
        self.assertEqual(self.values(numeric.series[0]), [500, 150])
        self.assertEqual(self.values(numeric.series[1]), [400, 200])

    def test_missing_middle_value_keeps_category_position_and_line_gap(self) -> None:
        self.write_source([
            ["人员A", "甲", 30, 12, 1],
            ["人员B", "乙", 20, None, 1],
            ["人员C", "丙", 10, 8, 1],
        ])
        _, book = self.open_result(self.options(chart_type="line"))
        chart = book["分析汇总"]._charts[0]
        self.assertIsInstance(chart, LineChart)
        self.assertIsNotNone(chart.legend)
        self.assertEqual(chart.display_blanks, "gap")
        self.assertEqual([self.series_name(s) for s in chart.series], ["计划回款", "实际回款"])
        self.assertEqual(self.categories(chart.series[1]), ["人员A", "人员B", "人员C"])
        self.assertEqual(self.values(chart.series[1]), [12, None, 8])
        self.assertEqual(chart.series[1].val.numRef.numCache.ptCount, 3)
        self.assertEqual([point.idx for point in chart.series[1].val.numRef.numCache.pt], [0, 2])

    def test_percentage_and_custom_units_are_separate_charts_with_shared_categories(self) -> None:
        metrics = (
            *self.amounts(), self.completion(),
            AnalysisMetric("销售数量", "sum", value_ref="C5", unit="件"),
        )
        result, book = self.open_result(self.options(metrics=metrics))
        charts = book["分析汇总"]._charts
        self.assertEqual(result.chart_count, 3)
        by_names = {tuple(self.series_name(s) for s in c.series): c for c in charts}
        self.assertEqual(set(by_names), {("计划回款", "实际回款"), ("回款完成率",), ("销售数量",)})
        percent_chart = by_names[("回款完成率",)]
        self.assertIn("%", percent_chart.y_axis.numFmt.formatCode)
        self.assertIn("件", self.title_text(by_names[("销售数量",)].y_axis.title))
        categories = self.categories(charts[0].series[0])
        for chart in charts:
            for series in chart.series:
                self.assertEqual(self.categories(series), categories)

    def test_pies_use_each_metric_own_top_groups_and_keep_each_total(self) -> None:
        self.write_source([
            ["人员A", "甲", 500, 10, 1],
            ["人员B", "乙", 100, 300, 1],
            ["人员C", "丙", 200, 30, 1],
        ])
        options = self.options(chart_type="pie", chart_limit=2)
        result, book = self.open_result(options)
        charts = book["分析汇总"]._charts
        self.assertEqual(result.chart_count, 2)
        self.assertEqual(len(charts), 2)
        by_name = {self.series_name(c.series[0]): c for c in charts}
        self.assertEqual(self.categories(by_name["计划回款"].series[0]), ["人员A", "其他（其余分组）"])
        self.assertEqual(self.categories(by_name["实际回款"].series[0]), ["人员B", "其他（其余分组）"])
        self.assertEqual(sum(self.values(by_name["计划回款"].series[0])), 800)
        self.assertEqual(sum(self.values(by_name["实际回款"].series[0])), 340)
        for chart in charts:
            self.assertIsInstance(chart, PieChart)
            self.assertEqual(len(chart.series), 1)
            self.assertTrue(chart.dataLabels.showCatName)
            self.assertTrue(chart.dataLabels.showPercent)
        # Excel must not replace the percentage precision with raw amount formats.
        with ZipFile(options.output) as archive:
            ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
            for filename in ("xl/charts/chart1.xml", "xl/charts/chart2.xml"):
                xml = ElementTree.fromstring(archive.read(filename))
                label_format = xml.find(".//c:dLbls/c:numFmt", ns)
                self.assertEqual(label_format.attrib, {"formatCode": "0.0%", "sourceLinked": "0"})
        self.assertEqual(len(self.table(book, ["计划回款", "实际回款"])), 3)

    def test_pie_rejects_non_additive_metrics_and_names_the_affected_metric(self) -> None:
        cases = (
            self.completion(),
            replace(self.completion(), as_percent=False),
            self.completion("average_of_ratios"),
            AnalysisMetric("平均回款", "average", value_ref="C4"),
            AnalysisMetric("比例合计", "sum", value_ref="C4", as_percent=True),
        )
        for index, invalid in enumerate(cases):
            with self.subTest(metric=invalid):
                options = self.options(f"invalid_pie_{index}.xlsx", metrics=(*self.amounts(), invalid), chart_type="pie")
                with self.assertRaises(ValueError) as raised:
                    analyze_to_workbook(options)
                self.assertIn(invalid.name, str(raised.exception))
                self.assertFalse(options.output.exists())

    def test_pie_rejects_negative_or_all_zero_metric_without_partial_output(self) -> None:
        for value in (-1, 0):
            with self.subTest(value=value):
                self.write_source([["人员A", "甲", 10, value, 1]])
                options = self.options(f"nonpositive_{value}.xlsx", chart_type="pie")
                with self.assertRaises(ValueError) as raised:
                    analyze_to_workbook(options)
                self.assertIn("实际回款", str(raised.exception))
                self.assertFalse(options.output.exists())

    def test_filter_variables_and_ratio_choices_coexist_without_modifying_source(self) -> None:
        before = sha256(self.source.read_bytes()).digest()
        variables = (CalculatedVariable("回款差额", "C3 - C4"),)
        metrics = (
            *self.amounts(),
            AnalysisMetric("差额", "sum", value_ref="V1", unit="元"),
            self.completion(),
            replace(self.completion("average_of_ratios"), name="逐笔完成率均值"),
        )
        result, book = self.open_result(self.options(
            metrics=metrics, variables=variables,
            rules=(FilterRule(column=1, operator="equals", values=("张三",)),),
        ))
        rows = self.table(book, [metric.name for metric in metrics])
        self.assertEqual(set(rows), {"张三"})
        self.assertEqual(rows["张三"][:3], (300, 180, 120))
        self.assertAlmostEqual(rows["张三"][3], 0.6)
        self.assertAlmostEqual(rows["张三"][4], (0.8 + 0.5) / 2)
        self.assertEqual(result.matched_rows, 2)
        self.assertIn("分析明细", book.sheetnames)
        self.assertIn("变量定义", book.sheetnames)
        self.assertIn("张三", book["分析汇总"]["B6"].value)
        self.assertEqual(sha256(self.source.read_bytes()).digest(), before)

    def test_empty_duplicate_and_invalid_sort_selection_fail_before_output(self) -> None:
        cases = (
            {"metrics": ()},
            {"metrics": (self.amounts()[0], self.amounts()[0])},
            {"sort_metric_index": -1},
            {"sort_metric_index": 2},
        )
        for index, changes in enumerate(cases):
            with self.subTest(changes=changes):
                options = self.options(f"bad_selection_{index}.xlsx", **changes)
                with self.assertRaises(ValueError):
                    analyze_to_workbook(options)
                self.assertFalse(options.output.exists())

    def test_cancel_preserves_source_and_existing_output(self) -> None:
        options = self.options()
        previous = b"Existing result must survive cancellation"
        options.output.write_bytes(previous)
        before = sha256(self.source.read_bytes()).digest()
        event = Event()
        event.set()
        with self.assertRaises(CancelledError):
            analyze_to_workbook(options, cancel_event=event)
        self.assertEqual(options.output.read_bytes(), previous)
        self.assertEqual(sha256(self.source.read_bytes()).digest(), before)


if __name__ == "__main__":
    unittest.main()
