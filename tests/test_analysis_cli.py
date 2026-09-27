"""Command-line analysis compatibility and multi-metric input validation."""

import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook

from main import _run_analysis


class AnalysisCliTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.source = self.directory / "source.xlsx"
        self.output = self.directory / "result.xlsx"
        self.status = self.directory / "status.json"
        self.metrics_file = self.directory / "metrics.json"
        book = Workbook()
        book.active.title = "data"
        book.active.append(["Salesperson", "Planned", "Actual"])
        book.active.append(["A", 100, 180])
        book.active.append(["B", 200, 80])
        book.save(self.source)
        book.close()
        self.base_args = [
            "--analyze", "--input", str(self.source), "--sheet", "data",
            "--output", str(self.output), "--group-column", "1", "--chart-type", "bar",
            "--status-file", str(self.status),
        ]

    def test_metrics_file_exports_multiple_series_and_chosen_sort(self):
        definitions = [
            {"name": "Planned", "mode": "sum", "value_ref": "C2", "unit": "元"},
            {"name": "Actual", "mode": "sum", "value_ref": "C3", "unit": "元"},
        ]
        self.metrics_file.write_text(json.dumps(definitions, ensure_ascii=False), encoding="utf-8-sig")
        result = _run_analysis(self.base_args + [
            "--metrics-file", str(self.metrics_file), "--sort-metric", "2", "--sort-by", "value_asc",
        ])
        self.assertEqual(result, 0, self.status.read_text(encoding="utf-8"))
        status = json.loads(self.status.read_text(encoding="utf-8"))
        self.assertEqual(status["metric_count"], 2)
        self.assertEqual(status["chart_count"], 1)
        book = load_workbook(self.output)
        self.addCleanup(book.close)
        summary = book["分析汇总"]
        self.assertEqual([summary.cell(7, c).value for c in range(1, 5)],
                         ["分组", "Planned", "Actual", "明细行数"])
        self.assertEqual([summary.cell(8, c).value for c in range(1, 5)], ["B", 200, 80, 1])
        self.assertEqual(len(summary._charts[0].series), 2)

    def test_legacy_single_metric_arguments_still_work(self):
        result = _run_analysis(self.base_args + [
            "--metric-name", "Actual", "--metric-mode", "sum", "--value-ref", "C3",
        ])
        self.assertEqual(result, 0, self.status.read_text(encoding="utf-8"))
        status = json.loads(self.status.read_text(encoding="utf-8"))
        self.assertEqual(status["metric_count"], 1)
        book = load_workbook(self.output, read_only=True)
        self.addCleanup(book.close)
        self.assertEqual(book["分析汇总"]["B8"].value, 180)

    def test_conflicting_metric_sources_do_not_silently_ignore_settings(self):
        self.metrics_file.write_text('[{"name":"Actual","mode":"sum","value_ref":"C3"}]', encoding="utf-8")
        result = _run_analysis(self.base_args + [
            "--metrics-file", str(self.metrics_file), "--metric-name", "Wrong",
        ])
        self.assertEqual(result, 1)
        self.assertFalse(self.output.exists())
        self.assertIn("不要同时", json.loads(self.status.read_text(encoding="utf-8"))["message"])

    def test_invalid_metric_definition_is_reported_in_status(self):
        self.metrics_file.write_text(
            '[{"name":"Actual","mode":"sum","value_ref":"C3","as_percent":"false"}]',
            encoding="utf-8",
        )
        result = _run_analysis(self.base_args + ["--metrics-file", str(self.metrics_file)])
        self.assertEqual(result, 1)
        self.assertFalse(self.output.exists())
        self.assertIn("as_percent", json.loads(self.status.read_text(encoding="utf-8"))["message"])


if __name__ == "__main__":
    unittest.main()
