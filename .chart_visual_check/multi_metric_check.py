"""Synthetic chart validation inputs for the multi-metric program upgrade."""
from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook
from analysis_core import AnalysisMetric, AnalysisOptions, analyze_to_workbook

directory = Path(__file__).resolve().parent
source = directory / "multi_source.xlsx"
book = Workbook()
sheet = book.active
sheet.title = "明细"
sheet.append(["业务员", "公司", "月份", "计划回款", "实际回款", "费用"])
for row in [
    ["张三", "甲公司", "2026-08", 100000, 80000, 3000],
    ["张三", "乙公司", "2026-09", 50000, 45000, 1500],
    ["李四", "甲公司", "2026-09", 200000, 180000, 5000],
    ["王五", "丙公司", "2026-09", 120000, None, 2000],
    ["赵六", "丁公司", "2026-09", 80000, 96000, 2200],
    ["孙七", "戊公司", "2026-09", 0, 1000, 200],
]:
    sheet.append(row)
book.save(source)
book.close()

definitions = [
    {"name": "计划回款", "mode": "sum", "value_ref": "C4", "unit": "元"},
    {"name": "实际回款", "mode": "sum", "value_ref": "C5", "unit": "元"},
    {"name": "完成率", "mode": "ratio_of_sums", "numerator_ref": "C5", "denominator_ref": "C4", "as_percent": True},
    {"name": "记录数", "mode": "count"},
]
(directory / "multi_metrics.json").write_text(json.dumps(definitions, ensure_ascii=False), encoding="utf-8")
metrics = tuple(AnalysisMetric(**item) for item in definitions)
for kind in ("bar", "line", "pie"):
    options = AnalysisOptions(
        source=source, sheet="明细", output=directory / f"multi_{kind}.xlsx",
        group_column=1, metrics=metrics[:2] if kind == "pie" else metrics,
        chart_type=kind, chart_limit=20,
    )
    result = analyze_to_workbook(options)
    print(kind, "groups=", result.group_count, "metrics=", result.metric_count, "charts=", result.chart_count)
    checked = load_workbook(options.output)
    try:
        assert len(checked["分析汇总"]._charts) == (2 if kind == "pie" else 3)
        if kind != "pie":
            assert len(checked["分析汇总"]._charts[0].series) == 2
        for chart in checked["分析汇总"]._charts:
            for series in chart.series:
                assert series.cat.strRef.strCache.ptCount > 0
    finally:
        checked.close()
