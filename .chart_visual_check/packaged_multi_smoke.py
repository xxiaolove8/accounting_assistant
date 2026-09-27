"""Verify the built Windows executable with synthetic multi-metric inputs."""
import json
from pathlib import Path
import subprocess
from zipfile import ZipFile

from openpyxl import load_workbook

directory = Path(__file__).resolve().parent
executable = directory.parent / "dist" / "ExcelFilterAssistant.exe"
definitions = json.loads((directory / "multi_metrics.json").read_text(encoding="utf-8"))
for kind, selected in (("bar", definitions), ("pie", definitions[:2])):
    metric_file = directory / f"packaged_{kind}_metrics.json"
    metric_file.write_text(json.dumps(selected, ensure_ascii=False), encoding="utf-8")
    output = directory / f"packaged_multi_{kind}.xlsx"
    status_file = directory / f"packaged_multi_{kind}_status.json"
    process = subprocess.run([
        str(executable), "--analyze", "--input", str(directory / "multi_source.xlsx"),
        "--sheet", "明细", "--output", str(output), "--group-column", "1",
        "--metrics-file", str(metric_file), "--chart-type", kind,
        "--status-file", str(status_file),
    ], timeout=60, creationflags=subprocess.CREATE_NO_WINDOW)
    status = json.loads(status_file.read_text(encoding="utf-8"))
    assert process.returncode == 0 and status["status"] == "ok", status
    assert status["metric_count"] == len(selected), status
    assert status["chart_count"] == (3 if kind == "bar" else 2), status
    book = load_workbook(output)
    try:
        sheet = book["分析汇总"]
        assert sheet["A8"].value == "李四"
        assert sheet["B8"].value == 200000 and sheet["C8"].value == 180000
        assert len(sheet._charts[0].series) == (2 if kind == "bar" else 1)
        assert sheet._charts[0].series[0].tx.strRef.strCache.pt[0].v == "计划回款"
    finally:
        book.close()
    if kind == "pie":
        with ZipFile(output) as archive:
            assert b'sourceLinked="0"' in archive.read("xl/charts/chart1.xml")
    print(f"PASS packaged {kind}: {status['metric_count']} metrics, {status['chart_count']} charts")
