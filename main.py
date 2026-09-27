"""Entry point for the Windows GUI and optional unattended exports."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys

from accounting import AccountingConfig
from analysis_core import AnalysisMetric, AnalysisOptions, CalculatedVariable, analyze_to_workbook
from merge_core import MergeOptions, merge_workbooks
from filter_core import (
    BatchJob,
    ExportOptions,
    FilterRule,
    export_batch,
    export_filtered,
    export_named_batch,
)


def _batch_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Excel 筛选助手批处理模式")
    parser.add_argument("--batch", action="store_true", required=True)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--output", type=Path, help="单文件导出的输出路径")
    parser.add_argument("--output-dir", type=Path, help="批量导出的输出父文件夹")
    parser.add_argument("--filter", action="append", default=[], metavar="列号:equals|contains:值")
    parser.add_argument("--all", action="store_true", help="不设置条件，导出全部明细")
    parser.add_argument("--split-column", type=int, action="append", default=[], help="按此列的每个值拆分")
    parser.add_argument("--jobs-file", type=Path, help="包含多套命名筛选条件的 JSON 文件")
    parser.add_argument("--prefix", help="自动拆分文件名前缀")
    parser.add_argument("--max-files", type=int, default=200)
    parser.add_argument("--header-start", type=int, default=1)
    parser.add_argument("--header-end", type=int, default=1)
    parser.add_argument("--data-start", type=int)
    parser.add_argument("--data-end", type=int)
    parser.add_argument("--no-stop-summary", action="store_true")
    parser.add_argument("--preserve-styles", action="store_true")
    parser.add_argument("--amount-column", type=int, action="append", default=[])
    parser.add_argument("--group-column", type=int)
    parser.add_argument("--required-column", type=int, action="append", default=[])
    parser.add_argument("--duplicate-column", type=int, action="append", default=[])
    parser.add_argument("--status-file", type=Path, help="将结果或错误写入 JSON 文件")
    return parser.parse_args(argv)


def _rules(expressions: list[str], export_all: bool) -> list[FilterRule]:
    if not expressions and not export_all:
        raise ValueError("请至少设置一个 --filter，或使用 --all。")
    grouped: dict[tuple[int, str], list[str]] = {}
    for expression in expressions:
        fields = expression.split(":", 2)
        if len(fields) != 3:
            raise ValueError(f"筛选表达式无效：{expression}")
        number, operator, value = fields
        try:
            column = int(number)
        except ValueError as exc:
            raise ValueError(f"筛选列号无效：{number}") from exc
        if operator not in {"equals", "contains"}:
            raise ValueError(f"筛选方式无效：{operator}")
        grouped.setdefault((column, operator), []).append(value)
    return [FilterRule(column, operator, tuple(values)) for (column, operator), values in grouped.items()]


def _write_status(path: Path | None, payload: dict) -> None:
    if path is not None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _analysis_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Excel 筛选助手数据分析与图表模式")
    parser.add_argument("--analyze", action="store_true", required=True)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--filter", action="append", default=[], metavar="列号:equals|contains:值")
    parser.add_argument("--group-column", type=int, required=True)
    parser.add_argument("--variable", action="append", default=[], metavar="名称=表达式")
    parser.add_argument("--metrics-file", type=Path, help="多个指标定义的 JSON 数组文件")
    parser.add_argument("--metric-name", help="单指标名称；与 --metrics-file 二选一")
    parser.add_argument(
        "--metric-mode",
        choices=("sum", "average", "count", "ratio_of_sums", "average_of_ratios"),
    )
    parser.add_argument("--value-ref")
    parser.add_argument("--numerator-ref")
    parser.add_argument("--denominator-ref")
    parser.add_argument("--metric-as-percent", action="store_true")
    parser.add_argument("--metric-unit", default="", help="指标单位，例如 元、万元、笔")
    parser.add_argument("--chart-type", choices=("bar", "pie", "line"), required=True)
    parser.add_argument("--chart-limit", type=int, default=20, help="图表展示的前 N 类；0 表示全部")
    parser.add_argument("--sort-by", choices=("value_desc", "value_asc", "group"), default="value_desc")
    parser.add_argument("--sort-metric", type=int, default=1, help="按数值排序所用的指标序号，从 1 开始")
    parser.add_argument("--no-details", action="store_true")
    parser.add_argument("--header-start", type=int, default=1)
    parser.add_argument("--header-end", type=int, default=1)
    parser.add_argument("--data-start", type=int)
    parser.add_argument("--data-end", type=int)
    parser.add_argument("--no-stop-summary", action="store_true")
    parser.add_argument("--status-file", type=Path)
    return parser.parse_args(argv)


def _load_analysis_metrics(path: Path) -> tuple[AnalysisMetric, ...]:
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(payload, list) or not payload:
        raise ValueError("指标文件必须是至少包含一个指标的 JSON 数组。")
    metrics = []
    allowed = {"name", "mode", "value_ref", "numerator_ref", "denominator_ref", "as_percent", "unit"}
    for index, item in enumerate(payload, 1):
        if not isinstance(item, dict):
            raise ValueError(f"第 {index} 个指标必须是 JSON 对象。")
        unexpected = set(item) - allowed
        if unexpected:
            raise ValueError(f"第 {index} 个指标包含未知设置：{'、'.join(sorted(unexpected))}")
        for field in ("name", "mode"):
            if not isinstance(item.get(field), str) or not item[field].strip():
                raise ValueError(f"第 {index} 个指标缺少有效的 {field}。")
        for field in ("value_ref", "numerator_ref", "denominator_ref"):
            if item.get(field) is not None and not isinstance(item[field], str):
                raise ValueError(f"第 {index} 个指标的 {field} 必须是字段引用文本。")
        if "as_percent" in item and not isinstance(item["as_percent"], bool):
            raise ValueError(f"第 {index} 个指标的 as_percent 必须为 true 或 false。")
        if "unit" in item and not isinstance(item["unit"], str):
            raise ValueError(f"第 {index} 个指标的 unit 必须是文本。")
        definition = dict(item)
        definition.setdefault("as_percent", item["mode"] in {"ratio_of_sums", "average_of_ratios"})
        metrics.append(AnalysisMetric(**definition))
    return tuple(metrics)


def _run_analysis(argv: list[str]) -> int:
    args = _analysis_args(argv)
    try:
        if args.metrics_file is not None:
            if any((args.metric_name, args.metric_mode, args.value_ref, args.numerator_ref,
                    args.denominator_ref, args.metric_as_percent, args.metric_unit)):
                raise ValueError("使用 --metrics-file 时，请将指标设置全部放入该文件，不要同时使用单指标参数。")
            metrics = _load_analysis_metrics(args.metrics_file)
            metric = None
        else:
            if not args.metric_name or not args.metric_mode:
                raise ValueError("请提供 --metrics-file，或同时提供 --metric-name 和 --metric-mode。")
            metrics = ()
            metric = AnalysisMetric(
                name=args.metric_name,
                mode=args.metric_mode,
                value_ref=args.value_ref,
                numerator_ref=args.numerator_ref,
                denominator_ref=args.denominator_ref,
                as_percent=args.metric_as_percent or args.metric_mode in {"ratio_of_sums", "average_of_ratios"},
                unit=args.metric_unit,
            )
        variables: list[CalculatedVariable] = []
        for item in args.variable:
            name, separator, expression = item.partition("=")
            if not separator or not name.strip() or not expression.strip():
                raise ValueError(f"计算变量应写为“名称=表达式”：{item}")
            variables.append(CalculatedVariable(name.strip(), expression.strip()))
        result = analyze_to_workbook(AnalysisOptions(
            source=args.input,
            sheet=args.sheet,
            output=args.output,
            group_column=args.group_column,
            metric=metric,
            metrics=metrics,
            chart_type=args.chart_type,
            variables=tuple(variables),
            rules=tuple(_rules(args.filter, True)),
            header_start=args.header_start,
            header_end=args.header_end,
            data_start=args.data_start,
            data_end=args.data_end,
            stop_at_summary=not args.no_stop_summary,
            chart_limit=None if args.chart_limit == 0 else args.chart_limit,
            sort_by=args.sort_by,
            sort_metric_index=args.sort_metric - 1,
            include_details=not args.no_details,
        ))
        _write_status(args.status_file, {
            "status": "ok",
            "scanned_rows": result.scanned_rows,
            "matched_rows": result.matched_rows,
            "group_count": result.group_count,
            "invalid_metric_rows": result.invalid_metric_rows,
            "zero_denominator_groups": result.zero_denominator_groups,
            "chart_points": result.chart_points,
            "metric_count": result.metric_count,
            "chart_count": result.chart_count,
            "output": str(result.output),
        })
        return 0
    except Exception as exc:
        _write_status(args.status_file, {"status": "error", "message": str(exc)})
        return 1


def _merge_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Excel 筛选助手多文件合并模式")
    parser.add_argument("--merge", action="store_true", required=True)
    parser.add_argument("--input", type=Path, action="append", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--sheet", help="每个文件中要合并的同名工作表；留空取首个工作表")
    parser.add_argument("--header-start", type=int, default=1)
    parser.add_argument("--header-end", type=int, default=1)
    parser.add_argument("--data-start", type=int)
    parser.add_argument("--data-end", type=int)
    parser.add_argument("--no-stop-summary", action="store_true")
    parser.add_argument("--source-column", action="store_true")
    parser.add_argument("--deduplicate", action="store_true")
    parser.add_argument("--match-mode", choices=("union", "strict"), default="union")
    parser.add_argument("--status-file", type=Path)
    return parser.parse_args(argv)


def _run_merge(argv: list[str]) -> int:
    args = _merge_args(argv)
    try:
        result = merge_workbooks(MergeOptions(
            sources=tuple(args.input),
            output=args.output,
            sheet_name=args.sheet,
            header_start=args.header_start,
            header_end=args.header_end,
            data_start=args.data_start,
            data_end=args.data_end,
            stop_at_summary=not args.no_stop_summary,
            add_source_column=args.source_column,
            deduplicate=args.deduplicate,
            match_mode=args.match_mode,
        ))
        _write_status(args.status_file, {
            "status": "ok",
            "files": result.files,
            "scanned_rows": result.scanned_rows,
            "merged_rows": result.merged_rows,
            "duplicates_skipped": result.duplicates_skipped,
            "sheets": result.sheets,
            "headers": list(result.headers),
            "output": str(result.output),
        })
        return 0
    except Exception as exc:
        _write_status(args.status_file, {"status": "error", "message": str(exc)})
        return 1


def _load_jobs(path: Path) -> list[BatchJob]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, list):
        raise ValueError("任务文件必须是 JSON 列表。")
    jobs = []
    for item in payload:
        if not isinstance(item, dict) or not isinstance(item.get("rules"), list):
            raise ValueError("每个任务必须包含 name 和 rules 列表。")
        rules = []
        for rule in item["rules"]:
            rules.append(FilterRule(
                column=int(rule["column"]),
                operator=str(rule["operator"]),
                values=tuple(str(value) for value in rule["values"]),
            ))
        jobs.append(BatchJob(name=str(item["name"]), rules=tuple(rules)))
    return jobs


def _run_batch(argv: list[str]) -> int:
    args = _batch_args(argv)
    try:
        batch_mode = bool(args.split_column) or args.jobs_file is not None
        if batch_mode and args.output_dir is None:
            raise ValueError("批量导出需要 --output-dir。")
        if not batch_mode and args.output is None:
            raise ValueError("单文件导出需要 --output。")
        if args.split_column and args.jobs_file is not None:
            raise ValueError("--split-column 和 --jobs-file 不能同时使用。")
        if args.jobs_file is not None and args.filter:
            raise ValueError("命名任务的筛选条件请写在 --jobs-file 中。")
        rules = _rules(args.filter, args.all or batch_mode)
        options = ExportOptions(
            source=args.input,
            sheet=args.sheet,
            output=args.output or args.output_dir / "_unused.xlsx",
            header_start=args.header_start,
            header_end=args.header_end,
            data_start=args.data_start,
            data_end=args.data_end,
            stop_at_summary=not args.no_stop_summary,
            preserve_styles=args.preserve_styles,
            accounting=AccountingConfig(
                amount_columns=tuple(args.amount_column),
                group_column=args.group_column,
                required_columns=tuple(args.required_column),
                duplicate_columns=tuple(args.duplicate_column),
            ),
        )
        if args.split_column:
            result = export_batch(
                options, rules, args.split_column, args.output_dir,
                prefix=args.prefix, max_files=args.max_files,
            )
        elif args.jobs_file is not None:
            result = export_named_batch(
                options, _load_jobs(args.jobs_file), args.output_dir,
                max_files=args.max_files,
            )
        else:
            result = export_filtered(options, rules)
        status = {
            "status": "ok",
            "scanned_rows": result.scanned_rows,
            "matched_rows": result.matched_rows,
            "stopped_at_row": result.stopped_at_row,
        }
        if batch_mode:
            status.update({
                "file_count": result.file_count,
                "output_dir": str(result.output_dir),
                "files": [str(path) for path in result.files],
            })
        else:
            status["output"] = str(result.output)
        _write_status(args.status_file, status)
        return 0
    except Exception as exc:
        _write_status(args.status_file, {"status": "error", "message": str(exc)})
        return 1


def main() -> int:
    if "--analyze" in sys.argv[1:]:
        return _run_analysis(sys.argv[1:])
    if "--merge" in sys.argv[1:]:
        return _run_merge(sys.argv[1:])
    if "--batch" in sys.argv[1:]:
        return _run_batch(sys.argv[1:])
    from excel_filter_gui import run_app

    run_app()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
