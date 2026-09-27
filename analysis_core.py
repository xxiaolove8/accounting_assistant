"""Grouped financial analysis with calculated fields and native Excel charts.

The source workbook is only read.  Results are saved to a temporary .xlsx and
atomically moved into place after a complete save, so cancellation leaves the
previous output untouched.
"""

from __future__ import annotations

import ast
from dataclasses import dataclass
from decimal import Decimal, DecimalException, InvalidOperation
import math
import os
from pathlib import Path
import re
import tempfile
from typing import Callable, Sequence

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.data_source import AxDataSource, NumData, NumVal, StrData, StrRef, StrVal
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from filter_core import (
    CancelledError,
    FilterRule,
    _check_layout,
    _check_source,
    _close_batch_workbook,
    _display_width,
    _is_blank_row,
    _is_summary_row,
    _matches,
    _open_sheet,
    _progress,
    _total_rows,
    _validate_rules,
    inspect_headers,
    normalize_value,
)


ProgressCallback = Callable[[str, int, int | None, int], None]
_REFERENCE = re.compile(r"([CV])([1-9][0-9]*)\Z", re.IGNORECASE)
_TEXT_NUMBER = re.compile(r"[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\Z|[+-]?\.\d+\Z")
_PERCENT_FORMAT = "0.00%"
_NUMBER_FORMAT = '#,##0.00;[Red](#,##0.00)'
_MAX_EXCEL_ROW = 1_048_576
_TABLE_HEADER_ROW = 7
_TABLE_FIRST_ROW = _TABLE_HEADER_ROW + 1


@dataclass(frozen=True)
class CalculatedVariable:
    """A named expression referring to C1..Cn and preceding V1..Vn."""

    name: str
    expression: str
    as_percent: bool = False


@dataclass(frozen=True)
class AnalysisMetric:
    name: str
    mode: str  # sum, average, count, ratio_of_sums, average_of_ratios
    value_ref: str | None = None
    numerator_ref: str | None = None
    denominator_ref: str | None = None
    as_percent: bool = False
    unit: str = ""


@dataclass(frozen=True)
class AnalysisOptions:
    source: Path
    sheet: str
    output: Path
    group_column: int
    metric: AnalysisMetric | None = None
    chart_type: str = "bar"  # bar, pie, line
    variables: tuple[CalculatedVariable, ...] = ()
    rules: tuple[FilterRule, ...] = ()
    header_start: int = 1
    header_end: int = 1
    data_start: int | None = None
    data_end: int | None = None
    stop_at_summary: bool = True
    chart_limit: int | None = 20
    sort_by: str = "value_desc"  # value_desc, value_asc, group
    include_details: bool = True
    metrics: tuple[AnalysisMetric, ...] = ()
    sort_metric_index: int = 0


@dataclass(frozen=True)
class AnalysisResult:
    scanned_rows: int
    matched_rows: int
    group_count: int
    invalid_metric_rows: int
    zero_denominator_groups: int
    chart_points: int
    output: Path
    metric_count: int = 1
    chart_count: int = 1


@dataclass
class _Group:
    key: str
    rows: int = 0
    valid: int = 0
    invalid: int = 0
    value_sum: Decimal = Decimal(0)
    numerator_sum: Decimal = Decimal(0)
    denominator_sum: Decimal = Decimal(0)


def _check_cancel(cancel_event) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise CancelledError()


def _number(value: object) -> Decimal | None:
    """Parse a deliberately small set of accounting number representations.

    Native Excel numbers, comma-grouped text, trailing percentages, and
    parenthesized negatives are recognized.  Currency symbols and arbitrary
    words are not silently coerced to zero.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            result = Decimal(str(value))
        except InvalidOperation:
            return None
        return result if result.is_finite() else None
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1].strip()
    percentage = text.endswith("%") or text.endswith("％")
    if percentage:
        text = text[:-1].strip()
    if not _TEXT_NUMBER.fullmatch(text):
        return None
    try:
        result = Decimal(text.replace(",", ""))
        if negative:
            result = -result
        if percentage:
            result /= 100
    except (InvalidOperation, DecimalException):
        return None
    return result if result.is_finite() else None


def _excel_number(value: Decimal | None) -> float | int | None:
    if value is None:
        return None
    try:
        converted = float(value)
    except (OverflowError, ValueError):
        raise ValueError("计算结果超过 Excel 可表示的数值范围。") from None
    if not math.isfinite(converted):
        raise ValueError("计算结果超过 Excel 可表示的数值范围。")
    return converted


def _parse_ref(reference: str | None, width: int, preceding_variables: int) -> tuple[str, int]:
    if not isinstance(reference, str):
        raise ValueError("数值字段必须选择源表列 C1、C2…或已定义变量 V1、V2…。")
    matched = _REFERENCE.fullmatch(reference.strip())
    if not matched:
        raise ValueError(f"无效的字段引用 {reference!r}；请使用 C1、C2…或 V1、V2…。")
    kind, raw_index = matched.groups()
    index = int(raw_index)
    if (kind.upper() == "C" and index > width) or (kind.upper() == "V" and index > preceding_variables):
        raise ValueError(f"字段引用 {reference!r} 超出可用列或已定义变量范围。")
    return kind.upper(), index - 1


def _compile_expression(expression: str, width: int, preceding_variables: int) -> ast.Expression:
    if not isinstance(expression, str) or not expression.strip():
        raise ValueError("新变量的计算表达式不能为空。")
    if len(expression) > 500:
        raise ValueError("新变量表达式过长，请简化到 500 个字符以内。")
    try:
        tree = ast.parse(expression.strip(), mode="eval")
    except (SyntaxError, ValueError) as exc:
        raise ValueError(f"新变量表达式语法错误：{exc}") from exc

    nodes = list(ast.walk(tree))
    if len(nodes) > 120:
        raise ValueError("新变量表达式过于复杂，请简化计算。")
    for node in nodes:
        if isinstance(node, ast.Expression):
            continue
        if isinstance(node, ast.BinOp):
            if type(node.op) not in {ast.Add, ast.Sub, ast.Mult, ast.Div}:
                raise ValueError("表达式仅支持 +、-、*、/ 和括号。")
            continue
        if isinstance(node, ast.UnaryOp):
            if type(node.op) not in {ast.UAdd, ast.USub}:
                raise ValueError("表达式仅支持 +、-、*、/ 和括号。")
            continue
        if isinstance(node, ast.Name):
            _parse_ref(node.id, width, preceding_variables)
            continue
        if isinstance(node, ast.Constant):
            if isinstance(node.value, bool) or not isinstance(node.value, (int, float)):
                raise ValueError("表达式中的常数必须是数字。")
            if _number(node.value) is None:
                raise ValueError("表达式中的常数必须是有限数字。")
            continue
        if isinstance(node, (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.UAdd, ast.USub, ast.Load)):
            continue
        raise ValueError("表达式仅支持列或变量引用、数字、+、-、*、/ 和括号。")
    return tree


def _eval_expression(
    node: ast.AST,
    source_values: Sequence[object],
    calculated_values: Sequence[Decimal | None],
    width: int,
) -> Decimal | None:
    if isinstance(node, ast.Expression):
        return _eval_expression(node.body, source_values, calculated_values, width)
    if isinstance(node, ast.Name):
        kind, index = _parse_ref(node.id, width, len(calculated_values))
        return _number(source_values[index]) if kind == "C" else calculated_values[index]
    if isinstance(node, ast.Constant):
        return _number(node.value)
    if isinstance(node, ast.UnaryOp):
        value = _eval_expression(node.operand, source_values, calculated_values, width)
        if value is None:
            return None
        return -value if isinstance(node.op, ast.USub) else value
    if isinstance(node, ast.BinOp):
        left = _eval_expression(node.left, source_values, calculated_values, width)
        right = _eval_expression(node.right, source_values, calculated_values, width)
        if left is None or right is None:
            return None
        try:
            if isinstance(node.op, ast.Add):
                result = left + right
            elif isinstance(node.op, ast.Sub):
                result = left - right
            elif isinstance(node.op, ast.Mult):
                result = left * right
            else:
                if right == 0:
                    return None
                result = left / right
        except (DecimalException, ZeroDivisionError):
            return None
        return result if result.is_finite() else None
    raise AssertionError("表达式已经验证，不应包含其他语法。")


def _resolve_value(reference: tuple[str, int], row_values: Sequence[object], variables: Sequence[Decimal | None]) -> Decimal | None:
    kind, index = reference
    return _number(row_values[index]) if kind == "C" else variables[index]


def _metric_description(metric: AnalysisMetric) -> str:
    if metric.mode == "count":
        return "按分组统计明细行数"
    if metric.mode == "sum":
        return f"按分组求和：{metric.value_ref}"
    if metric.mode == "average":
        return f"按分组计算有效数值平均值：{metric.value_ref}"
    if metric.mode == "ratio_of_sums":
        return f"先按组汇总，再相除：SUM({metric.numerator_ref}) / SUM({metric.denominator_ref})"
    return f"先逐行相除，再按组平均：AVG({metric.numerator_ref} / {metric.denominator_ref})"


def _group_value(group: _Group, mode: str) -> Decimal | None:
    if mode == "count":
        return Decimal(group.rows)
    if group.valid == 0:
        return None
    if mode == "sum":
        return group.value_sum
    if mode in {"average", "average_of_ratios"}:
        return group.value_sum / group.valid
    if group.denominator_sum == 0:
        return None
    return group.numerator_sum / group.denominator_sum


def _text_cell(sheet, value: object, *, header: bool = False, title: bool = False):
    cell = WriteOnlyCell(sheet, value=value)
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"
    if header or title:
        cell.font = Font(name="Microsoft YaHei", size=14 if title else 10, bold=True,
                         color="24476B" if title else "FFFFFF")
        if header:
            cell.fill = PatternFill(fill_type="solid", fgColor="24476B")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    return cell


def _numeric_cell(sheet, value: Decimal | None, *, percent: bool = False):
    number = _excel_number(value)
    cell = WriteOnlyCell(sheet, value=number)
    cell.number_format = _PERCENT_FORMAT if percent else _NUMBER_FORMAT
    return cell


def _wrapped_text_cell(sheet, value: str):
    cell = _text_cell(sheet, value)
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    return cell


def _detail_cell(sheet, cell):
    value = cell.value
    if isinstance(value, str) and value.startswith("="):
        out = WriteOnlyCell(sheet, value=value)
        out.data_type = "s"
        return out
    number_format = getattr(cell, "number_format", "General") or "General"
    if number_format != "General":
        out = WriteOnlyCell(sheet, value=value)
        out.number_format = number_format
        return out
    return value


def _effective_unit(metric: AnalysisMetric) -> str:
    if metric.mode == "count":
        return metric.unit.strip() or "条"
    if metric.as_percent:
        return "%"
    return metric.unit.strip() or "数值"


def _accumulate_metric(
    group: _Group,
    metric: AnalysisMetric,
    references: tuple[tuple[str, int] | None, tuple[str, int] | None, tuple[str, int] | None],
    row_values: Sequence[object],
    calculated: Sequence[Decimal | None],
) -> bool:
    """Accumulate one independent metric and return whether this row is valid."""
    value_ref, numerator_ref, denominator_ref = references
    group.rows += 1
    if metric.mode == "count":
        group.valid += 1
        return True
    if metric.mode in {"sum", "average"}:
        value = _resolve_value(value_ref, row_values, calculated)
        if value is None:
            group.invalid += 1
            return False
        group.valid += 1
        group.value_sum += value
        return True
    numerator = _resolve_value(numerator_ref, row_values, calculated)
    denominator = _resolve_value(denominator_ref, row_values, calculated)
    if numerator is None or denominator is None or (metric.mode == "average_of_ratios" and denominator == 0):
        group.invalid += 1
        return False
    group.valid += 1
    group.numerator_sum += numerator
    group.denominator_sum += denominator
    if metric.mode == "average_of_ratios":
        group.value_sum += numerator / denominator
    return True


def _chart_rows(
    rows: list[tuple[_Group, Decimal | None]],
    chart_type: str,
    chart_limit: int | None,
) -> tuple[list[tuple[str, Decimal]], str]:
    valid = [(group.key if group.key else "（空白）", value) for group, value in rows if value is not None]
    if chart_type == "pie":
        if any(value < 0 for _, value in valid):
            raise ValueError("饼图只适用于非负数值；当前分组中存在负数，请改用柱状图或折线图。")
        positive = sorted(((name, value) for name, value in valid if value > 0),
                          key=lambda item: (-item[1], item[0].casefold()))
        if not positive:
            raise ValueError("饼图需要至少一个大于零的分组数值。")
        if chart_limit is not None and len(positive) > chart_limit:
            if chart_limit < 2:
                raise ValueError("饼图分组超过显示上限时，上限至少需要设为 2，以保留“其他”汇总。")
            retained = positive[: chart_limit - 1]
            omitted = positive[chart_limit - 1 :]
            retained.append(("其他（其余分组）", sum((value for _, value in omitted), Decimal(0))))
            return retained, f"饼图展示前 {chart_limit - 1} 个分组，其余 {len(omitted)} 个合并为“其他”；完整数据见左侧汇总表。"
        return positive, f"饼图展示全部 {len(positive)} 个正值分组；零值分组仍保留在汇总表。"

    if not valid:
        raise ValueError("没有可作图的有效指标值；请检查选中的数值字段和除数。")
    selected = valid[:chart_limit] if chart_limit is not None else valid
    if len(valid) > len(selected):
        note = f"图表仅显示排序后的前 {len(selected)} 个有效分组；完整 {len(rows)} 个分组见左侧汇总表。"
    else:
        note = f"图表展示全部 {len(selected)} 个有效分组；完整 {len(rows)} 个分组见左侧汇总表。"
    return selected, note


def _make_chart(sheet, chart_type: str, metric: AnalysisMetric, group_label: str,
                chart_rows: Sequence[tuple[str, Decimal]]) -> None:
    point_count = len(chart_rows)
    data = Reference(sheet, min_col=10, min_row=_TABLE_HEADER_ROW, max_row=_TABLE_HEADER_ROW + point_count)
    categories = Reference(sheet, min_col=9, min_row=_TABLE_FIRST_ROW, max_row=_TABLE_HEADER_ROW + point_count)
    if chart_type == "bar":
        chart = BarChart()
        chart.type = "col"
        chart.grouping = "clustered"
        chart.y_axis.title = metric.name
        chart.y_axis.numFmt = "0.0%" if metric.as_percent else _NUMBER_FORMAT
        chart.x_axis.title = group_label
        if all(value >= 0 for _, value in chart_rows):
            chart.y_axis.scaling.min = 0
    elif chart_type == "line":
        chart = LineChart()
        chart.style = 13
        chart.y_axis.title = metric.name
        chart.y_axis.numFmt = "0.0%" if metric.as_percent else _NUMBER_FORMAT
        chart.x_axis.title = group_label
        chart.marker = True
    else:
        chart = PieChart()
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showCatName = True
        chart.dataLabels.showPercent = True
        chart.dataLabels.showSerName = False
        chart.dataLabels.showVal = False
        chart.dataLabels.showLeaderLines = True
    chart.title = metric.name
    chart.height = 11 if chart_type != "pie" else 10
    chart.width = min(38, max(19, 8 + point_count * 1.35)) if chart_type != "pie" else 19
    if chart_type != "pie":
        chart.layout = Layout(manualLayout=ManualLayout(
            layoutTarget="inner", x=0.16, y=0.08, w=0.79, h=0.55,
        ))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    # openpyxl's set_categories() stores a numeric reference even when the
    # category cells contain text.  Several spreadsheet viewers then omit
    # names from the chart.  Explicit text and numeric caches keep the labels
    # visible immediately, while the references remain linked to sheet cells.
    chart.series[0].cat = AxDataSource(strRef=StrRef(
        f=str(categories),
        strCache=StrData(
            ptCount=point_count,
            pt=[StrVal(idx=index, v=name) for index, (name, _) in enumerate(chart_rows)],
        ),
    ))
    chart.series[0].val.numRef.numCache = NumData(
        formatCode="0.0%" if metric.as_percent else _NUMBER_FORMAT,
        ptCount=point_count,
        pt=[NumVal(idx=index, v=str(_excel_number(value)))
            for index, (_, value) in enumerate(chart_rows)],
    )
    chart.series[0].tx.strRef.strCache = StrData(
        ptCount=1, pt=[StrVal(idx=0, v="图表数值")],
    )
    if chart_type != "pie":
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.x_axis.tickLblPos = "nextTo"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.tickLblSkip = 1
    chart.legend = None if chart_type != "pie" else chart.legend
    sheet.add_chart(chart, "L7")


_SERIES_COLORS = ("4472C4", "ED7D31", "70AD47", "A64D79", "5B9BD5", "FFC000", "8064A2", "008C95")


class _PercentDataLabels(DataLabelList):
    """Keep pie percentage precision independent of the amount cells' format."""

    __attrs__ = DataLabelList.__attrs__
    __elements__ = DataLabelList.__elements__
    __nested__ = DataLabelList.__nested__

    def to_tree(self, tagname=None, idx=None, namespace=None):
        node = super().to_tree(tagname, idx, namespace)
        # openpyxl's label numFmt descriptor exposes formatCode only. Excel
        # otherwise links the labels to source-cell formats and rounds to 0%.
        for child in node:
            if child.tag.rsplit("}", 1)[-1] == "numFmt":
                child.set("sourceLinked", "0")
        return node


def _make_multi_chart(
    summary, data_sheet, chart_type: str, metrics: Sequence[AnalysisMetric],
    metric_indices: Sequence[int], group_label: str,
    chart_rows: Sequence[tuple[str, Sequence[Decimal | None]]], anchor: str,
    *, data_columns: Sequence[int],
) -> None:
    """Create native chart series whose cells, cached labels and gaps agree."""
    point_count = len(chart_rows)
    unit = _effective_unit(metrics[metric_indices[0]])
    if chart_type == "bar":
        chart = BarChart()
        chart.type = "col"
        chart.grouping = "clustered"
        chart.overlap = 0
        chart.varyColors = False
        if all(value is None or value >= 0 for _, values in chart_rows for value in values):
            chart.y_axis.scaling.min = 0
    elif chart_type == "line":
        chart = LineChart()
        chart.marker = True
        chart.varyColors = False
    else:
        chart = PieChart()
        chart.varyColors = True
        chart.dataLabels = _PercentDataLabels()
        chart.dataLabels.numFmt = "0.0%"
        chart.dataLabels.showLegendKey = False
        chart.dataLabels.showCatName = True
        chart.dataLabels.showPercent = True
        chart.dataLabels.showSerName = False
        chart.dataLabels.showVal = False
        chart.dataLabels.showLeaderLines = True
    metric_title = "、".join(metrics[index].name.strip() for index in metric_indices)
    chart.title = metric_title if len(metric_title) <= 40 else f"{group_label}指标比较（{unit}）"
    chart.height = 11 if chart_type != "pie" else 10
    chart.width = min(38, max(19, 8 + point_count * 1.35)) if chart_type != "pie" else 19
    chart.display_blanks = "gap"
    if chart_type != "pie":
        chart.x_axis.title = group_label
        chart.y_axis.title = "百分比（%）" if unit == "%" else unit
        chart.y_axis.numFmt = "0.0%" if unit == "%" else ("0" if unit == "条" else _NUMBER_FORMAT)
        chart.x_axis.axPos = "b"
        chart.y_axis.axPos = "l"
        chart.x_axis.delete = chart.y_axis.delete = False
        chart.x_axis.tickLblPos = chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.tickLblSkip = 1
        chart.layout = Layout(manualLayout=ManualLayout(
            layoutTarget="inner", xMode="edge", yMode="edge",
            x=0.18, y=0.14, w=0.59, h=0.65,
        ))
        chart.legend.position = "r"
        chart.legend.overlay = False
    else:
        # Leave room above the pie for labels on very small slices.
        chart.layout = Layout(manualLayout=ManualLayout(
            layoutTarget="inner", xMode="edge", yMode="edge",
            x=0.12, y=0.20, w=0.60, h=0.65,
        ))
    categories = Reference(data_sheet, min_col=1, min_row=2, max_row=point_count + 1)
    for local_index, (metric_index, data_column) in enumerate(zip(metric_indices, data_columns)):
        metric = metrics[metric_index]
        data = Reference(data_sheet, min_col=data_column, min_row=1, max_row=point_count + 1)
        chart.add_data(data, titles_from_data=True)
        series = chart.series[-1]
        series.cat = AxDataSource(strRef=StrRef(
            f=str(categories), strCache=StrData(
                ptCount=point_count,
                pt=[StrVal(idx=index, v=name) for index, (name, _) in enumerate(chart_rows)],
            ),
        ))
        series.val.numRef.numCache = NumData(
            formatCode="0.0%" if _effective_unit(metric) == "%" else _NUMBER_FORMAT,
            ptCount=point_count,
            pt=[NumVal(idx=index, v=str(_excel_number(values[local_index])))
                for index, (_, values) in enumerate(chart_rows) if values[local_index] is not None],
        )
        series.tx.strRef.strCache = StrData(
            ptCount=1, pt=[StrVal(idx=0, v=metric.name.strip())],
        )
        if chart_type != "pie":
            color = _SERIES_COLORS[metric_index % len(_SERIES_COLORS)]
            series.graphicalProperties.solidFill = color
            series.graphicalProperties.line.solidFill = color
            if chart_type == "line":
                series.marker.symbol = "circle"
                series.marker.size = 6
                series.marker.graphicalProperties.solidFill = color
                series.marker.graphicalProperties.line.solidFill = color
    summary.add_chart(chart, anchor)


def _write_multi_results(
    result_book, summary, options: AnalysisOptions, metrics: Sequence[AnalysisMetric],
    metric_groups: Sequence[dict[str, _Group]],
    ordered_keys: Sequence[str], values_by_key: dict[str, tuple[Decimal | None, ...]],
    group_label: str, labels: Sequence[str], cancel_event=None,
) -> tuple[int, int]:
    """Write a wide comparison table and independent diagnostics for each metric."""
    metric_count = len(metrics)
    headers = ("分组", *(metric.name.strip() for metric in metrics), "明细行数")
    summary.append([_text_cell(summary, label, header=True) for label in headers])
    for key in ordered_keys:
        _check_cancel(cancel_event)
        summary.append([
            _text_cell(summary, key or "（空白）"),
            *(_numeric_cell(summary, value, percent=_effective_unit(metric) == "%")
              for metric, value in zip(metrics, values_by_key[key])),
            metric_groups[0][key].rows,
        ])
    summary.auto_filter.ref = f"A7:{get_column_letter(metric_count + 2)}{_TABLE_HEADER_ROW + len(ordered_keys)}"

    definitions = result_book.create_sheet("指标定义")
    definitions.freeze_panes = "A2"
    for column, width in {"A": 26, "B": 20, "C": 16, "D": 65, "E": 28}.items():
        definitions.column_dimensions[column].width = width
    definitions.append([_text_cell(definitions, name, header=True)
                        for name in ("指标", "统计口径", "单位", "计算规则", "显示格式")])
    mode_names = {"sum": "求和", "average": "平均值", "count": "明细行数",
                  "ratio_of_sums": "先汇总后相除", "average_of_ratios": "逐行相除后平均"}
    for metric in metrics:
        definitions.append([
            _text_cell(definitions, metric.name.strip()), _text_cell(definitions, mode_names[metric.mode]),
            _text_cell(definitions, _effective_unit(metric)), _text_cell(definitions, _metric_description(metric)),
            _text_cell(definitions, "百分比（原值不变）" if _effective_unit(metric) == "%" else "数值"),
        ])
    definitions.append([])
    definitions.append([_text_cell(definitions, "空值口径"),
                        _text_cell(definitions, "各指标独立判断有效行；缺失或无效值不计为零，图表对应位置留空。比例汇总仅使用分子、分母均有效的成对明细。")])
    definitions.append([_text_cell(definitions, "源字段引用", header=True), _text_cell(definitions, "源字段名", header=True)])
    for index, label in enumerate(labels, 1):
        definitions.append([_text_cell(definitions, f"C{index}"), _text_cell(definitions, label)])

    diagnostics = None
    diagnostic_rows = 0
    diagnostic_number = 0
    for key in ordered_keys:
        _check_cancel(cancel_event)
        for index, metric in enumerate(metrics):
            if diagnostics is None or diagnostic_rows >= _MAX_EXCEL_ROW - 1:
                diagnostic_number += 1
                name = "指标校验" if diagnostic_number == 1 else f"指标校验{diagnostic_number}"
                diagnostics = result_book.create_sheet(name)
                diagnostics.freeze_panes = "C2"
                for column in "ABCDEFGHI":
                    diagnostics.column_dimensions[column].width = 22
                diagnostics.append([_text_cell(diagnostics, label, header=True) for label in (
                    "分组", "指标", "指标值", "明细行数", "有效行数", "无效行数", "分子合计", "分母合计", "结果说明",
                )])
                diagnostic_rows = 0
            group = metric_groups[index][key]
            value = values_by_key[key][index]
            ratio = metric.mode in {"ratio_of_sums", "average_of_ratios"}
            zero = metric.mode == "ratio_of_sums" and group.valid > 0 and group.denominator_sum == 0
            note = "汇总分母为零，结果留空" if zero else ("无有效数值，结果留空" if value is None else "")
            diagnostics.append([
                _text_cell(diagnostics, key or "（空白）"), _text_cell(diagnostics, metric.name.strip()),
                _numeric_cell(diagnostics, value, percent=_effective_unit(metric) == "%"),
                group.rows, group.valid, group.invalid,
                _numeric_cell(diagnostics, group.numerator_sum if ratio else None),
                _numeric_cell(diagnostics, group.denominator_sum if ratio else None),
                _text_cell(diagnostics, note),
            ])
            diagnostic_rows += 1

    anchor_column = get_column_letter(max(12, metric_count + 4))
    if options.chart_type == "pie":
        largest_point_count = 0
        for index, metric in enumerate(metrics):
            _check_cancel(cancel_event)
            rows = [(metric_groups[index][key], values_by_key[key][index]) for key in ordered_keys]
            try:
                selected, _ = _chart_rows(rows, "pie", options.chart_limit)
            except ValueError as exc:
                raise ValueError(f"指标“{metric.name.strip()}”：{exc}") from exc
            data_sheet = result_book.create_sheet("图表数据" if index == 0 else f"图表数据{index + 1}")
            data_sheet.column_dimensions["A"].width = 30
            data_sheet.column_dimensions["B"].width = 22
            data_sheet.append([_text_cell(data_sheet, group_label, header=True),
                               _text_cell(data_sheet, metric.name.strip(), header=True)])
            for name, value in selected:
                data_sheet.append([_text_cell(data_sheet, name), _numeric_cell(data_sheet, value)])
            _make_multi_chart(summary, data_sheet, "pie", metrics, [index], group_label,
                              [(name, (value,)) for name, value in selected],
                              f"{anchor_column}{7 + index * 27}", data_columns=[2])
            largest_point_count = max(largest_point_count, len(selected))
        return largest_point_count, metric_count

    if not any(value is not None for values in values_by_key.values() for value in values):
        raise ValueError("没有可作图的有效指标值；请检查选中的数值字段和除数。")
    selected_keys = ordered_keys[:options.chart_limit] if options.chart_limit is not None else ordered_keys
    data_sheet = result_book.create_sheet("图表数据")
    data_sheet.freeze_panes = "B2"
    for column in range(1, metric_count + 2):
        data_sheet.column_dimensions[get_column_letter(column)].width = 24
    data_sheet.append([_text_cell(data_sheet, label, header=True) for label in headers[:-1]])
    for key in selected_keys:
        _check_cancel(cancel_event)
        data_sheet.append([
            _text_cell(data_sheet, key or "（空白）"),
            *(_numeric_cell(data_sheet, value, percent=_effective_unit(metric) == "%")
              for metric, value in zip(metrics, values_by_key[key])),
        ])
    units: dict[str, list[int]] = {}
    for index, metric in enumerate(metrics):
        units.setdefault(_effective_unit(metric), []).append(index)
    for chart_index, indices in enumerate(units.values()):
        _check_cancel(cancel_event)
        selected = [(key or "（空白）", tuple(values_by_key[key][index] for index in indices)) for key in selected_keys]
        _make_multi_chart(summary, data_sheet, options.chart_type, metrics, indices, group_label, selected,
                          f"{anchor_column}{7 + chart_index * 27}", data_columns=[index + 2 for index in indices])
    return len(selected_keys), len(units)


def analyze_to_workbook(
    options: AnalysisOptions,
    progress: ProgressCallback | None = None,
    cancel_event=None,
) -> AnalysisResult:
    """Analyze filtered detail rows and create a charted .xlsx workbook.

    Cn refers to a 1-based source column; Vn refers to the nth calculated
    variable. Variables can only reference variables defined before them.
    ``as_percent`` changes display formatting only; it never multiplies the
    value by 100. Missing or invalid operands remain blank, not zero.
    """
    first_data = _check_layout(options.header_start, options.header_end, options.data_start, options.data_end)
    source = _check_source(options.source)
    output = Path(options.output).expanduser().resolve()
    if output.suffix.lower() != ".xlsx":
        raise ValueError("分析结果必须保存为 .xlsx 文件。")
    if source.resolve() == output:
        raise ValueError("分析结果路径不能与源文件相同。")
    if not output.parent.is_dir():
        raise FileNotFoundError(f"输出文件夹不存在：{output.parent}")
    if options.chart_type not in {"bar", "pie", "line"}:
        raise ValueError("图表类型仅支持 bar、pie、line。")
    if options.sort_by not in {"value_desc", "value_asc", "group"}:
        raise ValueError("分组排序仅支持 value_desc、value_asc、group。")
    if options.chart_limit is not None and (
        not isinstance(options.chart_limit, int) or isinstance(options.chart_limit, bool)
        or options.chart_limit < 1 or options.chart_limit > _MAX_EXCEL_ROW - _TABLE_HEADER_ROW
    ):
        raise ValueError("图表显示分组上限必须是有效的正整数。")

    labels = inspect_headers(source, options.sheet, options.header_start, options.header_end)
    width = len(labels)
    if not isinstance(options.group_column, int) or isinstance(options.group_column, bool) or not 1 <= options.group_column <= width:
        raise ValueError("分组字段超出源表表头范围。")
    prepared_rules = _validate_rules(options.rules, width)

    compiled_variables: list[ast.Expression] = []
    variable_names: set[str] = set()
    for index, variable in enumerate(options.variables):
        name = variable.name.strip()
        if not name:
            raise ValueError(f"第 {index + 1} 个新变量缺少名称。")
        if name.casefold() in variable_names:
            raise ValueError(f"新变量名称重复：{name}")
        variable_names.add(name.casefold())
        try:
            compiled_variables.append(_compile_expression(variable.expression, width, index))
        except ValueError as exc:
            raise ValueError(f"新变量“{name}”：{exc}") from exc

    multi_metric = bool(options.metrics)
    if options.metric is not None and multi_metric:
        raise ValueError("metric 与 metrics 不能同时指定，请只使用一种指标设置。")
    metrics = tuple(options.metrics) if multi_metric else ((options.metric,) if options.metric is not None else ())
    if not metrics:
        raise ValueError("请至少选择一个指标。")
    if len(metrics) > 255:
        raise ValueError("每次最多选择 255 个指标，以满足 Excel 原生图表的数据系列上限。")
    if (not isinstance(options.sort_metric_index, int) or isinstance(options.sort_metric_index, bool)
            or not 0 <= options.sort_metric_index < len(metrics)):
        raise ValueError("排序指标超出已选择的指标范围。")
    metric_names: set[str] = set()
    metric_references = []
    for metric in metrics:
        if not isinstance(metric, AnalysisMetric) or not isinstance(metric.name, str) or not metric.name.strip():
            raise ValueError("指标名称不能为空。")
        name = metric.name.strip()
        if name.casefold() in metric_names:
            raise ValueError(f"指标名称重复：{name}")
        metric_names.add(name.casefold())
        if metric.mode not in {"sum", "average", "count", "ratio_of_sums", "average_of_ratios"}:
            raise ValueError(f"指标“{name}”：不支持的统计口径。")
        if not isinstance(metric.unit, str):
            raise ValueError(f"指标“{name}”：单位必须是文字。")
        if metric.mode == "count" and metric.as_percent:
            raise ValueError(f"指标“{name}”：明细行数不能按百分比显示。")
        if not metric.as_percent and metric.unit.strip() in {"%", "％", "百分比"}:
            raise ValueError(f"指标“{name}”：百分比单位请同时启用百分比显示。")
        if options.chart_type == "pie" and (metric.mode not in {"sum", "count"} or metric.as_percent):
            raise ValueError(f"指标“{name}”：饼图仅支持非负金额求和或行数统计；百分比、比例和平均值请用柱状图或折线图。")
        try:
            value_ref = _parse_ref(metric.value_ref, width, len(options.variables)) if metric.mode in {"sum", "average"} else None
            numerator_ref = _parse_ref(metric.numerator_ref, width, len(options.variables)) if metric.mode in {"ratio_of_sums", "average_of_ratios"} else None
            denominator_ref = _parse_ref(metric.denominator_ref, width, len(options.variables)) if metric.mode in {"ratio_of_sums", "average_of_ratios"} else None
        except ValueError as exc:
            raise ValueError(f"指标“{name}”：{exc}") from exc
        metric_references.append((value_ref, numerator_ref, denominator_ref))
    metric = metrics[0]

    full_group_label = labels[options.group_column - 1]
    group_label = full_group_label.partition(" · ")[2] or full_group_label
    metric_groups: list[dict[str, _Group]] = [{} for _ in metrics]
    groups = metric_groups[0]
    scanned = matched = invalid = 0
    temp_path: Path | None = None
    source_book = None
    result_book = Workbook(write_only=True)
    summary = result_book.create_sheet("分析汇总")
    definitions = result_book.create_sheet("变量定义") if options.variables else None
    detail = result_book.create_sheet("分析明细") if options.include_details else None
    detail_number = 1
    detail_rows = 0
    summary.freeze_panes = "B8"
    summary.column_dimensions["A"].width = min(max(_display_width(group_label) + 2, 18), 40)
    for column in "BCDEFG":
        summary.column_dimensions[column].width = 17
    summary.column_dimensions["I"].width = 30
    summary.column_dimensions["J"].width = 19
    if multi_metric:
        headers = ("分组", *(selected_metric.name.strip() for selected_metric in metrics), "明细行数")
        for column, label in enumerate(headers, 1):
            summary.column_dimensions[get_column_letter(column)].width = min(max(_display_width(label) + 2, 17), 38)
        summary.row_dimensions[7].height = 32
        for row_number in (3, 4, 5, 6):
            summary.merged_cells.add(f"B{row_number}:{get_column_letter(max(10, len(headers)))}{row_number}")
            summary.row_dimensions[row_number].height = 32 if row_number != 5 else 42
    if definitions is not None:
        definitions.freeze_panes = "A4"
        for column, width_in_chars in {"A": 18, "B": 28, "C": 48, "D": 25}.items():
            definitions.column_dimensions[column].width = width_in_chars
        definitions.append([_text_cell(definitions, "变量定义与源字段对照", title=True)])
        definitions.append([_text_cell(definitions, "C 表示原始列；V 表示本表前面已定义的变量。空值、无法识别的数字或除数为零时，计算结果保持空白。")])
        definitions.append([_text_cell(definitions, label, header=True)
                            for label in ("引用", "名称", "表达式", "显示格式")])
        for index, variable in enumerate(options.variables, 1):
            definitions.append([
                _text_cell(definitions, f"V{index}"),
                _text_cell(definitions, variable.name),
                _text_cell(definitions, variable.expression),
                _text_cell(definitions, "百分比显示（值不变）" if variable.as_percent else "数值"),
            ])
        definitions.append([])
        definitions.append([_text_cell(definitions, "源字段引用", header=True),
                            _text_cell(definitions, "源字段名", header=True)])
        for index, label in enumerate(labels, 1):
            definitions.append([_text_cell(definitions, f"C{index}"), _text_cell(definitions, label)])
    if detail is not None:
        detail.freeze_panes = "A2"
        for index, label in enumerate((*labels, *(v.name for v in options.variables)), 1):
            detail.column_dimensions[get_column_letter(index)].width = min(max(_display_width(label) + 2, 12), 38)
        detail.append([_text_cell(detail, label, header=True) for label in (*labels, *(v.name for v in options.variables))])

    try:
        _check_cancel(cancel_event)
        source_book, source_sheet = _open_sheet(source, options.sheet)
        total = _total_rows(source_sheet, first_data, options.data_end)
        _progress(progress, "analyze", 0, total, 0)
        for row_number, cells in enumerate(
            source_sheet.iter_rows(min_row=first_data, max_row=options.data_end, max_col=width),
            start=first_data,
        ):
            _check_cancel(cancel_event)
            if options.stop_at_summary and _is_summary_row(cells):
                break
            if _is_blank_row(cells):
                continue
            scanned += 1
            if not _matches(cells, prepared_rules):
                if scanned % 2000 == 0:
                    _progress(progress, "analyze", row_number - first_data + 1, total, matched)
                continue
            matched += 1
            row_values = tuple(cell.value for cell in cells)
            calculated: list[Decimal | None] = []
            for tree in compiled_variables:
                calculated.append(_eval_expression(tree, row_values, calculated, width))

            key = normalize_value(row_values[options.group_column - 1])
            if key not in groups:
                if len(groups) >= _MAX_EXCEL_ROW - _TABLE_HEADER_ROW:
                    raise ValueError("分组数量超过单个 Excel 工作表可容纳的行数。")
                for selected_groups in metric_groups:
                    selected_groups[key] = _Group(key)
            row_invalid = False
            for selected_metric, references, selected_groups in zip(metrics, metric_references, metric_groups):
                if not _accumulate_metric(selected_groups[key], selected_metric, references, row_values, calculated):
                    row_invalid = True
            invalid += int(row_invalid)

            if detail is not None:
                if detail_rows >= _MAX_EXCEL_ROW - 1:
                    detail_number += 1
                    detail = result_book.create_sheet(f"分析明细{detail_number}")
                    detail.freeze_panes = "A2"
                    detail.append([_text_cell(detail, label, header=True) for label in (*labels, *(v.name for v in options.variables))])
                    detail_rows = 0
                detail.append([
                    *(_detail_cell(detail, cell) for cell in cells),
                    *(_numeric_cell(detail, value, percent=variable.as_percent)
                      for variable, value in zip(options.variables, calculated)),
                ])
                detail_rows += 1
            if scanned % 2000 == 0:
                _progress(progress, "analyze", row_number - first_data + 1, total, matched)

        _progress(progress, "analyze", scanned, total, matched)
        _check_cancel(cancel_event)
        if not groups:
            raise ValueError("筛选条件没有命中任何明细行，无法生成分析图表。")

        values_by_key = {
            key: tuple(_group_value(selected_groups[key], selected_metric.mode)
                       for selected_metric, selected_groups in zip(metrics, metric_groups))
            for key in groups
        }
        ordered_keys = list(groups)
        sort_index = options.sort_metric_index
        if options.sort_by == "group":
            ordered_keys.sort(key=lambda key: (key.casefold(), key))
        elif options.sort_by == "value_desc":
            ordered_keys.sort(key=lambda key: (values_by_key[key][sort_index] is None,
                                              -(values_by_key[key][sort_index] or Decimal(0)), key.casefold()))
        else:
            ordered_keys.sort(key=lambda key: (values_by_key[key][sort_index] is None,
                                              values_by_key[key][sort_index] or Decimal(0), key.casefold()))
        rows = [(groups[key], values_by_key[key][0]) for key in ordered_keys]
        chart_rows = []
        if multi_metric:
            limit_note = f"前 {options.chart_limit} 个分组" if options.chart_limit is not None else "全部分组"
            if options.chart_type == "pie":
                chart_note = ("每个指标分别绘制饼图，仅展示正值分组；超过上限时，其余分组合并为“其他”。"
                              "完整数据见汇总表，逐图数据见图表数据工作表。")
            else:
                sort_note = "分组名称" if options.sort_by == "group" else f"指标“{metrics[sort_index].name.strip()}”"
                chart_note = (f"所有指标共同使用按{sort_note}排序的{limit_note}；不同单位分别成图。"
                              "空值保留对应分组位置；完整数据见汇总表，逐项校验见指标校验工作表。")
        else:
            chart_rows, chart_note = _chart_rows(rows, options.chart_type, options.chart_limit)
        zero_groups = sum(
            any(selected_metric.mode == "ratio_of_sums" and selected_groups[key].valid > 0
                and selected_groups[key].denominator_sum == 0
                for selected_metric, selected_groups in zip(metrics, metric_groups))
            for key in ordered_keys
        )

        summary.append([_text_cell(summary, "分组分析与图表", title=True)])
        summary.append([_text_cell(summary, "源文件"), _text_cell(summary, source.name),
                        _text_cell(summary, "工作表"), _text_cell(summary, options.sheet)])
        metric_names_text = "、".join(selected_metric.name.strip() for selected_metric in metrics)
        if multi_metric:
            summary.append([_text_cell(summary, "指标"),
                            _wrapped_text_cell(summary, f"{len(metrics)} 个指标：{metric_names_text}；各指标独立设置统计口径，详见指标定义工作表。")])
        else:
            summary.append([_text_cell(summary, "指标"), _text_cell(summary, metric.name),
                            _text_cell(summary, "统计口径"), _text_cell(summary, _metric_description(metric))])
        summary.append([_text_cell(summary, "明细概况"),
                        _wrapped_text_cell(summary, f"扫描 {scanned} 行，筛中 {matched} 行，{len(groups)} 个分组；无效指标行 {invalid}；汇总分母为零的分组 {zero_groups}")])
        summary.append([_text_cell(summary, "图表说明"), _wrapped_text_cell(summary, chart_note)])
        if options.rules:
            descriptions = []
            for rule in options.rules:
                field = labels[rule.column - 1].partition(" · ")[2] or labels[rule.column - 1]
                values = "、".join("（空白）" if not value else str(value) for value in rule.values)
                descriptions.append(f"{field}{'等于' if rule.operator == 'equals' else '包含'}{values}")
            filter_description = "；".join(descriptions)
            if len(filter_description) > 500:
                filter_description = filter_description[:497] + "…"
        else:
            filter_description = "全部明细"
        summary.append([_text_cell(summary, "筛选范围"), _wrapped_text_cell(summary, filter_description)])
        if multi_metric:
            chart_points, chart_count = _write_multi_results(
                result_book, summary, options, metrics, metric_groups, ordered_keys,
                values_by_key, group_label, labels, cancel_event=cancel_event,
            )
        else:
            headers = ("分组", "指标值", "明细行数", "有效行数", "无效行数", "分子合计", "分母合计")
            summary.append([*(_text_cell(summary, label, header=True) for label in headers),
                            None, _text_cell(summary, "图表分组", header=True),
                            _text_cell(summary, "图表数值", header=True)])
            for index in range(max(len(rows), len(chart_rows))):
                cells_out = []
                if index < len(rows):
                    group, value = rows[index]
                    cells_out.extend([
                        _text_cell(summary, group.key if group.key else "（空白）"),
                        _numeric_cell(summary, value, percent=metric.as_percent),
                        group.rows,
                        group.valid,
                        group.invalid,
                        _numeric_cell(summary, group.numerator_sum if metric.mode in {"ratio_of_sums", "average_of_ratios"} else None),
                        _numeric_cell(summary, group.denominator_sum if metric.mode in {"ratio_of_sums", "average_of_ratios"} else None),
                    ])
                else:
                    cells_out.extend([None] * 7)
                cells_out.append(None)
                if index < len(chart_rows):
                    name, value = chart_rows[index]
                    cells_out.extend([_text_cell(summary, name), _numeric_cell(summary, value, percent=metric.as_percent)])
                summary.append(cells_out)
            summary.auto_filter.ref = f"A7:G{_TABLE_HEADER_ROW + len(rows)}"
            _make_chart(summary, options.chart_type, metric, group_label, chart_rows)
            chart_points, chart_count = len(chart_rows), 1
        _progress(progress, "save", matched, matched, matched)
        _check_cancel(cancel_event)
        fd, temp_name = tempfile.mkstemp(prefix=".excel_analysis_", suffix=".xlsx", dir=output.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        result_book.save(temp_path)
        _check_cancel(cancel_event)
        os.replace(temp_path, output)
        _progress(progress, "done", scanned, total, matched)
        return AnalysisResult(scanned, matched, len(groups), invalid, zero_groups, chart_points, output, len(metrics), chart_count)
    finally:
        if source_book is not None:
            source_book.close()
        _close_batch_workbook(result_book)
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()
