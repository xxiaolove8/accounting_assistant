"""Focused hidden Tk smoke; run from any directory with the project Python."""
from pathlib import Path
import sys
import tempfile
import tkinter as tk
from tkinter import ttk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import Workbook, load_workbook
import excel_filter_gui as gui
from filter_core import ExportOptions


def descendants(widget):
    yield widget
    for child in widget.winfo_children():
        yield from descendants(child)


def field(parent, row, column):
    return next(widget for widget in parent.winfo_children()
                if widget.winfo_manager() == "grid"
                and int(widget.grid_info()["row"]) == row
                and int(widget.grid_info()["column"]) == column)


def main():
    warnings = []
    originals = {name: getattr(gui.messagebox, name) for name in (
        "showwarning", "showerror", "showinfo", "askyesno",
    )}
    for name in ("showwarning", "showerror", "showinfo"):
        setattr(gui.messagebox, name, lambda *args, **kwargs: warnings.append(args))
    gui.messagebox.askyesno = lambda *args, **kwargs: True
    app = gui.ExcelFilterApp()
    app.attributes("-alpha", 0)
    callback_errors = []
    app.report_callback_exception = lambda *args: callback_errors.append(args)
    try:
        with tempfile.TemporaryDirectory(prefix="excel-gui-metrics-") as temporary:
            source = Path(temporary) / "source.xlsx"
            workbook = Workbook()
            workbook.active.append(["业务员", "计划回款", "实际回款", "应收金额"])
            workbook.active.append(["张三", 100, 80, 1000])
            workbook.active.append(["张三", 200, 100, 1000])
            workbook.active.append(["李四", 100, 100, 1000])
            workbook.save(source)
            workbook.close()
            app.headers = ["A · 业务员", "B · 计划回款", "C · 实际回款", "D · 应收金额"]
            app._build_options = lambda **kwargs: ExportOptions(
                source=source, sheet="Sheet", output=source.with_name("result.xlsx"),
            )
            app._headers_match_options = lambda options: True
            results = []
            app._run_operation = lambda title, worker, callback: results.append(worker(None, None))
            app._show_analysis_dialog()
            dialog = next(widget for widget in reversed(app.winfo_children()) if isinstance(widget, tk.Toplevel))
            dialog.attributes("-alpha", 0)
            app.update()
            widgets = list(descendants(dialog))
            buttons = [widget for widget in widgets if isinstance(widget, ttk.Button)]
            metrics = next(widget for widget in widgets if isinstance(widget, ttk.Treeview))
            editor = next(widget for widget in widgets if isinstance(widget, ttk.LabelFrame)
                          and widget.cget("text").startswith("当前指标"))

            def invoke(text):
                next(button for button in buttons if button.cget("text") == text).invoke()
                app.update()
                assert not callback_errors, callback_errors

            def value(widget, text):
                app.setvar(widget.cget("textvariable"), text)

            assert all(metrics.item(item, "values")[0] == "□" for item in metrics.get_children())
            invoke("勾选计划 / 实际回款")
            assert metrics.item("C2", "values")[0] == "☑"
            assert metrics.item("C3", "values")[0] == "☑"
            metrics.selection_set("C2")
            app.update()
            value(field(editor, 1, 3), "%")
            invoke("生成分析 Excel")
            assert warnings and not results and dialog.winfo_exists()
            warnings.clear()
            value(field(editor, 1, 3), "")
            invoke("新增比值 / 计数")
            assert field(editor, 0, 1).get() == "回款完成率"
            assert field(editor, 2, 1).get().startswith("C3｜")
            assert field(editor, 2, 3).get().startswith("C2｜")
            invoke("应用指标设置")
            assert metrics.item("M1", "values")[0] == "☑"
            invoke("生成分析 Excel")
            assert len(results) == 1 and not warnings, warnings
            assert results[0].metric_count == 3 and results[0].chart_count == 2
            exported = load_workbook(results[0].output)
            charts = [chart for sheet in exported for chart in sheet._charts]
            assert sorted(len(chart.series) for chart in charts) == [1, 2]
            assert "指标校验" in exported.sheetnames
            exported.close()

            # Without actual + planned headers, do not label a guessed ratio as completion.
            app.headers = ["A · 业务员", "B · 预计回款", "C · 应收金额"]
            app._show_analysis_dialog()
            dialog = next(widget for widget in reversed(app.winfo_children()) if isinstance(widget, tk.Toplevel))
            dialog.attributes("-alpha", 0)
            app.update()
            widgets = list(descendants(dialog))
            buttons = [widget for widget in widgets if isinstance(widget, ttk.Button)]
            editor = next(widget for widget in widgets if isinstance(widget, ttk.LabelFrame)
                          and widget.cget("text").startswith("当前指标"))
            invoke("新增比值 / 计数")
            assert field(editor, 0, 1).get() == "自定义比值"
            assert field(editor, 2, 1).get().startswith("C2｜")
            assert field(editor, 2, 3).get().startswith("C3｜")
            print("PASS: default actual/planned ratio, fallback ratio naming, unit validation, real multi-metric export")
    finally:
        app.destroy()
        for name, value in originals.items():
            setattr(gui.messagebox, name, value)


if __name__ == "__main__":
    main()
