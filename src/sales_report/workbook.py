"""Write a styled Excel report from validated order lines."""

from collections.abc import Iterable, Sequence
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from .report import Report

NAVY = "18354B"
TEAL = "147E78"
PALE = "EDF5F4"
MONEY = "#,##0.00"


def add_row(sheet: Worksheet, row: int, values: Sequence[object]) -> None:
    """Write input text as literal cells, including text starting with '='."""
    for index, value in enumerate(values, 1):
        if isinstance(value, str):
            cell = sheet.cell(row, index, ILLEGAL_CHARACTERS_RE.sub("", value)[:32767])
            cell.data_type = "s"
        else:
            sheet.cell(row, index, value)


def populate_table(
    sheet: Worksheet,
    headers: Sequence[str],
    rows: Iterable[Sequence[object]],
    widths: Sequence[int],
    money_columns: Sequence[int] = (),
) -> None:
    add_row(sheet, 1, headers)
    for index, row in enumerate(rows, 2):
        add_row(sheet, index, row)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[sheet.cell(1, index).column_letter].width = width
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, color="FFFFFF")
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=PALE)
            if cell.column in money_columns:
                cell.number_format = MONEY


def create_workbook(report: Report) -> Workbook:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.sheet_view.showGridLines = False
    summary.merge_cells("A1:B1")
    summary["A1"] = "Sales report"
    summary["A1"].fill = PatternFill("solid", fgColor=NAVY)
    summary["B1"].fill = PatternFill("solid", fgColor=NAVY)
    summary["A1"].font = Font(size=20, color="FFFFFF", bold=True)
    summary["A1"].alignment = Alignment(vertical="center", indent=1)
    summary.row_dimensions[1].height = 38
    summary.column_dimensions["A"].width = 35
    summary.column_dimensions["B"].width = 26

    by_day, by_sku = report.breakdowns()
    completed = {line.order_id for line in report.lines if line.status == "completed"}
    cancelled = {line.order_id for line in report.lines if line.status == "cancelled"}
    pending = {line.order_id for line in report.lines if line.status == "pending"}
    metrics = [
        ("Currency", report.currency),
        ("Reporting timezone", report.timezone_name),
        ("Completed orders", len(completed)),
        ("Cancelled orders", len(cancelled)),
        ("Pending orders", len(pending)),
        ("Units sold", sum(total.units for total in by_day.values())),
        ("Completed revenue", sum(total.revenue_minor for total in by_day.values()) / 100),
        ("Accepted order lines", len(report.lines)),
        ("Duplicate rows skipped", report.duplicates),
        ("Rejected rows", len(report.rejected)),
    ]
    for row, (label, value) in enumerate(metrics, 3):
        summary.cell(row, 1, label).font = Font(bold=True, color=NAVY)
        value_cell = summary.cell(row, 2, value)
        if label == "Completed revenue":
            value_cell.number_format = MONEY
            value_cell.font = Font(bold=True, color=TEAL)
        if row % 2:
            for cell in summary[row][:2]:
                cell.fill = PatternFill("solid", fgColor=PALE)

    daily = workbook.create_sheet("Daily Sales")
    populate_table(
        daily,
        ("Day", "Units", f"Revenue ({report.currency})"),
        (
            (day.isoformat(), total.units, total.revenue_minor / 100)
            for day, total in sorted(by_day.items())
        ),
        (18, 14, 23),
        (3,),
    )
    if by_day:
        chart = BarChart()
        chart.title = "Daily completed sales"
        chart.y_axis.title = report.currency
        chart.height = 9
        chart.width = 19
        chart.add_data(
            Reference(daily, min_col=3, min_row=1, max_row=daily.max_row), titles_from_data=True
        )
        chart.set_categories(Reference(daily, min_col=1, min_row=2, max_row=daily.max_row))
        chart.series[0].graphicalProperties.solidFill = TEAL
        chart.legend = None
        summary.add_chart(chart, "D3")

    products = workbook.create_sheet("Product Sales")
    populate_table(
        products,
        ("SKU", "Units", f"Revenue ({report.currency})"),
        (
            (sku, total.units, total.revenue_minor / 100)
            for sku, total in sorted(
                by_sku.items(), key=lambda item: (-item[1].revenue_minor, item[0])
            )
        ),
        (24, 14, 23),
        (3,),
    )
    lines = workbook.create_sheet("Order Lines")
    populate_table(
        lines,
        (
            "Order ID",
            "Line ID",
            "Ordered at (UTC)",
            "Status",
            "SKU",
            "Product",
            "Quantity",
            "Unit price",
            "Line total",
            "Currency",
        ),
        (
            (
                line.order_id,
                line.line_id,
                line.ordered_at.isoformat(),
                line.status,
                line.sku,
                line.product_name,
                line.quantity,
                line.unit_price_minor / 100,
                line.amount_minor / 100,
                line.currency,
            )
            for line in report.lines
        ),
        (21, 18, 30, 16, 20, 29, 13, 16, 16, 12),
        (8, 9),
    )
    rejected = workbook.create_sheet("Rejected Rows")
    populate_table(
        rejected,
        ("Source file", "CSV line", "Order ID", "Line ID", "Reason"),
        (
            (
                entry.source.file,
                entry.source.row,
                entry.source.raw_order_id,
                entry.source.raw_line_id,
                entry.reason,
            )
            for entry in report.rejected
        ),
        (44, 13, 20, 20, 63),
    )
    return workbook


def save_workbook(report: Report, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook = create_workbook(report)
    with NamedTemporaryFile(dir=output.parent, suffix=".xlsx", delete=False) as temporary:
        staged = Path(temporary.name)
    try:
        workbook.save(staged)
        staged.replace(output)
    finally:
        workbook.close()
        staged.unlink(missing_ok=True)
