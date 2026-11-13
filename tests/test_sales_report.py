import csv
from pathlib import Path

from openpyxl import load_workbook

from sales_report.cli import main
from sales_report.report import COLUMNS, build_report
from sales_report.workbook import create_workbook

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def write_csv(path, rows):
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(COLUMNS)
        writer.writerows(rows)
    return path


def example_row(order_id="order-1", quantity=1, name="Tea", status="completed"):
    return (order_id, "1", "2026-01-10T22:30:00Z", status, "TEA", name, quantity, 1250, "EGP")


def test_cli_sample_has_reconciled_revenue_and_rejections(tmp_path):
    output = tmp_path / "nested" / "sales.xlsx"
    result = main(
        [
            str(EXAMPLES / "orders-a.csv"),
            str(EXAMPLES / "orders-b.csv"),
            "-o",
            str(output),
        ]
    )
    assert result == 0
    with output.open("rb") as file:
        workbook = load_workbook(file)
        assert workbook.sheetnames == [
            "Summary",
            "Daily Sales",
            "Product Sales",
            "Order Lines",
            "Rejected Rows",
        ]
        assert workbook["Summary"]["B5"].value == 2
        assert workbook["Summary"]["B6"].value == 1
        assert workbook["Summary"]["B7"].value == 1
        assert workbook["Summary"]["B9"].value == 265.50
        assert workbook["Summary"]["B10"].value == 5
        assert workbook["Summary"]["B11"].value == 1
        assert workbook["Summary"]["B12"].value == 2
        revenue = sum(
            row[2] for row in workbook["Daily Sales"].values if isinstance(row[2], (float, int))
        )
        assert revenue == 265.50
        workbook.close()


def test_conflicting_duplicate_excludes_the_order(tmp_path):
    first = write_csv(tmp_path / "a.csv", [example_row(), example_row(order_id="order-2")])
    second = write_csv(tmp_path / "b.csv", [example_row(quantity=2)])
    report = build_report([first, second], currency="EGP", timezone_name="UTC")
    assert [line.order_id for line in report.lines] == ["order-2"]
    assert len(report.rejected) == 2
    assert all("Conflicting rows" in rejected.reason for rejected in report.rejected)


def test_reporting_day_uses_requested_timezone(tmp_path):
    source = write_csv(tmp_path / "orders.csv", [example_row()])
    report = build_report([source], currency="EGP", timezone_name="Africa/Cairo")
    daily, _ = report.breakdowns()
    assert [day.isoformat() for day in daily] == ["2026-01-11"]


def test_spreadsheet_writes_imported_text_as_literal(tmp_path):
    source = write_csv(tmp_path / "orders.csv", [example_row(name="=1+1")])
    report = build_report([source], currency="EGP", timezone_name="UTC")
    workbook = create_workbook(report)
    assert workbook["Order Lines"]["F2"].value == "=1+1"
    assert workbook["Order Lines"]["F2"].data_type == "s"
    workbook.close()


def test_bad_header_does_not_overwrite_an_existing_report(tmp_path):
    source = tmp_path / "orders.csv"
    source.write_text("unexpected,header\n1,2\n", encoding="utf-8")
    output = tmp_path / "sales.xlsx"
    output.write_bytes(b"original report")
    assert main([str(source), "-o", str(output)]) == 1
    assert output.read_bytes() == b"original report"


def test_empty_input_with_valid_header_creates_zero_revenue_report(tmp_path):
    source = write_csv(tmp_path / "empty.csv", [])
    report = build_report([source], currency="EGP", timezone_name="UTC")
    workbook = create_workbook(report)
    assert workbook["Summary"]["B9"].value == 0
    assert workbook["Daily Sales"].max_row == 1
    workbook.close()
