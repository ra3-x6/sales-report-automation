"""Command-line entry point."""

import argparse
import csv
import sys
from pathlib import Path
from zoneinfo import ZoneInfoNotFoundError

from .report import build_report
from .workbook import save_workbook


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build an Excel sales report from order CSV files."
    )
    parser.add_argument("inputs", nargs="+", type=Path)
    parser.add_argument("-o", "--output", required=True, type=Path)
    parser.add_argument("--currency", choices=("EGP", "USD", "EUR"), default="EGP")
    parser.add_argument("--timezone", default="Africa/Cairo", help="IANA reporting timezone")
    args = parser.parse_args(argv)
    if args.output.suffix.lower() != ".xlsx":
        parser.error("--output must end in .xlsx")

    try:
        report = build_report(args.inputs, currency=args.currency, timezone_name=args.timezone)
        save_workbook(report, args.output)
    except (OSError, ValueError, csv.Error, ZoneInfoNotFoundError) as exc:
        print(f"sales-report: {exc}", file=sys.stderr)
        return 1

    print(
        f"Wrote {args.output} | {len(report.lines)} accepted, "
        f"{report.duplicates} duplicates, {len(report.rejected)} rejected"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
