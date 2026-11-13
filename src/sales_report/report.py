"""Read order exports and reconcile them before counting revenue."""

import csv
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import UTC, date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

COLUMNS = (
    "order_id",
    "line_id",
    "ordered_at",
    "status",
    "sku",
    "product_name",
    "quantity",
    "unit_price_minor",
    "currency",
)
STATUSES = {"pending", "completed", "cancelled"}
IDENTIFIER = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,63}")


@dataclass(frozen=True)
class OrderLine:
    order_id: str
    line_id: str
    ordered_at: datetime
    status: str
    sku: str
    product_name: str
    quantity: int
    unit_price_minor: int
    currency: str

    @property
    def amount_minor(self) -> int:
        return self.quantity * self.unit_price_minor


@dataclass(frozen=True)
class SourceRow:
    file: str
    row: int
    raw_order_id: str
    raw_line_id: str


@dataclass(frozen=True)
class RejectedRow:
    source: SourceRow
    reason: str


@dataclass
class SalesTotals:
    units: int = 0
    revenue_minor: int = 0

    def add(self, line: OrderLine) -> None:
        self.units += line.quantity
        self.revenue_minor += line.amount_minor


@dataclass
class Report:
    currency: str
    timezone_name: str
    lines: list[OrderLine] = field(default_factory=list)
    rejected: list[RejectedRow] = field(default_factory=list)
    duplicates: int = 0

    def breakdowns(self) -> tuple[dict[date, SalesTotals], dict[str, SalesTotals]]:
        by_day: dict[date, SalesTotals] = defaultdict(SalesTotals)
        by_sku: dict[str, SalesTotals] = defaultdict(SalesTotals)
        local_zone = ZoneInfo(self.timezone_name)
        for line in self.lines:
            if line.status == "completed":
                by_day[line.ordered_at.astimezone(local_zone).date()].add(line)
                by_sku[line.sku].add(line)
        return dict(by_day), dict(by_sku)


def read_line(values: dict[str, str], currency: str) -> OrderLine:
    for key in ("order_id", "line_id", "sku"):
        if not IDENTIFIER.fullmatch(values[key]):
            raise ValueError(f"Invalid {key}")

    try:
        ordered_at = datetime.fromisoformat(values["ordered_at"])
    except ValueError:
        raise ValueError("ordered_at must be an ISO 8601 timestamp") from None
    if ordered_at.utcoffset() is None:
        raise ValueError("ordered_at must include a UTC offset")

    if values["status"] not in STATUSES:
        raise ValueError("status must be pending, completed or cancelled")
    if values["currency"] != currency:
        raise ValueError(f"currency must be {currency}")
    if not 1 <= len(values["product_name"]) <= 120:
        raise ValueError("product_name must be 1 to 120 characters")
    if any(ord(char) < 32 for char in values["product_name"]):
        raise ValueError("product_name contains a control character")

    numbers = {}
    for field_name, minimum, maximum in (
        ("quantity", 1, 10_000),
        ("unit_price_minor", 0, 100_000_000),
    ):
        raw_value = values[field_name]
        if not re.fullmatch(r"[0-9]{1,9}", raw_value):
            raise ValueError(f"{field_name} must be an integer")
        value = int(raw_value)
        if not minimum <= value <= maximum:
            raise ValueError(f"{field_name} must be between {minimum} and {maximum}")
        numbers[field_name] = value

    return OrderLine(
        order_id=values["order_id"],
        line_id=values["line_id"],
        ordered_at=ordered_at.astimezone(UTC),
        status=values["status"],
        sku=values["sku"],
        product_name=values["product_name"],
        quantity=numbers["quantity"],
        unit_price_minor=numbers["unit_price_minor"],
        currency=currency,
    )


def build_report(inputs: list[Path], *, currency: str, timezone_name: str) -> Report:
    ZoneInfo(timezone_name)
    report = Report(currency=currency, timezone_name=timezone_name)
    grouped: dict[str, list[tuple[OrderLine, SourceRow]]] = defaultdict(list)
    invalid_orders: set[str] = set()

    for path in inputs:
        with path.open(encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file, strict=True)
            fields = reader.fieldnames or []
            if len(fields) != len(COLUMNS) or set(fields) != set(COLUMNS):
                raise ValueError(f"{path}: expected these CSV columns: {', '.join(COLUMNS)}")

            for record in reader:
                values = {key: (record[key] or "").strip() for key in COLUMNS}
                source = SourceRow(
                    str(path), reader.line_num, values["order_id"], values["line_id"]
                )
                try:
                    if None in record or any(record[key] is None for key in COLUMNS):
                        raise ValueError("Wrong number of CSV columns")
                    line = read_line(values, currency)
                except (ValueError, OverflowError) as exc:
                    report.rejected.append(RejectedRow(source, str(exc)))
                    invalid_orders.add(source.raw_order_id)
                else:
                    grouped[line.order_id].append((line, source))

    for order_id, entries in grouped.items():
        first = entries[0][0]
        unique: dict[str, OrderLine] = {}
        reason = None
        if order_id in invalid_orders:
            reason = "Another row in this order failed validation"
        for line, _ in entries:
            if (line.ordered_at, line.status) != (first.ordered_at, first.status):
                reason = reason or "Conflicting timestamps or statuses for the same order"
            previous = unique.setdefault(line.line_id, line)
            if previous != line:
                reason = reason or "Conflicting rows for the same order line"
        if reason:
            report.rejected.extend(RejectedRow(source, reason) for _, source in entries)
        else:
            report.lines.extend(unique.values())
            report.duplicates += len(entries) - len(unique)

    report.lines.sort(key=lambda line: (line.ordered_at, line.order_id, line.line_id))
    report.rejected.sort(key=lambda entry: (entry.source.file, entry.source.row))
    return report
