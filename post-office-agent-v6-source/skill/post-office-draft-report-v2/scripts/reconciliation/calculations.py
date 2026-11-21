from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal
from datetime import date, datetime
from typing import Any, Iterable, Sequence

from .matching import Issue
from .normalize import destination_name, normalize_text, route_candidates, tonnage_key


@dataclass
class ProfitMetrics:
    trips: int = 0
    load: float = 0.0
    freight: float = 0.0
    income: float = 0.0

    @property
    def average_load(self) -> float:
        return self.load / self.trips if self.trips else 0.0

    @property
    def profit(self) -> float:
        return self.income - self.freight

    @property
    def average_profit(self) -> float:
        return self.profit / self.trips if self.trips else 0.0


def apply_cross_tonnage_flags(rows: list[dict[str, Any]]) -> None:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[row.get("plate_key", "")].append(row)
        row.setdefault("issues", [])
        row["cross_tonnage"] = "否"
        row["previous_volume"] = None
        row["previous_time"] = None
    for plate_rows in grouped.values():
        ordered_rows = sorted(plate_rows, key=_tonnage_baseline_key)
        counts = Counter(_cross_basis_tonnage(row) for row in ordered_rows if _cross_basis_tonnage(row))
        if not counts:
            continue
        maximum = max(counts.values())
        modes = {key for key, count in counts.items() if count == maximum}
        if len(modes) > 1:
            baseline_row = next(row for row in ordered_rows if _cross_basis_tonnage(row))
            baseline = _cross_basis_tonnage(baseline_row)
            previous: dict[str, Any] | None = None
            for row in ordered_rows:
                current = _cross_basis_tonnage(row)
                if current and baseline and current != baseline and not row.get("downgraded"):
                    row["cross_tonnage"] = "是"
                    if previous:
                        row["previous_volume"] = previous.get("volume")
                        row["previous_time"] = previous.get("seal_time")
                    row["issues"].append(Issue("cross_tonnage", "yellow", "该记录吨位偏离同车牌本月首条有效吨位", "人工确认是否跨档"))
                previous = row
            continue
        mode = next(iter(modes))
        previous: dict[str, Any] | None = None
        for row in ordered_rows:
            current = _cross_basis_tonnage(row)
            if current and current != mode and not row.get("downgraded"):
                row["cross_tonnage"] = "是"
                if previous:
                    row["previous_volume"] = previous.get("volume")
                    row["previous_time"] = previous.get("seal_time")
                row["issues"].append(Issue("cross_tonnage", "yellow", "该记录吨位偏离同车牌众数吨位", "人工确认是否跨档"))
            previous = row


def _cross_basis_tonnage(row: dict[str, Any]) -> str:
    """Use the volume-derived model tier; settlement downgrades are separate."""
    return tonnage_key(row.get("model_tonnage") or row.get("tonnage"))


def _tonnage_baseline_key(row: dict[str, Any]) -> tuple[date, str, str, int]:
    value = row.get("execution_date") or row.get("seal_time")
    if isinstance(value, datetime):
        execution_date = value.date()
    elif isinstance(value, date):
        execution_date = value
    elif isinstance(value, str):
        try:
            execution_date = date.fromisoformat(value[:10])
        except ValueError:
            execution_date = date.max
    else:
        execution_date = date.max
    source_file = normalize_text(row.get("source_path") or row.get("source_file"))
    source_sheet = normalize_text(row.get("source_sheet"))
    try:
        source_row = int(row.get("source_row") or 0)
    except (TypeError, ValueError):
        source_row = 0
    return execution_date, source_file, source_sheet, source_row


def finance_tonnage(value: Any, available: Sequence[str]) -> str | None:
    key = tonnage_key(value)
    if not key:
        return None
    try:
        number = Decimal(key)
    except Exception:
        return key if key in available else None
    if number <= 0:
        return None
    numeric: list[tuple[Decimal, str]] = []
    for item in available:
        normalized = tonnage_key(item)
        if normalized in {"40A", "40B"}:
            numeric.append((Decimal(40), normalized))
        else:
            try:
                numeric.append((Decimal(normalized), normalized))
            except Exception:
                continue
    eligible = sorted((limit, label) for limit, label in numeric if limit >= number)
    if not eligible:
        return None
    # Numeric 40 maps to the operational 40A band; explicit 40B remains 40B above.
    for limit, label in eligible:
        if limit == 40 and label == "40A":
            return label
        if limit >= number:
            return label
    return None


def _canonical(route: Any) -> str:
    candidates = route_candidates(route)
    return candidates[-1] if candidates else normalize_text(route)


def _profit_route(route: Any, kind: str) -> str:
    canonical = _canonical(route)
    return canonical


def aggregate_profit(
    base_rows: Iterable[dict[str, Any]],
    finance_rows: Iterable[dict[str, Any]],
    kind: str,
    available_tonnages: Sequence[str],
) -> dict[tuple[str, str], ProfitMetrics]:
    result: dict[tuple[str, str], ProfitMetrics] = defaultdict(ProfitMetrics)
    for row in base_rows:
        row_kind = row.get("price_kind")
        if row_kind != kind:
            continue
        tonnage = tonnage_key(row.get("tonnage"))
        if not tonnage:
            continue
        key = (_profit_route(row.get("standard_route") or row.get("route"), kind), tonnage)
        result[key].load += float(row.get("transport_volume") or 0)
        result[key].freight += float(row.get("freight") or row.get("运价") or 0)
    for row in finance_rows:
        if "price_kind" in row and row.get("price_kind") != kind:
            continue
        canonical = _canonical(row.get("route"))
        if kind == "first" and not canonical.startswith("西安-"):
            continue
        tonnage = finance_tonnage(row.get("tonnage"), available_tonnages)
        if not tonnage:
            continue
        key = (_profit_route(canonical, kind), tonnage)
        result[key].trips += 1
        result[key].income += float(row.get("net_income") or 0)
    return dict(result)
