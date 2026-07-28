from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .normalize import normalize_text, tonnage_key


def _previous_month_label(target_month: str) -> str | None:
    try:
        year, month = map(int, target_month.split("-"))
    except (ValueError, AttributeError):
        return None
    prev_month = month - 1
    prev_year = year
    if prev_month == 0:
        prev_month = 12
        prev_year -= 1
    return f"{str(prev_year)[2:]}\u5e74{prev_month}\u6708"


def read_previous_month_totals(
    path: Path | str | None,
    kind: str,
    layout: dict[str, Any] | None = None,
) -> tuple[dict[tuple[str, str], dict[str, float]] | None, str | None]:
    if not path:
        return None, "\u4e0a\u6708\u76c8\u4e8f\u8868\uff08\u6682\u65e0\u6570\u636e\uff09"
    p = Path(path)
    if not p.is_file():
        return None, f"\u4e0a\u6708\u76c8\u4e8f\u8868\u4e0d\u5b58\u5728: {p.name}"
    # Temporary compatibility for callers using the pre-layout target-month API.
    # New phase-two code always passes the layout returned by discover_profit_template.
    legacy_label = _previous_month_label(kind) if layout is None else None
    try:
        wb = load_workbook(p, read_only=True, data_only=True, keep_links=False)
    except Exception as exc:
        return None, f"\u4e0a\u6708\u76c8\u4e8f\u8868\u65e0\u6cd5\u8bfb\u53d6: {exc}"
    try:
        if layout is None:
            if not legacy_label:
                return None, "\u65e0\u6cd5\u8bc6\u522b\u4e0a\u6708\u6708\u4efd\u6807\u7b7e"
            if legacy_label not in wb.sheetnames:
                return None, f"\u4e0a\u6708\u76c8\u4e8f\u8868\u7f3a\u5c11{legacy_label}Sheet"
            return None, "\u4e0a\u6708\u76c8\u4e8f\u8868\u7f3a\u5c11\u52a8\u6001\u5e03\u5c40"

        sheet_name = normalize_text(layout.get("sheet"))
        if not sheet_name or sheet_name not in wb.sheetnames:
            return None, f"\u4e0a\u6708\u76c8\u4e8f\u8868\u7f3a\u5c11{sheet_name or '有效月份'}Sheet"
        ws = wb[sheet_name]
        groups = [
            (int(group["start_column"]), tonnage_key(group.get("tonnage") or group.get("label")))
            for group in layout.get("groups", [])
            if isinstance(group, dict)
            and group.get("start_column")
            and tonnage_key(group.get("tonnage") or group.get("label"))
        ]
        if not groups:
            return None, f"\u4e0a\u6708\u76c8\u4e8f\u8868{sheet_name}\u65e0\u6709\u6548\u5428\u4f4d\u5206\u7ec4"
        totals: dict[tuple[str, str], dict[str, float]] = {}
        for route_item in layout.get("route_rows", []):
            if not isinstance(route_item, dict) or not route_item.get("row"):
                continue
            row = int(route_item["row"])
            route = normalize_text(route_item.get("route") or ws.cell(row, 1).value)
            if not route:
                continue
            for start_col, tonnage in groups:
                try:
                    trips = float(ws.cell(row, start_col).value or 0)
                    freight = float(ws.cell(row, start_col + 3).value or 0)
                    income = float(ws.cell(row, start_col + 4).value or 0)
                except (TypeError, ValueError):
                    continue
                # Presence is separate from value: an explicit all-zero group is
                # comparable, while a group absent from the workbook is not.
                totals[(route, tonnage)] = {"trips": trips, "freight": freight, "income": income}
        if not totals:
            return None, f"\u4e0a\u6708\u76c8\u4e8f\u8868{sheet_name}\u65e0\u6570\u636e\u884c"
        return totals, None
    except Exception as exc:
        return None, f"\u4e0a\u6708\u76c8\u4e8f\u8868\u89e3\u6790\u5931\u8d25: {exc}"
    finally:
        wb.close()
