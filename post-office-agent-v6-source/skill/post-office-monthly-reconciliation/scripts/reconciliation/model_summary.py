from __future__ import annotations

from collections import defaultdict
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import PatternFill

from .normalize import normalize_plate


SITE_ORDER = {"草滩": 0, "港务": 1}
HEADERS = [
    "序号", "测量日期", "场地", "邮路", "车牌号", "挂车号", "长", "宽", "高",
    "测量容积", "车型问题", "备注", "候选记录数", "来源文件", "来源Sheet", "来源行",
]
UNMEASURED_FILL = PatternFill(fill_type="solid", fgColor="FFFFC000")
INCONSISTENT_FILL = PatternFill(fill_type="solid", fgColor="FFFF0000")


def _numeric(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    try:
        return Decimal(str(value).replace(",", ""))
    except (InvalidOperation, ValueError):
        return Decimal("0")


def _measurement_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value)
        except ValueError:
            pass
    raise ValueError(f"在线记录缺少有效测量日期: {value!r}")


def _volume(row: dict[str, Any]) -> Any:
    return row.get("measured_volume", row.get("volume"))


def _sheet_order(row: dict[str, Any]) -> int:
    return int(row.get("source_sheet_order", row.get("source_order", 0)))


def _zero_value_note(selected: dict[str, Any], has_nonzero: bool) -> str | None:
    if has_nonzero:
        return None
    dimensions = (selected.get("length"), selected.get("width"), selected.get("height"))
    if all(value is None or value == "" for value in dimensions):
        return "在线记录确认未测量"
    return "在线零值与尺寸不一致"


def _safe_cell_value(value: Any) -> Any:
    """Prevent source formulas and broken image-formula text from entering output."""
    if not isinstance(value, str):
        return value
    normalized = value.strip().upper()
    if normalized.startswith("=") or "DISPIMG" in normalized or "#NAME?" in normalized:
        return None
    return value


def _build_summary_row(
    key: tuple[date, str], candidates: list[dict[str, Any]], selected: dict[str, Any], has_nonzero: bool,
) -> dict[str, Any]:
    result = {
        "measurement_date": key[0],
        "plate": key[1],
        "selected_source": selected["site"],
        "site": selected["site"],
        "route": selected.get("route"),
        "trailer": selected.get("trailer"),
        "length": selected.get("length"),
        "width": selected.get("width"),
        "height": selected.get("height"),
        "measured_volume": _volume(selected),
        "model_problem": selected.get("problem", selected.get("model_problem")),
        "remark": selected.get("remark"),
        "candidate_count": len(candidates),
        "source_file": selected.get("source_file"),
        "source_sheet": selected.get("source_sheet"),
        "source_row": selected.get("source_row"),
    }
    result["zero_value_note"] = _zero_value_note(selected, has_nonzero)
    return result


def merge_online_rows(rows: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Select one physical online record per date and normalized plate."""
    groups: dict[tuple[date, str], list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        groups[(_measurement_date(row["measurement_date"]), normalize_plate(row["plate"]))].append(row)

    result: list[dict[str, Any]] = []
    for key in sorted(groups):
        candidates = sorted(
            groups[key],
            key=lambda row: (SITE_ORDER[row["site"]], _sheet_order(row), row["source_row"]),
        )
        nonzero = [row for row in candidates if _numeric(_volume(row)) != 0]
        selected = (nonzero or candidates)[-1]
        result.append(_build_summary_row(key, candidates, selected, bool(nonzero)))
    return result


def write_model_summary_workbook(rows: Iterable[dict[str, Any]], path: Path | str, month: tuple[int, int]) -> None:
    """Write the value-only monthly vehicle-model summary workbook."""
    del month  # The caller owns month filtering; dates are retained from the selected rows.
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "车型汇总"
    worksheet.append(HEADERS)

    for sequence, row in enumerate(rows, 1):
        model_problem = row["zero_value_note"] or row.get("model_problem")
        worksheet.append([_safe_cell_value(value) for value in [
            sequence, row["measurement_date"], row.get("site"), row.get("route"), row["plate"], row.get("trailer"),
            row.get("length"), row.get("width"), row.get("height"), row.get("measured_volume"), model_problem,
            row.get("remark"), row["candidate_count"], row.get("source_file"), row.get("source_sheet"), row.get("source_row"),
        ]])
        if row["zero_value_note"]:
            fill = UNMEASURED_FILL if row["zero_value_note"] == "在线记录确认未测量" else INCONSISTENT_FILL
            for cell in worksheet[worksheet.max_row]:
                cell.fill = fill

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:P{worksheet.max_row}"
    worksheet.sheet_view.topLeftCell = "A1"
    worksheet.sheet_view.selection[0].activeCell = "A1"
    for cell in worksheet[1]:
        font = copy(cell.font)
        font.bold = True
        cell.font = font
    for cell in worksheet["B"][1:]:
        cell.number_format = "yyyy-mm-dd"
    for column in ("G", "H", "I", "J"):
        for cell in worksheet[column][1:]:
            cell.number_format = "0.00"
    worksheet.column_dimensions["A"].width = 8
    worksheet.column_dimensions["B"].width = 13
    worksheet.column_dimensions["C"].width = 10
    worksheet.column_dimensions["D"].width = 24
    for column in ("E", "F", "K", "L", "N", "O"):
        worksheet.column_dimensions[column].width = 18

    workbook.save(Path(path))
