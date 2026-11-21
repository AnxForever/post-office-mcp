from __future__ import annotations

from collections import defaultdict
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from .normalize import normalize_plate


# 「合并」是草滩+港务一次导出的容积测量表；没有更高优先级，
# 但它自己独占一个来源，不会和逐场地记录同时出现。
SITE_ORDER = {"草滩": 0, "港务": 1, "合并": 2}
HEADERS = [
    "序号", "测量日期", "场地", "邮路", "车牌号", "挂车号", "长", "宽", "高",
    "测量容积", "车型问题", "备注", "候选记录数", "来源文件", "来源Sheet", "来源行",
]
EVIDENCE_HEADERS = [
    "序号", "测量日期", "场地", "车牌号", "来源文件", "来源Sheet", "来源行", "容积单元格",
    "原公式/数值", "长", "宽", "高", "取整前容积", "计算表达式", "取整规则", "最终容积",
]
VOLUME_METHOD_LABELS = {
    "source_value_rounded": "读取源单元格数值后取整",
    "dimensions_formula_fallback": "公式无缓存，按长×宽×高重算",
    "dimensions_fallback": "容积为空，按长×宽×高重算",
    "missing": "缺少可用容积和完整长宽高",
}
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
        "volume_unrounded": selected.get("volume_unrounded"),
        "volume_method": selected.get("volume_method"),
        "volume_formula": selected.get("volume_formula"),
        "volume_cell": selected.get("volume_cell"),
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
    rows = list(rows)
    worksheet = workbook.active
    worksheet.title = "车型汇总"
    worksheet.append(HEADERS)

    for sequence, row in enumerate(rows, 1):
        model_problem = row.get("model_problem")
        remark = row.get("remark") or ""
        if row["zero_value_note"]:
            problem_parts = [p for p in [model_problem, row["zero_value_note"]] if p]
            model_problem = " | ".join(problem_parts)
            remark = "; ".join(problem_parts) if not remark else f"{remark}; {'; '.join(problem_parts)}"
        worksheet.append([_safe_cell_value(value) for value in [
            sequence, row["measurement_date"], row.get("site"), row.get("route"), row["plate"], row.get("trailer"),
            row.get("length"), row.get("width"), row.get("height"), row.get("measured_volume"), model_problem,
            remark or model_problem, row["candidate_count"], row.get("source_file"), row.get("source_sheet"), row.get("source_row"),
        ]])
        # 车型汇总表不设颜色标注，问题通过备注列的文字描述即可

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
            cell.number_format = "0.000" if column != "J" else "0"
    worksheet.column_dimensions["A"].width = 8
    worksheet.column_dimensions["B"].width = 13
    worksheet.column_dimensions["C"].width = 10
    worksheet.column_dimensions["D"].width = 24
    for column in ("E", "F", "K", "L", "N", "O"):
        worksheet.column_dimensions[column].width = 18

    evidence = workbook.create_sheet("容积取证")
    evidence.append(EVIDENCE_HEADERS)
    for sequence, row in enumerate(rows, 1):
        formula = row.get("volume_formula")
        raw_value = f"公式文本: {formula}" if formula else row.get("volume_unrounded")
        dimensions = (row.get("length"), row.get("width"), row.get("height"))
        calculation = None
        if row.get("volume_method") in {"dimensions_formula_fallback", "dimensions_fallback"} and all(
            value not in (None, "") for value in dimensions
        ):
            calculation = f"{dimensions[0]} x {dimensions[1]} x {dimensions[2]} = {row.get('volume_unrounded')}"
        elif row.get("volume_unrounded") is not None:
            calculation = f"源单元格值 {row.get('volume_unrounded')}"
        evidence.append([
            sequence,
            row.get("measurement_date"),
            row.get("site"),
            row.get("plate"),
            row.get("source_file"),
            row.get("source_sheet"),
            row.get("source_row"),
            row.get("volume_cell"),
            raw_value,
            *dimensions,
            row.get("volume_unrounded"),
            calculation,
            VOLUME_METHOD_LABELS.get(row.get("volume_method"), row.get("volume_method")),
            row.get("measured_volume"),
        ])
    for cell in evidence[1]:
        cell.font = Font(bold=True)
    for cell in evidence["B"][1:]:
        cell.number_format = "yyyy-mm-dd"
    for column in ("J", "K", "L", "M"):
        for cell in evidence[column][1:]:
            cell.number_format = "0.000000"
    for cell in evidence["P"][1:]:
        cell.number_format = "0"
    evidence.freeze_panes = "A2"
    evidence.auto_filter.ref = f"A1:P{evidence.max_row}"
    evidence.sheet_view.showGridLines = False
    for column, width in {"D": 16, "E": 22, "F": 16, "H": 12, "I": 24, "N": 40, "O": 34}.items():
        evidence.column_dimensions[column].width = width

    workbook.save(Path(path))
