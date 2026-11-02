# -*- coding: utf-8 -*-
"""02 基础报表初版（人工核对后）-> 03 基础报表拆分版。

独立技能：只认 02 里的「基础报表」和「拆分元数据」，不依赖 run-state。
拆分元数据由初版技能写入，记录每行对应的来源 ID、干线归属、标准邮路、
车型/结算吨位等；人工修订 02 后这些字段仍随行保存。
"""
from __future__ import annotations

from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import json
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from reconciliation.excel_io import read_price_catalog
from reconciliation.matching import _date_key, match_price
from reconciliation.normalize import normalize_plate, normalize_text
from reconciliation.template_discovery import discover_draft_template, discover_split_template
from reconciliation.templates import write_split_workbook

REVIEW_FIELD_HEADERS = {
    "company": "车辆承运单位",
    "dispatch_no": "派车单号",
    "route": "邮路名称",
    "seal_time": "封车时间",
    "plate": "车牌号",
    "volume": "容间",
    "tonnage": "吨位",
    "distance": "公里",
    "transport_volume": "运输量",
    "freight": "运价",
    "standard_price": "标准价",
    "remark": "备注",
    "cross_tonnage": "是否跨档",
    "previous_volume": "变动前容间",
    "previous_time": "变动前时间",
    "coefficient": "系数",
}
META_SHEET = "拆分元数据"
META_HEADERS = [
    "行号", "邮路名称", "车牌号", "派车单号", "干线归属", "标准邮路",
    "车型吨位", "结算吨位", "是否降吨位", "原运费", "差价", "来源ID",
]


def _load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(value, dict) and isinstance(value.get("aliases"), dict):
        return value["aliases"]
    return value if isinstance(value, dict) else {}


def _date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value.strip()[:10].replace("/", "-"))
        except ValueError:
            return None
    return None


def infer_month_from_review(path: Path | str) -> dict:
    """从 02 的封车时间推断月份。"""
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        sheet = workbook["基础报表"]
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, ()) or ()
        columns = {normalize_text(value): index for index, value in enumerate(header) if value}
        stamp_index = columns.get(normalize_text("封车时间"))
        counter: Counter[tuple[int, int]] = Counter()
        if stamp_index is not None:
            for row in rows:
                stamp = _date(row[stamp_index] if stamp_index < len(row) else None)
                if stamp:
                    counter[(stamp.year, stamp.month)] += 1
    finally:
        workbook.close()
    if not counter:
        return {"month": None, "questions": ["基础报表里没有可用的封车时间，请告诉我这是哪个月"]}
    if len(counter) > 1:
        return {
            "month": None,
            "questions": [
                "基础报表里出现了多个月份（"
                + "、".join(f"{y}年{m}月" for y, m in sorted(counter))
                + "），请确认要拆分的是哪个月"
            ],
        }
    (year, month), _ = counter.most_common(1)[0]
    return {"month": f"{year:04d}-{month:02d}"}


def _read_metadata(path: Path | str, sheet_names: list[str]) -> dict[int, dict[str, Any]]:
    if META_SHEET not in sheet_names:
        raise ValueError(
            f"基础报表初版缺少「{META_SHEET}」工作表。"
            "请用本流程的初版技能重新生成 02 后再拆分。"
        )
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        rows = workbook[META_SHEET].iter_rows(values_only=True)
        header = next(rows, ()) or ()
        columns = {normalize_text(value): index for index, value in enumerate(header) if value}
        result: dict[int, dict[str, Any]] = {}
        for row in rows:
            def get(name: str) -> Any:
                index = columns.get(normalize_text(name))
                return row[index] if index is not None and index < len(row) else None

            try:
                excel_row = int(get("行号"))
            except (TypeError, ValueError):
                continue
            result[excel_row] = {
                "meta_route": get("邮路名称"),
                "plate": get("车牌号"),
                "dispatch_no": get("派车单号"),
                "seal_time": get("封车时间"),
                "price_kind": normalize_text(get("干线归属")) or None,
                "standard_route": get("标准邮路") or None,
                "model_tonnage": get("车型吨位") or None,
                "settlement_tonnage": get("结算吨位"),
                "downgraded": bool(get("是否降吨位")),
                "original_freight": get("原运费"),
                "price_difference": get("差价"),
                "source_id": get("来源ID") or "",
            }
    finally:
        workbook.close()
    return result


def _apply_reviewed_classification(
    rows: list[dict[str, Any]],
    catalogs: dict[str, Any],
    aliases: dict[str, Any],
) -> None:
    """只重算路线归属，不动人工核对过的吨位和运费。

    02 是用当月价格表算出来的；拆分阶段如果再用内置价格表重算一遍，
    遇到合同价目不同（或用户改过吨位）就会把核对结果覆盖掉。所以这里
    只处理路线被改过、或元数据里没有标准邮路/干线归属的行。
    """
    for row in rows:
        row["plate_key"] = normalize_plate(row.get("plate"))
        route_changed = (
            row.get("meta_route") not in (None, "")
            and normalize_text(row.get("meta_route")) != normalize_text(row.get("route"))
        )
        if (
            not route_changed
            and row.get("price_kind") in {"first", "second"}
            and row.get("standard_route")
        ):
            continue
        round_trip = "往返邮路" in normalize_text(row.get("remark"))
        price = match_price(row.get("route"), row.get("volume") or 0, catalogs, aliases, round_trip=round_trip)
        row["price_kind"] = price.kind or row.get("price_kind")
        row["standard_route"] = price.route or row.get("standard_route")
        if row.get("model_tonnage") in (None, ""):
            row["model_tonnage"] = price.model_tonnage


def read_review_rows(path: Path | str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    review_path = Path(path)
    workbook = load_workbook(review_path, data_only=True, read_only=True, keep_links=False)
    try:
        sheet_names = list(workbook.sheetnames)
        if "基础报表" not in sheet_names:
            raise ValueError("基础报表初版缺少「基础报表」工作表")
        layout = discover_draft_template(review_path)
        metadata = _read_metadata(review_path, sheet_names)
        # 人工修订时可能插行/删行，行号会漂移；优先按（车牌, 封车时间, 派车单号）
        # 找元数据，行号只作为兜底。
        metadata_by_key: dict[tuple, list[dict[str, Any]]] = {}
        metadata_blank: dict[tuple, list[dict[str, Any]]] = {}
        for excel_row, entry in metadata.items():
            entry["_excel_row"] = excel_row
            plate = normalize_plate(entry.get("plate"))
            stamp = _date_key(entry.get("seal_time"))
            dispatch = normalize_text(entry.get("dispatch_no"))
            if not plate or not stamp:
                continue
            if dispatch:
                metadata_by_key.setdefault((plate, stamp, dispatch), []).append(entry)
            else:
                metadata_blank.setdefault((plate, stamp), []).append(entry)
        used_metadata: set[int] = set()

        def pick(
            candidates: list[dict[str, Any]], excel_row: int, route: Any
        ) -> dict[str, Any] | None:
            if not candidates:
                return None
            available = [item for item in candidates if id(item) not in used_metadata]
            if not available:
                available = candidates
            route_key = normalize_text(route)

            def score(item: dict[str, Any]) -> tuple[int, int]:
                # 人工插行后，元数据行号仍是旧值，光看行号会把同车同单的两趟
                # （去程/回程）配反。先用邮路名称对齐，再按行号就近兜底。
                same_route = 0 if route_key and normalize_text(item.get("meta_route")) == route_key else 1
                return same_route, abs(int(item["_excel_row"]) - excel_row)

            return min(available, key=score)

        metadata_missing = 0
        sheet = workbook["基础报表"]
        header_row = int(layout["header_row"])
        data_start_row = int(layout["data_start_row"])
        trailer_column = int(layout["headers"]["车牌号"]) + 1

        rows: list[dict[str, Any]] = []
        for excel_row, values in enumerate(
            sheet.iter_rows(min_row=data_start_row, values_only=True), data_start_row
        ):
            if not any(value not in (None, "") for value in values):
                continue

            def cell(header: str) -> Any:
                index = int(layout["headers"][header]) - 1
                return values[index] if index < len(values) else None

            row = {
                key: cell(header)
                for key, header in REVIEW_FIELD_HEADERS.items()
            }
            row["trailer"] = values[trailer_column - 1] if trailer_column <= len(values) else None
            row["settlement_tonnage"] = None
            row["source_id"] = ""
            row["price_kind"] = None
            row["standard_route"] = None
            row["model_tonnage"] = None
            row["downgraded"] = False
            row["original_coefficient"] = None
            row["original_freight"] = None
            row["price_difference"] = None

            plate = normalize_plate(row.get("plate"))
            stamp = _date_key(row.get("seal_time"))
            dispatch = normalize_text(row.get("dispatch_no"))
            entry = None
            if plate and stamp:
                route = row.get("route")
                if dispatch:
                    entry = pick(metadata_by_key.get((plate, stamp, dispatch), []), excel_row, route)
                else:
                    entry = pick(metadata_blank.get((plate, stamp), []), excel_row, route)
            if entry is None:
                candidate = metadata.get(excel_row)
                if candidate and (
                    id(candidate) in used_metadata
                    or (candidate.get("plate") and normalize_plate(candidate.get("plate")) != plate)
                    or normalize_text(candidate.get("dispatch_no")) != dispatch
                ):
                    candidate = None
                entry = candidate
            if entry is None:
                metadata_missing += 1
            else:
                used_metadata.add(id(entry))
                row.update(entry)
            rows.append(row)

        return rows, {
            "layout": layout,
            "sheet_names": sheet_names,
            "metadata_missing": metadata_missing,
            "metadata_unused": max(0, len(metadata) - len(used_metadata)),
        }
    finally:
        workbook.close()


def build_split_report(
    review_path: Path | str,
    output: Path | str,
    *,
    month: str,
    first_price: Path | str,
    second_price: Path | str,
    split_template: Path | str,
) -> dict[str, Any]:
    rows, info = read_review_rows(review_path)
    if not rows:
        raise ValueError("基础报表初版没有可拆分的数据行")
    aliases = _load_json(Path(__file__).resolve().parents[1] / "assets" / "route-aliases.json")
    catalogs = {
        "first": read_price_catalog(first_price, "first"),
        "second": read_price_catalog(second_price, "second"),
    }
    _apply_reviewed_classification(rows, catalogs, aliases)
    layout = discover_split_template(split_template, aliases)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    output, counts = write_split_workbook(rows, split_template, layout, output)
    total = sum(counts.values())
    if total != len(rows):
        output.unlink(missing_ok=True)
        raise ValueError(f"拆分版记录数不守恒：初版 {len(rows)} 行，拆分版 {total} 行")
    try:
        year, month_number = map(int, month.split("-"))
    except ValueError as exc:
        raise ValueError("month 必须使用 YYYY-MM") from exc
    final = output.with_name(f"{month_number}月_03_基础报表拆分版.xlsx")
    if final != output:
        output.replace(final)
    total_sheets = 0
    check = load_workbook(final, read_only=True, keep_links=False)
    try:
        total_sheets = len(check.sheetnames)
    finally:
        check.close()
    return {
        "output": final,
        "rows": len(rows),
        "sheets": counts,
        "total_sheets": total_sheets,
        "unmatched": counts.get("※未匹配目的地", 0),
        "metadata_missing": info["metadata_missing"],
        "metadata_unused": info["metadata_unused"],
        "month": f"{year:04d}-{month_number:02d}",
        "issue_rows": sum(1 for row in rows if row.get("price_kind") is None),
    }
