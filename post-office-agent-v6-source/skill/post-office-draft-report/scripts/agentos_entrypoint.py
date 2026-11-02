# -*- coding: utf-8 -*-
"""基础报表初版技能入口。

输入：01 车型汇总表 + 6 家车管数据（或 1 份车管平台汇总）+ 一干/二干价格表。
输出：N月_02_基础报表初版.xlsx（21 列 + 问题汇总 + 拆分元数据）。

    python agentos_entrypoint.py --workspace /workspace
    python agentos_entrypoint.py --workspace /workspace --month 2026-09
    python agentos_entrypoint.py --model-file 01.xlsx --vehicle a.xlsx ... \
        --first-price 一干.xlsx --second-price 二干.xlsx --output-dir out --month 2026-09

不依赖 run-state，可以单独任务运行；01 车型汇总表是上一阶段交付件。
"""
from __future__ import annotations

from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import sys
from collections import Counter
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from build_review_workbook import _is_current_month_vehicle_row, build_review_rows
from reconciliation.detect import describe_input_gaps, detect_for_phase, needs_input_message, detect_workspace
from reconciliation.excel_io import (
    file_sha256,
    load_company_map,
    read_price_catalog,
    read_vehicle_platform_rows,
    read_vehicle_rows,
)
from reconciliation.matching import _date_key
from reconciliation.normalize import normalize_plate, normalize_text
from reconciliation.paths import publish_files, resolve_input_dir
from reconciliation.template_discovery import discover_draft_template
from reconciliation.templates import write_review_workbook

MODEL_SHEET = "车型汇总"
META_SHEET = "拆分元数据"
META_HEADERS = [
    "行号", "邮路名称", "车牌号", "派车单号", "封车时间", "干线归属", "标准邮路",
    "车型吨位", "结算吨位", "是否降吨位", "原运费", "差价", "来源ID",
]
REVIEW_FIELD_HEADERS = {
    "volume": "容间",
    "tonnage": "吨位",
    "distance": "公里",
    "freight": "运价",
    "standard_price": "标准价",
    "remark": "备注",
    "cross_tonnage": "是否跨档",
    "previous_volume": "变动前容间",
    "previous_time": "变动前时间",
    "coefficient": "系数",
}


def _load_json(path: Path) -> dict[str, Any]:
    """读取资产 JSON；别名文件外层带 asset_version，只取 aliases 映射。"""
    value = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(value, dict) and isinstance(value.get("aliases"), dict):
        return value["aliases"]
    return value if isinstance(value, dict) else {}


def _number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        number = float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None
    # 01 表里的容积常写成整数；读回后保持整数，问题清单里的文案才和单块流程一致
    # （139 而不是 139.0）。
    return int(number) if number.is_integer() else number


def _date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()[:10].replace("/", "-")
        try:
            return date.fromisoformat(text)
        except ValueError:
            return None
    return None


def read_model_summary(path: Path | str) -> list[dict[str, Any]]:
    """把 01 车型汇总表读回成 review 阶段用的 model 行。"""
    workbook = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    try:
        if MODEL_SHEET not in workbook.sheetnames:
            raise ValueError(f"车型汇总表缺少 Sheet「{MODEL_SHEET}」")
        sheet = workbook[MODEL_SHEET]
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, ()) or ()
        columns = {normalize_text(value): index for index, value in enumerate(header) if value is not None}

        def get(row: tuple, name: str) -> Any:
            index = columns.get(normalize_text(name))
            return row[index] if index is not None and index < len(row) else None

        result: list[dict[str, Any]] = []
        for row in rows:
            measurement_date = _date(get(row, "测量日期"))
            plate = get(row, "车牌号")
            if not measurement_date or not plate:
                continue
            raw_problem = str(get(row, "车型问题") or "")
            raw_remark = str(get(row, "备注") or "")
            note = ""
            for candidate in ("在线记录确认未测量", "在线零值与尺寸不一致"):
                if candidate in raw_problem or candidate in raw_remark:
                    note = candidate
                    break
            problem = raw_problem.replace(f" | {note}", "").replace(note, "").strip(" |;")
            result.append({
                "measurement_date": measurement_date,
                "site": get(row, "场地"),
                "route": get(row, "邮路"),
                "plate": plate,
                "trailer": get(row, "挂车号"),
                "length": _number(get(row, "长")),
                "width": _number(get(row, "宽")),
                "height": _number(get(row, "高")),
                "measured_volume": _number(get(row, "测量容积")),
                "model_problem": problem or None,
                "remark": raw_remark or None,
                "zero_value_note": note or None,
                "candidate_count": int(_number(get(row, "候选记录数")) or 1),
                "source_file": get(row, "来源文件") or "",
                "source_sheet": get(row, "来源Sheet") or "",
                "source_row": get(row, "来源行") or "",
                "source_path": "",
            })
        return result
    finally:
        workbook.close()


def infer_month_from_model(path: Path | str) -> dict:
    """从 01 的测量日期推断月份。"""
    rows = read_model_summary(path)
    counter: Counter[tuple[int, int]] = Counter()
    for row in rows:
        stamp = _date_key(row.get("measurement_date"))
        if stamp:
            counter[(stamp.year, stamp.month)] += 1
    if not counter:
        return {"month": None, "questions": ["车型汇总表里没有可用的测量日期，请告诉我是哪个月"]}
    if len(counter) > 1:
        return {
            "month": None,
            "questions": [
                "车型汇总表里出现了多个月份（"
                + "、".join(f"{y}年{m}月" for y, m in sorted(counter))
                + "），请确认要处理的是哪个月"
            ],
        }
    (year, month), _ = counter.most_common(1)[0]
    return {"month": f"{year:04d}-{month:02d}", "evidence": {"测量日期": f"{year}年{month}月"}}


def _resolve_one(result: dict, role: str) -> Path | None:
    paths = result.get("resolved", {}).get(role, [])
    return Path(paths[0]) if len(paths) == 1 else None


def write_metadata(path: Path | str, rows: list[dict[str, Any]], template_layout: dict) -> None:
    workbook = load_workbook(path, keep_links=False)
    try:
        if META_SHEET in workbook.sheetnames:
            del workbook[META_SHEET]
        sheet = workbook.create_sheet(META_SHEET)
        sheet.append(META_HEADERS)
        data_start = int(template_layout.get("data_start_row", 2) or 2)
        for index, row in enumerate(rows, 1):
            sheet.append([
                data_start + index - 1,
                row.get("route"),
                row.get("plate"),
                row.get("dispatch_no"),
                row.get("seal_time"),
                row.get("price_kind") or "",
                row.get("standard_route") or "",
                row.get("model_tonnage") or "",
                row.get("settlement_tonnage") or "",
                1 if row.get("downgraded") else 0,
                row.get("original_freight"),
                row.get("price_difference"),
                row.get("source_id") or "",
            ])
        for cell in sheet[1]:
            font = copy(cell.font)
            font.bold = True
            cell.font = font
        for column, width in enumerate([8, 24, 14, 16, 18, 10, 24, 10, 12, 12, 12, 12, 28], 1):
            sheet.column_dimensions[sheet.cell(1, column).column_letter].width = width
        sheet.freeze_panes = "A2"
        workbook.save(path)
    finally:
        workbook.close()


def _emit(payload: dict) -> int:
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload.get("ok", True) else 2


PREVIOUS_OVERLAY_FIELDS = (
    "volume", "tonnage", "coefficient", "distance", "freight", "standard_price",
    "remark", "cross_tonnage", "previous_volume", "previous_time",
)


def _previous_review_index(path: Path | str) -> dict[tuple, list[dict[str, Any]]]:
    """读上一版人工核对过的 02，按（车牌, 日期, 派车单号）建索引。

    用于用户反复修正的场景：重新生成 02 时，把上一版改过的容间/吨位/运费等
    带回来，避免一次重跑就把人工核对结果抹掉。只认 02 的 21 列业务表，
    不依赖「拆分元数据」。

    同一辆车同一天同一张派车单可能有两趟（去程/回程），所以键后面挂列表，
    合并时再按邮路名称区分，不能后写覆盖前写。
    """
    layout = discover_draft_template(path)
    workbook = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    try:
        sheet = workbook["基础报表"]
        index: dict[tuple, list[dict[str, Any]]] = {}
        for values in sheet.iter_rows(min_row=int(layout["data_start_row"]), values_only=True):
            if not any(value not in (None, "") for value in values):
                continue

            def cell(header_name: str) -> Any:
                position = int(layout["headers"][header_name]) - 1
                return values[position] if position < len(values) else None

            stamp = _date_key(cell("封车时间"))
            plate = normalize_plate(cell("车牌号"))
            dispatch = normalize_text(cell("派车单号"))
            if not plate or not stamp:
                continue
            entry = {key: cell(header) for key, header in REVIEW_FIELD_HEADERS.items()}
            index.setdefault((plate, stamp, dispatch), []).append(entry)
            index.setdefault((plate, stamp, ""), []).append(entry)
        return index
    finally:
        workbook.close()


def _merge_previous_review(rows: list[dict[str, Any]], previous: dict[tuple, list[dict[str, Any]]]) -> int:
    if not previous:
        return 0
    used: set[int] = set()

    def pick(candidates: list[dict[str, Any]], route: Any) -> dict[str, Any] | None:
        if not candidates:
            return None
        available = [item for item in candidates if id(item) not in used]
        if not available:
            available = candidates
        route_key = normalize_text(route)
        same_route = [item for item in available if normalize_text(item.get("route")) == route_key]
        return (same_route or available)[0]

    merged = 0
    for row in rows:
        stamp = _date_key(row.get("seal_time"))
        plate = normalize_plate(row.get("plate"))
        dispatch = normalize_text(row.get("dispatch_no"))
        if not plate or not stamp:
            continue
        entry = pick(previous.get((plate, stamp, dispatch), []), row.get("route"))
        if not entry:
            entry = pick(previous.get((plate, stamp, ""), []), row.get("route"))
        if not entry:
            continue
        used.add(id(entry))
        merged += 1
        for field in PREVIOUS_OVERLAY_FIELDS:
            value = entry.get(field)
            if value not in (None, ""):
                row[field] = value
        row["previous_review_merged"] = True
    return merged


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成基础报表初版（02）")
    parser.add_argument("--workspace", type=Path, default=Path("/workspace"))
    parser.add_argument("--input-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--month", help="YYYY-MM，不传则从 01 的测量日期推断")
    parser.add_argument("--model-file", type=Path)
    parser.add_argument("--vehicle", type=Path, nargs="*")
    parser.add_argument("--vehicle-platform", type=Path)
    parser.add_argument("--first-price", type=Path)
    parser.add_argument("--second-price", type=Path)
    parser.add_argument("--draft-template", type=Path)
    parser.add_argument("--base-report", type=Path)
    parser.add_argument("--previous-review", type=Path, help="上一版人工核对过的 02，用于带回人工修改")
    parser.add_argument("--force", action="store_true", help="确认覆盖输出目录里已有的 02")
    args = parser.parse_args(argv)

    workspace = args.workspace.resolve()
    input_dir = args.input_dir.resolve() if args.input_dir else resolve_input_dir(workspace)
    output_dir = args.output_dir or (workspace / "output")
    assets = Path(__file__).resolve().parents[1] / "assets"

    model_file = args.model_file
    if not model_file:
        detected = detect_workspace(input_dir, args.month)
        models = [item for item in detected if item.role == "model_summary"]
        if len(models) == 1:
            model_file = models[0].path
        elif len(models) > 1:
            return _emit({
                "ok": True,
                "status": "needs_input",
                "phase": "awaiting_inputs",
                "missing_inputs": [{
                    "kind": "ambiguous",
                    "role": "model_summary",
                    "label": "车型汇总表",
                    "found": len(models),
                    "candidates": [item.path.name for item in models],
                    "action": "识别到多份车型汇总表，请只保留当月的一份，或显式指定 --model-file",
                }],
                "message": "识别到多份车型汇总表，无法确定用哪一份。",
                "next_action": "请只保留当月一份车型汇总表后重试。",
            })
        else:
            return _emit({
                "ok": True,
                "status": "needs_input",
                "phase": "awaiting_inputs",
                "missing_inputs": [{
                    "kind": "missing",
                    "role": "model_summary",
                    "label": "车型汇总表",
                    "action": "请上传上一阶段生成的 N月_01_车型汇总表.xlsx",
                }],
                "message": "当前上传的文件里没有车型汇总表。",
                "next_action": "请先跑车型汇总技能，或上传 01 车型汇总表后重试。",
            })

    month = args.month
    if not month:
        inference = infer_month_from_model(model_file)
        if not inference.get("month"):
            return _emit({
                "ok": True,
                "status": "needs_month",
                "phase": "needs_month",
                "questions": inference.get("questions", []),
                "next_action": "请把疑问转告用户，拿到月份后带 --month 重试。",
            })
        month = inference["month"]
    try:
        year, month_number = map(int, month.split("-"))
    except ValueError:
        return _emit({"ok": False, "errors": ["month 必须使用 YYYY-MM"], "next_action": "请修正月份"})

    result = detect_for_phase(input_dir, month, "draft")
    explicit_roles = set()
    if args.model_file:
        explicit_roles.add("model_summary")
    if args.vehicle:
        explicit_roles.add("vehicle")
    if args.vehicle_platform:
        explicit_roles.add("vehicle_platform")
    if args.first_price:
        explicit_roles.add("first_price")
    if args.second_price:
        explicit_roles.add("second_price")
    if args.draft_template:
        explicit_roles.add("draft_template")
    if args.base_report:
        explicit_roles.add("base_report")
    gaps = [gap for gap in describe_input_gaps(result) if gap.get("role") not in explicit_roles]
    if gaps:
        return _emit({
            "ok": True,
            "status": "needs_input",
            "phase": "awaiting_inputs",
            "target_month": month,
            "detected": result.get("files", []),
            "missing": result.get("missing", []),
            "ambiguous": result.get("ambiguous", []),
            "missing_inputs": gaps,
            "message": needs_input_message(gaps),
            "next_action": "请补充或确认输入文件后重试。",
        })

    if not model_file:
        model_file = _resolve_one(result, "model_summary")
    vehicle_paths = args.vehicle or []
    platform_path = args.vehicle_platform
    if not vehicle_paths and not platform_path:
        vehicle_paths = [Path(item) for item in result.get("resolved", {}).get("vehicle", [])]
        platform_path = _resolve_one(result, "vehicle_platform")
    first_price = args.first_price or _resolve_one(result, "first_price")
    second_price = args.second_price or _resolve_one(result, "second_price")
    draft_template = args.draft_template or _resolve_one(result, "draft_template")
    base_report = args.base_report or _resolve_one(result, "base_report")

    if not model_file or not first_price or not second_price or not draft_template:
        return _emit({
            "ok": False,
            "errors": ["缺少车型汇总表、价格表或初版模板"],
            "next_action": "请确认上述文件都已上传或显式指定",
        })
    if not vehicle_paths and not platform_path:
        return _emit({
            "ok": False,
            "errors": ["缺少车管数据（6 家公司）或车管平台汇总"],
            "next_action": "请上传 6 家车管文件或 1 份车管平台汇总后重试",
        })

    model_rows = read_model_summary(model_file)
    if not model_rows:
        return _emit({"ok": False, "errors": ["车型汇总表没有有效记录"], "next_action": "请检查 01 文件"})
    source_bundle = "merged" if platform_path else ("merged" if any(
        normalize_text(row.get("site")) == "合并" for row in model_rows
    ) else "split")
    if source_bundle == "merged" and not platform_path:
        return _emit({
            "ok": False,
            "errors": ["车型汇总表来自合并口径，但缺少车管平台汇总文件"],
            "next_action": "请上传车管平台汇总导出（vehicle_platform）后重试",
        })

    if source_bundle == "merged":
        vehicle_rows = [
            row for row in read_vehicle_platform_rows(platform_path)
            if _is_current_month_vehicle_row(row, year, month_number)
        ]
        vehicle_files = [platform_path]
    else:
        vehicle_rows = [
            row for row in read_vehicle_rows(vehicle_paths, load_company_map(assets / "company-map.json"))
            if _is_current_month_vehicle_row(row, year, month_number)
        ]
        vehicle_files = vehicle_paths
    if not vehicle_rows:
        return _emit({
            "ok": True,
            "status": "needs_input",
            "phase": "awaiting_inputs",
            "missing_inputs": [{
                "kind": "month_mismatch",
                "role": "vehicle",
                "label": "车管数据",
                "action": f"车管数据里没有 {month} 的派车记录，请确认月份或重新上传当月车管文件",
            }],
            "message": f"车管数据不包含 {month} 的派车记录。",
            "next_action": "请确认月份或上传当月车管数据后重试。",
        })

    aliases = _load_json(assets / "route-aliases.json")
    catalogs = {
        "first": read_price_catalog(first_price, "first"),
        "second": read_price_catalog(second_price, "second"),
    }
    source_hashes = {str(Path(path).resolve()): file_sha256(path) for path in vehicle_files}
    rows, model_matched, comparison_stats = build_review_rows(
        model_rows,
        vehicle_rows,
        catalogs,
        aliases,
        source_hashes,
        source_bundle=source_bundle,
        base_report_path=base_report,
    )
    if not rows:
        return _emit({"ok": False, "errors": ["没有生成任何初版行"], "next_action": "请检查输入文件"})

    layout = discover_draft_template(draft_template)
    output = Path(output_dir) / f"{month_number}月_02_基础报表初版.xlsx"
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists() and not args.force:
        return _emit({
            "ok": True,
            "status": "output_exists",
            "phase": "needs_confirm",
            "output": str(output.resolve()),
            "message": (
                "输出目录里已经有这个月的 02 初版，可能是上次人工核对过的版本。"
                "直接重新生成会覆盖人工修改。"
            ),
            "next_action": (
                "确认要覆盖时带 --force 重跑；"
                "如果要把上一版的人工修改带回来，再带 --previous-review <旧02>；"
                "只想保留旧版，请把输出目录换成一个新目录。"
            ),
        })
    previous_merged = 0
    if args.previous_review:
        if not Path(args.previous_review).is_file():
            return _emit({
                "ok": False,
                "errors": [f"上一版 02 不存在: {args.previous_review}"],
                "next_action": "请检查 --previous-review 路径",
            })
        previous = _previous_review_index(args.previous_review)
        previous_merged = _merge_previous_review(rows, previous)
    structure = {
        "sources": [
            {"role": "model_summary", "name": Path(model_file).name, "path": str(model_file)},
            *[{"role": "vehicle", "name": Path(path).name, "path": str(path)} for path in vehicle_files],
            {"role": "first_price", "name": Path(first_price).name, "path": str(first_price)},
            {"role": "second_price", "name": Path(second_price).name, "path": str(second_price)},
        ],
        "catalogs": catalogs,
    }
    write_review_workbook(rows, draft_template, output, (year, month_number), structure=structure)
    write_metadata(output, rows, layout)

    price_matched = sum(1 for row in rows if row.get("standard_price") is not None)
    issue_counts = Counter(issue.code for row in rows for issue in row.get("issues", []))
    warnings = []
    if price_matched / len(rows) < 0.90:
        warnings.append(f"价格匹配率为 {price_matched / len(rows):.1%}，低于90%")
    return _emit({
        "ok": True,
        "status": "draft_completed",
        "phase": "draft_completed",
        "target_month": month,
        "outputs": [str(output.resolve())],
        "published": publish_files([output]),
        "metrics": {
            "model_rows": len(model_rows),
            "vehicle_rows": len(vehicle_rows),
            "draft_rows": len(rows),
            "vehicle_matched": model_matched,
            "price_matched": price_matched,
            "price_unmatched": len(rows) - price_matched,
            "downgrade_count": sum(bool(row.get("downgraded")) for row in rows),
            "previous_review_merged": previous_merged,
            "problem_count": sum(len(row.get("issues", [])) for row in rows),
            "official_comparison": comparison_stats,
        },
        "issues": dict(issue_counts),
        "warnings": warnings,
        "sha256": file_sha256(output),
        "next_action": "请人工打开 02 初版核对颜色行和「问题汇总」，修订后再跑拆分技能",
    })


if __name__ == "__main__":
    try:
        code = main()
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 0
    raise SystemExit(code)
