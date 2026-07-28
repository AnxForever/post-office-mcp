from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Any, Mapping

from inspect_inputs import inspect_inputs
from reconciliation.calculations import apply_cross_tonnage_flags
from reconciliation.excel_io import (
    file_sha256,
    load_company_map,
    read_online_rows,
    read_price_catalog,
    read_vehicle_rows,
)
from reconciliation.matching import (
    Issue,
    apply_settlement_tonnage,
    index_model_summary,
    match_price,
    match_vehicle_row,
)
from reconciliation.integrity import PHASE_ONE_ASSETS, verify_state_files
from reconciliation.model_summary import merge_online_rows, write_model_summary_workbook
from reconciliation.normalize import normalize_plate, normalize_text
from reconciliation.templates import write_review_workbook
from reconciliation.template_discovery import discover_draft_template


_DYNAMIC_WARNING_PREFIXES = (
    "在线记录数与车型表差距为 ",
    "价格匹配率为 ",
    "车管匹配率为 ",
)

IMAGE_FORMULA = re.compile(r"^\s*=\s*(?:_xlfn\.)?dispimg\s*\(", re.IGNORECASE)
NUMERIC_MODEL_PROBLEM = re.compile(r"^\d+(?:\.0+)?$")
TRAILER_LIKE_MODEL_PROBLEM = re.compile(r"^[\u4e00-\u9fff][A-Z0-9]{2,}挂$", re.IGNORECASE)


def _load_json(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    value = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(value, dict) and isinstance(value.get("aliases"), dict):
        return value["aliases"]
    return value if isinstance(value, dict) else {}


def _save_state(path: Path, state: dict[str, Any]) -> None:
    temporary = path.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    temporary.replace(path)


def _source_entries(state: Mapping[str, Any], role: str) -> list[dict[str, Any]]:
    sources = state.get("sources", {})
    if isinstance(sources, Mapping):
        entries = sources.get(role, [])
        return [dict(item) for item in entries if isinstance(item, Mapping)]
    return [dict(item) for item in sources if isinstance(item, Mapping) and item.get("role") == role]


def _source_paths(state: Mapping[str, Any], role: str) -> list[Path]:
    return [Path(item["path"]) for item in _source_entries(state, role)]


def _vehicle_source_id(vehicle: Mapping[str, Any], source_hashes: Mapping[str, str]) -> str:
    source_path = str(Path(vehicle["source_path"]).resolve())
    source_sha256 = source_hashes.get(source_path)
    if not source_sha256:
        raise ValueError(f"车管源文件未在运行状态中锁定: {vehicle.get('source_file')}")
    return f"V:{source_sha256}:{vehicle['source_sheet']}:{vehicle['source_row']}"


def _unmatched_issue(vehicle: Mapping[str, Any]) -> Issue:
    execution_date = vehicle.get("execution_date")
    date_text = execution_date.isoformat() if isinstance(execution_date, date) else str(execution_date or "")
    plate = normalize_plate(vehicle.get("plate"))
    location = f"{vehicle.get('source_file')}/{vehicle.get('source_sheet')}/第{vehicle.get('source_row')}行"
    return Issue(
        "model_summary_missing",
        "red",
        f"未匹配车型汇总：精确匹配键 {date_text}+{plate}；来源 {location}",
        "核对该日期和车牌的在线测量记录，或人工补充容间及依赖字段",
    )


def _model_issues(model: Mapping[str, Any]) -> list[Issue]:
    note = model.get("zero_value_note")
    if note == "在线记录确认未测量":
        return [Issue("online_unmeasured", "orange", note, "人工确认是否保持0")]
    if note == "在线零值与尺寸不一致":
        return [Issue("online_zero_inconsistent", "red", note, "回查原始测量表")]
    return []


def _is_image_formula(value: Any) -> bool:
    return isinstance(value, str) and IMAGE_FORMULA.match(value) is not None


def _merge_remarks(*values: Any) -> str | None:
    result: list[str] = []
    for value in values:
        if _is_image_formula(value):
            continue
        text = normalize_text(value)
        if text and text not in result:
            result.append(text)
    return "；".join(result) or None


def _format_review_remark(
    model_problem: Any,
    model_remark: Any,
    issues: list[Issue],
) -> str | None:
    """只有有问题的行才写备注: 在线表原文 + 具体问题描述. 正常行不写."""
    if not issues:
        return None
    original = normalize_text(model_problem)
    if original and re.fullmatch(r"\d+(?:\.0+)?", original):
        n = int(float(original))
        original = f"{n}(疑似{n}轴车)"
    parts = []
    if original:
        parts.append(original)
    for issue in issues:
        text = normalize_text(issue.message)
        if text and text != original and text not in parts:
            parts.append(text)
    if model_remark:
        text = normalize_text(model_remark)
        if text and text != original and text not in parts and not _is_image_formula(model_remark):
            parts.append(text)
    return " | ".join(parts) if parts else None


def _is_current_month_vehicle_row(row: Mapping[str, Any], year: int, month: int) -> bool:
    execution_date = row.get("execution_date")
    return bool(
        isinstance(execution_date, date)
        and (execution_date.year, execution_date.month) == (year, month)
        and normalize_plate(row.get("plate"))
        and normalize_text(row.get("route"))
    )


def build_review_workbook(
    input_dir: Path | str,
    workspace_dir: Path | str,
    month: str | None = None,
    *,
    assets_dir: Path | str | None = None,
    force: bool = False,
) -> tuple[Path, Path]:
    input_dir, workspace_dir = Path(input_dir), Path(workspace_dir)
    package_dir = Path(__file__).resolve().parents[1]
    assets = Path(assets_dir) if assets_dir else package_dir / "assets"
    state_path = workspace_dir / "run-state.json"
    if not state_path.is_file():
        inspection = inspect_inputs(input_dir, workspace_dir, month, assets_dir=assets)
        if not inspection["ok"]:
            raise ValueError("输入预检失败: " + "；".join(inspection["fatal_errors"]))
    state = json.loads(state_path.read_text(encoding="utf-8"))
    integrity_errors = verify_state_files(state, assets, asset_names=PHASE_ONE_ASSETS)
    if integrity_errors:
        raise ValueError("同次运行完整性校验失败: " + "；".join(integrity_errors))
    if state.get("phase") == "awaiting_review" and state.get("review_file"):
        existing_summary = Path(state.get("model_summary_file", ""))
        existing_review = Path(state["review_file"])
        if existing_summary.is_file() and existing_review.is_file() and not force:
            return existing_summary, existing_review
        if force:
            state["phase"] = "inspected"
    if state.get("phase") != "inspected":
        raise ValueError(f"运行状态不允许生成初版: {state.get('phase')}")

    target_year, target_month = map(int, state["target_month"].split("-"))
    caotan_paths = _source_paths(state, "online_caotan")
    gangwu_paths = _source_paths(state, "online_gangwu")
    vehicle_paths = _source_paths(state, "vehicle")
    first_price_paths = _source_paths(state, "first_price")
    second_price_paths = _source_paths(state, "second_price")
    draft_template_paths = _source_paths(state, "draft_template")
    expected_vehicle_count = int(state.get("expected_role_counts", {}).get("vehicle", 6))
    if (
        len(caotan_paths) != 1
        or len(gangwu_paths) != 1
        or len(vehicle_paths) != expected_vehicle_count
        or len(first_price_paths) != 1
        or len(second_price_paths) != 1
        or len(draft_template_paths) != 1
    ):
        raise ValueError("运行状态中的输入角色数量无效，请重新预检")

    online_rows = read_online_rows(caotan_paths[0], (target_year, target_month), "草滩")
    online_rows += read_online_rows(gangwu_paths[0], (target_year, target_month), "港务")
    model_summary_rows = merge_online_rows(online_rows)
    model_index = index_model_summary(model_summary_rows)
    company_map = load_company_map(assets / "company-map.json")
    vehicle_rows = [
        row
        for row in read_vehicle_rows(vehicle_paths, company_map)
        if _is_current_month_vehicle_row(row, target_year, target_month)
    ]
    source_hashes = {
        str(Path(item["path"]).resolve()): item["sha256"]
        for item in _source_entries(state, "vehicle")
        if item.get("path") and item.get("sha256")
    }
    catalogs = {
        "first": read_price_catalog(first_price_paths[0], "first"),
        "second": read_price_catalog(second_price_paths[0], "second"),
    }
    draft_layout = discover_draft_template(draft_template_paths[0])
    aliases = _load_json(assets / "route-aliases.json")

    review_rows: list[dict[str, Any]] = []
    model_matched = 0
    for vehicle in vehicle_rows:
        model = match_vehicle_row(vehicle, model_index)
        if model is not None:
            model_matched += 1
        issues = _model_issues(model) if model is not None else [_unmatched_issue(vehicle)]
        if model is not None and model.get("_cross_station"):
            issues.append(Issue("cross_station", "blue", "串行：同车牌跨站点/跨日期取容积", "核对容积是否适用"))
        volume = model.get("measured_volume") if model is not None else None
        # Even when the vehicle has no model-summary match, resolve the route
        # with a zero volume. This preserves the official price-table distance
        # and route type in the issue row while leaving tonnage/freight blank.
        price = match_price(
            vehicle.get("route"),
            volume if model is not None else 0,
            catalogs,
            aliases,
        )
        if price.kind in catalogs:
            if model is not None:
                price = apply_settlement_tonnage(
                    price,
                    vehicle.get("settlement_tonnage"),
                    catalogs[price.kind],
                )
        issues.extend(price.issues)
        source_id = _vehicle_source_id(vehicle, source_hashes)
        review_rows.append({
            "source_id": source_id,
            "source_path": vehicle.get("source_path"),
            "source_file": vehicle.get("source_file"),
            "source_sheet": vehicle.get("source_sheet"),
            "source_row": vehicle["source_row"],
            "company": vehicle.get("company"),
            "dispatch_no": vehicle.get("dispatch_no"),
            "route": vehicle.get("route"),
            "seal_time": vehicle.get("execution_date"),
            "plate": vehicle.get("plate"),
            "trailer": vehicle.get("trailer"),
            "plate_key": vehicle.get("plate_key"),
            "trailer_key": vehicle.get("trailer_key"),
            "volume": volume,
            "tonnage": price.tonnage if price is not None else None,
            "model_tonnage": price.model_tonnage if price is not None else None,
            "settlement_tonnage": vehicle.get("settlement_tonnage"),
            "distance": price.distance if price is not None else None,
            "transport_volume": vehicle.get("transport_volume"),
            "freight": price.freight if price is not None else None,
            "standard_price": price.standard_price if price is not None else None,
            "coefficient": price.coefficient if price is not None else None,
            "price_kind": price.kind if price is not None else None,
            "standard_route": price.route if price is not None else None,
            "downgraded": bool(price and price.downgraded),
            "original_coefficient": price.original_coefficient if price is not None else None,
            "original_freight": price.original_freight if price is not None else None,
            "price_difference": price.price_difference if price is not None else None,
            "remark": _format_review_remark(
                model.get("model_problem") if model is not None else None,
                model.get("remark") if model is not None else None,
                issues,
            ),
            "issues": issues,
        })
    apply_cross_tonnage_flags(review_rows)

    model_summary_output = workspace_dir / f"{target_month}月_01_车型汇总表.xlsx"
    output = workspace_dir / f"{target_month}月_02_基础报表初版.xlsx"
    write_model_summary_workbook(model_summary_rows, model_summary_output, (target_year, target_month))
    write_review_workbook(
        review_rows,
        draft_template_paths[0],
        output,
        (target_year, target_month),
        structure={"sources": state.get("sources", []), "catalogs": catalogs},
    )
    issue_counter = Counter(issue.code for row in review_rows for issue in row.get("issues", []))
    price_matched = sum(1 for row in review_rows if row.get("standard_price") is not None)
    problem_count = sum(len(row.get("issues", [])) for row in review_rows)
    nonzero_selections = sum(
        row.get("measured_volume") not in (None, "", 0, 0.0)
        for row in model_summary_rows
    )
    count = len(review_rows)
    warnings = [
        warning for warning in state.get("warnings", [])
        if not str(warning).startswith(_DYNAMIC_WARNING_PREFIXES)
    ]
    online_gap = abs(len(model_summary_rows) - count) / count if count else 0
    if online_gap >= 0.10:
        warnings.append(f"在线记录数与车型表差距为 {online_gap:.1%}")
    if count and price_matched / count < 0.90:
        warnings.append(f"价格匹配率为 {price_matched / count:.1%}，低于90%")
    warnings = list(dict.fromkeys(warnings))
    state.update({
        "phase": "awaiting_review",
        "status": "awaiting_review",
        "model_summary_file": str(model_summary_output.resolve()),
        "model_summary_sha256": file_sha256(model_summary_output),
        "review_file": str(output.resolve()),
        "review_sha256": file_sha256(output),
        "output_paths": {
            "model_summary": str(model_summary_output.resolve()),
            "review": str(output.resolve()),
        },
        "output_hashes": {
            model_summary_output.name: file_sha256(model_summary_output),
            output.name: file_sha256(output),
        },
        "template_layouts": {
            **state.get("template_layouts", {}),
            "draft": draft_layout,
        },
        "row_count": count,
        "row_ids": [
            {
                "source_id": row["source_id"],
                "source_file_sha256": row["source_id"].split(":", 2)[1],
                "source_sheet": vehicle["source_sheet"],
                "source_row": vehicle["source_row"],
                "company": row["company"],
                "dispatch_no": row["dispatch_no"],
                "execution_date": row["seal_time"],
                "route": row["route"],
                "plate": row["plate"],
                "trailer": row["trailer"],
                "settlement_tonnage": vehicle.get("settlement_tonnage"),
            }
            for row, vehicle in zip(review_rows, vehicle_rows)
        ],
        "metrics": {
            "model_rows": len(model_summary_rows),
            "online_rows": len(online_rows),
            "online_matched": model_matched,
            "vehicle_matched": model_matched,
            "price_matched": price_matched,
            "model_summary_candidates": len(online_rows),
            "model_summary_unique_keys": len(model_summary_rows),
            "model_summary_nonzero_selections": nonzero_selections,
            "model_summary_all_zero": len(model_summary_rows) - nonzero_selections,
            "authoritative_vehicle_rows": count,
            "model_summary_matched": model_matched,
            "model_summary_unmatched": count - model_matched,
            "price_unmatched": count - price_matched,
            "downgrade_count": sum(bool(row.get("downgraded")) for row in review_rows),
            "problem_count": problem_count,
        },
        "issue_counts": dict(issue_counter),
        "warnings": warnings,
    })
    _save_state(state_path, state)
    return model_summary_output, output


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成月度基础报表初版并进入人工复核闸门")
    parser.add_argument("--input-dir", required=True, type=Path)
    parser.add_argument("--workspace-dir", required=True, type=Path)
    parser.add_argument("--month", help="目标月份 YYYY-MM")
    parser.add_argument("--force", action="store_true", help="重新生成已有初版；会覆盖尚未提交的人工修改")
    args = parser.parse_args(argv)
    try:
        model_summary, review = build_review_workbook(
            args.input_dir, args.workspace_dir, args.month, force=args.force
        )
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        return 2
    state = json.loads((args.workspace_dir / "run-state.json").read_text(encoding="utf-8"))
    print(json.dumps({
        "ok": True,
        "phase": state["phase"],
        "outputs": [str(model_summary), str(review)],
        "metrics": state.get("metrics"),
        "issues": state.get("issue_counts"),
        "warnings": state.get("warnings"),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
