from __future__ import annotations

import argparse
import json
import random
import re
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from reconciliation.normalize import tonnage_key
from reconciliation.calculations import aggregate_profit
from reconciliation.excel_io import read_finance_rows, read_price_catalog
from reconciliation.integrity import file_sha256, validate_dynamic_state, verify_state_files
from reconciliation.model_summary import EVIDENCE_HEADERS, HEADERS as MODEL_SUMMARY_HEADERS
from reconciliation.profit_audit import audit_profit_workbook
from reconciliation.template_discovery import discover_draft_template, discover_profit_template, discover_split_template
from reconciliation.templates import ISSUE_HEADERS, PROFIT_SUMMARY_HEADERS, REVIEW_HEADERS


def _formula_errors(wb) -> list[str]:
    errors: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("=") and "#REF!" in cell.value.upper():
                    errors.append(f"{ws.title}!{cell.coordinate} 包含 #REF! 公式")
    return errors


def _review_formula_errors(wb, limit: int = 20) -> list[str]:
    errors: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                is_formula = cell.data_type == "f" or (
                    isinstance(cell.value, str) and cell.value.lstrip().startswith("=")
                )
                if is_formula:
                    errors.append(f"{ws.title}!{cell.coordinate} 不允许包含公式")
                    if len(errors) >= limit:
                        return errors
    return errors


def _validate_model_summary(state: dict[str, Any], errors: list[str], checks: dict[str, Any]) -> None:
    requires_summary = bool(state.get("model_summary_file"))
    if not requires_summary:
        return
    path = Path(state.get("model_summary_file", ""))
    if not path.is_file():
        errors.append("车型汇总表不存在")
        return
    expected_hash = state.get("model_summary_sha256")
    if not expected_hash:
        errors.append("运行状态缺少车型汇总表哈希")
    elif file_sha256(path) != expected_hash:
        errors.append("车型汇总表哈希不一致")
    try:
        wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except Exception as exc:
        errors.append(f"车型汇总表无法读取: {exc}")
        return
    try:
        checks["model_summary_sheets"] = wb.sheetnames
        if wb.sheetnames != ["车型汇总", "容积取证"]:
            errors.append("车型汇总表结构无效：必须依次包含车型汇总和容积取证 Sheet")
            return
        headers = [cell.value for cell in wb["车型汇总"][1]]
        if headers != MODEL_SUMMARY_HEADERS:
            errors.append("车型汇总表结构无效：表头不一致")
        evidence = wb["容积取证"]
        if [cell.value for cell in evidence[1]] != EVIDENCE_HEADERS:
            errors.append("容积取证表结构无效：表头不一致")
        summary_rows = max(0, wb["车型汇总"].max_row - 1)
        evidence_rows = max(0, evidence.max_row - 1)
        checks["model_summary_rows"] = summary_rows
        checks["volume_evidence_rows"] = evidence_rows
        if summary_rows != evidence_rows:
            errors.append("容积取证记录数与车型汇总记录数不一致")
        errors.extend(_review_formula_errors(wb))
    finally:
        wb.close()


def _validate_review(state: dict[str, Any], errors: list[str], checks: dict[str, Any]) -> None:
    path = Path(state.get("review_file", ""))
    if not path.is_file():
        errors.append("基础报表初版不存在")
        return
    expected_hash = state.get("review_sha256")
    if not expected_hash:
        errors.append("运行状态缺少初版文件哈希")
    elif file_sha256(path) != expected_hash:
        errors.append("初版文件哈希不一致")
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        checks["review_sheets"] = wb.sheetnames
        if wb.sheetnames[:2] != ["结构发现", "问题汇总"] or "基础报表" not in wb.sheetnames:
            errors.append("初版前两页必须依次为结构发现、问题汇总，且必须包含基础报表")
            if "结构发现" not in wb.sheetnames:
                errors.append("初版缺少结构发现 Sheet")
            if "问题汇总" not in wb.sheetnames:
                errors.append("初版缺少问题汇总 Sheet")
        if "问题汇总" in wb.sheetnames:
            issue_ws = wb["问题汇总"]
            headers = [cell.value for cell in issue_ws[1]]
            if headers != ISSUE_HEADERS:
                errors.append("问题汇总必须严格使用已确认的10列表头")
            # 问题汇总页签不需要标记行颜色
        if "基础报表" in wb.sheetnames:
            ws = wb["基础报表"]
            layout = state.get("template_layouts", {}).get("draft", {})
            header_row = int(layout.get("header_row", 0) or 0)
            data_start_row = int(layout.get("data_start_row", 0) or 0)
            header_columns = layout.get("headers", {})
            checks["review_header_row"] = header_row
            if header_row < 1 or data_start_row != header_row + 1 or not isinstance(header_columns, dict):
                errors.append("运行状态中的draft动态模板布局无效")
                review_rows = max(0, ws.max_row - max(header_row, 1))
            else:
                wrong_headers = [
                    header
                    for header, column in header_columns.items()
                    if ws.cell(header_row, int(column)).value != header
                ]
                if wrong_headers:
                    errors.append("基础报表表头与上传的初版模板发现布局不一致")
                review_rows = max(0, ws.max_row - data_start_row + 1)
            checks["review_rows"] = review_rows
            if review_rows != state.get("row_count"):
                errors.append("基础报表有效记录数与目标月车管权威行数不一致")
        errors.extend(_review_formula_errors(wb))
    finally:
        wb.close()


def _validate_split(path: Path, state: dict[str, Any], errors: list[str], checks: dict[str, Any]) -> None:
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        checks["split_sheets"] = wb.sheetnames
        raw_counts = {ws.title: max(0, ws.max_row - 2) for ws in wb.worksheets}
        writable_sheets: list[str] = []
        assets_dir = Path(state.get("assets_dir", ""))
        template_path = assets_dir / "split-report-template.xlsx"
        if template_path.is_file():
            layout = discover_split_template(template_path, {})
            template_sheets = layout["sheets"]
            writable_sheets.extend(layout["writable_sheets"])
            checks["template_sheet_count"] = len(template_sheets)
            if wb.sheetnames[:len(template_sheets)] != template_sheets:
                errors.append("拆分版缺少原模板Sheet或原模板Sheet顺序发生变化")

        try:
            output_layout = discover_split_template(path, {})
        except ValueError:
            output_layout = {"writable_sheets": []}
        for name in output_layout["writable_sheets"]:
            if name not in writable_sheets:
                writable_sheets.append(name)

        expected_counts = state.get("expected_split_sheet_counts", {})
        if not isinstance(expected_counts, dict):
            expected_counts = {}
        for name in expected_counts:
            if name not in writable_sheets:
                writable_sheets.append(name)

        actual_counts = {name: raw_counts.get(name, 0) for name in writable_sheets}
        data_rows = sum(actual_counts.values())
        checks["split_rows"] = data_rows
        if expected_counts:
            per_sheet: dict[str, dict[str, Any]] = {}
            for name in writable_sheets:
                expected = expected_counts.get(name, 0)
                actual = actual_counts.get(name, 0)
                per_sheet[name] = {"expected": expected, "actual": actual, "ok": expected == actual}
            expected_total = sum(expected_counts.values())
            checks["split_counts"] = {
                "expected_total": expected_total,
                "actual_total": data_rows,
                "per_sheet": per_sheet,
            }
            for name, info in per_sheet.items():
                if not info["ok"]:
                    errors.append(f"拆分版页签{name}记录数不一致：期望{info['expected']}，实际{info['actual']}")
            if data_rows != expected_total:
                errors.append(f"拆分版可写页总记录数不一致：期望{expected_total}，实际{data_rows}")
            if state.get("row_count") is not None and expected_total != state.get("row_count"):
                errors.append(
                    f"拆分版预期记录总数{expected_total}与修订初版{state.get('row_count')}不一致"
                )
        else:
            if data_rows != state.get("row_count"):
                errors.append(f"拆分版记录数{data_rows}与修订初版{state.get('row_count')}不一致")
        errors.extend(_formula_errors(wb))
    finally:
        wb.close()


def _validate_profit(path: Path, state: dict[str, Any], label: str, errors: list[str], checks: dict[str, Any]) -> None:
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    try:
        expected = f"{state['target_month'][2:4]}年{int(state['target_month'][5:7])}月"
        if wb.sheetnames != [expected]:
            errors.append(f"{label}盈亏表必须且只能保留一个当月Sheet {expected}")
        try:
            layout = discover_profit_template(path)
        except ValueError as exc:
            errors.append(f"{label}盈亏表动态布局无效: {exc}")
            errors.extend(_formula_errors(wb))
            return
        ws = wb[layout["sheet"]]
        groups = layout.get("groups", [])
        starts = [int(group["start_column"]) for group in groups]
        complete = bool(groups) and all(
            int(group.get("width", 0)) == 7
            and int(group["end_column"]) == int(group["start_column"]) + 6
            for group in groups
        ) and all(right == left + 7 for left, right in zip(starts, starts[1:]))
        if not complete:
            errors.append(f"{label}盈亏表吨位区不是完整7列分组")
        else:
            tonnages = [tonnage_key(group.get("tonnage") or group.get("label")) for group in groups]
            if any(not group for group in tonnages):
                errors.append(f"{label}盈亏表存在缺少吨位标题的分组")
            checks[f"{label}_tonnage_groups"] = tonnages
            group_row = int(layout["header_rows"][0])
            summary_start = int(groups[-1]["end_column"]) + 1
            actual_summary = [ws.cell(group_row, summary_start + offset).value for offset in range(4)]
            if actual_summary != PROFIT_SUMMARY_HEADERS:
                errors.append(f"{label}盈亏表汇总表头不正确")
        errors.extend(_formula_errors(wb))
    finally:
        wb.close()


def _source_paths(state: dict[str, Any], role: str) -> list[Path]:
    sources = state.get("sources", {})
    entries = sources.get(role, []) if isinstance(sources, dict) else [
        item for item in sources if isinstance(item, dict) and item.get("role") == role
    ]
    return [Path(item.get("path", "")) for item in entries if isinstance(item, dict)]


def _validate_layout_snapshot(
    state: dict[str, Any],
    role: str,
    layout_name: str,
    discover,
    errors: list[str],
) -> Path | None:
    paths = _source_paths(state, role)
    if len(paths) != 1 or not paths[0].is_file():
        return None
    try:
        discovered = discover(paths[0])
    except ValueError as exc:
        errors.append(f"动态模板{role}无法重新发现布局: {exc}")
        return paths[0]
    locked = state.get("template_layouts", {}).get(layout_name)
    if locked != discovered:
        errors.append(f"运行状态template_layouts.{layout_name}与动态模板重新发现结果不一致")
    return paths[0]


def _recompute_profit_context(state: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], list[str]] | None:
    review_path = Path(state.get("corrected_review_file", ""))
    assets = Path(state.get("assets_dir", ""))
    finance_paths = _source_paths(state, "finance")
    first_price_paths = _source_paths(state, "first_price")
    second_price_paths = _source_paths(state, "second_price")
    required_paths = finance_paths + first_price_paths + second_price_paths
    if (
        not review_path.is_file()
        or not assets.is_dir()
        or any(len(paths) != 1 for paths in (finance_paths, first_price_paths, second_price_paths))
        or not all(path.is_file() for path in required_paths)
    ):
        return None
    from build_final_reports import _classify_finance_rows, _load_json, _recalculate_derived_fields, read_corrected_review

    rows, review_errors = read_corrected_review(review_path, state)
    if review_errors:
        return {}, {}, review_errors
    aliases = _load_json(assets / "route-aliases.json")
    catalogs = {
        "first": read_price_catalog(first_price_paths[0], "first"),
        "second": read_price_catalog(second_price_paths[0], "second"),
    }
    _recalculate_derived_fields(rows, catalogs, aliases)
    finance_rows = read_finance_rows(finance_paths[0])
    _classify_finance_rows(finance_rows, catalogs, aliases)
    first_tonnages = sorted({item["key"] for item in catalogs["first"].get("tiers", [])} | {str(row["tonnage"]) for row in rows if row.get("price_kind") == "first"})
    second_tonnages = sorted({item["key"] for item in catalogs["second"].get("tiers", [])} | {str(row["tonnage"]) for row in rows if row.get("price_kind") == "second"})
    return (
        aggregate_profit(rows, finance_rows, "first", first_tonnages),
        aggregate_profit(rows, finance_rows, "second", second_tonnages),
        [],
    )


def _validate_final_hashes(outputs: list[Path], state: dict[str, Any], errors: list[str]) -> None:
    expected = state.get("final_output_hashes", {})
    if not expected:
        errors.append("运行状态缺少最终输出哈希")
        return
    for path in outputs:
        if path.is_file() and expected.get(path.name) != file_sha256(path):
            errors.append(f"最终输出文件哈希不一致: {path.name}")


def validate_outputs(workspace_dir: Path | str, phase: str | None = None) -> dict[str, Any]:
    workspace_dir = Path(workspace_dir)
    state_path = workspace_dir / "run-state.json"
    errors: list[str] = []
    warnings: list[str] = []
    checks: dict[str, Any] = {}
    if not state_path.is_file():
        return {"ok": False, "phase": phase, "errors": ["缺少 run-state.json"], "warnings": [], "checks": {}}
    state = json.loads(state_path.read_text(encoding="utf-8"))
    phase = phase or ("review" if state.get("phase") == "awaiting_review" else "final")
    errors.extend(validate_dynamic_state(state, phase))
    assets = Path(state.get("assets_dir", ""))
    errors.extend(verify_state_files(state, assets))
    if phase == "review":
        if state.get("phase") != "awaiting_review":
            errors.append(f"状态 {state.get('phase')} 不是阶段一待复核状态")
        _validate_model_summary(state, errors, checks)
        _validate_layout_snapshot(state, "draft_template", "draft", discover_draft_template, errors)
        _validate_review(state, errors, checks)
    elif phase == "final":
        if state.get("phase") != "final_completed":
            errors.append(f"状态 {state.get('phase')} 不是最终完成状态")
        outputs = [Path(item) for item in state.get("final_outputs", [])]
        if len(outputs) != 3:
            errors.append("最终输出文件必须为3个")
        missing = [str(path) for path in outputs if not path.is_file()]
        if missing:
            errors.append("最终输出文件缺失: " + ", ".join(missing))
        _validate_final_hashes(outputs, state, errors)
        if len(outputs) >= 1 and outputs[0].is_file():
            _validate_split(outputs[0], state, errors, checks)
        if len(outputs) >= 2 and outputs[1].is_file():
            _validate_profit(outputs[1], state, "一干", errors, checks)
        if len(outputs) >= 3 and outputs[2].is_file():
            _validate_profit(outputs[2], state, "二干", errors, checks)
        context = _recompute_profit_context(state)
        if context is not None and len(outputs) >= 3 and outputs[1].is_file() and outputs[2].is_file():
            first_metrics, second_metrics, context_errors = context
            errors.extend(context_errors)
            if not context_errors:
                year, month = map(int, state["target_month"].split("-"))
                first_template = _validate_layout_snapshot(
                    state, "first_profit_template", "first_profit", discover_profit_template, errors
                )
                second_template = _validate_layout_snapshot(
                    state, "second_profit_template", "second_profit", discover_profit_template, errors
                )
                if first_template is None or second_template is None:
                    errors.append("运行状态缺少可读取的一干/二干动态盈亏模板")
                    first_audit = {"errors": [], "spot_candidates": []}
                    second_audit = {"errors": [], "spot_candidates": []}
                else:
                    first_audit = audit_profit_workbook(outputs[1], first_template, first_metrics, (year, month), "first")
                    second_audit = audit_profit_workbook(outputs[2], second_template, second_metrics, (year, month), "second")
                errors.extend(first_audit["errors"])
                errors.extend(second_audit["errors"])
                candidates = first_audit["spot_candidates"] + second_audit["spot_candidates"]
                candidates = sorted(candidates, key=lambda item: (item["kind"], item["route"], item["tonnage"]))
                if len(candidates) > 3:
                    candidates = random.Random(state["target_month"]).sample(candidates, 3)
                checks["profit_spot_checks"] = candidates
                checks["profit_value_groups_checked"] = len(first_audit["spot_candidates"]) + len(second_audit["spot_candidates"])
    else:
        errors.append("phase 只能是 review 或 final")
    warnings.extend(state.get("warnings", []))
    return {"ok": not errors, "phase": phase, "errors": errors, "warnings": warnings, "checks": checks}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="独立验证月度核对输出")
    parser.add_argument("--workspace-dir", required=True, type=Path)
    parser.add_argument("--phase", choices=("review", "final"))
    args = parser.parse_args(argv)
    result = validate_outputs(args.workspace_dir, args.phase)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
