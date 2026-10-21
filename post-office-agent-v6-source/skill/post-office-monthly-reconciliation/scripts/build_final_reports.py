from __future__ import annotations
from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import sys
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from reconciliation.calculations import aggregate_profit, apply_cross_tonnage_flags
from reconciliation.excel_io import file_sha256, read_finance_rows, read_price_catalog
from reconciliation.integrity import REQUIRED_ASSETS, validate_authoritative_rows, validate_authoritative_state, verify_state_files
from reconciliation.matching import (
    _alias_kind,
    _alias_route,
    _tonnage_number,
    apply_settlement_tonnage,
    match_price,
)
from reconciliation.normalize import normalize_plate, normalize_text, route_candidates, tonnage_key
from reconciliation.template_discovery import discover_profit_template, discover_split_template
from reconciliation.templates import REVIEW_HEADERS, write_profit_workbook, write_split_workbook


REQUIRED_REVIEW_COLUMNS = {
    "company": 2,
    "dispatch_no": 3,
    "route": 4,
    "seal_time": 6,
    "plate": 8,
    "trailer": 9,
    "volume": 10,
    "tonnage": 11,
    "distance": 12,
    "transport_volume": 13,
    "freight": 15,
    "standard_price": 16,
    "remark": 17,
    "cross_tonnage": 18,
    "previous_volume": 19,
    "previous_time": 20,
    "coefficient": 21,
}
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
RED_ISSUE_FILL = "FFFFC7CE"


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


def _state_source_entries(state: dict[str, Any], role: str) -> list[dict[str, Any]]:
    sources = state.get("sources", {})
    if isinstance(sources, dict):
        return [item for item in sources.get(role, []) if isinstance(item, dict)]
    return [item for item in sources if isinstance(item, dict) and item.get("role") == role]


def _previous_profit_input(path: Path | str | None) -> dict[str, Any]:
    if path is None:
        return {"status": "not_provided", "path": None, "sha256": None}
    resolved = Path(path).resolve()
    if not resolved.is_file():
        return {"status": "missing", "path": str(resolved), "sha256": None}
    return {"status": "present", "path": str(resolved), "sha256": file_sha256(resolved)}


def _previous_profit_inputs(
    previous_first_profit: Path | str | None,
    previous_second_profit: Path | str | None,
) -> dict[str, dict[str, Any]]:
    return {
        "first": _previous_profit_input(previous_first_profit),
        "second": _previous_profit_input(previous_second_profit),
    }


def _is_red_issue_cell(cell) -> bool:
    color = cell.fill.fgColor
    return color.type == "rgb" and str(color.rgb).upper() == RED_ISSUE_FILL


def _issue_results(ws) -> tuple[dict[int, list[str]], list[str]]:
    headers = {normalize_text(cell.value): cell.column for cell in ws[1] if cell.value is not None}
    row_col = headers.get("行号") or headers.get("基础报表行号")
    result_col = headers.get("处理结果")
    if not all((row_col, result_col)):
        raise ValueError("问题汇总缺少基础报表行号或处理结果列")
    results: dict[int, list[str]] = defaultdict(list)
    unresolved: list[str] = []
    for row, cells in enumerate(ws.iter_rows(min_row=2, values_only=False), 2):
        base_cell = cells[row_col - 1]
        base_row = base_cell.value
        if not isinstance(base_row, (int, float)):
            continue
        result = normalize_text(cells[result_col - 1].value)
        if result:
            results[int(base_row)].append(result)
    return dict(results), unresolved


def read_corrected_review(path: Path | str, state: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    wb = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    errors: list[str] = []
    try:
        current = ["结构发现", "问题汇总", "基础报表"]
        legacy = ["问题汇总", "基础报表"]
        has_current_prefix = wb.sheetnames[:len(current)] == current
        has_legacy_prefix = wb.sheetnames[:len(legacy)] == legacy
        if not (has_current_prefix or has_legacy_prefix):
            required = set(current) if "结构发现" in wb.sheetnames else set(legacy)
            if required.issubset(wb.sheetnames):
                raise ValueError("修订初版工作表顺序不正确")
            raise ValueError("修订初版必须包含问题汇总和基础报表，可包含结构发现")
        issue_results, unresolved = _issue_results(wb["问题汇总"])
        errors.extend(unresolved)
        ws = wb["基础报表"]
        draft_layout = state.get("template_layouts", {}).get("draft", {})
        header_row = int(draft_layout.get("header_row", 1) or 1)
        data_start_row = int(draft_layout.get("data_start_row", header_row + 1) or header_row + 1)
        discovered_headers = draft_layout.get("headers")
        if isinstance(discovered_headers, dict) and discovered_headers:
            header_columns = {header: int(column) for header, column in discovered_headers.items()}
            columns = {key: header_columns[header] for key, header in REVIEW_FIELD_HEADERS.items()}
            columns["trailer"] = columns["plate"] + 1
            headers = list(next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), ()))
            invalid_headers = any(
                column > len(headers) or headers[column - 1] != header
                for header, column in header_columns.items()
            )
        else:
            columns = dict(REQUIRED_REVIEW_COLUMNS)
            headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ()))
            invalid_headers = headers[:len(REVIEW_HEADERS)] != REVIEW_HEADERS
        if data_start_row != header_row + 1 or invalid_headers:
            return [], errors + ["人工修订初版基础报表结构无效：必须保留上传模板的业务表头布局"]
        state_errors = validate_authoritative_state(state)
        if state_errors:
            return [], errors + state_errors
        stable = state["row_ids"]
        actual_row_count = max(0, ws.max_row - data_start_row + 1)
        if actual_row_count != state["row_count"]:
            return [], errors + [
                f"人工修订初版行数变化：期望 {state['row_count']}，实际 {actual_row_count}"
            ]
        rows: list[dict[str, Any]] = []
        authoritative_rows: list[dict[str, Any]] = []
        for excel_row, cells in enumerate(ws.iter_rows(min_row=data_start_row, values_only=True), data_start_row):
            index = excel_row - data_start_row
            row_id = stable[index] if index < len(stable) else {}
            values = {
                key: cells[column - 1] if column <= len(cells) else None
                for key, column in columns.items()
                if key not in {"remark", "cross_tonnage", "previous_volume", "previous_time"}
            }
            authoritative_rows.append({
                "company": values["company"],
                "dispatch_no": values["dispatch_no"],
                "route": values["route"],
                "execution_date": values["seal_time"],
                "plate": values["plate"],
                "trailer": values["trailer"],
            })
            rows.append({
                "source_id": row_id["source_id"],
                **values,
                "settlement_tonnage": row_id.get("settlement_tonnage"),
                "remark": cells[columns["remark"] - 1],
                "cross_tonnage": cells[columns["cross_tonnage"] - 1],
                "previous_volume": cells[columns["previous_volume"] - 1],
                "previous_time": cells[columns["previous_time"] - 1],
                "manual_results": issue_results.get(excel_row, []),
            })
        errors.extend(validate_authoritative_rows(authoritative_rows, state))
    finally:
        wb.close()
    return rows, errors


def _split_record_count(path: Path) -> int:
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        return sum(max(0, ws.max_row - 2) for ws in wb.worksheets)
    finally:
        wb.close()


def _classify_finance_route(route: Any, catalogs: dict[str, Any], aliases: dict[str, Any]) -> tuple[str | None, str | None]:
    matches: list[tuple[str, str]] = []
    for candidate in route_candidates(route, aliases):
        # 与 match_price 一致：别名可以声明干线类型（如固原/平凉/天水/庆阳
        # 属于二干，但价格取自一干价格表），不能只看价格表里有没有这条路线。
        preferred = _alias_kind(candidate, aliases)
        if preferred in catalogs:
            return preferred, _alias_route(candidate, aliases) or candidate
        for kind, catalog in catalogs.items():
            if candidate in catalog.get("routes", {}):
                matches.append((kind, candidate))
        if matches:
            break
    kinds = {kind for kind, _ in matches}
    if len(kinds) != 1:
        return None, None
    kind = next(iter(kinds))
    standard = next(candidate for item_kind, candidate in matches if item_kind == kind)
    return kind, standard


def _recalculate_derived_fields(
    rows: list[dict[str, Any]],
    catalogs: dict[str, Any],
    aliases: dict[str, Any],
) -> None:
    for row in rows:
        row["plate_key"] = normalize_plate(row.get("plate"))
        previous_tonnage = tonnage_key(row.get("tonnage"))
        supplied_standard_price = row.get("standard_price")
        volume = row.get("volume")
        if volume is None or volume == "":
            volume = 0
        price = match_price(row.get("route"), volume, catalogs, aliases)
        row["price_kind"] = price.kind
        row["standard_route"] = price.route
        row["model_tonnage"] = price.model_tonnage
        row["downgraded"] = price.downgraded
        row["original_coefficient"] = price.original_coefficient
        row["original_freight"] = price.original_freight
        row["price_difference"] = price.price_difference
        if not price.tonnage:
            for field in ("tonnage", "coefficient", "distance", "standard_price", "freight"):
                if row.get(field) in (None, ""):
                    row[field] = None
            continue

        row["tonnage"] = price.tonnage
        row["coefficient"] = price.coefficient
        row["distance"] = price.distance
        if price.standard_price is not None:
            row["standard_price"] = price.standard_price
            row["freight"] = price.freight
        elif previous_tonnage == price.tonnage and supplied_standard_price not in (None, ""):
            row["standard_price"] = supplied_standard_price
            try:
                freight = Decimal(str(price.coefficient)) * Decimal(str(supplied_standard_price))
                row["freight"] = float(freight) if freight % 1 else int(freight)
            except Exception:
                row["freight"] = None
        else:
            row["standard_price"] = None
            row["freight"] = None
    apply_cross_tonnage_flags(rows)


def _classify_finance_rows(
    rows: list[dict[str, Any]],
    catalogs: dict[str, Any],
    aliases: dict[str, Any],
) -> dict[str, Any]:
    summary = Counter(total=len(rows))
    unclassified = Counter()
    unclassified_income = Counter()
    for row in rows:
        key = tonnage_key(row.get("tonnage"))
        try:
            valid_tonnage = Decimal(key) > 0
        except Exception:
            valid_tonnage = bool(key)
        if valid_tonnage:
            summary["valid_tonnage"] += 1
        row_kind, standard_route = _classify_finance_route(row.get("route"), catalogs, aliases)
        row["price_kind"] = row_kind
        if standard_route:
            row["route"] = standard_route
            summary[f"classified_{row_kind}"] += 1
            continue
        route = normalize_text(row.get("route"))
        summary["unclassified"] += 1
        if route.startswith("西安"):
            summary["unclassified_forward_xian"] += 1
        unclassified[route] += 1
        unclassified_income[route] += float(row.get("net_income") or 0)
    return {
        "total": summary["total"],
        "valid_tonnage": summary["valid_tonnage"],
        "classified_first": summary["classified_first"],
        "classified_second": summary["classified_second"],
        "unclassified": summary["unclassified"],
        "unclassified_forward_xian": summary["unclassified_forward_xian"],
        "top_unclassified": [
            {"route": route, "rows": count, "net_income": unclassified_income[route]}
            for route, count in unclassified.most_common(20)
        ],
    }


def build_final_reports(
    review_path: Path | str,
    workspace_dir: Path | str,
    *,
    assets_dir: Path | str | None = None,
    previous_first_profit: Path | str | None = None,
    previous_second_profit: Path | str | None = None,
) -> list[Path]:
    review_path, workspace_dir = Path(review_path), Path(workspace_dir)
    state_path = workspace_dir / "run-state.json"
    if not state_path.is_file():
        raise ValueError("缺少同一任务的 run-state.json，不能执行第二阶段")
    if not review_path.is_file():
        raise ValueError("人工修订初版不存在")
    state = json.loads(state_path.read_text(encoding="utf-8"))
    package_dir = Path(__file__).resolve().parents[1]
    assets = Path(assets_dir) if assets_dir else package_dir / "assets"
    integrity_errors = verify_state_files(state, assets, asset_names=REQUIRED_ASSETS)
    if integrity_errors:
        raise ValueError("同次运行完整性校验失败: " + "；".join(integrity_errors))
    previous_profit_inputs = _previous_profit_inputs(previous_first_profit, previous_second_profit)
    if state.get("phase") == "final_completed":
        outputs = [Path(item) for item in state.get("final_outputs", [])]
        expected_hashes = state.get("final_output_hashes", {})
        if (
            file_sha256(review_path) == state.get("corrected_review_sha256")
            and previous_profit_inputs == state.get("previous_profit_inputs")
            and len(outputs) == 3
            and all(path.is_file() for path in outputs)
            and all(file_sha256(path) == expected_hashes.get(path.name) for path in outputs)
        ):
            return outputs
    if state.get("phase") not in {"awaiting_review", "final_completed"}:
        raise ValueError(f"运行状态不允许生成最终文件: {state.get('phase')}")
    rows, errors = read_corrected_review(review_path, state)
    if errors:
        raise ValueError("第二阶段验证失败: " + "；".join(errors))
    aliases = _load_json(assets / "route-aliases.json")
    first_price_paths = [Path(item["path"]) for item in _state_source_entries(state, "first_price")]
    second_price_paths = [Path(item["path"]) for item in _state_source_entries(state, "second_price")]
    first_profit_templates = [Path(item["path"]) for item in _state_source_entries(state, "first_profit_template")]
    second_profit_templates = [Path(item["path"]) for item in _state_source_entries(state, "second_profit_template")]
    if any(len(paths) != 1 for paths in (
        first_price_paths,
        second_price_paths,
        first_profit_templates,
        second_profit_templates,
    )):
        raise ValueError("运行状态中缺少唯一的一干/二干价格表或盈亏模板，不能执行第二阶段")
    catalogs = {
        "first": read_price_catalog(first_price_paths[0], "first"),
        "second": read_price_catalog(second_price_paths[0], "second"),
    }
    _recalculate_derived_fields(rows, catalogs, aliases)

    finance_paths = [Path(item["path"]) for item in _state_source_entries(state, "finance")]
    if len(finance_paths) != 1 or not finance_paths[0].is_file():
        raise ValueError("运行状态中缺少可读取的财务收入表")
    finance_rows = read_finance_rows(finance_paths[0])
    finance_classification = _classify_finance_rows(finance_rows, catalogs, aliases)
    year, month = map(int, state["target_month"].split("-"))
    split_template = assets / "split-report-template.xlsx"
    split_layout = discover_split_template(split_template, aliases)
    split_path = workspace_dir / f"{month}月_03_基础报表拆分版.xlsx"
    first_path = workspace_dir / f"{month}月_04_一干盈亏表.xlsx"
    second_path = workspace_dir / f"{month}月_04_二干盈亏表.xlsx"
    split_path, expected_split_counts = write_split_workbook(rows, split_template, split_layout, split_path)
    if sum(expected_split_counts.values()) != len(rows):
        split_path.unlink(missing_ok=True)
        raise ValueError("拆分版记录数与修订初版不守恒")
    unmatched_split = expected_split_counts.get("※未匹配目的地", 0)
    # These tabs are generated from explicit business classifications. The
    # two-trip tab remains manual-pending when the corrected workbook contains
    # no explicit merge instruction; it must never be inferred from a plate or
    # route alone.
    special_sheet_status = {
        "二干基础": "populated" if any(row.get("price_kind") == "second" for row in rows) else "manual_pending",
        "一干汇总": "populated" if any(row.get("price_kind") == "first" for row in rows) else "manual_pending",
        "两趟合一趟": (
            "populated"
            if any(any(token in str(row.get("remark") or "") for token in ("两辆", "两趟合一趟", "合并为一辆")) for row in rows)
            else "manual_pending"
        ),
    }

    first_tonnages = sorted({item["key"] for item in catalogs["first"].get("tiers", [])} | {str(row["tonnage"]) for row in rows if row.get("price_kind") == "first"})
    second_tonnages = sorted({item["key"] for item in catalogs["second"].get("tiers", [])} | {str(row["tonnage"]) for row in rows if row.get("price_kind") == "second"})
    first_metrics = aggregate_profit(rows, finance_rows, "first", first_tonnages)
    second_metrics = aggregate_profit(rows, finance_rows, "second", second_tonnages)
    first_profit_layout = discover_profit_template(first_profit_templates[0])
    second_profit_layout = discover_profit_template(second_profit_templates[0])

    def previous_totals(path: Path | str | None, kind: str):
        label = "一干" if kind == "first" else "二干"
        if not path:
            return None, f"上月{label}盈亏表（暂无数据）"
        p = Path(path)
        if not p.is_file():
            return None, f"上月{label}盈亏表不存在: {p.name}"
        try:
            wb = load_workbook(p, read_only=False, data_only=True, keep_links=False)
        except Exception as exc:
            return None, f"上月{label}盈亏表无法读取: {exc}"
        try:
            # 找上月的 sheet（不依赖 discover_profit_template，它可能返回错误的 sheet）
            # 上月 = 目标月份 - 1
            target_month = state.get("target_month", "")
            try:
                ty, tm = map(int, target_month.split("-"))
                pm = tm - 1
                py = ty
                if pm == 0:
                    pm = 12
                    py -= 1
            except:
                pm = None

            ws = None
            if pm:
                # 尝试多种格式："3月", "26年3月", "2026年3月"
                for candidate in [f"{pm}月", f"{str(py)[2:]}年{pm}月", f"{py}年{pm}月"]:
                    for sn in wb.sheetnames:
                        if normalize_text(sn) == normalize_text(candidate):
                            ws = wb[sn]
                            break
                    if ws:
                        break
            if ws is None:
                # fallback: 用第一个有"合计"行的 sheet
                for sn in wb.sheetnames:
                    test_ws = wb[sn]
                    for row in range(1, min(test_ws.max_row + 1, 100)):
                        if normalize_text(test_ws.cell(row, 1).value) == "合计":
                            ws = test_ws
                            break
                    if ws:
                        break
            if ws is None:
                return None, f"上月{label}盈亏表找不到有效数据Sheet"

            # 找合计行
            total_row = None
            for row in range(4, ws.max_row + 1):
                if normalize_text(ws.cell(row, 1).value) == "合计":
                    total_row = row
                    break
            if total_row is None:
                return None, f"上月{label}盈亏表找不到合计行"

            # 找汇总表头行（包含"委办运费"或"总盈亏"的行）
            summary_header_row = None
            for row in range(1, min(5, ws.max_row + 1)):
                row_text = " ".join(normalize_text(ws.cell(row, c).value) for c in range(1, ws.max_column + 1))
                if "委办运费" in row_text or "总盈亏" in row_text:
                    summary_header_row = row
                    break

            field_map = {}
            if summary_header_row:
                for c in range(1, ws.max_column + 1):
                    h = normalize_text(ws.cell(summary_header_row, c).value)
                    if not h:
                        continue
                    if "趟次" in h:
                        field_map["trips"] = c
                    if "运费" in h:
                        field_map["freight"] = c
                    if "收入" in h:
                        field_map["income"] = c
                    if "盈亏" in h:
                        field_map["pnl"] = c

            def _num(col):
                try:
                    return float(ws.cell(total_row, col).value or 0)
                except:
                    return 0.0

            trips = _num(field_map.get("trips", 0))
            freight = _num(field_map.get("freight", 0))
            income = _num(field_map.get("income", 0))
            pnl = _num(field_map.get("pnl", 0))

            # 运输量：扫描合计行，每7列一组的 col+1(装载量) 累加
            load = 0.0
            col = 2
            while col + 6 <= ws.max_column:
                l = _num(col + 1)
                t = _num(col)
                if t > 0:
                    load += l
                col += 7

            return {"trips": trips, "load": load, "freight": freight, "income": income, "pnl": pnl}, None
        except Exception as exc:
            return None, f"上月{label}盈亏表解析失败: {exc}"
        finally:
            wb.close()

    prev_first, warn1 = previous_totals(previous_first_profit, "first")
    prev_second, warn2 = previous_totals(previous_second_profit, "second")
    prev_warnings: list[str] = []
    if warn1:
        prev_warnings.append(warn1)
    if warn2:
        prev_warnings.append(warn2)
    write_profit_workbook(first_profit_templates[0], first_profit_layout, first_metrics, first_path, (year, month), "first", previous_totals=prev_first, warnings=prev_warnings)
    write_profit_workbook(second_profit_templates[0], second_profit_layout, second_metrics, second_path, (year, month), "second", previous_totals=prev_second, warnings=prev_warnings)
    outputs = [split_path, first_path, second_path]
    warnings = [
        item
        for item in state.get("warnings", [])
        if not str(item).startswith(("财务记录未分类", "上月", "官方特殊Sheet未自动填充"))
    ] + prev_warnings
    if finance_classification["unclassified"]:
        warnings.append(
            f"财务记录未分类{finance_classification['unclassified']}条，其中西安正向路线"
            f"{finance_classification['unclassified_forward_xian']}条；未分类记录未进入盈亏表"
        )
    pending_special = [name for name, status in special_sheet_status.items() if status != "populated"]
    if pending_special:
        warnings.append(
            "官方特殊Sheet未自动填充，需人工补齐后才能作为完整官方交付件: "
            + "、".join(pending_special)
        )
    if unmatched_split:
        warnings.append(
            f"拆分版有{unmatched_split}条记录无法按路线名匹配官方线路Sheet，"
            "已放入「※未匹配目的地」，需人工指定归属后交付"
        )
    previous_profit_files = {
        kind: item["path"]
        for kind, item in previous_profit_inputs.items()
        if item["path"] is not None
    }
    previous_profit_hashes = {
        kind: item["sha256"]
        for kind, item in previous_profit_inputs.items()
        if item["status"] == "present"
    }
    state.update({
        "phase": "final_completed",
        "corrected_review_file": str(review_path.resolve()),
        "corrected_review_sha256": file_sha256(review_path),
        "final_outputs": [str(path.resolve()) for path in outputs],
        "final_output_hashes": {path.name: file_sha256(path) for path in outputs},
        "finance_classification": finance_classification,
        "expected_split_sheet_counts": expected_split_counts,
        "special_sheet_status": special_sheet_status,
        "template_layouts": {
            **state.get("template_layouts", {}),
            "first_profit": first_profit_layout,
            "second_profit": second_profit_layout,
        },
        "previous_profit_files": previous_profit_files,
        "previous_profit_hashes": previous_profit_hashes,
        "previous_profit_inputs": previous_profit_inputs,
        "final_metrics": {
            "row_count": len(rows),
            "downgrade_count": sum(bool(row.get("downgraded")) for row in rows),
            "finance_rows": len(finance_rows),
            "finance_unclassified": finance_classification["unclassified"],
            "first_groups": len(first_metrics),
            "second_groups": len(second_metrics),
            "special_sheets_pending": pending_special,
        },
        "warnings": warnings,
        "assets_dir": str(assets.resolve()),
    })
    _save_state(state_path, state)
    return outputs


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="读取人工修订初版并生成三个最终月度报表")
    parser.add_argument("--review-file", required=True, type=Path)
    parser.add_argument("--workspace-dir", required=True, type=Path)
    parser.add_argument("--previous-first-profit", type=Path, default=None)
    parser.add_argument("--previous-second-profit", type=Path, default=None)
    args = parser.parse_args(argv)
    try:
        outputs = build_final_reports(args.review_file, args.workspace_dir, previous_first_profit=args.previous_first_profit, previous_second_profit=args.previous_second_profit)
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        return 2
    print(json.dumps({"ok": True, "outputs": [str(path) for path in outputs]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
