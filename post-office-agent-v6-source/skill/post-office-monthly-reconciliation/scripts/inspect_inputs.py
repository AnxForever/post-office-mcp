from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from reconciliation.detect import refine_profiles
from reconciliation.excel_io import (
    classify_workbook,
    discover_workbooks,
    file_sha256,
    load_company_map,
    read_online_rows,
    read_vehicle_rows,
)
from reconciliation.integrity import PHASE_ONE_ASSETS, asset_hashes, validate_assets, validate_vehicle_sources
from reconciliation.normalize import parse_month_token


ROLE_LABELS = {
    "online_caotan": "草滩在线测量表",
    "online_gangwu": "港务在线测量表",
    "vehicle": "车管基础数据",
    "finance": "财务收入表",
    "draft_template": "初版生成模板",
    "first_price": "一干价格表",
    "second_price": "二干价格表",
    "first_profit_template": "一干委办盈亏模板",
    "second_profit_template": "二干委办盈亏模板",
    "price_candidate": "价格表（一干/二干待定）",
    "profit_candidate": "盈亏模板（一干/二干待定）",
}
REQUIRED_ROLES = {
    "online_caotan": 1,
    "online_gangwu": 1,
    "vehicle": 6,
    "finance": 1,
    "first_price": 1,
    "second_price": 1,
    "first_profit_template": 1,
    "second_profit_template": 1,
}
PACKAGE_VERSION = "6.2.0"


def _parse_explicit_month(value: str | None) -> tuple[int, int] | None:
    if not value:
        return None
    try:
        parsed = datetime.strptime(value, "%Y-%m")
    except ValueError as exc:
        raise ValueError("--month 必须使用 YYYY-MM") from exc
    return parsed.year, parsed.month


def _month_text(month: tuple[int, int]) -> str:
    return f"{month[0]:04d}-{month[1]:02d}"


def _source_entry(profile: Any, explicit: tuple[int, int] | None) -> dict[str, Any]:
    return {
        "path": str(profile.path.resolve()),
        "name": profile.path.name,
        "role": profile.role,
        "role_label": ROLE_LABELS.get(profile.role, profile.role),
        "sha256": file_sha256(profile.path),
        "sheets": profile.sheet_names,
        "target_sheets": [
            name for name in profile.sheet_names
            if profile.role not in {"online_caotan", "online_gangwu"}
            or (
                (token := parse_month_token(name)) is not None
                and explicit is not None
                and token[1] == explicit[1]
                and (token[0] is None or token[0] == explicit[0])
            )
        ],
        "header_row": profile.header_row,
        "field_map": profile.field_map,
        "months": profile.months,
        "discovery": {},
    }


def _count_online_rows(profiles: list[Any], target: tuple[int, int]) -> tuple[int, int]:
    current = other = 0
    for profile in profiles:
        site = "草滩" if profile.role == "online_caotan" else "港务"
        current += len(read_online_rows(profile.path, target, site))
        for year, month in profile.months:
            candidate = (year or target[0], month)
            if candidate != target:
                other += len(read_online_rows(profile.path, candidate, site))
    return current, other


def inspect_inputs(
    input_dir: Path | str,
    workspace_dir: Path | str,
    month: str | None = None,
    *,
    check_assets: bool = True,
    assets_dir: Path | str | None = None,
) -> dict[str, Any]:
    input_dir, workspace_dir = Path(input_dir), Path(workspace_dir)
    explicit = _parse_explicit_month(month)
    files = discover_workbooks(input_dir)
    profiles = [classify_workbook(path) for path in files]
    # 用结构识别细化角色：草滩/港务看日 Sheet 命名，价格表/盈亏模板看表内标题。
    # 细化结果写回 profile，后续所有读取都基于这个角色。
    detected = refine_profiles(profiles, explicit)
    for item in detected:
        if item.profile is not None and item.role != item.profile.role:
            item.profile.role = item.role
            item.profile.site = (
                "港务" if item.role == "online_gangwu"
                else "草滩" if item.role == "online_caotan" else None
            )
    fatal: list[str] = []
    warnings: list[str] = []

    damaged = [profile.path.name for profile in profiles if profile.role == "damaged"]
    unknown = [profile.path.name for profile in profiles if profile.role in {"unknown", "price_candidate", "profit_candidate"}]
    if damaged:
        fatal.append(f"工作簿损坏或无法读取: {', '.join(damaged)}")
    if unknown:
        warnings.append(f"未识别且未参与核对的工作簿: {', '.join(unknown)}")
    for item in detected:
        if item.confidence == "low" and item.role in {"online_caotan", "online_gangwu"}:
            warnings.append(f"在线表站点无法从内容确认，请人工复核: {item.path.name}")

    by_role: dict[str, list[Any]] = {}
    for profile in profiles:
        by_role.setdefault(profile.role, []).append(profile)
    # 初版模板是可选的：没有上传时由 build_review_workbook 回落到内置 assets/draft-template.xlsx。
    # integrity.DYNAMIC_ROLE_COUNTS 要求该角色必须存在，所以这里补一条指向内置模板的条目。
    package_dir = Path(__file__).resolve().parents[1]
    assets_root = Path(assets_dir) if assets_dir else package_dir / "assets"
    if "draft_template" not in by_role:
        bundled = assets_root / "draft-template.xlsx"
        if bundled.is_file():
            by_role["draft_template"] = [classify_workbook(bundled)]
    roles_required = dict(REQUIRED_ROLES, draft_template=1)
    active_profiles = {role: by_role.get(role, []) for role in roles_required}
    ignored_models = by_role.get("model", [])
    company_map: dict[str, Any] | None = None
    expected_role_counts = dict(roles_required)
    try:
        company_map = load_company_map(assets_root / "company-map.json")
        expected_role_counts["vehicle"] = len(company_map["companies"])
    except Exception as exc:
        fatal.append(f"company-map.json 无法读取: {exc}")
    if ignored_models:
        warnings.append("外部车型表已忽略，不参与本次月度运输核对")
    for role, count in expected_role_counts.items():
        actual = len(by_role.get(role, []))
        if actual < count:
            fatal.append(f"缺少{ROLE_LABELS[role]}: 需要{count}个，识别到{actual}个")
        elif actual > count:
            fatal.append(f"{ROLE_LABELS[role]}角色冲突: 需要{count}个，识别到{actual}个")
    vehicle_paths = [item.path for item in active_profiles["vehicle"]]
    if vehicle_paths and company_map is not None:
        fatal.extend(validate_vehicle_sources(vehicle_paths, company_map))

    # 草滩/港务已按日 Sheet 命名（草滩「8月1日」/ 港务「8.1」）识别，
    # 不再依据文件名。无法区分时 refine_profiles 会把置信度标为 low，
    # 上面已经转成警告，由人工复核。

    observed: list[tuple[int | None, int, str]] = []
    online_profiles: list[Any] = []
    for profile in profiles:
        if profile.role not in REQUIRED_ROLES:
            continue
        if profile.role in {"online_caotan", "online_gangwu"}:
            online_profiles.append(profile)
            continue
        for token in profile.months:
            observed.append((token[0], token[1], profile.path.name))
    if not explicit:
        full = {(year, mon) for year, mon, _ in observed if year is not None}
        months = {mon for _, mon, _ in observed}
        if len(full) == 1:
            explicit = next(iter(full))
        elif len(months) == 1:
            explicit = (datetime.now().year, next(iter(months)))
            warnings.append("输入未提供可确认年份，暂用当前年份；建议显式传入 --month")
        elif observed:
            fatal.append("月份不一致: 输入文件或工作表包含多个业务月份")
        else:
            fatal.append("无法从输入推断目标月份，请使用 --month YYYY-MM")

    if explicit:
        for profile in online_profiles:
            has_target_month = any(
                observed_month == explicit[1] and (year is None or year == explicit[0])
                for year, observed_month in profile.months
            )
            if not has_target_month:
                fatal.append(f"月份不一致: {profile.path.name} 不包含目标月份 {_month_text(explicit)} 的工作表")

    month_counts = {
        "vehicle_current": 0,
        "vehicle_other": 0,
        "online_current": 0,
        "online_other": 0,
    }
    if explicit and len(vehicle_paths) == expected_role_counts["vehicle"] and not any(
        "车管" in item for item in fatal
    ):
        try:
            vehicle_rows = read_vehicle_rows(vehicle_paths, company_map)
        except Exception as exc:
            fatal.append(f"车管基础数据无法读取: {exc}")
        else:
            invalid_dates = sum(row["execution_date"] is None for row in vehicle_rows)
            rows_by_file: dict[str, list[dict[str, Any]]] = {}
            for row in vehicle_rows:
                rows_by_file.setdefault(row["source_file"], []).append(row)
            for source_file, rows in rows_by_file.items():
                if rows and not any(row["execution_date"] is not None for row in rows):
                    fatal.append(f"车管基础数据 {source_file} 的业务记录日期均无法解析")
            for row in vehicle_rows:
                execution_date = row["execution_date"]
                if execution_date is None:
                    continue
                if (execution_date.year, execution_date.month) == explicit:
                    month_counts["vehicle_current"] += 1
                else:
                    month_counts["vehicle_other"] += 1
            if invalid_dates >= 3 and invalid_dates * 2 >= len(vehicle_rows):
                fatal.append("车管基础数据日期大范围无法解析，无法生成可靠月度核对")
            elif invalid_dates:
                warnings.append(f"车管基础数据有{invalid_dates}条日期无法解析，已从目标月份核对中排除")
            if month_counts["vehicle_other"]:
                warnings.append(f"车管基础数据有{month_counts['vehicle_other']}条非目标月份记录，已排除")
            if not month_counts["vehicle_current"]:
                fatal.append(f"车管基础数据不包含目标月份 {_month_text(explicit)} 的有效记录")
    if explicit and not fatal:
        try:
            month_counts["online_current"], month_counts["online_other"] = _count_online_rows(
                active_profiles["online_caotan"] + active_profiles["online_gangwu"], explicit
            )
        except Exception as exc:
            fatal.append(f"在线测量表无法读取: {exc}")

    if check_assets:
        fatal.extend(validate_assets(assets_root, PHASE_ONE_ASSETS))

    result: dict[str, Any] = {
        "ok": not fatal,
        "target_month": _month_text(explicit) if explicit else None,
        "files_found": len(files),
        "roles": {role: [str(item.path) for item in items] for role, items in by_role.items() if role not in {"unknown", "damaged"}},
        "sources": {role: [_source_entry(item, explicit) for item in items] for role, items in active_profiles.items()},
        "ignored_sources": {"model": [_source_entry(item, explicit) for item in ignored_models]},
        "month_counts": month_counts,
        "fatal_errors": fatal,
        "warnings": warnings,
        "profiles": [
            {
                "path": str(profile.path),
                "role": profile.role,
                "site": profile.site,
                "sheets": profile.sheet_names,
                "header_row": profile.header_row,
                "field_map": profile.field_map,
                "months": profile.months,
                "errors": profile.errors,
            }
            for profile in profiles
        ],
    }
    if fatal:
        return result

    workspace_dir.mkdir(parents=True, exist_ok=True)
    state = {
        "package_version": PACKAGE_VERSION,
        "phase": "inspected",
        "target_month": result["target_month"],
        "created_at": datetime.now(timezone.utc).isoformat(),
        "input_dir": str(input_dir.resolve()),
        "assets_dir": str(assets_root.resolve()),
        "asset_hashes": asset_hashes(assets_root, PHASE_ONE_ASSETS) if check_assets else {},
        "expected_role_counts": expected_role_counts,
        "sources": result["sources"],
        "ignored_sources": result["ignored_sources"],
        "month_counts": month_counts,
        "warnings": warnings,
    }
    temp = workspace_dir / "run-state.json.tmp"
    temp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(workspace_dir / "run-state.json")
    result["state_file"] = str(workspace_dir / "run-state.json")
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="预检邮局月度委办运输报表输入")
    parser.add_argument("--input-dir", required=True, type=Path)
    parser.add_argument("--workspace-dir", required=True, type=Path)
    parser.add_argument("--month", help="目标月份，格式 YYYY-MM")
    args = parser.parse_args(argv)
    result = inspect_inputs(args.input_dir, args.workspace_dir, args.month)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
