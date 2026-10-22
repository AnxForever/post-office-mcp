from __future__ import annotations
from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from reconciliation.detect import bundled_asset_path, refine_profiles
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
# 版式模板和合同价格表随包内置，角色固定；模板含全年页签、价格表只有单程/往返两页，
# 都不参与目标月份推断，也不该因为客户没上传而被判成「缺文件」。
MONTH_AGNOSTIC_ROLES = (
    "draft_template",
    "first_profit_template",
    "second_profit_template",
    "first_price",
    "second_price",
)
PRICE_ROLES = ("first_price", "second_price")
# 一干/二干成套使用：用户传了其中一份或一份分不出干线的候选时，
# 不再用内置的另一份补齐，避免两版合同混用。
BUNDLED_PAIRS = {
    "first_price": ("first_price", "second_price", "price_candidate"),
    "second_price": ("first_price", "second_price", "price_candidate"),
    "first_profit_template": (
        "first_profit_template",
        "second_profit_template",
        "profit_candidate",
    ),
    "second_profit_template": (
        "first_profit_template",
        "second_profit_template",
        "profit_candidate",
    ),
}
PACKAGE_VERSION = "6.2.5"


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
    # 输入不全不是执行失败：这些条目会让上层返回 status=needs_input（退出码 0），
    # 由 Agent 请用户补充文件，而不是抛异常或报“执行失败”。
    missing_inputs: list[dict[str, Any]] = []

    damaged = [profile.path.name for profile in profiles if profile.role == "damaged"]
    unknown = [profile.path.name for profile in profiles if profile.role in {"unknown", "price_candidate", "profit_candidate"}]
    if damaged:
        fatal.append(f"工作簿损坏或无法读取: {', '.join(damaged)}")
        for name in damaged:
            missing_inputs.append({
                "kind": "unreadable",
                "role": "damaged",
                "label": name,
                "action": f"文件无法读取，请重新上传可打开的工作簿：{name}",
            })
    if unknown:
        warnings.append(f"未识别且未参与核对的工作簿: {', '.join(unknown)}")
    for item in detected:
        if item.confidence == "low" and item.role in {"online_caotan", "online_gangwu"}:
            warnings.append(f"在线表站点无法从内容确认，请人工复核: {item.path.name}")

    by_role: dict[str, list[Any]] = {}
    for profile in profiles:
        by_role.setdefault(profile.role, []).append(profile)
    # 版式模板和合同价格表是可选的：没有上传时用内置 assets 里的版本
    # （初版模板、一干/二干盈亏模板、一干/二干价格表）。
    # integrity.DYNAMIC_ROLE_COUNTS 要求这些角色必须存在，所以缺哪个就补一条指向内置资产的条目；
    # 用户上传了则以用户的为准（可能是新版式或新合同）。
    package_dir = Path(__file__).resolve().parents[1]
    assets_root = Path(assets_dir) if assets_dir else package_dir / "assets"
    # 与 detect 一致：只看用户实际上传的角色，刚补进去的内置资产不算已上传。
    uploaded_roles = {profile.role for profile in profiles}
    for role in MONTH_AGNOSTIC_ROLES:
        if by_role.get(role):
            continue
        related = BUNDLED_PAIRS.get(role)
        if related and uploaded_roles & set(related):
            # 用户已上传这一套里的某一份（或无法区分干线的候选），
            # 让缺少/歧义照常报出来，不要用内置版本掩盖。
            continue
        bundled = bundled_asset_path(role, assets_root)
        if bundled is None:
            continue
        profile = classify_workbook(bundled)
        # 内置资产的角色是确定的，不靠表内标题猜一干/二干。
        profile.role = role
        by_role[role] = [profile]
        if role in PRICE_ROLES:
            warnings.append(
                f"未上传{ROLE_LABELS[role]}，使用 Skill 内置价格表（{bundled.name}）；"
                "合同调价后请上传新版覆盖"
            )
        elif role != "draft_template":
            warnings.append(f"未上传{ROLE_LABELS[role]}，使用 Skill 内置模板")
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
            label = ROLE_LABELS[role]
            fatal.append(f"缺少{label}: 需要{count}个，识别到{actual}个")
            missing_inputs.append({
                "kind": "missing",
                "role": role,
                "label": label,
                "required": count,
                "found": actual,
                "action": f"请上传{label}（需要 {count} 个，当前识别到 {actual} 个）",
            })
        elif actual > count:
            label = ROLE_LABELS[role]
            fatal.append(f"{label}角色冲突: 需要{count}个，识别到{actual}个")
            names = [item.path.name for item in by_role.get(role, [])]
            missing_inputs.append({
                "kind": "ambiguous",
                "role": role,
                "label": label,
                "required": count,
                "found": actual,
                "candidates": names,
                "action": f"请确认使用哪一个{label}（识别到 {actual} 个：{'、'.join(names)}），"
                          "删除多余文件后重试",
            })
    vehicle_paths = [item.path for item in active_profiles["vehicle"]]
    if vehicle_paths and company_map is not None:
        vehicle_errors = validate_vehicle_sources(vehicle_paths, company_map)
        fatal.extend(vehicle_errors)
        for message in vehicle_errors:
            # 公司缺文件/重复/未映射都是用户可补充的输入问题，不是执行失败。
            if "车管" in message:
                missing_inputs.append({
                    "kind": "vehicle_mapping",
                    "role": "vehicle",
                    "label": "车管基础数据",
                    "action": f"{message}；请补齐或替换对应的车管文件后重试",
                })

    # 草滩/港务已按日 Sheet 命名（草滩「8月1日」/ 港务「8.1」）识别，
    # 不再依据文件名。无法区分时 refine_profiles 会把置信度标为 low，
    # 上面已经转成警告，由人工复核。

    observed: list[tuple[int | None, int, str]] = []
    online_profiles: list[Any] = []
    for profile in profiles:
        if profile.role not in REQUIRED_ROLES:
            continue
        # 模板类工作簿自带全年页签，不能拿来推断目标月份。
        if profile.role in MONTH_AGNOSTIC_ROLES:
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
            missing_inputs.append({
                "kind": "month_mismatch",
                "role": "month",
                "label": "目标月份",
                "action": "输入文件包含多个业务月份，请确认只上传同一个月的文件，或显式指定 --month",
            })
        else:
            fatal.append("无法从输入推断目标月份，请使用 --month YYYY-MM")
            missing_inputs.append({
                "kind": "month_unknown",
                "role": "month",
                "label": "目标月份",
                "action": "无法从输入推断目标月份，请补充当月文件或显式指定 --month YYYY-MM",
            })

    if explicit:
        for profile in online_profiles:
            has_target_month = any(
                observed_month == explicit[1] and (year is None or year == explicit[0])
                for year, observed_month in profile.months
            )
            if not has_target_month:
                label = ROLE_LABELS.get(profile.role, "在线测量表")
                fatal.append(f"月份不一致: {profile.path.name} 不包含目标月份 {_month_text(explicit)} 的工作表")
                missing_inputs.append({
                    "kind": "month_mismatch",
                    "role": profile.role,
                    "label": f"{label}（{profile.path.name}）",
                    "action": f"请上传统计月份为 {_month_text(explicit)} 的{label}",
                })

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
            missing_inputs.append({
                "kind": "unreadable",
                "role": "vehicle",
                "label": "车管基础数据",
                "action": f"车管基础数据无法读取，请重新上传：{exc}",
            })
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
                missing_inputs.append({
                    "kind": "unreadable",
                    "role": "vehicle",
                    "label": "车管基础数据",
                    "action": "车管基础数据的「邮路执行日期」大范围无法解析，请检查或重新上传",
                })
            elif invalid_dates:
                warnings.append(f"车管基础数据有{invalid_dates}条日期无法解析，已从目标月份核对中排除")
            if month_counts["vehicle_other"]:
                warnings.append(f"车管基础数据有{month_counts['vehicle_other']}条非目标月份记录，已排除")
            if not month_counts["vehicle_current"]:
                fatal.append(f"车管基础数据不包含目标月份 {_month_text(explicit)} 的有效记录")
                missing_inputs.append({
                    "kind": "month_mismatch",
                    "role": "vehicle",
                    "label": "车管基础数据",
                    "action": f"请上传统计月份为 {_month_text(explicit)} 的车管基础数据",
                })
    if explicit and not fatal:
        try:
            month_counts["online_current"], month_counts["online_other"] = _count_online_rows(
                active_profiles["online_caotan"] + active_profiles["online_gangwu"], explicit
            )
        except Exception as exc:
            fatal.append(f"在线测量表无法读取: {exc}")
            missing_inputs.append({
                "kind": "unreadable",
                "role": "online",
                "label": "在线测量表",
                "action": f"在线测量表无法读取，请重新上传：{exc}",
            })

    if check_assets:
        fatal.extend(validate_assets(assets_root, PHASE_ONE_ASSETS))

    result: dict[str, Any] = {
        "ok": not fatal,
        "status": "ok" if not fatal else ("needs_input" if missing_inputs else "failed"),
        "target_month": _month_text(explicit) if explicit else None,
        "files_found": len(files),
        "missing_inputs": missing_inputs,
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
    # 缺文件/月份不符/文件损坏属于「请用户补充」，不是执行失败：退出码 0。
    return 0 if result["ok"] or result.get("missing_inputs") else 2


if __name__ == "__main__":
    sys.exit(main())
