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
    read_measure_rows,
    read_online_rows,
    read_vehicle_platform_rows,
    read_vehicle_rows,
)
from reconciliation.integrity import (
    DYNAMIC_ROLE_COUNTS,
    PHASE_ONE_ASSETS,
    asset_hashes,
    validate_assets,
    validate_vehicle_sources,
)
from reconciliation.normalize import parse_month_token


ROLE_LABELS = {
    "online_caotan": "草滩在线测量表",
    "online_gangwu": "港务在线测量表",
    "online_measure": "车辆容积测量表",
    "vehicle": "车管基础数据",
    "vehicle_platform": "车管平台汇总导出",
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
# 输入可以按「两套口径」之一提供，二选一：
#   逐场地口径：草滩在线表 + 港务在线表 + 6 家公司车管导出
#   合并口径  ：容积测量表（一场一表合并导出）+ 车管平台汇总导出
# 两套等价，不能混着用（混用会让容间与车管行对不上）。
SOURCE_BUNDLES = {
    "split": {
        "label": "逐场地/逐公司导出",
        "roles": {"online_caotan": 1, "online_gangwu": 1, "vehicle": 6},
        "pattern_roles": ("online_caotan", "online_gangwu", "vehicle"),
    },
    "merged": {
        "label": "合并导出",
        "roles": {"online_measure": 1, "vehicle_platform": 1},
        "pattern_roles": ("online_measure", "vehicle_platform"),
    },
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
PACKAGE_VERSION = "6.3.1"


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
        if profile.role == "online_measure":
            current += len(read_measure_rows(profile.path, target))
            continue
        site = "草滩" if profile.role == "online_caotan" else "港务"
        current += len(read_online_rows(profile.path, target, site))
        for year, month in profile.months:
            candidate = (year or target[0], month)
            if candidate != target:
                other += len(read_online_rows(profile.path, candidate, site))
    return current, other


def _select_source_bundle(by_role: dict[str, list[Any]]) -> tuple[str, str]:
    """判断这份输入走「逐场地/逐公司」还是「合并导出」口径。

    两套数据不能混用：合并口径的容间与逐场地测量按不同规则取整，
    混着读会让同一趟在两套取数之间来回跳。所以按上传的角色整包选一套，
    另一套需要的角色直接不再要求，避免误报「缺 6 个车管文件」。
    """
    merged_roles = set(SOURCE_BUNDLES["merged"]["pattern_roles"])
    # 逐场地口径只看草滩/港务：单独一家车管文件在合并口径里也常见，
    # 拿它当「混用」会把正常的合并输入误判掉。
    split_roles = {"online_caotan", "online_gangwu"}
    has_merged = bool(merged_roles & set(by_role))
    has_split = bool(split_roles & set(by_role))
    if has_merged and has_split:
        return "split", (
            "同时识别到逐场地/逐公司导出与合并导出两类文件，按逐场地口径处理；"
            "合并导出的文件已忽略"
        )
    if has_merged:
        return "merged", (
            f"本次按「{SOURCE_BUNDLES['merged']['label']}」口径核对："
            "容积测量表 + 车管平台汇总导出"
        )
    return "split", ""


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
    source_bundle, bundle_note = _select_source_bundle(by_role)
    if bundle_note:
        warnings.append(bundle_note)
    if source_bundle == "merged":
        # 合并口径：一张容积测量表 + 一张车管汇总，替代原来的 2+6 个文件。
        for role in SOURCE_BUNDLES["split"]["roles"]:
            roles_required.pop(role, None)
        roles_required.update(SOURCE_BUNDLES["merged"]["roles"])
    active_profiles = {role: by_role.get(role, []) for role in roles_required}
    ignored_models = by_role.get("model", [])
    company_map: dict[str, Any] | None = None
    # 全部动态角色都写进期望表（不需要的记 0），校验时只按这张表核对，
    # 不去猜「这个口径是不是还该有草滩/港务/6 家车管文件」。
    expected_role_counts = {role: roles_required.get(role, 0) for role in DYNAMIC_ROLE_COUNTS}
    try:
        company_map = load_company_map(assets_root / "company-map.json")
        # 只有逐公司口径才需要按公司数校验文件个数；
        # 合并口径只有一张平台汇总，写回 6 会凭空要求 6 个车管文件。
        if source_bundle == "split":
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
    vehicle_paths = [item.path for item in active_profiles.get("vehicle", [])]
    # 逐公司口径才检查公司映射；合并口径只有一张平台汇总，没有公司维度。
    if vehicle_paths and company_map is not None and source_bundle == "split":
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
        if profile.role not in roles_required:
            continue
        # 模板类工作簿自带全年页签，不能拿来推断目标月份。
        if profile.role in MONTH_AGNOSTIC_ROLES:
            continue
        if profile.role in {"online_caotan", "online_gangwu", "online_measure"}:
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
            if profile.role == "online_measure":
                # 合并测量表没有日页签，月份只能靠文件名；文件名不可靠时
                # 用「读出来的行数」判断，而不是硬拦。
                continue
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
    if (
        explicit
        and source_bundle == "split"
        and len(vehicle_paths) == expected_role_counts.get("vehicle", 0)
        and not any("车管" in item for item in fatal)
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
    # 合并口径只有一张车管汇总，公司映射不适用；这里单独校验能否读出当月行。
    merged_vehicle_paths = [item.path for item in active_profiles.get("vehicle_platform", [])]
    if explicit and source_bundle == "merged" and merged_vehicle_paths and not any(
        "车管" in item for item in fatal
    ):
        try:
            platform_rows = read_vehicle_platform_rows(merged_vehicle_paths[0])
        except Exception as exc:
            fatal.append(f"车管平台汇总导出无法读取: {exc}")
            missing_inputs.append({
                "kind": "unreadable",
                "role": "vehicle_platform",
                "label": "车管平台汇总导出",
                "action": f"车管平台汇总导出无法读取，请重新上传：{exc}",
            })
        else:
            current = sum(
                1 for row in platform_rows
                if row["execution_date"] and (row["execution_date"].year, row["execution_date"].month) == explicit
            )
            month_counts["vehicle_current"] = current
            month_counts["vehicle_other"] = len(platform_rows) - current
            if not current:
                fatal.append(f"车管平台汇总导出不包含目标月份 {_month_text(explicit)} 的有效记录")
                missing_inputs.append({
                    "kind": "month_mismatch",
                    "role": "vehicle_platform",
                    "label": "车管平台汇总导出",
                    "action": f"请上传统计月份为 {_month_text(explicit)} 的车管平台汇总导出",
                })
            elif month_counts["vehicle_other"]:
                warnings.append(f"车管平台汇总导出有{month_counts['vehicle_other']}条非目标月份记录，已排除")
    if explicit and not fatal:
        try:
            online_sources = (
                active_profiles.get("online_caotan", [])
                + active_profiles.get("online_gangwu", [])
                + active_profiles.get("online_measure", [])
            )
            month_counts["online_current"], month_counts["online_other"] = _count_online_rows(
                online_sources, explicit
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
        "sources": {
            **{role: [_source_entry(item, explicit) for item in items] for role, items in active_profiles.items()},
            # 基础报表原版是可选输入：传了就用它逐单对价，没传不影响流程，
            # 所以不进 roles_required，只在这里带上（有才带）。
            **(
                {"base_report": [_source_entry(item, explicit) for item in by_role["base_report"]]}
                if by_role.get("base_report") else {}
            ),
        },
        "ignored_sources": {"model": [_source_entry(item, explicit) for item in ignored_models]},
        "month_counts": month_counts,
        "source_bundle": source_bundle,
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
        "source_bundle": source_bundle,
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
