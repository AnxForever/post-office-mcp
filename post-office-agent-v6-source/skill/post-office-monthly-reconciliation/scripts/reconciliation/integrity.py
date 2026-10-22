from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Sequence as SequenceABC
from pathlib import Path
from typing import Any, Mapping, Sequence

from openpyxl import load_workbook

from .excel_io import parse_date
from .normalize import normalize_dispatch_no, normalize_plate, normalize_text, normalize_trailer


REQUIRED_ASSETS = (
    "split-report-template.xlsx",
    "route-aliases.json",
    "company-map.json",
    # 版式模板和合同价格表随包内置：客户不上传时直接用它，上传了则用客户的。
    "draft-template.xlsx",
    "first-profit-template.xlsx",
    "second-profit-template.xlsx",
    "first-price-table.xlsx",
    "second-price-table.xlsx",
)

PHASE_ONE_ASSETS = REQUIRED_ASSETS

DYNAMIC_ROLE_COUNTS = {
    "online_caotan": 1,
    "online_gangwu": 1,
    "vehicle": 6,
    "finance": 1,
    "draft_template": 1,
    "first_price": 1,
    "second_price": 1,
    "first_profit_template": 1,
    "second_profit_template": 1,
}

PHASE_TEMPLATE_LAYOUTS = {
    "review": ("draft",),
    "final": ("draft", "first_profit", "second_profit"),
}

AUTHORITATIVE_FIELDS = (
    "company",
    "dispatch_no",
    "route",
    "execution_date",
    "plate",
    "trailer",
)


def _authoritative_value(field: str, value: Any) -> str:
    if field == "dispatch_no":
        return normalize_dispatch_no(value)
    if field == "plate":
        return normalize_plate(value)
    if field == "trailer":
        return normalize_trailer(value)
    if field == "execution_date":
        parsed = parse_date(value)
        return parsed.date().isoformat() if parsed else normalize_text(value)
    return normalize_text(value)


def validate_authoritative_state(state: Mapping[str, Any]) -> list[str]:
    """Fail closed when the phase-one authority snapshot is not usable."""
    details: list[str] = []
    if not isinstance(state, Mapping):
        return ["第一阶段状态不完整，请重新执行第一阶段：state不是对象"]

    raw_rows = state.get("row_ids")
    if isinstance(raw_rows, (str, bytes, bytearray)) or not isinstance(raw_rows, SequenceABC):
        details.append("row_ids必须是非字符串序列")
        rows: list[Any] = []
    else:
        rows = list(raw_rows)

    row_count = state.get("row_count")
    if isinstance(row_count, bool) or not isinstance(row_count, int) or row_count < 0:
        details.append("row_count必须是有效整数")
    elif row_count != len(rows):
        details.append(f"row_count与row_ids数量不一致：{row_count}!={len(rows)}")

    target_month = normalize_text(state.get("target_month"))
    if not re.fullmatch(r"20\d{2}-(?:0[1-9]|1[0-2])", target_month):
        details.append("target_month缺失或格式无效")

    source_ids: set[str] = set()
    for index, row in enumerate(rows, 1):
        if not isinstance(row, Mapping):
            details.append(f"权威记录第{index}行不是对象")
            continue
        source_id = normalize_text(row.get("source_id"))
        if not source_id:
            details.append(f"权威记录第{index}行缺少source_id")
        elif source_id in source_ids:
            details.append(f"权威记录第{index}行source_id重复：{source_id}")
        else:
            source_ids.add(source_id)
        missing_fields = [field for field in AUTHORITATIVE_FIELDS if field not in row]
        if missing_fields:
            details.append(f"权威记录第{index}行缺少{','.join(missing_fields)}")

    if not details:
        return []
    return ["第一阶段状态不完整，请重新执行第一阶段：" + "；".join(details)]


def validate_dynamic_state(state: Mapping[str, Any], phase: str) -> list[str]:
    """Validate locked monthly inputs and discovered layouts."""
    errors: list[str] = []
    sources = state.get("sources")
    if isinstance(sources, Mapping):
        by_role = sources
    elif isinstance(sources, SequenceABC) and not isinstance(sources, (str, bytes, bytearray)):
        by_role = {
            role: [item for item in sources if isinstance(item, Mapping) and item.get("role") == role]
            for role in DYNAMIC_ROLE_COUNTS
        }
    else:
        by_role = {}
        errors.append("运行状态sources必须是动态输入角色映射")

    expected_role_counts = state.get("expected_role_counts")
    if not isinstance(expected_role_counts, Mapping):
        expected_role_counts = DYNAMIC_ROLE_COUNTS
    for role, default_count in DYNAMIC_ROLE_COUNTS.items():
        raw_expected = expected_role_counts.get(role, default_count)
        expected_count = raw_expected if isinstance(raw_expected, int) and raw_expected > 0 else default_count
        entries = by_role.get(role, []) if isinstance(by_role, Mapping) else []
        if isinstance(entries, (str, bytes, bytearray)) or not isinstance(entries, SequenceABC):
            errors.append(f"运行状态动态输入角色{role}不是文件序列")
            continue
        entries = list(entries)
        if len(entries) != expected_count:
            errors.append(f"运行状态动态输入角色{role}应为{expected_count}个，实际{len(entries)}个")
            continue
        for index, item in enumerate(entries, 1):
            if not isinstance(item, Mapping):
                errors.append(f"运行状态动态输入角色{role}第{index}项不是对象")
                continue
            if not normalize_text(item.get("path")):
                errors.append(f"运行状态动态输入角色{role}第{index}项缺少path")
            digest = normalize_text(item.get("sha256"))
            if not re.fullmatch(r"[0-9a-fA-F]{64}", digest):
                errors.append(f"运行状态动态输入角色{role}第{index}项缺少有效sha256")

    layouts = state.get("template_layouts")
    if not isinstance(layouts, Mapping):
        layouts = {}
        errors.append("运行状态缺少template_layouts动态模板发现结果")
    for name in PHASE_TEMPLATE_LAYOUTS.get(phase, ()):
        layout = layouts.get(name)
        if not isinstance(layout, Mapping) or not layout:
            errors.append(f"运行状态template_layouts缺少{name}动态模板发现结果")
    return errors


def validate_authoritative_rows(
    workbook_rows: Sequence[Mapping[str, Any]],
    state: Mapping[str, Any],
    *,
    max_error_details: int = 20,
) -> list[str]:
    """Validate user-visible rows against the ordered phase-one authority snapshot."""
    state_errors = validate_authoritative_state(state)
    if state_errors:
        return state_errors
    state_rows = list(state["row_ids"])
    expected_count = state["row_count"]
    if len(workbook_rows) != expected_count:
        return [f"人工修订初版行数变化：期望 {expected_count}，实际 {len(workbook_rows)}"]

    errors: list[str] = []
    target_month = normalize_text(state["target_month"])
    actual_months = sorted({
        parsed.strftime("%Y-%m")
        for row in workbook_rows
        if (parsed := parse_date(row.get("execution_date"))) is not None
    })
    if target_month and actual_months and actual_months != [target_month]:
        errors.append(
            f"人工修订初版月份与当前运行月份{target_month}不一致：实际月份{','.join(actual_months)}"
        )
    elif target_month and not actual_months:
        errors.append(f"人工修订初版没有可解析的执行日期，无法校验当前运行月份{target_month}")

    mismatches: list[str] = []
    for index, (actual, expected) in enumerate(zip(workbook_rows, state_rows), 2):
        for field in AUTHORITATIVE_FIELDS:
            expected_value = _authoritative_value(field, expected.get(field))
            actual_value = _authoritative_value(field, actual.get(field))
            if actual_value == expected_value:
                continue
            mismatches.append(
                f"基础报表Excel第{index}行字段{field}与车管权威派车记录不一致："
                f"期望 {expected_value!r}，实际 {actual_value!r}"
            )
    detail_limit = max(0, int(max_error_details))
    errors.extend(mismatches[:detail_limit])
    if len(mismatches) > detail_limit:
        errors.append(f"车管权威派车记录校验共{len(mismatches)}处不一致，仅展示前{detail_limit}处")
    return errors


def file_sha256(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_template_inventory(path: Path | str) -> list[str]:
    value = json.loads(Path(path).read_text(encoding="utf-8"))
    sheets = value.get("sheets") if isinstance(value, dict) else None
    if not isinstance(sheets, list) or not sheets or not all(isinstance(item, str) and item for item in sheets):
        raise ValueError("基础模板Sheet清单无效")
    return sheets


def asset_hashes(
    assets_dir: Path | str,
    required_assets: Sequence[str] = REQUIRED_ASSETS,
) -> dict[str, str]:
    root = Path(assets_dir)
    return {name: file_sha256(root / name) for name in required_assets if (root / name).is_file()}


def validate_vehicle_sources(paths: list[Path | str], company_map: Mapping[str, Any]) -> list[str]:
    paths = [Path(item) for item in paths]
    errors: list[str] = []
    companies = company_map.get("companies") if isinstance(company_map, Mapping) else None
    if not isinstance(companies, list) or not companies:
        return ["company-map.json 必须定义至少一家公司"]
    entries = [
        (normalize_text(item.get("source_token")), normalize_text(item.get("company")))
        for item in companies
        if isinstance(item, Mapping)
    ]
    if len(entries) != len(companies) or any(not token or not company for token, company in entries):
        return ["company-map.json 公司定义无效"]
    if len(paths) != len(entries):
        errors.append(f"车管基础数据需要公司映射中{len(entries)}家公司各1个文件，当前为{len(paths)}个")

    mapped: dict[str, list[Path]] = {token: [] for token, _ in entries}
    for path in paths:
        matches = [(token, company) for token, company in entries if token in normalize_text(path.stem)]
        if not matches:
            errors.append(f"车管文件未映射到 company-map.json: {path.name}")
        elif len(matches) > 1:
            errors.append(f"车管文件公司映射不唯一: {path.name} ({', '.join(token for token, _ in matches)})")
        else:
            mapped[matches[0][0]].append(path)
    missing = [company for token, company in entries if not mapped[token]]
    duplicate = [company for token, company in entries if len(mapped[token]) > 1]
    if missing:
        errors.append("车管文件缺少公司: " + ", ".join(missing))
    if duplicate:
        errors.append("车管文件公司重复: " + ", ".join(duplicate))
    existing = [path for path in paths if path.is_file()]
    digests = [file_sha256(path) for path in existing]
    if len(digests) != len(set(digests)):
        errors.append("车管基础数据存在内容重复的工作簿")
    return errors


def validate_assets(
    assets_dir: Path | str,
    required_assets: Sequence[str] = REQUIRED_ASSETS,
) -> list[str]:
    root = Path(assets_dir)
    errors: list[str] = []
    required = tuple(required_assets)
    missing = [name for name in required if not (root / name).is_file()]
    if missing:
        return ["Skill静态资产缺失: " + ", ".join(missing)]
    for name in ("route-aliases.json", "company-map.json"):
        if name not in required:
            continue
        try:
            value = json.loads((root / name).read_text(encoding="utf-8"))
            if not isinstance(value, dict):
                raise ValueError
        except Exception:
            errors.append(f"Skill静态资产JSON无效: {name}")
    if "split-report-template.xlsx" in required:
        try:
            wb = load_workbook(root / "split-report-template.xlsx", read_only=True, data_only=False, keep_links=False)
            try:
                if not wb.sheetnames:
                    errors.append("拆分母版没有可用Sheet")
            finally:
                wb.close()
        except Exception as exc:
            errors.append(f"拆分母版无效: {exc}")
    return errors


def verify_state_files(
    state: Mapping[str, Any],
    assets_dir: Path | str,
    *,
    include_sources: bool = True,
    asset_names: Sequence[str] | None = None,
) -> list[str]:
    errors: list[str] = []
    if include_sources:
        sources = state.get("sources", [])
        if isinstance(sources, Mapping):
            source_items = [
                item
                for role_items in sources.values()
                if isinstance(role_items, list)
                for item in role_items
            ]
        else:
            source_items = sources
        for item in source_items:
            if not isinstance(item, Mapping):
                continue
            path = Path(item.get("path", ""))
            expected = item.get("sha256")
            if not path.is_file():
                errors.append(f"运行源文件不存在: {item.get('name') or path.name}")
            elif expected and file_sha256(path) != expected:
                errors.append(f"运行源文件已变化: {item.get('name') or path.name}")
    assets_root = Path(assets_dir)
    allowed_assets = set(asset_names) if asset_names is not None else set(REQUIRED_ASSETS)
    locked_assets = state.get("asset_hashes", {})
    if not isinstance(locked_assets, Mapping):
        locked_assets = {}
    if asset_names is not None:
        for name in asset_names:
            expected = normalize_text(locked_assets.get(name))
            if not re.fullmatch(r"[0-9a-fA-F]{64}", expected):
                errors.append(f"运行状态缺少Skill静态资产哈希: {name}")
    for name, expected in locked_assets.items():
        if name not in allowed_assets:
            continue
        path = assets_root / name
        if not path.is_file():
            errors.append(f"Skill静态资产不存在: {name}")
        elif expected and file_sha256(path) != expected:
            errors.append(f"Skill静态资产已变化: {name}")
    return errors
