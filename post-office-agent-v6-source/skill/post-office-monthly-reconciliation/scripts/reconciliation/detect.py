# -*- coding: utf-8 -*-
"""按结构识别工作区里的 Excel 文件角色。

原则：
  1. 先看表结构（Sheet 名、表头字段、日 Sheet 命名、档位特征），不靠文件名；
  2. 文件名只作为最后的兜底线索；
  3. 每个判断都带证据和置信度，识别不准时报 ambiguous，不硬猜。

业务契约（这些是固定的接口，不是写死）：
  - 财务收入表：表头含 邮路名称 / 额定吨位 / 结算收入 / 普服收入
  - 基础报表原版：含 Sheet「一干汇总」「二干基础」
  - 基础报表初版：21 列表头
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from .excel_io import (
    FileProfile,
    _pair_feature,
    classify_workbook,
    discover_workbooks,
    normalize_text,
    parse_month_token,
)

BASE_SHEETS = ("一干汇总", "二干基础")
# 草滩「4月1」/「4月1日」，港务「4.1」
DAILY_CAOTAN = re.compile(r"^\d+月\d+日?$")
DAILY_GANGWU = re.compile(r"^\d+\.\d+$")

# 这些角色按月区分版本；价格表/盈亏模板/初版模板跟随文档，不按月份淘汰。
MONTH_SCOPED = {"base_report", "finance", "online_caotan", "online_gangwu", "vehicle"}

ROLE_LABELS = {
    "base_report": "基础报表原版",
    "finance": "财务收入表",
    "online_caotan": "草滩在线测量表",
    "online_gangwu": "港务在线测量表",
    "vehicle": "车管基础数据",
    "first_price": "一干价格表",
    "second_price": "二干价格表",
    "price_candidate": "价格表（一干/二干待定）",
    "first_profit_template": "一干委办盈亏模板",
    "second_profit_template": "二干委办盈亏模板",
    "profit_candidate": "盈亏模板（一干/二干待定）",
    "draft_template": "基础报表初版模板",
    "model": "车型表",
    "damaged": "损坏文件",
    "unknown": "未识别",
}


@dataclass
class Detected:
    path: Path
    role: str
    confidence: str = "medium"
    evidence: list[str] = field(default_factory=list)
    months: list[tuple[int | None, int]] = field(default_factory=list)
    sheets: list[str] = field(default_factory=list)
    profile: FileProfile | None = field(default=None, repr=False)

    def to_dict(self) -> dict:
        return {
            "path": str(self.path),
            "name": self.path.name,
            "role": self.role,
            "role_label": ROLE_LABELS.get(self.role, self.role),
            "confidence": self.confidence,
            "evidence": self.evidence,
            "months": [f"{y or '?'}-{m:02d}" for y, m in self.months],
            "sheets": self.sheets[:12],
        }


def _finance_months(path: Path, header_row: int | None = None) -> list[tuple[int | None, int]]:
    """财务收入表没有日期 Sheet，用「流水号」前 4 位取**出现最多的月份**。

    财务表里会夹杂少量跨月记录（次日结算、补录），所以不能取全部月份，
    否则 5 月的表会因为几条 4 月记录也被判成 4 月。

    header_row 由 classify_workbook 提供，避免重复打开工作簿
    （Windows 上 openpyxl 只读模式重复打开同一文件有临时文件竞争）。
    """
    counts: dict[int, int] = {}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return []
    try:
        for ws in wb.worksheets:
            if hasattr(ws, "reset_dimensions"):
                ws.reset_dimensions()
            stream = ws.iter_rows(values_only=True)
            serial = None
            start_row = header_row
            if start_row:
                # 已经知道表头行，直接跳到表头读取列位置
                for offset, header in enumerate(stream, 1):
                    if offset == start_row:
                        columns = {
                            normalize_text(v): i for i, v in enumerate(header) if v is not None
                        }
                        serial = columns.get("流水号")
                        break
            else:
                for offset, header in enumerate(stream, 1):
                    if offset > 30:
                        break
                    columns = {
                        normalize_text(v): i for i, v in enumerate(header) if v is not None
                    }
                    if "流水号" in columns:
                        start_row, serial = offset, columns["流水号"]
                        break
            if serial is None:
                continue
            for index, row in enumerate(stream, (start_row or 1) + 1):
                if index > (start_row or 1) + 800:
                    break
                if serial >= len(row):
                    continue
                token = str(row[serial] or "").strip()
                if len(token) >= 4 and token[:2].isdigit() and token[2:4].isdigit():
                    month = int(token[:2])
                    if 1 <= month <= 12:
                        counts[month] = counts.get(month, 0) + 1
            if counts:
                break
    finally:
        wb.close()
    if not counts:
        return []
    dominant = max(counts, key=lambda month: counts[month])
    return [(None, dominant)]


def _base_report_months(path: Path) -> list[tuple[int | None, int]]:
    """基础报表原版：从「一干汇总」「二干基础」的封车时间采样。"""
    found: set[tuple[int, int]] = set()
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return []
    try:
        for sheet in BASE_SHEETS:
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            stream = ws.iter_rows(values_only=True)
            header = next(stream, None) or ()
            columns = {normalize_text(v): i for i, v in enumerate(header) if v is not None}
            date_col = columns.get("封车时间")
            if date_col is None:
                continue
            for index, row in enumerate(stream, 2):
                if index > 300:
                    break
                value = str(row[date_col] or "")[:10]
                match = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", value)
                if not match:
                    match = re.match(r"(\d{4})/(\d{1,2})/(\d{1,2})", value)
                if match:
                    found.add((int(match.group(1)), int(match.group(2))))
    finally:
        wb.close()
    return sorted(found)


def _probe_pair_features(path: Path) -> tuple[str | None, list[str]]:
    """结构像价格表/盈亏模板，但一干/二干待定时，用档位特征再判一次。"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None, []
    try:
        for ws in wb.worksheets:
            feature = _pair_feature(ws)
            if feature == "first":
                return "first", ["表内标题标注「一干」"]
            if feature == "second":
                return "second", ["表内标题标注「二干」"]
        return None, []
    finally:
        wb.close()


def _daily_sheet_counts(sheets: list[str]) -> tuple[int, int]:
    caotan = sum(1 for name in sheets if DAILY_CAOTAN.match(name.strip()))
    gangwu = sum(1 for name in sheets if DAILY_GANGWU.match(name.strip()))
    return caotan, gangwu


def detect_file(path: Path, month: tuple[int, int] | None = None) -> Detected:
    return _detect_from_profile(classify_workbook(path), month)


def refine_profiles(
    profiles: list[FileProfile], month: tuple[int, int] | None = None
) -> list[Detected]:
    """已有 FileProfile 时复用，避免重复打开工作簿。"""
    return [_detect_from_profile(profile, month) for profile in profiles]


def _detect_from_profile(profile: FileProfile, month: tuple[int, int] | None = None) -> Detected:
    path = profile.path
    role = profile.role
    evidence: list[str] = []
    confidence = "medium"
    sheets = list(profile.sheet_names)

    base_sheets = [name for name in sheets if name.strip() in BASE_SHEETS]
    if base_sheets:
        role, confidence = "base_report", "high"
        evidence.append("含 Sheet「%s」" % "」「".join(base_sheets))
    elif role in ("online_caotan", "online_gangwu"):
        caotan, gangwu = _daily_sheet_counts(sheets)
        if gangwu > caotan:
            role, confidence = "online_gangwu", "high"
            evidence.append(f"日 Sheet：港务格式 {gangwu} 个、草滩格式 {caotan} 个")
        elif caotan > gangwu:
            role, confidence = "online_caotan", "high"
            evidence.append(f"日 Sheet：草滩格式 {caotan} 个、港务格式 {gangwu} 个")
        elif caotan and caotan == gangwu:
            confidence = "low"
            evidence.append(f"两种日 Sheet 各 {caotan} 个，无法区分草滩/港务")
        else:
            evidence.append("无日 Sheet，保留分类器结果")
    elif role in ("first_price", "second_price"):
        confidence = "high"
        evidence.append("表头含 路线 + 公里 + 吨位/系数")
        evidence.append("表内标题或文件名标注「一干」" if role == "first_price"
                        else "表内标题或文件名标注「二干」")
    elif role in ("first_profit_template", "second_profit_template"):
        confidence = "high"
        evidence.append("含 7 列盈亏子表头")
        evidence.append("表内标题或文件名标注「一干」" if role == "first_profit_template"
                        else "表内标题或文件名标注「二干」")
    elif role in ("price_candidate", "profit_candidate"):
        feature_role, feature_evidence = _probe_pair_features(path)
        if feature_role == "first":
            role = "first_price" if role == "price_candidate" else "first_profit_template"
            confidence = "medium"
            evidence = feature_evidence
        elif feature_role == "second":
            role = "second_price" if role == "price_candidate" else "second_profit_template"
            confidence = "medium"
            evidence = feature_evidence
        else:
            confidence = "low"
            evidence.append("结构与价格表/盈亏模板一致，但表内标题和文件名都没有标注一干/二干")
    elif role == "unknown":
        confidence = "low"
        evidence.append("表结构未匹配任何已知角色")
    elif role == "damaged":
        confidence = "low"
        evidence.append("文件无法打开：" + ("；".join(profile.errors) or "未知原因"))
    elif role == "draft_template":
        confidence = "high"
        evidence.append("表头匹配 21 列基础报表格式")
    elif role == "vehicle":
        confidence = "high"
        evidence.append("表头含 邮路名称 + 邮路执行日期 + 派车单号 + 邮运量")
    elif role == "finance":
        confidence = "high"
        evidence.append("表头含 邮路名称 + 结算收入 + 普服收入")

    months = list(profile.months)
    if role == "finance" and not months:
        months = _finance_months(path, profile.header_row)
    elif role == "base_report" and not months:
        months = _base_report_months(path)
    if not months:
        token = parse_month_token(path.name)
        if token:
            months = [token]

    detected = Detected(
        path=path,
        role=role,
        confidence=confidence,
        evidence=evidence,
        months=months,
        sheets=sheets,
        profile=profile,
    )
    if month and months and role in MONTH_SCOPED:
        compatible = any(
            (year is None or year == month[0]) and mon == month[1]
            for year, mon in months
        )
        if not compatible:
            confidence = "low"
            detected.evidence.append(
                "文件月份 %s 与目标月 %04d-%02d 不一致"
                % ("、".join(f"{y or '?'}-{m:02d}" for y, m in months), month[0], month[1])
            )
    detected.confidence = confidence
    return detected


def detect_workspace(workspace: Path | str, month: str | None = None) -> list[Detected]:
    target = None
    if month:
        try:
            year, mon = (int(part) for part in str(month).split("-")[:2])
            target = (year, mon)
        except (TypeError, ValueError):
            target = None
    return [detect_file(path, target) for path in discover_workbooks(workspace)]


def resolve(detected: list[Detected], roles: dict[str, int]) -> dict:
    """按角色取文件。返回 {resolved, missing, ambiguous}。"""
    resolved: dict[str, list[str]] = {}
    missing: list[str] = []
    ambiguous: list[dict] = []
    for role, quantity in roles.items():
        candidates = [item for item in detected if item.role == role]
        # 月份对不上的候选只在没有其他选择时使用
        preferred = [item for item in candidates if item.confidence != "low"]
        if len(preferred) >= quantity:
            candidates = preferred
        if len(candidates) < quantity:
            missing.append(role)
            resolved[role] = [str(item.path) for item in candidates]
        elif len(candidates) == quantity:
            resolved[role] = [str(item.path) for item in candidates]
        else:
            ambiguous.append({
                "role": role,
                "role_label": ROLE_LABELS.get(role, role),
                "reason": f"需要 {quantity} 个，识别到 {len(candidates)} 个",
                "candidates": [item.to_dict() for item in candidates],
            })
            resolved[role] = [str(item.path) for item in candidates]
    return {"resolved": resolved, "missing": missing, "ambiguous": ambiguous}


def describe_input_gaps(result: dict) -> list[dict]:
    """把 missing / ambiguous / 损坏文件转成给用户看的补充清单。

    这是「还缺什么」，不是「执行失败」：入口在存在这些条目时返回
    ``status=needs_input``、退出码 0，Agent 应请用户补齐文件后重试。
    """
    gaps: list[dict] = []
    required = result.get("required_roles", {}) or {}
    resolved = result.get("resolved", {}) or {}
    for role in result.get("missing", []) or []:
        quantity = int(required.get(role, 1) or 1)
        found = len(resolved.get(role, []) or [])
        label = ROLE_LABELS.get(role, role)
        if quantity > 1:
            action = f"请上传{label}（需要 {quantity} 个，当前识别到 {found} 个）"
        else:
            action = f"请上传{label}"
        gaps.append({
            "kind": "missing",
            "role": role,
            "label": label,
            "required": quantity,
            "found": found,
            "action": action,
        })
    for entry in result.get("ambiguous", []) or []:
        role = entry.get("role", "")
        label = entry.get("role_label") or ROLE_LABELS.get(role, role)
        candidates = [item.get("name", "") for item in entry.get("candidates", [])]
        gaps.append({
            "kind": "ambiguous",
            "role": role,
            "label": label,
            "required": int(required.get(role, 1) or 1),
            "found": len(candidates),
            "candidates": candidates,
            "reason": entry.get("reason", ""),
            "action": f"请确认使用哪一个{label}（识别到 {len(candidates)} 个：{'、'.join(candidates)}），"
                      "删除多余文件后重试",
        })
    for item in result.get("files", []) or []:
        if item.get("role") != "damaged":
            continue
        gaps.append({
            "kind": "unreadable",
            "role": "damaged",
            "label": item.get("name") or "工作簿",
            "action": f"文件无法读取，请重新上传可打开的工作簿：{item.get('name', '')}",
        })
    return gaps


def needs_input_message(gaps: list[dict]) -> str:
    """给用户的一句话说明，只说缺什么/要确认什么，不说「执行失败」。"""
    labels: list[str] = []
    for gap in gaps:
        label = gap.get("label") or gap.get("role") or "输入文件"
        if label not in labels:
            labels.append(label)
    if not labels:
        return "当前输入不完整，请补充后重试。"
    return "当前上传的文件不完整，需要补充或确认：" + "、".join(labels) + "。"


def _vehicle_count() -> int:
    """车管文件数量跟 assets/company-map.json 的公司数一致，不写死 6。"""
    path = Path(__file__).resolve().parents[2] / "assets" / "company-map.json"
    try:
        return len(json.loads(path.read_text(encoding="utf-8"))["companies"])
    except Exception:
        return 6


def _phase_roles(phase: str) -> dict[str, int] | None:
    if phase == "profit":
        return {"base_report": 1, "finance": 1}
    if phase == "summary":
        return {"online_caotan": 1, "online_gangwu": 1}
    if phase == "review":
        return {
            "online_caotan": 1,
            "online_gangwu": 1,
            "vehicle": _vehicle_count(),
            "finance": 1,
            "draft_template": 1,
            "first_price": 1,
            "second_price": 1,
            "first_profit_template": 1,
            "second_profit_template": 1,
        }
    return None


# 版式模板和合同价格表都是半年甚至整年才换一次，客户上传了就用客户的（可能更新），
# 没上传就用包里的。动态数据（在线表、车管、财务）仍然必须每月上传。
BUNDLED_INPUT_ASSETS = {
    "draft_template": "draft-template.xlsx",
    "first_profit_template": "first-profit-template.xlsx",
    "second_profit_template": "second-profit-template.xlsx",
    "first_price": "first-price-table.xlsx",
    "second_price": "second-price-table.xlsx",
}
# 旧名字保留给已有调用方（inspect_inputs 只需 templates 子集）。
BUNDLED_TEMPLATE_ASSETS = {
    role: name
    for role, name in BUNDLED_INPUT_ASSETS.items()
    if role.endswith("_template")
}
PRICE_ROLES = ("first_price", "second_price")
# 一干/二干是成套的：只要用户传了其中一份（或传了一份分不出干线的候选），
# 就不能再拿内置的另一份凑数，否则等于把两版合同混着用。
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


def bundled_asset_path(role: str, assets_dir: Path | str | None = None) -> Path | None:
    """返回内置资产路径；该角色没有内置版本时返回 None。"""
    name = BUNDLED_INPUT_ASSETS.get(role)
    if not name:
        return None
    root = Path(assets_dir) if assets_dir else Path(__file__).resolve().parents[2] / "assets"
    path = root / name
    return path if path.is_file() else None


def _with_bundled_inputs(
    detected: list,
    roles: dict[str, int],
    assets_dir: Path | str | None = None,
    include_prices: bool = True,
) -> tuple[list, list[str]]:
    """为缺失的模板/价格表角色补上内置资产，用户不必每月重复上传。"""
    merged = list(detected)
    bundled: list[str] = []
    # 只看用户实际上传了什么：刚补进去的内置资产不算「已上传」，否则一干补齐后
    # 二干就会被当成「用户已有」而漏补。
    uploaded_roles = {item.role for item in detected}
    for role in roles:
        if role not in BUNDLED_INPUT_ASSETS:
            continue
        if not include_prices and role in PRICE_ROLES:
            continue
        if any(item.role == role for item in merged):
            continue
        related = BUNDLED_PAIRS.get(role)
        if related and uploaded_roles & set(related):
            # 用户已经传了这一套里的某一份（或一份分不出干线的候选），
            # 交给 resolve 报缺/报歧义，不要用内置版本悄悄补齐。
            continue
        path = bundled_asset_path(role, assets_dir)
        if path is None:
            continue
        item = detect_file(path)
        # 内置资产的角色是确定的，不再靠表内标题猜一干/二干。
        item.role = role
        item.confidence = "high"
        item.evidence.append(
            "使用 Skill 内置价格表" if role in PRICE_ROLES else "使用 Skill 内置模板"
        )
        merged.append(item)
        bundled.append(path.name)
    return merged, bundled


def detect_for_phase(workspace: Path | str, month: str | None, phase: str) -> dict:
    roles = _phase_roles(phase)
    if roles is None:
        raise ValueError(f"不支持的阶段: {phase}")
    uploaded = detect_workspace(workspace, month)
    detected, bundled = _with_bundled_inputs(uploaded, roles)
    resolution = resolve(detected, roles)
    bundled_prices = [
        name for name in bundled
        if name in {BUNDLED_INPUT_ASSETS[role] for role in PRICE_ROLES}
    ]
    return {
        "phase": phase,
        "target_month": month,
        "workspace": str(Path(workspace).resolve()),
        "file_count": len(uploaded),
        "bundled_inputs": bundled,
        "bundled_templates": [name for name in bundled if name not in bundled_prices],
        "bundled_price_tables": bundled_prices,
        "required_roles": {role: count for role, count in roles.items()},
        "files": [item.to_dict() for item in detected],
        **resolution,
    }
