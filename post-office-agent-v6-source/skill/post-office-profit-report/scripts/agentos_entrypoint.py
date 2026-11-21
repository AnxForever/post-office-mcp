# -*- coding: utf-8 -*-
"""盈亏表技能入口。

只做一件事：找到「基础报表原版」和「财务收入表」，按月生成一干/二干盈亏表。

用法：
    python agentos_entrypoint.py --phase profit --workspace /workspace
    python agentos_entrypoint.py --phase profit --base-file A.xlsx --finance-file B.xlsx --month 2026-09

不传 --month 时按基础报表的业务日期推断；推断不出来返回 status=needs_month，由上层问用户。
文件不按文件名认，按表结构认：含「一干汇总/二干基础」的是基础报表，
表头同时有「结算收入」「普服收入」的是财务收入表。
"""
from __future__ import annotations

import sys
from pathlib import Path

# 独立运行时不依赖调用方设置 sys.path：把本脚本目录放到最前面，
# 保证 reconciliation / build_* 等本地模块优先从本技能包导入。
_SCRIPTS_DIR = Path(__file__).resolve().parent
if not (_SCRIPTS_DIR / "build_profit_tables.py").is_file():
    # 平台脚本执行器可能把入口文件暂存到别的目录；按技能名回到真实技能包。
    for _base in (Path("/workspace/skills"), Path("/workspace")):
        _candidate = _base / "post-office-profit-report" / "scripts"
        if (_candidate / "build_profit_tables.py").is_file():
            _SCRIPTS_DIR = _candidate
            break
_SCRIPTS_DIR = str(_SCRIPTS_DIR)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import re
import runpy
from collections import Counter

import openpyxl

from paths import resolve_input_dir

BASE_SHEETS = ("一干汇总", "二干基础")
FINANCE_HEADERS = ("结算收入", "普服收入")
DATE_HEADERS = ("封车时间", "邮路执行日期", "邮路执行时间")
# 文件名里出现这些字样的，通常是客户上传的财务收入表（运输局导出）。
FINANCE_NAME_HINTS = ("运输局", "财务", "收入")


def _excel_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(root.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        if {"skills", "output", "__pycache__"} & set(path.parts):
            continue
        files.append(path)
    return files


def _first_header(ws) -> list[str]:
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return [str(v or "").strip() for v in row]
    return []


def _sheet_header(ws) -> list[str]:
    """读第一行表头；只读到一格时丢掉工作簿里写坏的维度缓存再读一次。

    客户从邮政系统导出的运输局表常带 `<dimension ref="A1"/>`，
    read_only 模式信任这个标签，只吐出 A1 一个格子（表头变成 ['账']），
    于是财务收入表被当成认不出来的文件跳过。2026-11 线上 7/8 月两次失败都是这个原因。
    """
    header = _first_header(ws)
    if sum(1 for value in header if value) > 1:
        return header
    try:
        ws.reset_dimensions()
    except Exception:  # 非 read_only 工作表的 reset_dimensions 是空实现，失败也不影响
        return header
    return _first_header(ws)


def _column_index(header: list[str], *names: str) -> int | None:
    for i, value in enumerate(header):
        if value in names:
            return i
    return None


def _numbers_by_column(ws, index: int, limit: int = 4000) -> list[str]:
    """扫一列文本值，用来判断财务收入表属于哪个月（流水号前 4 位是 MMDD）。"""
    values: list[str] = []
    for row in ws.iter_rows(min_row=2, min_col=index + 1, max_col=index + 1, values_only=True):
        values.append(str((row or [""])[0] or "").strip())
        if len(values) >= limit:
            break
    return values


def probe_finance_month(path: Path) -> dict:
    """从财务收入表的流水号（MMDD）推断它覆盖的月份，用于和基础报表对月。

    只读流水号一列，不解压整张表；认不出来就返回 {"month": None}，不拦截。
    """
    try:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
    except Exception as exc:
        return {"file": path.name, "month": None, "error": str(exc)}
    try:
        sheet = workbook[workbook.sheetnames[0]]
        header = _sheet_header(sheet)
        index = _column_index(header, "流水号")
        if index is None:
            return {"file": path.name, "month": None}
        counter: Counter[str] = Counter()
        for value in _numbers_by_column(sheet, index):
            if len(value) >= 4 and value[:4].isdigit():
                counter[value[:2]] += 1
        if not counter:
            return {"file": path.name, "month": None}
        mmdd, count = counter.most_common(1)[0]
        if int(mmdd) < 1 or int(mmdd) > 12 or count < 10:
            return {"file": path.name, "month": None, "samples": dict(counter.most_common(4))}
        return {
            "file": path.name,
            "month": int(mmdd),
            "rows": count,
            "samples": dict(counter.most_common(4)),
        }
    finally:
        workbook.close()


def _pick_base(candidates: list[Path], month: str | None) -> Path:
    """多份基础报表时挑一份：优先业务日期就是目标月份的，其次最近上传的。"""
    if len(candidates) == 1:
        return candidates[0]
    if month:
        wanted = str(month)
        for path in candidates:
            if infer_month(path).get("month") == wanted:
                return path
    return max(candidates, key=lambda path: path.stat().st_mtime)


def _pick_finance(candidates: list[Path], month: str, probes: list[dict]) -> Path | None:
    """多份财务收入表时按流水号月份挑；没有一张对得上就返回 None，绝不随便顶替。"""
    try:
        target = int(str(month).split("-")[1])
    except (IndexError, ValueError):
        target = None
    ordered = sorted(candidates, key=lambda path: path.stat().st_mtime, reverse=True)
    for path in ordered:
        probes.append(probe_finance_month(path))
    if target is None:
        return ordered[0] if ordered else None
    matches = [p for p in probes if p.get("month") == target]
    if matches:
        wanted = {p["file"] for p in matches}
        return next(path for path in ordered if path.name in wanted)
    # 月份读不出来（缺流水号列等）的表不能一票否决，交给出表前的对账闸门兜底。
    unknown = [p for p in probes if p.get("month") is None]
    if unknown:
        wanted = {p["file"] for p in unknown}
        return next(path for path in ordered if path.name in wanted)
    return None


def identify(path: Path) -> tuple[str | None, dict]:
    """按表结构判断这个工作簿是 base（基础报表原版）还是 finance（财务收入表）。"""
    try:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
    except Exception as exc:  # 损坏、加密、被占用
        return None, {"file": path.name, "error": str(exc)}
    try:
        names = set(workbook.sheetnames)
        if all(sheet in names for sheet in BASE_SHEETS):
            return "base_report", {"file": path.name, "sheets": list(names)[:8]}
        sheet = workbook[workbook.sheetnames[0]]
        header = _sheet_header(sheet)
        if all(any(h == name for h in header) for name in FINANCE_HEADERS):
            return "finance", {"file": path.name, "headers": header[:12]}
        return None, {"file": path.name, "headers": header[:12]}
    finally:
        workbook.close()


def infer_month(base_path: Path) -> dict:
    """从基础报表的业务日期推断目标月。返回 {"month": "2026-09"} 或 {"month": None, "questions": [...]}。"""
    counter: Counter[tuple[int, int]] = Counter()
    try:
        workbook = openpyxl.load_workbook(base_path, read_only=True, data_only=True, keep_links=False)
    except Exception:
        workbook = None
    if workbook is not None:
        try:
            for sheet_name in BASE_SHEETS:
                if sheet_name not in workbook.sheetnames:
                    continue
                sheet = workbook[sheet_name]
                rows = sheet.iter_rows(values_only=True)
                header = [str(v or "").strip() for v in next(rows, ())]
                # 写坏的 <dimension ref="A1"/> 会让表头只读出一格，丢掉维度缓存再读。
                if sum(1 for value in header if value) <= 1:
                    try:
                        sheet.reset_dimensions()
                    except Exception:
                        pass
                    else:
                        rows = sheet.iter_rows(values_only=True)
                        header = [str(v or "").strip() for v in next(rows, ())]
                col = next((i for i, name in enumerate(header) if name in DATE_HEADERS), None)
                if col is None:
                    continue
                for row in rows:
                    if not (row and isinstance(row[0], int)) or col >= len(row):
                        continue
                    stamp = _year_month(row[col])
                    if stamp:
                        counter[stamp] += 1
        finally:
            workbook.close()
    if counter:
        (year, month), _count = counter.most_common(1)[0]
        return {"month": f"{year:04d}-{month:02d}", "evidence": {"业务日期": dict(counter)}}

    match = re.search(r"(20\d{2})\D{0,2}(\d{1,2})\s*月", base_path.name) or \
        re.search(r"(\d{1,2})\s*月", base_path.name)
    if match:
        groups = match.groups()
        if len(groups) == 2:
            return {"month": f"{int(groups[0]):04d}-{int(groups[1]):02d}", "evidence": {"文件名": base_path.name}}
        return {
            "month": None,
            "questions": [f"文件名里只看到「{groups[0]}月」，是哪一年？"],
            "evidence": {"文件名": base_path.name},
        }
    return {
        "month": None,
        "questions": ["这些报表是哪一年哪一個月的？请告诉我月份（例如 2026 年 9 月）。"],
        "evidence": {},
    }


def _year_month(value):
    if hasattr(value, "year") and hasattr(value, "month"):
        return (value.year, value.month)
    text = str(value or "").strip()
    match = re.match(r"(\d{4})\s*[-/.年]\s*(\d{1,2})", text)
    if match:
        return (int(match.group(1)), int(match.group(2)))
    return None


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="生成一干/二干盈亏表")
    parser.add_argument("--phase", default="profit", choices=["profit"])
    parser.add_argument("--workspace", type=Path, default=Path("/workspace"))
    parser.add_argument("--month", help="YYYY-MM，不传则从基础报表推断")
    parser.add_argument("--base-file", type=Path)
    parser.add_argument("--finance-file", type=Path)
    parser.add_argument("--out", type=Path, help="输出目录，默认 <workspace>/output")
    parser.add_argument("--tpl-dir", type=Path, help="可选：盈亏模板目录，只用于补历史路线占位")
    parser.add_argument("--no-template-routes", action="store_true")
    parser.add_argument("--force", action="store_true", help="确认覆盖输出目录里已有的盈亏表")
    args = parser.parse_args(argv)

    scripts_dir = Path(__file__).resolve().parent
    input_dir = resolve_input_dir(args.workspace)

    base_file, finance_file = args.base_file, args.finance_file
    detected: list[dict] = []
    base_candidates: list[Path] = []
    finance_candidates: list[Path] = []
    unrecognized: list[Path] = []
    if not base_file or not finance_file:
        for path in _excel_files(input_dir):
            role, info = identify(path)
            detected.append({"role": role, **info})
            if role == "base_report":
                base_candidates.append(path)
            elif role == "finance":
                finance_candidates.append(path)
            elif any(hint in path.name for hint in FINANCE_NAME_HINTS):
                unrecognized.append(path)

    if not base_file and base_candidates:
        base_file = _pick_base(base_candidates, args.month)

    missing = []
    if not base_file:
        missing.append({"role": "base_report", "label": "基础报表原版",
                        "action": "请上传基础报表原版（含「一干汇总」和「二干基础」两个工作表）"})
    if not finance_file and not finance_candidates:
        missing.append({"role": "finance", "label": "财务收入表",
                        "action": "请上传财务收入表（表头要有「结算收入」和「普服收入」）"})
    if missing:
        print(json.dumps({
            "ok": True, "status": "needs_input", "phase": "awaiting_inputs",
            "input_dir": str(input_dir), "detected": detected,
            "unrecognized_files": [path.name for path in unrecognized],
            "missing_inputs": missing,
            "message": "当前上传的文件不完整，需要补充：" + "、".join(m["label"] for m in missing) + "。",
            "next_action": "请补充后重试。",
        }, ensure_ascii=False, indent=2))
        return 0

    month = args.month
    if not month:
        inference = infer_month(Path(base_file))
        if inference.get("month") is None:
            print(json.dumps({
                "ok": True, "status": "needs_month", "phase": "needs_month",
                "questions": inference["questions"],
                "month_evidence": inference.get("evidence", {}),
            }, ensure_ascii=False, indent=2))
            return 0
        month = inference["month"]

    # 基础报表的月份定下来之后再挑财务表：输入目录里往往还留着别的月份的运输局表，
    # 只按文件名先后取第一张，会拿上个月的表硬配当月的基础报表，
    # 最后在出表闸门报「两条数据对不上」，把文件选错说成数据问题。
    finance_probes: list[dict] = []
    if not finance_file and finance_candidates:
        finance_file = _pick_finance(finance_candidates, month, finance_probes)
        if finance_file is None:
            print(json.dumps({
                "ok": False, "status": "finance_month_mismatch", "phase": "needs_input",
                "month": month,
                "base_file": Path(base_file).name,
                "finance_candidates": finance_probes,
                "unrecognized_files": [path.name for path in unrecognized],
                "message": f"输入目录里的财务收入表没有一张是 {month} 的，"
                           f"和基础报表原版「{Path(base_file).name}」对不上月份。",
                "next_action": f"请上传 {month} 的运输局（财务收入）表；"
                               "目录里其他月份的表不会被拿来顶替。",
            }, ensure_ascii=False, indent=2))
            return 2
        chosen = next((p for p in finance_probes if p["file"] == Path(finance_file).name), None)
        if chosen and chosen.get("month"):
            print(f"[OK] 财务收入表: {Path(finance_file).name}（流水号月份 {chosen['month']:02d}）")

    output_dir = args.out or (args.workspace / "output")
    try:
        month_number = int(str(month).split("-")[1])
    except (IndexError, ValueError):
        print(json.dumps({
            "ok": False, "errors": ["month 必须使用 YYYY-MM"],
            "next_action": "请传入正确的月份",
        }, ensure_ascii=False, indent=2))
        return 0
    existing = [
        output_dir / f"{month_number}月_04_{level}盈亏表.xlsx"
        for level in ("一干", "二干")
    ]
    if not args.force and any(path.exists() for path in existing):
        found = [str(path.resolve()) for path in existing if path.exists()]
        print(json.dumps({
            "ok": True,
            "status": "output_exists",
            "phase": "needs_confirm",
            "outputs": found,
            "message": "输出目录里已经有这个月的盈亏表，可能是上次人工调整过的版本。直接重新生成会覆盖。",
            "next_action": "确认要覆盖时带 --force 重跑；只想保留旧版，请把输出目录换成一个新目录。",
        }, ensure_ascii=False, indent=2))
        return 0
    child = [
        "build_profit_tables.py",
        "--base", str(base_file),
        "--finance", str(finance_file),
        "--month", month,
        "--out", str(output_dir),
    ]
    if args.tpl_dir:
        child += ["--tpl-dir", str(args.tpl_dir)]
    if args.no_template_routes:
        child += ["--no-template-routes"]

    sys.argv[:] = child
    runpy.run_path(str(scripts_dir / "build_profit_tables.py"), run_name="__main__")
    return 0


if __name__ == "__main__":
    try:
        code = main()
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 0
    raise SystemExit(code)
