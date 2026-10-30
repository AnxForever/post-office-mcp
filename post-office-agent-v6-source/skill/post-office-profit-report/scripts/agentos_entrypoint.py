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

import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import re
import runpy
import sys
from collections import Counter
from pathlib import Path

import openpyxl

from paths import resolve_input_dir

BASE_SHEETS = ("一干汇总", "二干基础")
FINANCE_HEADERS = ("结算收入", "普服收入")
DATE_HEADERS = ("封车时间", "邮路执行日期", "邮路执行时间")


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
        header = _first_header(sheet)
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
                rows = workbook[sheet_name].iter_rows(values_only=True)
                header = [str(v or "").strip() for v in next(rows)]
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
    args = parser.parse_args(argv)

    scripts_dir = Path(__file__).resolve().parent
    input_dir = resolve_input_dir(args.workspace)

    base_file, finance_file = args.base_file, args.finance_file
    detected: list[dict] = []
    if not base_file or not finance_file:
        for path in _excel_files(input_dir):
            role, info = identify(path)
            detected.append({"role": role, **info})
            if role == "base_report" and not base_file:
                base_file = path
            elif role == "finance" and not finance_file:
                finance_file = path

    missing = []
    if not base_file:
        missing.append({"role": "base_report", "label": "基础报表原版",
                        "action": "请上传基础报表原版（含「一干汇总」和「二干基础」两个工作表）"})
    if not finance_file:
        missing.append({"role": "finance", "label": "财务收入表",
                        "action": "请上传财务收入表（表头要有「结算收入」和「普服收入」）"})
    if missing:
        print(json.dumps({
            "ok": True, "status": "needs_input", "phase": "awaiting_inputs",
            "input_dir": str(input_dir), "detected": detected,
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

    output_dir = args.out or (args.workspace / "output")
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
