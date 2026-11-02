# -*- coding: utf-8 -*-
"""基础报表拆分技能入口。

输入：人工核对后的 02 基础报表初版（含「基础报表」「拆分元数据」）。
输出：N月_03_基础报表拆分版.xlsx。

    python agentos_entrypoint.py --workspace /workspace
    python agentos_entrypoint.py --review 4月_02_基础报表初版.xlsx --month 2026-04
    python agentos_entrypoint.py --input-dir in --output-dir out

不依赖 run-state，可以单独任务运行；02 是上一阶段交付件。
"""
from __future__ import annotations

from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
from pathlib import Path

from openpyxl import load_workbook

from build_split_report import build_split_report, infer_month_from_review
from reconciliation.excel_io import file_sha256
from reconciliation.paths import resolve_input_dir


def _emit(payload: dict) -> int:
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload.get("ok", True) else 2


def discover_review_files(input_dir: Path) -> list[Path]:
    """按表结构找 02：同时含「基础报表」和「拆分元数据」。"""
    candidates: list[Path] = []
    for path in sorted(input_dir.rglob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        try:
            workbook = load_workbook(path, read_only=True, keep_links=False)
        except Exception:
            continue
        try:
            sheets = {str(name).strip() for name in workbook.sheetnames}
        finally:
            workbook.close()
        if {"基础报表", "拆分元数据"}.issubset(sheets):
            candidates.append(path)
    return candidates


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="把基础报表初版拆成路线页签（03）")
    parser.add_argument("--workspace", type=Path, default=Path("/workspace"))
    parser.add_argument("--input-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--review", type=Path, help="人工核对后的 02 基础报表初版")
    parser.add_argument("--month", help="YYYY-MM，不传则从 02 的封车时间推断")
    parser.add_argument("--first-price", type=Path)
    parser.add_argument("--second-price", type=Path)
    parser.add_argument("--force", action="store_true", help="确认覆盖输出目录里已有的 03")
    args = parser.parse_args(argv)

    workspace = args.workspace.resolve()
    input_dir = args.input_dir.resolve() if args.input_dir else resolve_input_dir(workspace)
    output_dir = args.output_dir or (workspace / "output")
    assets = Path(__file__).resolve().parents[1] / "assets"

    review = args.review
    if not review:
        candidates = discover_review_files(input_dir)
        if len(candidates) > 1:
            return _emit({
                "ok": True,
                "status": "needs_input",
                "phase": "awaiting_inputs",
                "missing_inputs": [{
                    "kind": "ambiguous",
                    "role": "review",
                    "label": "基础报表初版",
                    "found": len(candidates),
                    "candidates": [path.name for path in candidates],
                    "action": "识别到多份基础报表初版，请只保留当月一份，或显式指定 --review",
                }],
                "message": "识别到多份基础报表初版，无法确定拆分哪一份。",
                "next_action": "请只保留当月一份 02 后重试。",
            })
        if not candidates:
            return _emit({
                "ok": True,
                "status": "needs_input",
                "phase": "awaiting_inputs",
                "missing_inputs": [{
                    "kind": "missing",
                    "role": "review",
                    "label": "基础报表初版",
                    "action": "请上传上一阶段人工核对后的 N月_02_基础报表初版.xlsx",
                }],
                "message": "当前上传的文件里没有基础报表初版（需要含「拆分元数据」）。",
                "next_action": "请先跑基础报表初版技能，或上传 02 后重试。",
            })
        review = candidates[0]
    review = Path(review).resolve()
    if not review.is_file():
        return _emit({"ok": False, "errors": [f"基础报表初版不存在: {review}"], "next_action": "请检查文件路径"})

    month = args.month
    if not month:
        inference = infer_month_from_review(review)
        if not inference.get("month"):
            return _emit({
                "ok": True,
                "status": "needs_month",
                "phase": "needs_month",
                "questions": inference.get("questions", []),
                "next_action": "请把疑问转告用户，拿到月份后带 --month 重试。",
            })
        month = inference["month"]

    first_price = args.first_price or (assets / "first-price-table.xlsx")
    second_price = args.second_price or (assets / "second-price-table.xlsx")
    split_template = assets / "split-report-template.xlsx"
    for role, path in (("first_price", first_price), ("second_price", second_price), ("split_template", split_template)):
        if not Path(path).is_file():
            return _emit({"ok": False, "errors": [f"缺少{role}: {path}"], "next_action": "请检查技能包资产"})

    try:
        year, month_number = map(int, month.split("-"))
    except ValueError:
        return _emit({"ok": False, "errors": ["month 必须使用 YYYY-MM"], "next_action": "请传入正确的月份"})
    expected = Path(output_dir) / f"{month_number}月_03_基础报表拆分版.xlsx"
    if expected.exists() and not args.force:
        return _emit({
            "ok": True,
            "status": "output_exists",
            "phase": "needs_confirm",
            "output": str(expected.resolve()),
            "message": "输出目录里已经有这个月的 03 拆分版，可能是上次人工调整过的版本。直接重新生成会覆盖。",
            "next_action": "确认要覆盖时带 --force 重跑；只想保留旧版，请把输出目录换成一个新目录。",
        })

    try:
        result = build_split_report(
            review,
            Path(output_dir) / f"{review.stem}.xlsx",
            month=month,
            first_price=first_price,
            second_price=second_price,
            split_template=split_template,
        )
    except Exception as exc:
        return _emit({"ok": False, "errors": [str(exc)], "next_action": "请检查 02 是否被插删行或表头被改"})

    output = Path(result["output"])
    return _emit({
        "ok": True,
        "status": "split_completed",
        "phase": "split_completed",
        "target_month": month,
        "outputs": [str(output.resolve())],
        "metrics": {
            "draft_rows": result["rows"],
            "split_rows": sum(result["sheets"].values()),
            "unmatched": result["unmatched"],
            "metadata_missing": result["metadata_missing"],
            "metadata_unused": result["metadata_unused"],
            "unclassified_rows": result["issue_rows"],
            "sheet_count": result["total_sheets"],
        },
        "sheets": result["sheets"],
        "sha256": file_sha256(output),
        "next_action": "拆分版可直接作为盈亏技能的输入；请抽查「※未匹配目的地」页签。",
    })


if __name__ == "__main__":
    try:
        code = main()
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 0
    raise SystemExit(code)
