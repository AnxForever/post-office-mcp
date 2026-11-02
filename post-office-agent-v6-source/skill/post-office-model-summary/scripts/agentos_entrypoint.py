# -*- coding: utf-8 -*-
"""车型汇总表技能入口。

只做一件事：草滩在线测量表 + 港务在线测量表 -> N月_01_车型汇总表.xlsx。

    python agentos_entrypoint.py --workspace /workspace
    python agentos_entrypoint.py --workspace /workspace --month 2026-09
    python agentos_entrypoint.py --caotan A.xlsx --gangwu B.xlsx --output-dir out --month 2026-09

不传 --month 时从文件名/日 Sheet 名推断；推断不确定时返回 needs_month/needs_confirm，
由上层用大白话问用户，不硬猜。
"""
from __future__ import annotations

from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import sys
from pathlib import Path

from build_model_summary import build_model_summary
from month import infer_month_from_inputs
from reconciliation.detect import describe_input_gaps, detect_for_phase, needs_input_message
from reconciliation.paths import resolve_input_dir


def _emit(payload: dict) -> int:
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="草滩+港务在线表生成车型汇总表")
    parser.add_argument("--workspace", type=Path, default=Path("/workspace"))
    parser.add_argument("--input-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--month", help="YYYY-MM，不传则从文件推断")
    parser.add_argument("--caotan", type=Path)
    parser.add_argument("--gangwu", type=Path)
    args = parser.parse_args(argv)

    workspace = args.workspace.resolve()
    input_dir = args.input_dir.resolve() if args.input_dir else resolve_input_dir(workspace)
    month = args.month

    if not month:
        inference = infer_month_from_inputs(
            input_dir,
            roles={"online_caotan", "online_gangwu", "online_measure"},
        )
        if inference["month"] is None:
            return _emit({
                "ok": True,
                "status": "needs_month",
                "phase": "needs_month",
                "questions": inference["questions"],
                "month_evidence": inference.get("evidence", {}),
                "conflicts": inference.get("conflicts", []),
                "next_action": "请把疑问转告用户，拿到月份后重试。",
            })
        month = inference["month"]
        if inference.get("confidence") != "high":
            return _emit({
                "ok": True,
                "status": "needs_confirm",
                "phase": "needs_confirm",
                "target_month": month,
                "questions": inference.get("questions", []),
                "month_evidence": inference.get("evidence", {}),
                "conflicts": inference.get("conflicts", []),
                "next_action": "月份是按文件推断的，请与用户确认后再执行。",
            })

    caotan, gangwu = args.caotan, args.gangwu
    if not caotan or not gangwu:
        result = detect_for_phase(input_dir, month, "summary")
        resolved = result.get("resolved", {})
        if not caotan and len(resolved.get("online_caotan", [])) == 1:
            caotan = Path(resolved["online_caotan"][0])
        if not gangwu and len(resolved.get("online_gangwu", [])) == 1:
            gangwu = Path(resolved["online_gangwu"][0])
        if not caotan or not gangwu:
            gaps = describe_input_gaps(result)
            if not gaps:
                gaps = [{
                    "kind": "ambiguous",
                    "role": "online",
                    "label": "草滩/港务在线测量表",
                    "action": "请分别上传草滩和港务两张在线测量表",
                }]
            return _emit({
                "ok": True,
                "status": "needs_input",
                "phase": "awaiting_inputs",
                "task_phase": "summary",
                "target_month": month,
                "detected": result.get("files", []),
                "missing": result.get("missing", []),
                "ambiguous": result.get("ambiguous", []),
                "missing_inputs": gaps,
                "message": needs_input_message(gaps),
                "next_action": "请补充或确认草滩/港务在线测量表后重试。",
            })

    output_dir = args.output_dir or (workspace / "output")
    result = build_model_summary(caotan, gangwu, output_dir, month)
    result.setdefault("status", result.get("phase"))
    return _emit(result)


if __name__ == "__main__":
    try:
        code = main()
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 0
    raise SystemExit(code)
