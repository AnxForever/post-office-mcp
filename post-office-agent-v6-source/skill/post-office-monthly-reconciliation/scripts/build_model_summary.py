from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from reconciliation.excel_io import file_sha256, read_online_rows
from reconciliation.model_summary import merge_online_rows, write_model_summary_workbook


def build_model_summary(
    caotan: Path | str,
    gangwu: Path | str,
    output_dir: Path | str,
    month: str,
) -> dict[str, object]:
    try:
        parsed = datetime.strptime(month, "%Y-%m")
    except ValueError as exc:
        raise ValueError("month 必须使用 YYYY-MM") from exc
    caotan_path, gangwu_path = Path(caotan), Path(gangwu)
    if not caotan_path.is_file() or not gangwu_path.is_file():
        raise ValueError("草滩和港务在线测量表必须同时存在")

    target_month = (parsed.year, parsed.month)
    online_rows = read_online_rows(caotan_path, target_month, "草滩")
    online_rows += read_online_rows(gangwu_path, target_month, "港务")
    if not online_rows:
        raise ValueError(f"在线测量表不包含目标月份 {month} 的有效记录")
    summary_rows = merge_online_rows(online_rows)
    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    output = output_root / f"{parsed.month}月_01_车型汇总表.xlsx"
    write_model_summary_workbook(summary_rows, output, target_month)
    nonzero = sum(row.get("measured_volume") not in (None, "", 0, 0.0) for row in summary_rows)
    return {
        "ok": True,
        "phase": "summary_completed",
        "target_month": month,
        "outputs": [str(output.resolve())],
        "metrics": {
            "online_rows": len(online_rows),
            "summary_rows": len(summary_rows),
            "nonzero_volume": nonzero,
            "unmeasured_or_inconsistent": len(summary_rows) - nonzero,
        },
        "sha256": file_sha256(output),
        "next_action": "如需完整月度报表，请提供全部当月输入并执行 review 阶段",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="仅使用草滩和港务在线表生成车型汇总表")
    parser.add_argument("--caotan", required=True, type=Path)
    parser.add_argument("--gangwu", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--month", required=True)
    try:
        args = parser.parse_args(argv)
        result = build_model_summary(args.caotan, args.gangwu, args.output_dir, args.month)
    except Exception as exc:
        result = {"ok": False, "phase": "failed", "errors": [str(exc)], "next_action": "请修正在线表或目标月份"}
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    sys.exit(main())
