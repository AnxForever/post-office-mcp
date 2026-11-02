# -*- coding: utf-8 -*-
"""四个独立技能串跑测试：01 -> 02 -> 03 -> 盈亏识别。

    python post-office-agent-v6-source/tests/test_independent_skills.py

依赖本地 4 月数据包（不入库）。脚本输出 [OK]/[FAIL]，退出码 0 表示全部通过。
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

PACKAGE = Path(__file__).resolve().parents[1]
ROOT = PACKAGE.parent
SKILLS = PACKAGE / "skill"
APRIL = ROOT / "4月数据包"
ENV = dict(
    os.environ,
    PYTHONIOENCODING="utf-8",
    PYTHONUTF8="1",
    POST_OFFICE_HARD_EXIT="1",
    PYTHONFAULTHANDLER="1",
)
NATIVE_CRASH = {-1073741819, -1073741571, -1073740791, 3221225477, 3221225725, 3221226505}

RESULTS: list[tuple[str, bool, str]] = []


def record(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, ok, detail))
    print(f"[{'OK' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail and not ok else ""))


def run(skill: str, *args: str, expect_code: int = 0) -> tuple[int, dict]:
    entry = SKILLS / skill / "scripts" / "agentos_entrypoint.py"
    for attempt in range(1, 9):
        proc = subprocess.run(
            [sys.executable, str(entry), *args],
            cwd=str(entry.parent),
            env=ENV,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if proc.returncode not in NATIVE_CRASH:
            break
        print(f"[WARN] 原生崩溃，重跑一次: {skill} (exit {proc.returncode})")
    text = proc.stdout.strip()
    start = text.find("{")
    payload = {}
    if start >= 0:
        try:
            payload = json.loads(text[start:])
        except json.JSONDecodeError:
            payload = {"_raw": text[:400]}
    else:
        payload = {"_raw": text[:400], "_stderr": proc.stderr[-400:]}
    record(f"{skill} exit {expect_code}", proc.returncode == expect_code, f"got {proc.returncode}; {payload}")
    return proc.returncode, payload


def main() -> int:
    if not APRIL.is_dir():
        print(f"[FAIL] 缺少测试数据目录: {APRIL}")
        return 2
    with tempfile.TemporaryDirectory(prefix="post-office-independent-") as temp:
        work = Path(temp)

        # 01 车型汇总
        _, summary = run(
            "post-office-model-summary",
            "--caotan", str(APRIL / "草滩在线表.xlsx"),
            "--gangwu", str(APRIL / "港务在线表.xlsx"),
            "--month", "2026-04",
            "--output-dir", str(work / "01"),
        )
        summary_file = work / "01" / "4月_01_车型汇总表.xlsx"
        record("01 生成", summary.get("phase") == "summary_completed" and summary_file.is_file())
        record("01 行数", summary.get("metrics", {}).get("summary_rows") == 3224, str(summary.get("metrics")))

        # 02 基础报表初版
        _, draft = run(
            "post-office-draft-report",
            "--input-dir", str(APRIL),
            "--model-file", str(summary_file),
            "--month", "2026-04",
            "--output-dir", str(work / "02"),
        )
        draft_file = work / "02" / "4月_02_基础报表初版.xlsx"
        record("02 生成", draft.get("status") == "draft_completed" and draft_file.is_file())
        record("02 行数", draft.get("metrics", {}).get("draft_rows") == 4309, str(draft.get("metrics")))
        record("02 价格匹配", draft.get("metrics", {}).get("price_matched") == 3628, str(draft.get("metrics")))

        # 已有 02 时不自动覆盖，避免把人工核对结果冲掉。
        _, draft_again = run(
            "post-office-draft-report",
            "--input-dir", str(APRIL),
            "--model-file", str(summary_file),
            "--month", "2026-04",
            "--output-dir", str(work / "02"),
        )
        record("02 不覆盖已有输出", draft_again.get("status") == "output_exists", str(draft_again)[:300])

        # 03 拆分
        _, split = run(
            "post-office-split-report",
            "--review", str(draft_file),
            "--month", "2026-04",
            "--output-dir", str(work / "03"),
        )
        split_file = work / "03" / "4月_03_基础报表拆分版.xlsx"
        record("03 生成", split.get("status") == "split_completed" and split_file.is_file())
        record("03 记录守恒", split.get("metrics", {}).get("split_rows") == 4309, str(split.get("metrics")))
        record(
            "03 元数据全部用上",
            split.get("metrics", {}).get("metadata_missing") == 0
            and split.get("metrics", {}).get("metadata_unused") == 0,
            str(split.get("metrics")),
        )

        # 用户在 02 里插一行，03 仍应保持同样的分组（元数据按业务键+邮路匹配）。
        inserted = work / "02_ins" / draft_file.name
        inserted.parent.mkdir(parents=True, exist_ok=True)
        workbook = load_workbook(draft_file)
        workbook["基础报表"].insert_rows(3)
        workbook.save(inserted)
        workbook.close()
        _, split_ins = run(
            "post-office-split-report",
            "--review", str(inserted),
            "--month", "2026-04",
            "--output-dir", str(work / "03_ins"),
        )
        normal_metrics = split.get("metrics", {})
        inserted_metrics = split_ins.get("metrics", {})
        record(
            "03 容忍 02 插行",
            split_ins.get("status") == "split_completed"
            and inserted_metrics.get("unmatched") == normal_metrics.get("unmatched")
            and inserted_metrics.get("unclassified_rows") == normal_metrics.get("unclassified_rows")
            and inserted_metrics.get("metadata_missing") == 0,
            str(inserted_metrics),
        )

        # 已有 03 时不自动覆盖。
        _, split_again = run(
            "post-office-split-report",
            "--review", str(draft_file),
            "--month", "2026-04",
            "--output-dir", str(work / "03"),
        )
        record("03 不覆盖已有输出", split_again.get("status") == "output_exists", str(split_again)[:300])

        # 盈亏技能要能认出 03。未核对初版缺吨位，预期被硬闸门拦下而不是静默出表。
        _, profit = run(
            "post-office-profit-report",
            "--phase", "profit",
            "--base-file", str(split_file),
            "--finance-file", str(APRIL / "2026.04财务收入表.xlsx"),
            "--month", "2026-04",
            "--out", str(work / "04"),
            expect_code=2,
        )
        record(
            "04 认出 03 并拦住缺吨位行",
            profit.get("phase") == "base_report_check" and profit.get("status") == "base_rows_dropped",
            str(profit)[:300],
        )

    failed = [name for name, ok, _ in RESULTS if not ok]
    print(f"\n{len(RESULTS) - len(failed)}/{len(RESULTS)} 通过")
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
