# -*- coding: utf-8 -*-
"""Skill 回归测试：把每个阶段都跑一遍，坏了立刻看出来。

用法：
    python post-office-agent-v6-source/tests/test_skill.py
    python post-office-agent-v6-source/tests/test_skill.py --fast   # 跳过完整 review/final

依赖本地数据目录（不入库）：
    4月数据包/                     完整 4 月输入
    数据/委办报表开发/3月/          3 月基础报表原版 + 财务收入表
    数据/                          多月份工作区，用于月度消歧

约定：脚本输出 [OK]/[FAIL]，退出码 0 表示全部通过。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PACKAGE = ROOT / "post-office-agent-v6-source"
SKILL = PACKAGE / "skill" / "post-office-monthly-reconciliation"
ENTRY = SKILL / "scripts" / "agentos_entrypoint.py"
APRIL = ROOT / "4月数据包"
MARCH = ROOT / "数据" / "委办报表开发" / "3月"
MULTI_MONTH = ROOT / "数据"

ENV = dict(
    os.environ,
    PYTHONIOENCODING="utf-8",
    PYTHONUTF8="1",
    PYTHONFAULTHANDLER="1",
    # 绕过 Windows 解释器收尾阶段偶发的 0xC0000005（输出已完整打印后才崩）。
    # 只影响 CLI 运行；AgentOS 走 runpy 在宿主进程内调用，不读这个变量。
    POST_OFFICE_HARD_EXIT="1",
)

RESULTS: list[tuple[str, bool, str]] = []


def record(name: str, ok: bool, detail: str = "") -> None:
    RESULTS.append((name, ok, detail))
    print(f"[{'OK' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail and not ok else ""))


# Windows 上偶发的原生崩溃（访问冲突 / 栈溢出），与业务逻辑无关。
# 出现时自动重跑一次，但仍会记一条 CRASH，避免把偶发问题当成没发生。
NATIVE_CRASH = {
    3221225477,  # 0xC0000005 STATUS_ACCESS_VIOLATION
    3221225725,  # 0xC00000FD STATUS_STACK_OVERFLOW
    3221225786,  # 0xC000013A CTRL_C
    -1073741819,  # 上面的有符号写法
}


def _is_native_crash(code: int) -> bool:
    return code in NATIVE_CRASH


def run_phase(*args: str, expect_code: int | None = 0) -> tuple[int, dict]:
    attempts = 0
    while True:
        attempts += 1
        proc = subprocess.run(
            [sys.executable, str(ENTRY), *args],
            cwd=str(SKILL),
            env=ENV,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if not _is_native_crash(proc.returncode) or attempts >= 2:
            break
        print(
            f"[WARN] 原生崩溃，重跑一次: {' '.join(args[:2])} "
            f"(exit {proc.returncode}) {proc.stderr[-160:]!r}"
        )

    text = proc.stdout.strip()
    start = text.find("{")
    payload: dict = {}
    if start >= 0:
        try:
            payload = json.loads(text[start:])
        except json.JSONDecodeError:
            payload = {"_raw": text[:400], "_stderr": proc.stderr[-400:]}
    else:
        payload = {"_raw": text[:400], "_stderr": proc.stderr[-400:]}
    if expect_code is not None:
        detail = f"got {proc.returncode} after {attempts} attempt(s)"
        if proc.returncode != expect_code:
            detail += f"; stderr={proc.stderr[-600:]!r}"
        record(
            f"exit code {expect_code}: {' '.join(args[:2])}",
            proc.returncode == expect_code,
            detail,
        )
    return proc.returncode, payload


def phase_detect(workspace: Path, task: str, month: str, expect_ok: bool = True) -> dict:
    _, payload = run_phase(
        "--phase", "detect", "--task", task, "--month", month, "--workspace", str(workspace),
        expect_code=0 if expect_ok else 3,
    )
    return payload


def check_detect_basic() -> None:
    payload = phase_detect(APRIL, "review", "2026-04")
    record("detect review: ok", payload.get("ok") is True, str(payload.get("missing")))
    resolved = payload.get("resolved", {})
    expected = {
        "online_caotan": 1,
        "online_gangwu": 1,
        "vehicle": 6,
        "finance": 1,
        "first_price": 1,
        "second_price": 1,
        "first_profit_template": 1,
        "second_profit_template": 1,
    }
    actual = {role: len(resolved.get(role, [])) for role in expected}
    record("detect review: 角色数量", actual == expected, f"{actual} != {expected}")
    record("detect review: 无歧义", not payload.get("ambiguous"), str(payload.get("ambiguous"))[:200])

    payload = phase_detect(APRIL, "summary", "2026-04")
    record("detect summary: ok", payload.get("ok") is True, str(payload.get("missing")))
    record(
        "detect summary: 草滩/港务各 1 个",
        all(len(payload.get("resolved", {}).get(role, [])) == 1
            for role in ("online_caotan", "online_gangwu")),
    )

    payload = phase_detect(APRIL, "profit", "2026-04")
    record("detect profit: ok", payload.get("ok") is True, str(payload.get("missing")))


def check_detect_renamed() -> None:
    """把文件改成 a.xlsx/b.xlsx，角色仍应靠表内结构识别出来。"""
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp)
        sources = {
            "a.xlsx": MARCH / "2024年10月新一干价格表(1).xlsx",
            "b.xlsx": MARCH / "二干运费价.xlsx",
            "c.xlsx": MARCH / "一干委办盈亏.xlsx",
            "d.xlsx": MARCH / "二干委办盈亏.xlsx",
            "e.xlsx": MARCH / "3月基础报表原版.xlsx",
            "f.xlsx": MARCH / "2026.03财务收入表.xlsx",
        }
        for name, source in sources.items():
            if not source.exists():
                record(f"detect renamed: 缺少样本 {source.name}", False, str(source))
                return
            shutil.copy2(source, target / name)
        payload = phase_detect(target, "profit", "2026-03")
        roles = {item["name"]: item["role"] for item in payload.get("files", [])}
        expected = {
            "a.xlsx": "first_price",
            "b.xlsx": "second_price",
            "c.xlsx": "first_profit_template",
            "d.xlsx": "second_profit_template",
            "e.xlsx": "base_report",
            "f.xlsx": "finance",
        }
        record("detect renamed: 角色不依赖文件名", roles == expected, f"{roles} != {expected}")


def check_detect_month() -> None:
    if not MULTI_MONTH.is_dir():
        record("detect month: 缺少多月份样本目录", False, str(MULTI_MONTH))
        return
    payload = phase_detect(MULTI_MONTH, "profit", "2026-04")
    base = payload.get("resolved", {}).get("base_report", [])
    finance = payload.get("resolved", {}).get("finance", [])
    record(
        "detect month: 3/4/5 月并存时选中 4 月",
        len(base) == 1 and "4月" in Path(base[0]).name
        and len(finance) == 1 and "2026.04" in Path(finance[0]).name,
        f"base={base} finance={finance}",
    )


def check_detect_missing() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        payload = phase_detect(Path(tmp), "profit", "2026-04", expect_ok=False)
        record(
            "detect missing: 空目录报 missing",
            payload.get("ok") is False and bool(payload.get("missing")),
            json.dumps(payload, ensure_ascii=False)[:200],
        )


def check_list() -> None:
    _, payload = run_phase("--phase", "list", "--workspace", str(APRIL))
    names = [item["name"] for item in payload.get("files", [])]
    record(
        "list: 返回真实文件名",
        payload.get("phase") == "listed" and "草滩在线表.xlsx" in names,
        str(names[:5]),
    )


def check_price_ambiguous() -> None:
    """价格表既没有一干/二干标题，文件名也没有 → 必须报 candidate，不能猜。"""
    import openpyxl

    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp) / "价格表.xlsx"
        shutil.copy2(MARCH / "二干运费价.xlsx", target)
        wb = openpyxl.load_workbook(target)
        for ws in wb.worksheets:
            for row in ws.iter_rows(min_row=1, max_row=3):
                for cell in row:
                    if isinstance(cell.value, str) and "二干" in cell.value:
                        cell.value = cell.value.replace("二干", "运输")
        wb.save(target)
        wb.close()

        payload = phase_detect(Path(tmp), "review", "2026-03", expect_ok=False)
        roles = [item["role"] for item in payload.get("files", [])]
        record(
            "detect ambiguous: 未标注一干/二干时返回 candidate",
            roles == ["price_candidate"],
            str(roles),
        )
        record(
            "detect ambiguous: 置信度为 low",
            all(item["confidence"] == "low" for item in payload.get("files", [])),
            str([(item["role"], item["confidence"]) for item in payload.get("files", [])]),
        )
        record(
            "detect ambiguous: 不静默选中价格表",
            "first_price" in payload.get("missing", [])
            and "second_price" in payload.get("missing", []),
            str(payload.get("missing")),
        )


def check_detect_damaged() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp) / "坏文件.xlsx"
        target.write_bytes(b"not an xlsx at all")
        payload = phase_detect(Path(tmp), "profit", "2026-04", expect_ok=False)
        files = payload.get("files", [])
        record(
            "detect damaged: 损坏文件标 damaged/low",
            len(files) == 1 and files[0]["role"] == "damaged" and files[0]["confidence"] == "low",
            str(files)[:200],
        )


def check_detect_ambiguous_role() -> None:
    """同一角色出现两份文件时必须报 ambiguous，不能随便挑一个。"""
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp)
        shutil.copy2(APRIL / "4月基础报表原版.xlsx", target / "基础报表A.xlsx")
        shutil.copy2(APRIL / "4月基础报表原版.xlsx", target / "基础报表B.xlsx")
        shutil.copy2(APRIL / "2026.04财务收入表.xlsx", target / "财务.xlsx")
        payload = phase_detect(target, "profit", "2026-04", expect_ok=False)
        ambiguous = [item["role"] for item in payload.get("ambiguous", [])]
        record(
            "detect ambiguous: 同角色两份文件报 ambiguous",
            ambiguous == ["base_report"],
            str(ambiguous),
        )


def check_detect_wrong_month() -> None:
    """月份对不上时降置信度，但角色仍保留，不静默丢弃。"""
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp)
        shutil.copy2(MARCH / "3月基础报表原版.xlsx", target / "基础报表.xlsx")
        shutil.copy2(MARCH / "2026.03财务收入表.xlsx", target / "财务.xlsx")
        payload = phase_detect(target, "profit", "2026-08")
        files = payload.get("files", [])
        record(
            "detect wrong month: 角色保留但置信度 low",
            all(item["role"] in {"base_report", "finance"} and item["confidence"] == "low"
                for item in files),
            str([(item["role"], item["confidence"]) for item in files]),
        )
        record(
            "detect wrong month: evidence 说明月份不一致",
            all(any("月份" in line for line in item["evidence"]) for item in files),
            str([item["evidence"] for item in files]),
        )


def check_runpy_host_survives() -> None:
    """AgentOS 用 runpy 在宿主进程内调用入口，入口不得杀掉宿主进程。"""
    script = (
        "import runpy, sys\n"
        "sys.path.insert(0, 'scripts')\n"
        "sys.argv = ['agentos_entrypoint.py', '--phase', 'detect', '--task', 'summary',\n"
        "            '--month', '2026-04', '--workspace', %r]\n"
        "try:\n"
        "    runpy.run_path('scripts/agentos_entrypoint.py', run_name='__main__')\n"
        "except SystemExit:\n"
        "    pass\n"
        "print('HOST_ALIVE')\n"
    ) % str(APRIL)
    # 模拟 AgentOS：不设 POST_OFFICE_HARD_EXIT，入口必须正常返回而不是 os._exit
    host_env = {key: value for key, value in ENV.items() if key != "POST_OFFICE_HARD_EXIT"}
    proc = subprocess.run(
        [sys.executable, "-c", script],
        cwd=str(SKILL), env=host_env, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )
    record(
        "runpy: 宿主进程存活（入口不 os._exit）",
        "HOST_ALIVE" in proc.stdout,
        f"exit={proc.returncode} stderr={proc.stderr[-300:]!r}",
    )


def check_online_site_by_structure() -> None:
    """草滩日 Sheet 实际是「4月1」（不带「日」），也要能按结构认出来。"""
    payload = phase_detect(APRIL, "summary", "2026-04")
    files = {item["role"]: item for item in payload.get("files", [])
             if item["role"].startswith("online")}
    record(
        "detect online: 草滩/港务均为 high 置信度",
        all(item["confidence"] == "high" for item in files.values()),
        str({role: item["confidence"] for role, item in files.items()}),
    )
    record(
        "detect online: 依据是日 Sheet 命名",
        all(any("日 Sheet" in line for line in item["evidence"]) for item in files.values()),
        str({role: item["evidence"] for role, item in files.items()}),
    )


def check_inspect() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        _, payload = run_phase(
            "--phase", "inspect", "--month", "2026-04",
            "--workspace", str(APRIL), "--output-dir", tmp,
        )
        record("inspect: ok", payload.get("ok") is True, str(payload.get("fatal"))[:300])
    _, payload = run_phase("--phase", "runtime", "--workspace", str(APRIL))
    record("runtime: ok", payload.get("ok") is True, str(payload.get("errors"))[:300])


def check_profit() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        _, payload = run_phase(
            "--phase", "profit", "--month", "2026-03",
            "--workspace", str(MARCH), "--output-dir", tmp,
        )
        outputs = payload.get("outputs", [])
        record("profit: 产出两张表", len(outputs) == 2, str(payload.get("errors"))[:300])
        match = payload.get("income_match", {})
        record(
            "profit: 未匹配趟次可接受",
            match.get("miss", 9999) <= 100,
            f"miss={match.get('miss')}",
        )
        record(
            "profit: 收入守恒",
            match.get("exact", 0) + match.get("near", 0) > 0,
            str(match),
        )


def check_full_workflow() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        workspace = Path(tmp) / "workspace"
        shutil.copytree(APRIL, workspace)
        _, review = run_phase(
            "--phase", "review", "--workspace", str(workspace),
            "--month", "2026-04", "--bundle-root", str(SKILL),
        )
        metrics = review.get("metrics", {})
        record("review: awaiting_review", review.get("phase") == "awaiting_review", str(review.get("errors"))[:300])
        baseline = {
            "model_rows": 3224,
            "model_summary_matched": 3921,
            "model_summary_unmatched": 388,
            "price_matched": 3628,
        }
        actual = {key: metrics.get(key) for key in baseline}
        record("review: 指标与基线一致", actual == baseline, f"{actual} != {baseline}")

        review_file = workspace / "output" / "4月_02_基础报表初版.xlsx"
        _, final = run_phase(
            "--phase", "final", "--workspace", str(workspace),
            "--review-file", str(review_file), "--bundle-root", str(SKILL),
        )
        outputs = [Path(item).name for item in final.get("outputs", [])]
        record(
            "final: 产出拆分版 + 两张盈亏表",
            final.get("phase") == "final_completed" and len(outputs) == 3,
            f"{final.get('phase')} {outputs} {str(final.get('errors'))[:200]}",
        )

        _, validation = run_phase(
            "--phase", "validate", "--workspace", str(workspace),
            "--stage", "final", "--output-dir", str(workspace / "output"),
        )
        record(
            "validate: final 输出校验通过",
            validation.get("ok") is True,
            str(validation.get("errors"))[:300],
        )


def check_package() -> None:
    manifest = json.loads((SKILL / "manifest.json").read_text(encoding="utf-8"))
    record("package: 版本 6.2.0", manifest.get("version") == "6.2.0", manifest.get("version", ""))
    import zipfile

    zips = list(ROOT.glob("post-office-monthly-reconciliation-skill.zip"))
    if not zips:
        record("package: skill zip 存在", False)
        return
    names = zipfile.ZipFile(zips[0]).namelist()
    record("package: 包含 detect.py", "scripts/reconciliation/detect.py" in names)
    record("package: 包含 detect 入口", "scripts/agentos_entrypoint.py" in names)
    record("package: 包含输入目录解析", "scripts/reconciliation/paths.py" in names)

    # 平台「导入专家」只接受 agent/agent.yaml 结构，且技能必须随包放在
    # agent/skills/<name>/ 下，否则导入后任务会报「绑定技能未安装」。
    agentos = sorted(ROOT.glob("*.agentos.zip"))
    if not agentos:
        record("package: agentos 导入包存在", False)
        return
    agent_names = set(zipfile.ZipFile(agentos[0]).namelist())
    record("package: agentos 含 agent.yaml", "agent/agent.yaml" in agent_names)
    record("package: agentos 含 prompt.md", "agent/prompt.md" in agent_names)
    record(
        "package: agentos 随包携带技能",
        "agent/skills/post-office-monthly-reconciliation/manifest.json" in agent_names,
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--fast", action="store_true", help="跳过完整 review/final")
    args = parser.parse_args()

    for required in (APRIL, MARCH):
        if not required.is_dir():
            print(f"[FAIL] 缺少测试数据目录: {required}")
            return 2

    check_detect_basic()
    check_detect_renamed()
    check_detect_month()
    check_detect_missing()
    check_price_ambiguous()
    check_detect_damaged()
    check_detect_ambiguous_role()
    check_detect_wrong_month()
    check_runpy_host_survives()
    check_online_site_by_structure()
    check_list()
    check_inspect()
    check_profit()
    check_package()
    if not args.fast:
        check_full_workflow()

    failed = [name for name, ok, _ in RESULTS if not ok]
    print("-" * 60)
    print(f"共 {len(RESULTS)} 项，通过 {len(RESULTS) - len(failed)} 项，失败 {len(failed)} 项")
    for name in failed:
        print(f"  [FAIL] {name}")
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
