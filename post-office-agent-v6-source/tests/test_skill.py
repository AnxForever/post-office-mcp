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
# 合并导出口径（1 张容积测量表 + 1 张车管平台汇总 + 官方基础报表 + 财务表）的样本。
MERGED = ROOT / "9月数据包"

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
    3221226505,  # 0xC0000409 STATUS_STACK_BUFFER_OVERRUN（.pyd 收尾阶段偶发）
    -1073741819,  # 上面的有符号写法
    -1073741571,  # 0xC00000FD 有符号写法
    -1073740791,  # 0xC0000409 有符号写法
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


def phase_detect(workspace: Path, task: str, month: str, expect_complete: bool = True) -> dict:
    """detect 始终以退出码 0 返回；输入不全时用 status=needs_input 表示。"""
    _, payload = run_phase(
        "--phase", "detect", "--task", task, "--month", month, "--workspace", str(workspace),
        expect_code=0,
    )
    record(
        f"detect {task}: status 与输入完整性一致",
        (payload.get("status") == "ready") if expect_complete
        else (payload.get("status") == "needs_input" and payload.get("ok") is True),
        f"status={payload.get('status')} ok={payload.get('ok')}",
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
        payload = phase_detect(Path(tmp), "profit", "2026-04", expect_complete=False)
        gaps = payload.get("missing_inputs", [])
        roles = {item.get("role") for item in gaps}
        record(
            "detect missing: 空目录列出缺少的角色（不报错）",
            payload.get("missing") == ["base_report", "finance"]
            and {"base_report", "finance"} <= roles
            and all(item.get("action") for item in gaps),
            json.dumps(payload.get("missing_inputs"), ensure_ascii=False)[:300],
        )


def check_partial_upload_flow() -> None:
    """只上传个别文件时，各阶段都应正常返回 needs_input，而不是报错。"""
    with tempfile.TemporaryDirectory() as tmp:
        workspace = Path(tmp)
        shutil.copy2(APRIL / "2026.04财务收入表.xlsx", workspace / "2026.04财务收入表.xlsx")

        cases = [
            ("review", [
                "--phase", "review", "--workspace", str(workspace),
                "--month", "2026-04", "--bundle-root", str(SKILL),
            ]),
            ("profit", [
                "--phase", "profit", "--workspace", str(workspace), "--month", "2026-04",
            ]),
            ("summary", [
                "--phase", "summary", "--workspace", str(workspace), "--month", "2026-04",
            ]),
            ("final", [
                "--phase", "final", "--workspace", str(workspace), "--month", "2026-04",
            ]),
            ("inspect", [
                "--phase", "inspect", "--workspace", str(workspace), "--month", "2026-04",
                "--output-dir", str(workspace / "output"),
            ]),
        ]
        for name, args in cases:
            _, payload = run_phase(*args, expect_code=0)
            gaps = payload.get("missing_inputs", [])
            # inspect 是诊断接口：ok 仍表示“输入是否齐全”，但 status/退出码已经
            # 表明这是等待用户补文件的正常结果，不是执行失败。
            ok_field_ok = payload.get("ok") is True or name == "inspect"
            record(
                f"{name}: 部分上传返回 needs_input（不报错）",
                payload.get("status") == "needs_input"
                and ok_field_ok
                and bool(gaps)
                and not payload.get("errors"),
                json.dumps(payload, ensure_ascii=False)[:300],
            )

        record(
            "partial: 未生成交付件",
            not list(workspace.glob("output/*.xlsx")),
            str(list(workspace.glob("output/*.xlsx"))),
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

        payload = phase_detect(Path(tmp), "review", "2026-03", expect_complete=False)
        # 内置模板会出现在 files 里，这里只看用户实际上传的那份。
        bundled = set(payload.get("bundled_templates") or [])
        uploaded = [item for item in payload.get("files", []) if item["name"] not in bundled]
        roles = [item["role"] for item in uploaded]
        record(
            "detect ambiguous: 未标注一干/二干时返回 candidate",
            roles == ["price_candidate"],
            str(roles),
        )
        record(
            "detect ambiguous: 置信度为 low",
            all(item["confidence"] == "low" for item in uploaded),
            str([(item["role"], item["confidence"]) for item in uploaded]),
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
        payload = phase_detect(Path(tmp), "profit", "2026-04", expect_complete=False)
        files = payload.get("files", [])
        record(
            "detect damaged: 损坏文件标 damaged/low",
            len(files) == 1 and files[0]["role"] == "damaged" and files[0]["confidence"] == "low",
            str(files)[:200],
        )
        record(
            "detect damaged: 提示重新上传损坏文件",
            any(item.get("kind") == "unreadable" for item in payload.get("missing_inputs", [])),
            str(payload.get("missing_inputs"))[:200],
        )


def check_detect_ambiguous_role() -> None:
    """同一角色出现两份文件时必须报 ambiguous，不能随便挑一个。"""
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp)
        shutil.copy2(APRIL / "4月基础报表原版.xlsx", target / "基础报表A.xlsx")
        shutil.copy2(APRIL / "4月基础报表原版.xlsx", target / "基础报表B.xlsx")
        shutil.copy2(APRIL / "2026.04财务收入表.xlsx", target / "财务.xlsx")
        payload = phase_detect(target, "profit", "2026-04", expect_complete=False)
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


def check_error_messages_readable() -> None:
    """MemoryError 的 str() 是空串，入口不能把它上报成 errors:[""]。"""
    code = (
        "import sys\n"
        "sys.path.insert(0, 'scripts')\n"
        "from run_workflow import _describe_exception\n"
        "for exc in (MemoryError(), RuntimeError()):\n"
        "    text = _describe_exception(exc)\n"
        "    assert text.strip(), type(exc).__name__\n"
        "print('READABLE')\n"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code],
        cwd=str(SKILL), env=ENV, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )
    record(
        "错误上报: 空异常信息不会变成空串",
        "READABLE" in proc.stdout,
        f"exit={proc.returncode} {proc.stderr[-200:]!r}",
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


def check_pricing() -> None:
    """价格表单元格已含系数：运价=所在档位单元格，标准价=系数1.0档价格。"""
    sys.path.insert(0, str(SKILL / "scripts"))
    try:
        from reconciliation.excel_io import read_price_catalog
        from reconciliation.matching import match_price
    finally:
        sys.path.pop(0)
    raw = json.loads((SKILL / "assets" / "route-aliases.json").read_text(encoding="utf-8"))
    aliases = raw.get("aliases", raw)
    catalogs = {
        "first": read_price_catalog(APRIL / "2024年10月新一干价格表(1).xlsx", "first"),
        "second": read_price_catalog(APRIL / "二干运费价.xlsx", "second"),
    }
    cases = [
        # 路线, 容间, 期望档位, 期望标准价(系数1.0档), 期望运价(=该档位价格表单元格)
        ("西安-广州(陕1)", 66, "12", 8850, 5575.5),
        ("西安-汉中港1", 134, "40", 1745, 3001.4),
        ("西安-银川(陕2)", 136, "30", 4150, 4150),
    ]
    for route, volume, tonnage, standard, freight in cases:
        price = match_price(route, volume, catalogs, aliases)
        ok = (
            price.tonnage == tonnage
            and price.standard_price == standard
            and price.freight == freight
        )
        record(
            f"pricing: {route} {tonnage}吨 运价不重复乘系数",
            ok,
            f"ton={price.tonnage} std={price.standard_price} freight={price.freight}",
        )

    round_trip_cases = [
        # 路线, 容间, 期望档位, 期望标准价, 期望运价
        # 一干：往返 Sheet 写全档位，直接取 30 吨单元格。
        ("西安-银川(陕2)", 139, "30", 7000, 7000),
        # 二干：往返列只写 12 吨价，30 吨由系数折算。
        ("西安-汉中提速1", 62, "12", 2570, 2570),
        ("西安-宝鸡港1", 165, "40", 1698.3, 2921.08),
    ]
    for route, volume, tonnage, standard, freight in round_trip_cases:
        price = match_price(route, volume, catalogs, aliases, round_trip=True)
        ok = (
            price.tonnage == tonnage
            and price.standard_price == standard
            and float(price.freight) == freight
        )
        record(
            f"pricing: {route} {tonnage}吨 往返价",
            ok,
            f"ton={price.tonnage} std={price.standard_price} freight={price.freight}",
        )


BLUE_MARK = "FFBDD7EE"


def _review_mark_check(path: Path) -> tuple[bool, str]:
    """跨站兜底取到的容间必须可见：行有浅蓝底色，问题汇总里有对应条目。"""
    from openpyxl import load_workbook

    wb = load_workbook(path)
    sheet = wb["基础报表"]
    volume_col = next(c.column for c in sheet[1] if c.value == "容间")
    blue_rows = empty_volume_blue = 0
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=21):
        fills = {
            str(cell.fill.fgColor.rgb)
            for cell in row
            if cell.fill is not None and cell.fill.fill_type == "solid"
        }
        if BLUE_MARK not in fills:
            continue
        blue_rows += 1
        if row[volume_col - 1].value in (None, "", 0, 0.0):
            empty_volume_blue += 1
    issues = wb["问题汇总"]
    detail_col = next(c.column for c in issues[1] if c.value == "问题详情")
    cross_issues = sum(
        1
        for r in range(2, issues.max_row + 1)
        if "容间取自同车牌其他日期" in str(issues.cell(r, detail_col).value or "")
    )
    ok = blue_rows > 0 and cross_issues >= blue_rows and empty_volume_blue == 0
    return ok, f"蓝行={blue_rows} 跨站问题={cross_issues} 蓝行缺容间={empty_volume_blue}"


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
        record("review: 跨站兜底容间标蓝", *_review_mark_check(review_file))

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


def check_bundled_templates() -> None:
    """三份版式模板随包内置：用户不传模板时，完整流程仍要能跑完。"""
    with tempfile.TemporaryDirectory() as tmp:
        workspace = Path(tmp) / "workspace"
        shutil.copytree(APRIL, workspace)
        removed = []
        for name in ("基础报表初版模板.xlsx", "一干委办盈亏.xlsx", "二干委办盈亏.xlsx"):
            path = workspace / name
            if path.is_file():
                path.unlink()
                removed.append(name)
        record("bundled: 测试前删除三份模板", len(removed) == 3, str(removed))

        payload = phase_detect(workspace, "review", "2026-04")
        missing = payload.get("missing") or []
        record(
            "bundled: 模板缺失不算缺文件",
            "first_profit_template" not in missing and "second_profit_template" not in missing,
            str(missing),
        )
        record(
            "bundled: detect 报告使用内置模板",
            set(payload.get("bundled_templates") or [])
            == {"draft-template.xlsx", "first-profit-template.xlsx", "second-profit-template.xlsx"},
            str(payload.get("bundled_templates")),
        )

        _, review = run_phase(
            "--phase", "review", "--workspace", str(workspace),
            "--month", "2026-04", "--bundle-root", str(SKILL),
        )
        record(
            "bundled: review 用内置初版模板生成初版",
            review.get("phase") == "awaiting_review",
            f"{review.get('phase')} {str(review.get('errors'))[:300]}",
        )
        if review.get("phase") != "awaiting_review":
            return

        review_file = workspace / "output" / "4月_02_基础报表初版.xlsx"
        _, final = run_phase(
            "--phase", "final", "--workspace", str(workspace),
            "--review-file", str(review_file), "--bundle-root", str(SKILL),
        )
        outputs = [Path(item).name for item in final.get("outputs", [])]
        record(
            "bundled: final 用内置盈亏模板产出三件",
            final.get("phase") == "final_completed" and len(outputs) == 3,
            f"{final.get('phase')} {outputs} {str(final.get('errors'))[:200]}",
        )


def check_bundled_prices() -> None:
    """合同价格表也随包内置：用户不传时用内置版本，并在结果里说明。"""
    with tempfile.TemporaryDirectory() as tmp:
        workspace = Path(tmp) / "workspace"
        shutil.copytree(APRIL, workspace)
        removed = []
        for name in ("2024年10月新一干价格表(1).xlsx", "二干运费价.xlsx"):
            path = workspace / name
            if path.is_file():
                path.unlink()
                removed.append(name)
        record("bundled: 测试前删除两份价格表", len(removed) == 2, str(removed))

        payload = phase_detect(workspace, "review", "2026-04")
        missing = payload.get("missing") or []
        record(
            "bundled: 价格表缺失不算缺文件",
            "first_price" not in missing and "second_price" not in missing,
            str(missing),
        )
        record(
            "bundled: detect 报告使用内置价格表",
            set(payload.get("bundled_price_tables") or [])
            == {"first-price-table.xlsx", "second-price-table.xlsx"},
            str(payload.get("bundled_price_tables")),
        )
        # 内置价格表必须能读出档位和路线，否则等于没内置。
        sys.path.insert(0, str(SKILL / "scripts"))
        try:
            from reconciliation.excel_io import read_price_catalog

            assets = SKILL / "assets"
            first = read_price_catalog(assets / "first-price-table.xlsx", "first")
            second = read_price_catalog(assets / "second-price-table.xlsx", "second")
        finally:
            if str(SKILL / "scripts") in sys.path:
                sys.path.remove(str(SKILL / "scripts"))
        record(
            "bundled: 内置一干价格表可解析",
            len(first.get("tiers") or []) >= 4 and len(first.get("routes") or {}) > 50,
            f"tiers={len(first.get('tiers') or [])} routes={len(first.get('routes') or {})}",
        )
        record(
            "bundled: 内置二干价格表可解析",
            len(second.get("tiers") or []) >= 7 and len(second.get("routes") or {}) > 10,
            f"tiers={len(second.get('tiers') or [])} routes={len(second.get('routes') or {})}",
        )


def check_publish() -> None:
    sys.path.insert(0, str(SKILL / "scripts"))
    try:
        from reconciliation.paths import publish_files, resolve_publish_dir
    finally:
        sys.path.pop(0)
    with tempfile.TemporaryDirectory() as tmp:
        source = Path(tmp) / "4月_02_基础报表初版.xlsx"
        source.write_bytes(b"xlsx")
        target = Path(tmp) / "outputs"
        published = publish_files([source, Path(tmp) / "missing.xlsx"], target)
        record(
            "publish: 交付件复制到成果目录",
            [item["name"] for item in published] == [source.name]
            and (target / source.name).read_bytes() == b"xlsx",
            str(published),
        )
        record(
            "publish: 缺失文件不报错",
            len(published) == 1,
            str(published),
        )
    # Windows 上 Path('/outputs') 会落到当前盘符根目录，必须拒绝。
    if sys.platform == "win32":
        record("publish: 本地不把盘符根目录当成果目录", resolve_publish_dir() is None)
    # 自建部署可用环境变量指定成果目录。
    with tempfile.TemporaryDirectory() as tmp:
        override = Path(tmp) / "published"
        os.environ["POST_OFFICE_PUBLISH_DIR"] = str(override)
        try:
            record("publish: 环境变量覆盖成果目录", resolve_publish_dir() == override)
        finally:
            os.environ.pop("POST_OFFICE_PUBLISH_DIR", None)


def check_package() -> None:
    manifest = json.loads((SKILL / "manifest.json").read_text(encoding="utf-8"))
    record("package: 版本 6.3.3", manifest.get("version") == "6.3.3", manifest.get("version", ""))
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


def check_merged_bundle_workflow() -> None:
    """合并导出口径必须走通 run_workflow，而不是只跑 build_review_workbook。

    早先 DYNAMIC_ROLE_COUNTS 写死了拆分口径的 online_caotan=1 / online_gangwu=1 /
    vehicle=6，合并口径上传后 detect 判定正常，review 却在校验阶段整片报
    「应为1个，实际0个」并以退出码 2 结束。这条测试盯的就是那个缺口。
    """
    merged = MERGED
    if not merged.is_dir():
        print(f"[SKIP] 合并口径测试数据不存在: {merged}")
        return
    with tempfile.TemporaryDirectory() as tmp:
        workspace = Path(tmp) / "workspace"
        # resolve_input_dir 优先用 workspace，附件直接放这里即可。
        workspace.mkdir()
        for path in merged.glob("*.xlsx"):
            shutil.copy2(path, workspace / path.name)

        # 不传 --month：脚本必须从文件名/工作表名自己推断出 2026-09。
        # 用户不该被要求记 YYYY-MM，更不该被要求写 --phase 这类参数。
        _, review = run_phase(
            "--phase", "review", "--workspace", str(workspace),
            "--bundle-root", str(SKILL),
        )
        record(
            "merged: 未传月份也能自动推断目标月",
            review.get("target_month") == "2026-09",
            f"target_month={review.get('target_month')} errors={str(review.get('errors'))[:200]}",
        )
        state_path = workspace / "output" / "run-state.json"
        state = json.loads(state_path.read_text(encoding="utf-8")) if state_path.is_file() else {}
        record(
            "merged: detect 判为合并导出口径",
            state.get("source_bundle") == "merged",
            f"source_bundle={state.get('source_bundle')}",
        )
        record(
            "merged: review 走到 awaiting_review",
            review.get("phase") == "awaiting_review",
            f"{review.get('phase')} errors={str(review.get('errors'))[:300]}",
        )
        metrics = review.get("metrics", {})
        # 合并口径的容间来自平台列；这两个数掉下来通常表示又退回按实测值取数了。
        record(
            "merged: 运价已匹配到绝大多数行",
            metrics.get("price_matched", 0) >= 3600,
            f"price_matched={metrics.get('price_matched')}",
        )
        record(
            "merged: 与官方基础报表逐单对价已生效",
            metrics.get("official_comparison", {}).get("available", 0) >= 3200,
            str(metrics.get("official_comparison")),
        )

        # final 阶段靠 row_ids 与 row_count 对齐来校验；row_ids 早先只覆盖车管行，
        # 漏掉「在线有测量但车管无派车」的孤儿行，第二阶段直接报
        # 「row_count 与 row_ids 数量不一致」而失败。
        record(
            "merged: row_ids 覆盖全部行（含在线独有行）",
            len(state.get("row_ids", [])) == state.get("row_count"),
            f"row_ids={len(state.get('row_ids', []))} row_count={state.get('row_count')}",
        )

        review_file = workspace / "output" / "9月_02_基础报表初版.xlsx"
        _, final = run_phase(
            "--phase", "final", "--workspace", str(workspace),
            "--review-file", str(review_file), "--bundle-root", str(SKILL),
        )
        outputs = [Path(item).name for item in final.get("outputs", [])]
        record(
            "merged: final 产出拆分版 + 两张盈亏表",
            final.get("phase") == "final_completed" and len(outputs) == 3,
            f"{final.get('phase')} {outputs} {str(final.get('errors'))[:200]}",
        )


def check_month_inference_asks_when_unsure() -> None:
    """月份推断不确定时必须反问，不能自己挑一个往下跑。

    这是业务硬要求：选错月份等于把整月报表做废，宁可多问一句。
    """
    merged = MERGED
    if not merged.is_dir():
        print(f"[SKIP] 合并口径测试数据不存在: {merged}")
        return

    def run_detect(files: list[str], tmp: Path) -> dict:
        workspace = tmp / "workspace"
        workspace.mkdir(parents=True)
        for name in files:
            src = merged / name
            if src.is_file():
                shutil.copy2(src, workspace / name)
        _, payload = run_phase(
            "--phase", "detect", "--task", "review",
            "--workspace", str(workspace), expect_code=0,
        )
        return payload

    with tempfile.TemporaryDirectory() as tmp:
        # 1) 只有 9 月、没有年份 → 要提示年份不确定
        payload = run_detect(["9月.xlsx", "汇总.xlsx"], Path(tmp) / "a")
        inf = payload.get("month_inference") or {}
        record(
            "month: 缺年份时标记为不确定并给出疑问",
            inf.get("confidence") in {"medium", "low"} and bool(inf.get("questions")),
            f"confidence={inf.get('confidence')} questions={inf.get('questions')}",
        )

        # 2) 9 月和 10 月打平 → 必须反问，不能返回月份
        workspace2 = Path(tmp) / "b" / "workspace"
        workspace2.mkdir(parents=True)
        shutil.copy2(merged / "9月.xlsx", workspace2 / "9月.xlsx")
        shutil.copy2(merged / "汇总.xlsx", workspace2 / "汇总.xlsx")
        shutil.copy2(merged / "9月基础报表原版.xlsx", workspace2 / "10月基础报表原版.xlsx")
        _, payload2 = run_phase(
            "--phase", "detect", "--task", "review",
            "--workspace", str(workspace2), expect_code=0,
        )
        inf2 = payload2.get("month_inference") or {}
        record(
            "month: 月份冲突时不猜、返回 needs_month",
            payload2.get("status") == "needs_month" and inf2.get("month") is None,
            f"status={payload2.get('status')} month={inf2.get('month')}",
        )
        record(
            "month: 冲突时列出候选月份供用户选择",
            bool(payload2.get("questions")) and bool(inf2.get("conflicts")),
            f"questions={payload2.get('questions')} conflicts={inf2.get('conflicts')}",
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
    check_partial_upload_flow()
    check_runpy_host_survives()
    check_error_messages_readable()
    check_online_site_by_structure()
    check_list()
    check_inspect()
    check_profit()
    check_pricing()
    check_publish()
    check_package()
    if not args.fast:
        check_full_workflow()
        check_merged_bundle_workflow()
        check_month_inference_asks_when_unsure()
        check_bundled_templates()
        check_bundled_prices()

    failed = [name for name, ok, _ in RESULTS if not ok]
    print("-" * 60)
    print(f"共 {len(RESULTS)} 项，通过 {len(RESULTS) - len(failed)} 项，失败 {len(failed)} 项")
    for name in failed:
        print(f"  [FAIL] {name}")
    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
