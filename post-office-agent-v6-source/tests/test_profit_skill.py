# -*- coding: utf-8 -*-
"""盈亏表技能回归测试：容间分档边界 + 合计不被分档影响。

不依赖任何外部数据，自己造两张小表跑一遍。

    python post-office-agent-v6-source/tests/test_profit_skill.py

输出 [OK]/[FAIL]，退出码 0 表示全部通过。
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

PACKAGE = Path(__file__).resolve().parents[1]
SCRIPTS = PACKAGE / "skill" / "post-office-profit-report" / "scripts"
BUILD = SCRIPTS / "build_profit_tables.py"

# 容间 -> 应该在的表头档位（含 180 整的归属）
CASES = [
    (35, "小于55"),
    (54.9, "小于55"),
    (55, "[55,100)"),
    (99.9, "[55,100)"),
    (100, "[100,130)"),
    (129.9, "[100,130)"),
    (130, "[130,150)"),
    (149.9, "[130,150)"),
    (150, "[150,180)"),
    (179.9, "[150,180)"),
    (180, "大于180"),
    (260, "大于180"),
]


def make_base(path: Path) -> None:
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    header = ["序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
              "车牌号", "", "容间", "运输量", "运价", "备注"]
    for sheet_name, cases, offset in (
        ("一干汇总", CASES[:6], 0),
        ("二干基础", CASES[6:], 6),
    ):
        sheet = workbook.create_sheet(sheet_name)
        sheet.append(header)
        for index, (volume, _bin) in enumerate(cases, 1):
            number = offset + index
            sheet.append([
                index, "测试承运", f"P{number:03d}", f"西安-测试{number}", "L1",
                "2026-09-10", "", f"陕A{number:04d}", "", volume,
                100 * number, 1000 * number, "",
            ])
    workbook.save(path)
    workbook.close()


def make_finance(path: Path, base_path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["邮路级别", "自委办", "结算收入", "普服收入", "邮路名称", "车牌号", "流水号"])
    for index, _case in enumerate(CASES, 1):
        route = f"西安-测试{index}"
        sheet.append(["一干" if index <= 6 else "二干", "委办", 950.0 * index, 0,
                      route, f"陕A{index:04d}", "0910"])
    workbook.save(path)
    workbook.close()


def run(base: Path, finance: Path, out: Path) -> dict:
    command = [sys.executable, str(BUILD), "--base", str(base), "--finance", str(finance),
               "--month", "2026-09", "--out", str(out)]
    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                            errors="replace", cwd=str(SCRIPTS))
    payload = None
    for line in reversed((result.stdout or "").splitlines()):
        if line.strip().startswith("{"):
            payload = json.loads(line)
            break
    return {"code": result.returncode, "report": payload,
            "stdout": result.stdout, "stderr": result.stderr}


def check_bins(out: Path) -> list[str]:
    problems: list[str] = []
    for filename, cases in (("9月_04_一干盈亏表.xlsx", CASES[:6]),
                            ("9月_04_二干盈亏表.xlsx", CASES[6:])):
        workbook = openpyxl.load_workbook(out / filename, data_only=True)
        sheet = workbook["26年9月"]
        headers = {}
        for column in range(2, sheet.max_column + 1):
            value = sheet.cell(2, column).value
            if value:
                headers.setdefault(str(value).strip(), column)
        for volume, expected in cases:
            column = headers.get(expected)
            if column is None:
                problems.append(f"{filename} 表头缺少档位 {expected}")
                continue
            trips = sum(sheet.cell(row, column).value or 0
                        for row in range(4, sheet.max_row + 1)
                        if isinstance(sheet.cell(row, column).value, (int, float)))
            if not trips:
                problems.append(f"{filename} 容间 {volume} 没有落到 {expected}")
        workbook.close()
    return problems


def check_totals(out: Path) -> list[str]:
    """合计必须等于所有行加起来：分档不该改变总数。"""
    problems: list[str] = []
    for filename, expected_trips, expected_freight in (
        ("9月_04_一干盈亏表.xlsx", 6, 21000),
        ("9月_04_二干盈亏表.xlsx", 6, 57000),
    ):
        workbook = openpyxl.load_workbook(out / filename, data_only=True)
        sheet = workbook["26年9月"]
        total_row = next(row for row in range(4, sheet.max_row + 1)
                         if str(sheet.cell(row, 1).value or "").strip() == "合计")
        trips = freight = None
        for column in range(2, sheet.max_column + 1):
            label = str(sheet.cell(2, column).value or "").strip()
            sub = str(sheet.cell(3, column).value or "").strip()
            if label == "趟次" and sub in ("", "None", "趟次(趟)"):
                trips = sheet.cell(total_row, column).value
            if label == "委办运费(元)":
                freight = sheet.cell(total_row, column).value
        if trips != expected_trips:
            problems.append(f"{filename} 合计趟次 {trips}，应为 {expected_trips}")
        if round(float(freight or 0)) != expected_freight:
            problems.append(f"{filename} 合计运费 {freight}，应为 {expected_freight}")
        if "问题清单" not in workbook.sheetnames:
            problems.append(f"{filename} 缺少第二页「问题清单」")
        workbook.close()
    return problems


def main() -> int:
    failures: list[str] = []
    with tempfile.TemporaryDirectory(prefix="profit-skill-") as tmp:
        root = Path(tmp)
        base, finance, out = root / "基础报表原版.xlsx", root / "财务收入表.xlsx", root / "out"
        make_base(base)
        make_finance(finance, base)
        result = run(base, finance, out)
        if result["code"] != 0:
            print(result["stdout"][-2000:])
            print(result["stderr"][-2000:])
            print("[FAIL] 脚本退出码", result["code"])
            return 1
        if not (result["report"] or {}).get("ok"):
            failures.append("报告里 ok 不是 true")
        failures += check_bins(out)
        failures += check_totals(out)

    if failures:
        for item in failures:
            print("[FAIL]", item)
        return 1
    print(f"[OK] 容间分档边界 {len(CASES)} 例、合计口径、问题清单页 全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
