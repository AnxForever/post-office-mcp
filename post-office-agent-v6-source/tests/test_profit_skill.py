# -*- coding: utf-8 -*-
"""盈亏表技能回归测试：一干容间分档、二干吨位分档、备注优先、问题清单。

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

# 一干：容间 -> 应该在的表头档位（含区间边界和 180 整的归属）
VOLUME_CASES = [
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

# 二干：吨位, 容间, 备注。备注写了结算吨位的，要按备注落档。
TONNAGE_CASES = [
    ("3", 16, ""),
    ("12", 62, ""),
    ("20", 128, ""),
    ("30", 130, ""),
    ("40A", 153, ""),
    ("40B", 183, ""),
    ("12", 63, "两辆12吨的车合并一辆30吨的车"),   # -> 30
    ("40A", 153, "调度通知按30吨结算"),          # -> 30
    ("40", 160, ""),                              # -> 40（40A）
    ("40B", 183, "调度通知按40A结算"),            # -> 40（40A）
    ("12", 63, "两辆12吨的车合并一辆40A的车"),    # -> 40（40A）
]

# 一干汇总里 5 条由财务收入表改判到二干的邮路：没有吨位列，按容间 63 推回 12 吨。
REROUTE_COUNT = 5

# 客户确认的单趟结算吨位：只对 2026-09-08 那一趟生效，容间 35 方按 3 吨。
CONFIRMED_FREIGHT = 2250
CONFIRMED_PLATE = "苏E6W073"
# 同车同邮路、但不是确认的那一天，仍按容间推成 5 吨。
OTHER_FREIGHT = 2250

# 收入没对上的趟次 / 财务表里有收入但基础表没有对应趟次：都不进结果表，只上问题清单。
UNMATCHED_PLATE = "陕Z9999"
ORPHAN_PLATE = "陕Y8888"
ORPHAN_INCOME = 777
UNMATCHED_FREIGHT = 999
UNUSED_PLATE = "陕C0001"
UNUSED_INCOME = 555

# 二干表头 -> 期望趟次（11 条二干记录 + 1 条确认趟按 3 吨 + 1 条同车其他日期按 5 吨 + 5 条改判）
EXPECTED_SECOND = {"3": 2, "5": 1, "12": 6, "20": 1, "30": 3, "40（40A）": 4, "40B": 1}


def make_base(path: Path) -> None:
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    first = workbook.create_sheet("一干汇总")
    first.append(["序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
                  "车牌号", "", "容间", "运输量", "运价", "备注"])
    for index, (volume, _bin) in enumerate(VOLUME_CASES, 1):
        first.append([index, "测试承运", f"P{index:03d}", f"西安-测试{index}", "L1",
                      "2026-09-10", "", f"陕A{index:04d}", "", volume,
                      100 * index, 1000 * index, ""])
    for offset in range(1, REROUTE_COUNT + 1):
        index = len(VOLUME_CASES) + offset
        first.append([index, "测试承运", f"P{index:03d}", f"西安-改判{offset}", "L1",
                      "2026-09-10", "", f"陕B{offset:04d}", "", 63,
                      100, 5000 + offset, ""])

    second = workbook.create_sheet("二干基础")
    second.append(["序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
                   "车牌号", "", "容间", "吨位", "运输量", "运价", "备注"])
    for index, (tonnage, volume, remark) in enumerate(TONNAGE_CASES, 1):
        second.append([index, "测试承运", f"Q{index:03d}", f"西安-二干{index}", "L1",
                       "2026-09-10", "", f"陕C{index:04d}", "", volume, tonnage,
                       100 * index, 2000 * index, remark])
    # 容间 35 方、没有吨位，但客户确认 2026-09-08 这一趟按 3 吨结算（不是 5 吨）。
    second.append([len(TONNAGE_CASES) + 1, "测试承运", "Q900", "西安-陇南(陕T1)", "L1",
                   "2026-09-08", "", CONFIRMED_PLATE, "", 35, "",
                   300, CONFIRMED_FREIGHT, ""])
    # 同车同邮路、另一天：不在客户确认范围内，仍按容间推成 5 吨。
    second.append([len(TONNAGE_CASES) + 2, "测试承运", "Q901", "西安-陇南(陕T1)", "L1",
                   "2026-09-10", "", CONFIRMED_PLATE, "", 35, "",
                   300, OTHER_FREIGHT, ""])
    # 这一趟在财务收入表里没有收入：不能进结果表，要单独列在问题清单。
    second.append([len(TONNAGE_CASES) + 3, "测试承运", "Q902", "西安-未匹配", "L1",
                   "2026-09-11", "", UNMATCHED_PLATE, "", 63, "12",
                   300, UNMATCHED_FREIGHT, ""])

    workbook.save(path)
    workbook.close()


def make_finance(path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["邮路级别", "自委办", "结算收入", "普服收入", "邮路名称", "车牌号", "流水号"])
    for index in range(1, len(VOLUME_CASES) + 1):
        sheet.append(["一干", "委办", 950.0 * index, 0,
                      f"西安-测试{index}", f"陕A{index:04d}", "0910"])
    for offset in range(1, REROUTE_COUNT + 1):
        sheet.append(["二干", "委办", 7000.0 + offset, 0,
                      f"西安-改判{offset}", f"陕B{offset:04d}", "0910"])
    for index in range(1, len(TONNAGE_CASES) + 1):
        sheet.append(["二干", "委办", 1500.0 * index, 0,
                      f"西安-二干{index}", f"陕C{index:04d}", "0910"])
    sheet.append(["二干", "委办", 2000, 0, "西安-陇南", CONFIRMED_PLATE, "0908"])
    sheet.append(["二干", "委办", 2000, 0, "西安-陇南", CONFIRMED_PLATE, "0910"])
    # 财务收入表里有收入，但基础报表里没有这台车/这一趟：只上问题清单。
    sheet.append(["二干", "委办", ORPHAN_INCOME, 0, "西安-孤儿", ORPHAN_PLATE, "0910"])
    # 车牌在基础表里，但日期差得太远，没对到具体趟次：也只上问题清单。
    sheet.append(["二干", "委办", UNUSED_INCOME, 0, "西安-二干1", UNUSED_PLATE, "0925"])
    workbook.save(path)
    workbook.close()


def make_swapped_base(path: Path) -> None:
    """反向情况：一干带吨位、二干没有吨位。分档要跟着数据走，不能写死。"""
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    first = workbook.create_sheet("一干汇总")
    first.append(["序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
                  "车牌号", "", "容间", "吨位", "运输量", "运价", "备注"])
    for index, tonnage in enumerate(("3", "12", "30", "40A"), 1):
        first.append([index, "测试承运", f"P{index:03d}", f"西安-一干{index}", "L1",
                      "2026-09-10", "", f"陕D{index:04d}", "", 16 * index,
                      tonnage, 100 * index, 1000 * index, ""])

    second = workbook.create_sheet("二干基础")
    second.append(["序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
                   "车牌号", "", "容间", "运输量", "运价", "备注"])
    for index, volume in enumerate((35, 63, 180), 1):
        second.append([index, "测试承运", f"Q{index:03d}", f"西安-二干{index}", "L1",
                       "2026-09-10", "", f"陕E{index:04d}", "", volume,
                       100 * index, 2000 * index, ""])

    workbook.save(path)
    workbook.close()


def make_swapped_finance(path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["邮路级别", "自委办", "结算收入", "普服收入", "邮路名称", "车牌号", "流水号"])
    for index in range(1, 5):
        sheet.append(["一干", "委办", 900.0 * index, 0,
                      f"西安-一干{index}", f"陕D{index:04d}", "0910"])
    for index in range(1, 4):
        sheet.append(["二干", "委办", 800.0 * index, 0,
                      f"西安-二干{index}", f"陕E{index:04d}", "0910"])
    workbook.save(path)
    workbook.close()


def make_dropped_base(path: Path) -> None:
    """有一行既没吨位也没容间：分不出车型，必须停下，不能静默丢行。"""
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    first = workbook.create_sheet("一干汇总")
    first.append(["序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
                  "车牌号", "", "容间", "吨位", "运输量", "运价", "备注"])
    first.append([1, "测试承运", "P001", "西安-正常", "L1", "2026-09-10", "",
                  "陕A0001", "", 63, "", 100, 1000, ""])
    first.append([2, "测试承运", "P002", "西安-缺档", "L1", "2026-09-10", "",
                  "陕A0002", "", "", "", 100, 2000, ""])

    second = workbook.create_sheet("二干基础")
    second.append(["序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
                   "车牌号", "", "容间", "吨位", "运输量", "运价", "备注"])
    second.append([1, "测试承运", "Q001", "西安-正常二", "L1", "2026-09-10", "",
                   "陕B0001", "", 63, "12", 100, 3000, ""])

    workbook.save(path)
    workbook.close()


def run(base: Path, finance: Path, out: Path) -> dict:
    command = [sys.executable, str(BUILD), "--base", str(base), "--finance", str(finance),
               "--month", "2026-09", "--out", str(out)]
    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                            errors="replace", cwd=str(SCRIPTS))
    payload = None
    lines = (result.stdout or "").splitlines()
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped == "{":
            candidate = "\n".join(lines[index:])
        elif stripped.startswith("{"):
            candidate = stripped
        else:
            continue
        try:
            payload, _ = json.JSONDecoder().raw_decode(candidate)
        except json.JSONDecodeError:
            continue
        break
    return {"code": result.returncode, "report": payload,
            "stdout": result.stdout, "stderr": result.stderr}


def header_trips(sheet) -> dict[str, int]:
    """表头档位 -> 合计行里的趟次。"""
    total_row = next(row for row in range(4, sheet.max_row + 1)
                     if str(sheet.cell(row, 1).value or "").strip() == "合计")
    result: dict[str, int] = {}
    for column in range(2, sheet.max_column + 1):
        label = str(sheet.cell(2, column).value or "").strip()
        sub = str(sheet.cell(3, column).value or "").strip()
        if label and sub == "趟次(趟)":
            result[label] = sheet.cell(total_row, column).value or 0
    return result


def check_first(out: Path) -> list[str]:
    problems: list[str] = []
    workbook = openpyxl.load_workbook(out / "9月_04_一干盈亏表.xlsx", data_only=True)
    sheet = workbook["26年9月"]
    trips = header_trips(sheet)
    for _volume, expected in VOLUME_CASES:
        if trips.get(expected) != 2:
            problems.append(f"一干 {expected} 趟次 {trips.get(expected)}，应为 2")
    if sum(trips.values()) != len(VOLUME_CASES):
        problems.append(f"一干 合计趟次 {sum(trips.values())}，应为 {len(VOLUME_CASES)}")

    issue = workbook["问题清单"]
    text = "\n".join(str(cell.value) for row in issue.iter_rows() for cell in row if cell.value)
    if "按容间汇总" not in text:
        problems.append("一干问题清单缺少「按容间汇总」")
    if "西安-改判1" in text:
        problems.append("改判到二干的邮路不应出现在一干问题清单")
    workbook.close()
    return problems


def check_second(out: Path) -> list[str]:
    problems: list[str] = []
    workbook = openpyxl.load_workbook(out / "9月_04_二干盈亏表.xlsx", data_only=True)
    sheet = workbook["26年9月"]
    trips = header_trips(sheet)
    if trips != EXPECTED_SECOND:
        problems.append(f"二干 各车型趟次 {trips}，应为 {EXPECTED_SECOND}")
    if sum(trips.values()) != len(TONNAGE_CASES) + 2 + REROUTE_COUNT:
        problems.append(f"二干 合计趟次 {sum(trips.values())}，"
                        f"应为 {len(TONNAGE_CASES) + 2 + REROUTE_COUNT}")

    issue = workbook["问题清单"]
    text = "\n".join(str(cell.value) for row in issue.iter_rows() for cell in row if cell.value)
    if "按吨位汇总" not in text:
        problems.append("二干问题清单缺少「按吨位汇总」")
    if "您确认过的" not in text or CONFIRMED_PLATE not in text:
        problems.append("二干问题清单没有说明客户确认的结算吨位")
    if "9月8日" not in text:
        problems.append("二干问题清单没有写明确认的是 9月8日 那一趟")
    if UNMATCHED_PLATE not in text or ORPHAN_PLATE not in text:
        problems.append("二干问题清单没有列出未匹配的趟次/收入")
    if UNUSED_PLATE not in text or "日期/趟次没对上" not in text:
        problems.append("二干问题清单没有列出没对到趟次的收入")
    if "没有算进本表" not in text:
        problems.append("二干问题清单没有说明未匹配项没有算进本表")

    routes = [str(sheet.cell(row, 1).value or "").strip()
              for row in range(4, sheet.max_row + 1)]
    if any("西安-未匹配" in route for route in routes):
        problems.append("收入没对上的趟次不应出现在结果表")
    for offset in range(1, REROUTE_COUNT + 1):
        if f"西安-改判{offset}" not in text:
            problems.append(f"二干问题清单没有列出改判邮路 西安-改判{offset}")
    workbook.close()
    return problems


def check_swapped(out: Path) -> list[str]:
    """一干带吨位、二干没有吨位时，分档要跟着数据走。"""
    problems: list[str] = []

    workbook = openpyxl.load_workbook(out / "9月_04_一干盈亏表.xlsx", data_only=True)
    sheet = workbook["26年9月"]
    trips = header_trips(sheet)
    if set(trips) != {"3", "12", "30", "40A"}:
        problems.append(f"一干带吨位列时表头 {sorted(trips)}，应为 3/12/30/40A")
    if sum(trips.values()) != 4:
        problems.append(f"一干带吨位列时合计趟次 {sum(trips.values())}，应为 4")
    issue = workbook["问题清单"]
    text = "\n".join(str(cell.value) for row in issue.iter_rows() for cell in row if cell.value)
    if "按吨位汇总" not in text or "填了「吨位」" not in text:
        problems.append("一干带吨位列时，问题清单没有说明按吨位分车型")
    workbook.close()

    workbook = openpyxl.load_workbook(out / "9月_04_二干盈亏表.xlsx", data_only=True)
    sheet = workbook["26年9月"]
    trips = header_trips(sheet)
    if set(trips) != {"小于55", "[55,100)", "大于180"}:
        problems.append(f"二干没有吨位列时表头 {sorted(trips)}，应为容间档位")
    if sum(trips.values()) != 3:
        problems.append(f"二干没有吨位列时合计趟次 {sum(trips.values())}，应为 3")
    issue = workbook["问题清单"]
    text = "\n".join(str(cell.value) for row in issue.iter_rows() for cell in row if cell.value)
    if "按容间汇总" not in text or "没有填「吨位」" not in text:
        problems.append("二干没有吨位列时，问题清单没有说明按容间分类")
    workbook.close()
    return problems


def check_totals(out: Path) -> list[str]:
    """合计必须等于所有行加起来：分档不该改变总数。"""
    problems: list[str] = []
    first_freight = sum(1000 * index for index in range(1, len(VOLUME_CASES) + 1))
    second_freight = (sum(2000 * index for index in range(1, len(TONNAGE_CASES) + 1))
                      + CONFIRMED_FREIGHT + OTHER_FREIGHT
                      + sum(5000 + offset for offset in range(1, REROUTE_COUNT + 1)))
    for filename, expected_trips, expected_freight in (
        ("9月_04_一干盈亏表.xlsx", len(VOLUME_CASES), first_freight),
        ("9月_04_二干盈亏表.xlsx", len(TONNAGE_CASES) + 2 + REROUTE_COUNT, second_freight),
    ):
        workbook = openpyxl.load_workbook(out / filename, data_only=True)
        sheet = workbook["26年9月"]
        total_row = next(row for row in range(4, sheet.max_row + 1)
                         if str(sheet.cell(row, 1).value or "").strip() == "合计")
        trips = freight = None
        for column in range(2, sheet.max_column + 1):
            label = str(sheet.cell(2, column).value or "").strip()
            if label == "趟次":
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
        make_finance(finance)
        result = run(base, finance, out)
        if result["code"] != 0:
            print(result["stdout"][-2000:])
            print(result["stderr"][-2000:])
            print("[FAIL] 脚本退出码", result["code"])
            return 1
        if not (result["report"] or {}).get("ok"):
            failures.append("报告里 ok 不是 true")
        report = result["report"] or {}
        if report.get("unmatched") != 1:
            failures.append(f"报告里未匹配趟次 {report.get('unmatched')}，应为 1")
        if report.get("orphan_income") != 1:
            failures.append(f"报告里孤儿收入 {report.get('orphan_income')}，应为 1")
        if report.get("unused_income") != 1:
            failures.append(f"报告里没对到趟次的收入 {report.get('unused_income')}，应为 1")
        if report.get("excluded_rows") != 1:
            failures.append(f"报告里剔除趟次 {report.get('excluded_rows')}，应为 1")
        rec = report.get("reconciliation") or {}
        if rec.get("matched_rows", 0) + rec.get("excluded_rows", 0) != rec.get("base_rows"):
            failures.append("报告里 matched + excluded != base_rows")
        if not rec.get("freight_ok"):
            failures.append("报告里运费核对没有通过")
        if not rec.get("income_ok"):
            failures.append("报告里收入核对没有通过")
        failures += check_first(out)
        failures += check_second(out)
        failures += check_totals(out)

        # 反向情况：一干带吨位、二干只有容间。
        swapped_root = root / "swapped"
        swapped_root.mkdir()
        base2 = swapped_root / "基础报表原版.xlsx"
        finance2 = swapped_root / "财务收入表.xlsx"
        out2 = swapped_root / "out"
        make_swapped_base(base2)
        make_swapped_finance(finance2)
        result2 = run(base2, finance2, out2)
        if result2["code"] != 0:
            print(result2["stdout"][-2000:])
            print(result2["stderr"][-2000:])
            print("[FAIL] 反向分档脚本退出码", result2["code"])
            return 1
        failures += check_swapped(out2)

        # 有一行既没吨位也没容间：必须停下并说明，不能静默丢行。
        dropped_root = root / "dropped"
        dropped_root.mkdir()
        base3 = dropped_root / "基础报表原版.xlsx"
        finance3 = dropped_root / "财务收入表.xlsx"
        out3 = dropped_root / "out"
        make_dropped_base(base3)
        make_finance(finance3)
        result3 = run(base3, finance3, out3)
        if result3["code"] != 2:
            failures.append(f"有行分不出车型时退出码 {result3['code']}，应为 2")
        elif (result3["report"] or {}).get("status") != "base_rows_dropped":
            failures.append("有行分不出车型时，没有返回 base_rows_dropped")

    if failures:
        for item in failures:
            print("[FAIL]", item)
        return 1
    print(f"[OK] 一干容间分档 {len(VOLUME_CASES)} 例、二干吨位分档与备注、"
          f"问题清单、合计口径、反向分档（一干吨位/二干容间）全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
