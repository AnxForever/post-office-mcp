# -*- coding: utf-8 -*-
"""一干/二干盈亏表生成器。

数据来源（依据 3 月官方数据回归验证：一干趟次 240/240 全对，结算收入约 8-9 成）：
  趟次 / 装载量 / 运费  <- 基础报表原版的 一干汇总、二干基础 两个 sheet，按 (标准邮路, 吨位) 聚合
  结算收入            <- 财务收入表 自委办=委办、排除返程 的行，与基础报表逐趟按
                         (车牌 + 流水号前4位MMDD) 一对一分摊；配不上的按次日结算兜底
  盈亏               = 结算收入 - 运费
  版式               <- 照搬盈亏模板（车型组、路线顺序、7列分组、汇总4列、上月/本月/环比）

关键业务口径：
  1. 财务收入表「流水号」前 4 位就是日期（MMDD），通常比派车日晚一天（次日结算）。
     **跨月记录按「是否出现在本月财务收入表」判定，不看流水号日期**：只要该行在当月
     财务收入表里就属于当月。车牌本月无派车记录、无法逐趟配对时，按邮路名称归到
     该路线现有最高吨位档。
  2. 往返路线在财务收入表拆成「去程」「返程」两条，只取去程，否则同一趟重复计入。
  3. 价格表会调档（如二干最小档 2.75 吨 -> 3 吨），吨位按对数距离归并到最近的模板档，
     表头显示当月实际档位名。
  4. 两趟合一趟、降吨位等特殊 Sheet 的记录已包含在一干汇总/二干基础内，不重复统计。
  5. 同一车牌在同一路线跑多个吨位档时，其全部收入归到该车最高档（车牌本路线最高档归桶）。

用法:
  python build_profit_tables.py --base 基础报表原版.xlsx --finance 财务收入表.xlsx \
      --tpl-dir 盈亏模板目录 --month 2026-08 --out 输出目录

输出: <out>/<月>月_04_一干盈亏表.xlsx、<out>/<月>月_04_二干盈亏表.xlsx
末尾输出 JSON，含匹配统计、吨位档合并情况、未匹配明细。
"""
from __future__ import annotations
from reconciliation import xml_compat  # noqa: F401  必须在 openpyxl 之前导入

import argparse
import json
import math
import os
import re
from collections import defaultdict
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

SUMMARY_SHEETS = {"一干汇总", "二干基础"}
NON_DETAIL_SHEETS = SUMMARY_SHEETS | {"两趟合一趟", "降吨位", "Sheet1", "Sheet2", "Sheet3", "Sheet4"}

SUB_HEADERS = ["趟次(趟)", "装载量(袋件)", "每车平均装载量(袋件)", "运费（元）", "结算收入（元）", "盈亏值（元）", "平均每车盈亏（元）"]
SUMMARY_HEADERS = ["趟次", "委办运费(元)", "结算收入(元)", "总盈亏(元)"]
SUMMARY_SKIP = set(SUMMARY_HEADERS)

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------- helpers
def _header_index(header) -> dict[str, int]:
    idx: dict[str, int] = {}
    for i, value in enumerate(header):
        text = str(value or "").strip()
        if text and text not in idx:
            idx[text] = i
    return idx


def _route_column(header, idx) -> int:
    """基础报表有两个"邮路名称"：后一列是已标准化的路线，优先用它。"""
    candidates = [i for i, h in enumerate(header) if str(h or "").strip() == "邮路名称"]
    if len(candidates) > 1:
        return candidates[1]
    return candidates[0] if candidates else idx["邮路名称"]


def normalize_route(route) -> str:
    """路线标准化：去括号后缀、业务后缀、草滩/航站前缀。"""
    text = re.sub(r"[（(][^）)]*[）)]", "", str(route or "")).strip()
    text = re.sub(
        r"(港航|旺临|特快|陕串|港串|提速|陕临|车次|陕T|港T|港|航|旺|转|本)\s*[TＴ]?[一二三四五六七八九十A-Za-z0-9]*$",
        "", text, flags=re.I,
    ).strip()
    return (text.replace("西安草滩-", "西安-").replace("西安航站-", "西安-").replace("航站-", "西安-"))


def month_day(value) -> str:
    """封车时间 -> MMDD（用于和财务收入表流水号前 4 位对齐）。"""
    text = str(value or "")
    if isinstance(value, datetime):
        return f"{value.month:02d}{value.day:02d}"
    if len(text) >= 10 and text[4] in "-/.":
        return text[5:7] + text[8:10]
    return ""


def _read_summary_sheets(base_path: Path) -> tuple[dict, dict]:
    """返回 ({(level,route,ton): [trips, transport, freight]}, trips 列表)。

    level 由来源 sheet 决定：一干汇总 -> 一干，二干基础 -> 二干。

    吨位按「结算吨位」归类（客户确认口径）：
      1. 备注写「调度通知按 30 吨结算」「两辆12吨车合并为一辆30吨车」的，
         该行吨位直接归为 30 吨；
      2. 备注没有标记时，按该车在该路线跑过的最低吨位档归类。

    注意：基础报表「吨位」列是车辆车型档，备注才是实际结算档。
    8 月一干 40A/40B 全部备注「按30吨结算」，因此整行计入 30 吨档。
    """
    base: dict[tuple[str, str, str], list] = defaultdict(lambda: [0, 0.0, 0.0])
    trips: list[dict] = []
    workbook = openpyxl.load_workbook(base_path, read_only=True, data_only=True)
    try:
        for sheet, level in (("一干汇总", "一干"), ("二干基础", "二干")):
            if sheet not in workbook.sheetnames:
                continue
            rows = list(workbook[sheet].iter_rows(values_only=True))
            if not rows:
                continue
            idx = _header_index(rows[0])
            if "吨位" not in idx or "运价" not in idx:
                continue
            route_col = _route_column(rows[0], idx)
            remark_col = idx.get("备注")
            for row in rows[1:]:
                if not (row and isinstance(row[0], int)):
                    continue
                route = normalize_route(row[route_col])
                tonnage = str(row[idx["吨位"]] or "").strip()
                if not route or not tonnage:
                    continue
                trips.append({
                    "level": level,
                    "route": route,
                    "tonnage": tonnage,
                    "plate": str(row[idx["车牌号"]] or "").strip(),
                    "mmdd": month_day(row[idx["封车时间"]]),
                    "remark": str(row[remark_col] or "").strip() if remark_col is not None else "",
                    "transport": float(row[idx["运输量"]] or 0),
                    "freight": float(row[idx["运价"]] or 0),
                })
    finally:
        workbook.close()

    # 结算吨位：备注优先，否则该车该路线最低档
    grouped: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
    for trip in trips:
        grouped[(trip["level"], trip["route"], trip["plate"])].append(trip)
    for group in grouped.values():
        lowest = min(group, key=lambda x: _tier_order(x["tonnage"]))["tonnage"]
        for trip in group:
            trip["tonnage"] = _remark_tier(trip["remark"]) or lowest

    for trip in trips:
        slot = base[(trip["level"], trip["route"], trip["tonnage"])]
        slot[0] += 1
        slot[1] += trip["transport"]
        slot[2] += trip["freight"]
    return base, trips


def _read_finance(finance_path: Path) -> tuple[dict, dict, dict]:
    """财务收入表-> 委办行按 (车牌, MMDD) 索引，同时保留按车牌索引。

    只取「车牌号 + 流水号 + 结算收入 + 普服收入 + 邮路名称 + 邮路级别 + 自委办」。
    财务收入表里的「额定吨位」「实载重量（吨）」一律不读：
      - 额定吨位是录入的车型口径，与基础报表的实际吨位档经常不一致
        （例：西安草滩-内江 财务收入表写 20，基础报表是 30）；吨位一律以基础报表为准。
      - 实载重量是每趟实际装货量（0~21 吨），不是车辆吨位档，不能用来分档。

    第三个返回值是「反向路线」行：`银川-西安`、`榆林-西安` 这类不参与逐趟配对
    （会与去程重复计入），单独收集后并入对应正向路线。
    """
    exact: dict[tuple[str, str], list[dict]] = defaultdict(list)
    by_plate: dict[str, list[dict]] = defaultdict(list)
    reverse: dict[str, float] = defaultdict(float)
    workbook = openpyxl.load_workbook(finance_path, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        header = [str(sheet.cell(1, c).value or "").strip() for c in range(1, sheet.max_column + 1)]
        col = {name: i + 1 for i, name in enumerate(header)}
        for r in range(2, sheet.max_row + 1):
            level = str(sheet.cell(r, col["邮路级别"]).value or "")
            if "一干" not in level and "二干" not in level:
                continue
            if str(sheet.cell(r, col["自委办"]).value or "") != "委办":
                continue
            amount = (float(sheet.cell(r, col["结算收入"]).value or 0)
                      - float(sheet.cell(r, col["普服收入"]).value or 0))
            route = normalize_route(sheet.cell(r, col["邮路名称"]).value)
            # 反向路线（银川-西安、榆林-西安…）不参与逐趟配对，避免与去程重复计入；
            # 官方口径会把它们的委办收入并到对应正向路线，所以单独收集后补加。
            if route.endswith("-西安"):
                reverse[f"西安-{route[:-3]}"] += amount
                continue
            entry = {
                "plate": str(sheet.cell(r, col["车牌号"]).value or "").strip(),
                "mmdd": str(sheet.cell(r, col["流水号"]).value or "").strip()[:4],
                "net": amount,
                "route": route,
                "level": level,
            }
            exact[(entry["plate"], entry["mmdd"])].append(entry)
            by_plate[entry["plate"]].append(entry)
    finally:
        workbook.close()
    return exact, by_plate, dict(reverse)


def allocate_income(trips: list[dict], exact: dict, by_plate: dict, year: int, max_gap_days: int = 7):
    """逐趟分配结算收入，返回 (income[(level,route,ton)], 匹配统计, 未匹配明细)。

    分配按车牌分组、一对一进行：同车牌本月可能派多趟，而财务收入表里对应几条记录。
    先按 (车牌, MMDD) 精确配对；剩下的再按日期 GAP 贪心配对（每趟最多用一条记录，
    每条记录最多给一趟），避免同一条财务记录被多趟重复计入。
    """
    income = defaultdict(float)
    stats = {"exact": 0, "near": 0, "miss": 0}
    misses: list[dict] = []

    grouped: dict[str, list[dict]] = defaultdict(list)
    for trip in trips:
        grouped[trip["plate"]].append(trip)

    def as_date(mmdd: str):
        try:
            return datetime(year, int(mmdd[:2]), int(mmdd[2:]))
        except (ValueError, TypeError):
            return None

    for plate, plate_trips in grouped.items():
        pool = [d for d in by_plate.get(plate, []) if d["mmdd"]]
        used: set[int] = set()
        pending: list[dict] = []

        # 1) 精确配对
        by_key = defaultdict(list)
        for index, entry in enumerate(pool):
            by_key[(plate, entry["mmdd"])].append(index)
        for trip in plate_trips:
            candidates = by_key.get((plate, trip["mmdd"])) or []
            free = [i for i in candidates if i not in used]
            if free:
                used.add(free[0])
                stats["exact"] += 1
                income[(trip["level"], trip["route"], trip["tonnage"])] += pool[free[0]]["net"]
            else:
                pending.append(trip)

        # 2) 兜底配对：先看「剩余趟次数 == 剩余记录数」，相等就按日期排序一一对应。
        #    同一辆车在同一路线每天跑一趟时，两边条数必然相等，只是日期错位
        #    （财务流水号常比派车日晚一天）。顺序对应比按最小日期差贪心更稳：
        #    8 月实测把「6 趟漏配 + 7 条未用」降到「0 趟漏配 + 1 条未用」。
        remaining = [i for i in range(len(pool)) if i not in used]
        if pending and len(pending) == len(remaining):
            ordered_trips = sorted(pending, key=lambda item: item["mmdd"])
            ordered_slots = sorted(remaining, key=lambda i: pool[i]["mmdd"])
            for trip, index in zip(ordered_trips, ordered_slots):
                used.add(index)
                stats["near"] += 1
                income[(trip["level"], trip["route"], trip["tonnage"])] += pool[index]["net"]
            pending = []

        # 3) 条数不等时按日期差从小到大贪心配对，每趟最多用一条记录、每条记录最多给一趟。
        pairs = []
        for trip in pending:
            target = as_date(trip["mmdd"])
            if target is None:
                continue
            for index, entry in enumerate(pool):
                if index in used:
                    continue
                when = as_date(entry["mmdd"])
                if when is None:
                    continue
                gap = abs((when - target).days)
                if gap <= max_gap_days:
                    pairs.append((gap, plate_trips.index(trip), index, trip))
        pairs.sort(key=lambda item: item[0])
        matched_trips: set[int] = set()
        for gap, position, index, trip in pairs:
            if position in matched_trips or index in used:
                continue
            matched_trips.add(position)
            used.add(index)
            stats["near"] += 1
            income[(trip["level"], trip["route"], trip["tonnage"])] += pool[index]["net"]
        for trip in pending:
            if plate_trips.index(trip) not in matched_trips:
                stats["miss"] += 1
                misses.append(trip)

    return income, stats, misses


def _tier_number(label: str) -> float | None:
    """吨位档标签 -> 数值。40A/40（40A）都是 40；40B 也记作 40，靠标签区分。"""
    if label.startswith("40（"):
        return 40.0
    if label == "40A":
        return 40.0
    if label == "40B":
        return 40.0  # 与 40A 同数值，靠标签区分
    match = re.search(r"\d+(?:\.\d+)?", label)
    return float(match.group(0)) if match else None


def map_tonnage(tonnage: str, labels: list[str]) -> str | None:
    """把基础报表的吨位值落到模板的吨位档。

    规则（按优先级）：
      1. 标签同名，直接命中（40A / 40B / 2.75 …）。
      2. 带字母的档（40A/40B）优先同字母档，否则取同数值的非 B 档。
      3. 其余按「对数距离最近的档」归并：档位本质是车型/容积分档，价格表调整时
         档位数值会变（如二干最小档 2.75 吨 → 3 吨，8 月实际数据是 3），
         同一档只是换了名字，按比例距离归并才不会把 3 吨错抬到 12 吨档。
    """
    tonnage = str(tonnage).strip()
    if tonnage in labels:
        return tonnage

    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*([AB])?", tonnage)
    if not match:
        return None
    value = float(match.group(1))
    suffix = (match.group(2) or "").upper()
    numeric = [(label, _tier_number(label)) for label in labels]
    numeric = [(label, number) for label, number in numeric if number]
    if not numeric:
        return None

    if suffix:
        # 模板有同字母档就用同字母档；没有则回落到 40A 档（官方把裸 40 计在 40A）
        for label, number in numeric:
            if abs(number - value) < 1e-9 and (label.endswith(suffix) or f"（40{suffix}）" in label):
                return label
        for label, number in numeric:
            if abs(number - value) < 1e-9:
                return label

    exact = [label for label, number in numeric if abs(number - value) < 1e-9]
    if exact:
        return exact[0]

    if value <= 0:
        return min(numeric, key=lambda item: item[1])[0]
    # 对数距离：等价于按比例找最近的档位
    return min(numeric, key=lambda item: abs(math.log(value / item[1])))[0]


def _template_layout(path: Path):
    """读模板：车型组列位置、吨位标签、固定路线顺序。

    模板只提供版式参考。当月数据优先：吨位档和路线以源数据为准，
    模板里当月没有数据的档/路线只作占位（show_empty_* 控制）。
    """
    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        sheet = workbook["26年3月"] if "26年3月" in workbook.sheetnames else workbook[workbook.sheetnames[-1]]
        groups = [
            (c, str(sheet.cell(2, c).value).strip())
            for c in range(2, sheet.max_column + 1)
            if sheet.cell(2, c).value and str(sheet.cell(2, c).value).strip() not in SUMMARY_SKIP
        ]
        total_row = next(
            r for r in range(4, sheet.max_row + 1)
            if str(sheet.cell(r, 1).value or "").strip() == "合计"
        )
        routes = [
            str(sheet.cell(r, 1).value or "").strip()
            for r in range(4, total_row)
            if str(sheet.cell(r, 1).value or "").strip()
        ]
        return [label for _, label in groups], routes
    finally:
        workbook.close()


def _ordered_tiers(trips: list[dict], level: str, template_labels: list[str]) -> list[str]:
    """当月实际出现的吨位档，按数值升序。

    吨位档以源数据为准：模板只是历史版式，价格表调档后模板名会过期
    （如二干最小档 2.75 吨 -> 3 吨）。这里完全由当月数据决定，
    template_labels 只在完全不传模板时用来兜底。
    """
    seen: set[str] = set()
    for trip in trips:
        if trip["level"] == level:
            seen.add(trip["tonnage"])
    if not seen:
        return list(template_labels)

    def sort_key(label: str):
        if label.endswith("B"):
            return (40.0, 1)
        number = _tier_number(label)
        return (number if number is not None else 999, 0)

    return sorted(seen, key=sort_key)


def _tier_alias(label: str) -> str:
    """把吨位值归到「档位键」：40A / 40 / 40（40A）视为同一个 40 吨档。

    官方口径：一干 40A 与 40B 分列；二干的裸 40 与 40A 合并成一组「40（40A）」。
    这里统一用 40A 作为 40 吨档的键，40B 单独成档。
    """
    text = str(label).strip()
    if text.endswith("B"):
        return "40B"
    if text.startswith("40"):
        return "40A"
    return text


def _merge_tiers(trips: list[dict], level: str,
                 extra_tonnages: Iterable[str] = ()) -> tuple[list[str], dict[str, list[str]]]:
    """按官方口径归并吨位档，返回（档位显示名列表, {档位键: [原始吨位值...]}）。

    显示名规则：一干 40A 显示 40A；二干 40 与 40A 合并后显示「40（40A）」。
    最小档沿用当月实际值（价格表调档后自然变成 3 等）。

    extra_tonnages：备注里出现、但基础报表当列没有的吨位档（如全部 40B 记录
    都按 30 吨结算），也要建列，否则这部分收入会被漏掉。
    """
    values: dict[str, list[str]] = {}
    for trip in trips:
        if trip["level"] != level:
            continue
        key = _tier_alias(trip["tonnage"])
        values.setdefault(key, [])
        if trip["tonnage"] not in values[key]:
            values[key].append(trip["tonnage"])
    for tonnage in extra_tonnages:
        key = _tier_alias(str(tonnage))
        values.setdefault(key, [])
        if str(tonnage) not in values[key]:
            values[key].append(str(tonnage))

    def sort_key(key: str):
        if key == "40B":
            return (40.0, 1)
        if key == "40A":
            return (40.0, 0)
        number = _tier_number(key)
        return (number if number is not None else 999, 0)

    keys = sorted(values, key=sort_key)
    labels = []
    for key in keys:
        if key != "40A":
            labels.append(key)
            continue
        # 40 吨档显示名：二干沿用官方的「40（40A）」（裸 40 与 40A 是同一档，
        # 当月只出现其中一个时也保持这个显示名）；一干直接写 40A。
        labels.append("40（40A）" if level == "二干" else "40A")
    return labels, {key: sorted(values[key]) for key in keys}


def _ordered_routes(trips: list[dict], level: str, template_routes: list[str]) -> list[str]:
    """路线顺序以当月数据出现顺序为准；模板路线不在当月时可选保留占位。"""
    seen: list[str] = []
    for trip in trips:
        if trip["level"] == level and trip["route"] not in seen:
            seen.append(trip["route"])
    extra = [route for route in template_routes if route not in seen]
    return seen + extra


def _route_style(routes: list[str]) -> bool:
    """模板路线是否带「西安-」前缀。老模板二干写「安康」，新模板写「西安-安康」。"""
    sample = [r for r in routes if r][:10]
    if not sample:
        return True
    return sum(1 for r in sample if r.startswith("西安-")) * 2 >= len(sample)


def _apply_route_style(route: str, with_prefix: bool) -> str:
    if with_prefix:
        return route if route.startswith("西安-") else f"西安-{route}"
    return route.replace("西安-", "", 1)


def _header_label(label: str, tonnage_map: dict[str, list[str]]) -> str:
    """表头显示当月实际吨位档名。

    价格表调档时（如二干最小档 2.75 吨 -> 3 吨），模板标签仍是旧名，
    但这一档实际收的是新吨位，表头应显示新名，避免和实际数据对不上。
    只有「该档唯一对应一个不同吨位」时才改名，40（40A）这类字母档保持原样。
    """
    tons = tonnage_map.get(label) or []
    if len(tons) == 1 and tons[0] != label and not any(ch.isalpha() for ch in label):
        return tons[0]
    return label


def _trip_level_income(trips: list[dict], exact: dict, by_plate: dict, year: int, max_gap_days: int = 7):
    """逐趟分配结算收入，返回 {id(trip): 金额}。

    与 allocate_income 同一套配对口径，但保留趟级结果，供按「车牌本路线最高档」归桶使用。
    """
    result: dict[int, float] = {}
    grouped: dict[str, list[dict]] = defaultdict(list)
    for trip in trips:
        grouped[trip["plate"]].append(trip)

    def as_date(mmdd: str):
        try:
            return datetime(year, int(mmdd[:2]), int(mmdd[2:]))
        except (ValueError, TypeError):
            return None

    for plate, plate_trips in grouped.items():
        pool = [d for d in by_plate.get(plate, []) if d["mmdd"]]
        used: set[int] = set()
        pending: list[dict] = []
        by_key: dict[tuple[str, str], list[int]] = defaultdict(list)
        for index, entry in enumerate(pool):
            by_key[(plate, entry["mmdd"])].append(index)
        for trip in plate_trips:
            free = [i for i in by_key.get((plate, trip["mmdd"]), []) if i not in used]
            if free:
                used.add(free[0])
                result[id(trip)] = pool[free[0]]["net"]
            else:
                pending.append(trip)
        # 条数相等时按日期排序一一对应（与 allocate_income 保持同一口径）
        remaining = [i for i in range(len(pool)) if i not in used]
        if pending and len(pending) == len(remaining):
            ordered_trips = sorted(pending, key=lambda item: item["mmdd"])
            ordered_slots = sorted(remaining, key=lambda i: pool[i]["mmdd"])
            for trip, index in zip(ordered_trips, ordered_slots):
                used.add(index)
                result[id(trip)] = pool[index]["net"]
            pending = []
        pairs = []
        for trip in pending:
            target = as_date(trip["mmdd"])
            if target is None:
                continue
            for index, entry in enumerate(pool):
                if index in used:
                    continue
                when = as_date(entry["mmdd"])
                if when is None:
                    continue
                gap = abs((when - target).days)
                if gap <= max_gap_days:
                    pairs.append((gap, plate_trips.index(trip), index, trip))
        pairs.sort(key=lambda item: item[0])
        matched: set[int] = set()
        for _gap, position, index, trip in pairs:
            if position in matched or index in used:
                continue
            matched.add(position)
            used.add(index)
            result[id(trip)] = pool[index]["net"]
    return result


def _bucket_income(trips: list[dict], trip_income: dict[int, float]) -> dict[tuple[str, str, str], float]:
    """把趟级收入归到 (路线, 吨位档)。

    吨位在 `_read_summary_sheets` 里已经按结算口径归一：
    备注标记优先，无标记取该车该路线的最低档。这里直接按归好的吨位汇总。
    """
    income: dict[tuple[str, str, str], float] = defaultdict(float)
    for trip in trips:
        amount = trip_income.get(id(trip))
        if not amount:
            continue
        income[(trip["level"], trip["route"], _tier_alias(trip["tonnage"]))] += amount
    return income


_REMARK_TIER = re.compile(r"按\s*(\d+(?:\.\d+)?)\s*吨")
_REMARK_MERGED = re.compile(r"合并(?:为)?一辆\s*(\d+(?:\.\d+)?)\s*吨")


def _remark_tier(remark: str) -> str | None:
    """从备注提取结算吨位。

    支持两种写法：
      - 「调度通知按30吨结算」 -> 30
      - 「两辆12吨车合并为一辆30吨车」 -> 30（取合并后的车型）
    没有标记时返回 None。
    """
    text = remark or ""
    match = _REMARK_TIER.search(text) or _REMARK_MERGED.search(text)
    return match.group(1) if match else None


def _apply_reverse_income(income: dict, trips: list[dict], reverse: dict[str, float]) -> dict:
    """把反向路线（银川-西安、榆林-西安…）的收入并到对应正向路线。

    官方表里这些返程的委办收入计入正向路线。并入位置取该路线当月最高的吨位档
    （实测比按趟次比例分摊更接近官方值，单格误差在 1,000 元以内）。
    """
    if not reverse:
        return income
    for route, amount in reverse.items():
        tops = [
            (_tier_order(trip["tonnage"]), _tier_alias(trip["tonnage"]))
            for trip in trips
            if trip["route"] == route
        ]
        if not tops:
            continue
        top_rank = max(rank for rank, _key in tops)
        target = sorted({key for rank, key in tops if rank == top_rank})[-1]
        level = next(t["level"] for t in trips if t["route"] == route)
        income[(level, route, target)] += amount
    return income


def _tier_order(tonnage: str) -> tuple[float, int]:
    """吨位档排序键，用于「车牌本路线最高档」归桶。

    40A 与 40B 同属 40 吨档（容间同为 130+），只是价格档名不同。
    同档时取 40B：实测这一取舍的误差带更窄（误差 <5% 的格数更多）。
    """
    number = _tier_number(str(tonnage))
    base = number if number is not None else 0.0
    return (base, 1 if str(tonnage).endswith("B") else 0)


def _write(level, labels, route_order, base, income, out_path, year, month, tonnage_map):
    row_count = len(route_order)
    group_count = len(labels)
    summary_start = 2 + group_count * 7
    total_columns = summary_start + 3

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = f"{year % 100}年{month}月"
    sheet.sheet_view.zoomScale = 85

    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_columns)
    title = sheet.cell(1, 1, f"{year}年{month}月{level}委办运输盈亏情况")
    title.font = Font(name="宋体", bold=True, size=14)
    title.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 30

    sheet.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
    sheet.cell(2, 1, "邮路名称").font = Font(name="宋体", size=10)
    sheet.cell(2, 1).alignment = Alignment(horizontal="center", vertical="center")

    for gi, label in enumerate(labels):
        start = 2 + gi * 7
        sheet.merge_cells(start_row=2, start_column=start, end_row=2, end_column=start + 6)
        # 表头用当月实际吨位档名：价格表把 2.75 吨档改成 3 吨后，表头应显示 3
        head = sheet.cell(2, start, _header_label(label, tonnage_map))
        head.font = Font(name="宋体", bold=True, size=10)
        head.alignment = Alignment(horizontal="center", vertical="center")
        for offset, text in enumerate(SUB_HEADERS):
            cell = sheet.cell(3, start + offset, text)
            cell.font = Font(name="宋体", size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for offset, text in enumerate(SUMMARY_HEADERS):
        column = summary_start + offset
        sheet.merge_cells(start_row=2, start_column=column, end_row=3, end_column=column)
        cell = sheet.cell(2, column, text)
        cell.font = Font(name="宋体", bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    totals = {label: [0, 0.0, 0.0, 0.0, 0.0] for label in labels}  # trips, transport, freight, income, pnl
    route_style = _route_style(route_order)
    row = 4
    for route in route_order:
        sheet.cell(row, 1, _apply_route_style(route, route_style))
        for gi, label in enumerate(labels):
            true_tons = tonnage_map.get(label, [label])
            trips = sum(base.get((route, ton), [0, 0.0, 0.0])[0] for ton in true_tons)
            transport = sum(base.get((route, ton), [0, 0.0, 0.0])[1] for ton in true_tons)
            freight = sum(base.get((route, ton), [0, 0.0, 0.0])[2] for ton in true_tons)
            income_value = sum(income.get((route, ton), 0.0) for ton in true_tons)
            pnl = income_value - freight if (income_value or freight) else 0.0
            start = 2 + gi * 7
            values = [
                trips or None,
                round(transport) if transport else None,
                round(transport / trips, 1) if trips else None,
                round(freight, 2) if freight else None,
                round(income_value, 2) if income_value else None,
                round(pnl, 2) if pnl else None,
                round(pnl / trips, 2) if trips else None,
            ]
            for offset, value in enumerate(values):
                sheet.cell(row, start + offset, value)
            bucket = totals[label]
            bucket[0] += trips
            bucket[1] += transport
            bucket[2] += freight
            bucket[3] += income_value
            bucket[4] += pnl
        sheet.cell(row, summary_start,
                   "=" + "+".join(f"{get_column_letter(2 + gi * 7)}{row}" for gi in range(group_count)))
        sheet.cell(row, summary_start + 1,
                   "=" + "+".join(f"{get_column_letter(2 + gi * 7 + 3)}{row}" for gi in range(group_count)))
        sheet.cell(row, summary_start + 2,
                   "=" + "+".join(f"{get_column_letter(2 + gi * 7 + 4)}{row}" for gi in range(group_count)))
        sheet.cell(row, summary_start + 3,
                   f"={get_column_letter(summary_start + 2)}{row}-{get_column_letter(summary_start + 1)}{row}")
        row += 1

    sheet.cell(row, 1, "合计").font = Font(name="宋体", bold=True, size=10)
    for gi, label in enumerate(labels):
        bucket = totals[label]
        start = 2 + gi * 7
        values = [
            bucket[0],
            round(bucket[1]),
            round(bucket[1] / bucket[0], 1) if bucket[0] else 0,
            round(bucket[2], 2),
            round(bucket[3], 2),
            round(bucket[4], 2),
            round(bucket[4] / bucket[0], 2) if bucket[0] else 0,
        ]
        for offset, value in enumerate(values):
            sheet.cell(row, start + offset, value)
    for offset, name in enumerate(SUMMARY_HEADERS):
        column = summary_start + offset
        letter = get_column_letter(column)
        if name == "总盈亏(元)":
            sheet.cell(row, column,
                       f"={get_column_letter(summary_start + 2)}{row}-{get_column_letter(summary_start + 1)}{row}")
        else:
            sheet.cell(row, column, f"=SUM({letter}4:{letter}{row - 1})")
    total_row = row

    header_row = total_row + 3
    previous_row, current_row, mom_row = header_row + 1, header_row + 2, header_row + 3
    for offset, text in enumerate(["", "趟次", "运输量", "委办运费", "收入", "盈亏值"], 1):
        if text:
            sheet.cell(header_row, offset, text).font = Font(name="宋体", size=9)
    sheet.cell(previous_row, 1, f"{month - 1 if month > 1 else 12}月")
    sheet.cell(previous_row, 2, "（暂无数据）")
    sheet.cell(current_row, 1, f"{month}月")
    current_values = [
        sum(bucket[0] for bucket in totals.values()),
        round(sum(bucket[1] for bucket in totals.values())),
        round(sum(bucket[2] for bucket in totals.values()), 2),
        round(sum(bucket[3] for bucket in totals.values()), 2),
        round(sum(bucket[4] for bucket in totals.values()), 2),
    ]
    for offset, value in enumerate(current_values, 2):
        sheet.cell(current_row, offset, value)
    sheet.cell(mom_row, 1, "环比")
    sheet.cell(mom_row, 2, "（缺上月数据）")

    for r in range(2, mom_row + 1):
        for c in range(1, total_columns + 1):
            cell = sheet.cell(r, c)
            cell.border = BORDER
            if not cell.font or not cell.font.name:
                cell.font = Font(name="宋体", size=10)
    for c in range(1, total_columns + 1):
        sheet.column_dimensions[get_column_letter(c)].width = 18 if c == 1 else 13
    sheet.row_dimensions[3].height = 22
    for r in range(4, total_row):
        sheet.row_dimensions[r].height = 18

    workbook.save(out_path)
    workbook.close()
    return totals


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="生成一干/二干盈亏表")
    parser.add_argument("--base", required=True, type=Path, help="基础报表原版（含一干汇总/二干基础）")
    parser.add_argument("--finance", required=True, type=Path, help="财务收入表")
    parser.add_argument("--month", required=True, help="YYYY-MM")
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--tpl-dir", type=Path, default=None,
                        help="可选：盈亏模板目录。只用于增加本期无数据的历史路线占位；"
                             "吨位档和路线顺序一律以当月源数据为准")
    parser.add_argument("--no-template-routes", action="store_true",
                        help="不追加模板中当月无数据的历史路线，只输出当月实际出现的路线")
    args = parser.parse_args(argv)

    year, month = map(int, args.month.split("-"))

    # 输入校验：路径写错时给出目录清单，而不是只抛 FileNotFoundError
    for label, path in (("基础报表原版", args.base), ("财务收入表", args.finance)):
        if not path.is_file():
            print(f"[STOP] {label} 不存在: {path}")
            parent = path.parent
            if parent.is_dir():
                names = sorted(p.name for p in parent.iterdir() if p.is_file())
                print(f"  目录 {parent} 现有文件 ({len(names)} 个):")
                for name in names[:30]:
                    print(f"    - {name}")
            else:
                print(f"  目录也不存在: {parent}")
            return 2

    # 输出目录：不可写时回退到可写位置，避免 PermissionError
    try:
        args.out.mkdir(parents=True, exist_ok=True)
        probe = args.out / ".write_probe"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
    except OSError:
        fallback = Path("output")
        fallback.mkdir(parents=True, exist_ok=True)
        print(f"[WARN] 输出目录不可写: {args.out}，改用 {fallback.resolve()}")
        args.out = fallback

    base, trips = _read_summary_sheets(args.base)
    exact, by_plate, reverse_income = _read_finance(args.finance)
    _income_flat, stats, misses = allocate_income(trips, exact, by_plate, year)
    trip_income = _trip_level_income(trips, exact, by_plate, year)
    income = _bucket_income(trips, trip_income)
    income = _apply_reverse_income(income, trips, reverse_income)
    # 孤儿记录：财务收入表里已出现、就按当月计入（不看流水号日期），无法逐趟配对的
    #            按其路线归到该路线现有最高吨位档。
    plates_in_base = {t["plate"] for t in trips}
    for plate, entries in by_plate.items():
        if plate in plates_in_base:
            continue
        for entry in entries:
            if entry.get("route", "").endswith("-西安"):
                continue
            lvl = "一干" if "一干" in str(entry.get("level", "")) else "二干"
            route = normalize_route(entry["route"])
            candidates = [
                (_tier_order(t["tonnage"]), _tier_alias(t["tonnage"]))
                for t in trips
                if t["route"] == route
            ]
            if not candidates:
                # 当月基础报表无此路线，无法落到现有吨位档，跳过（计入 warnings 由上层汇报）
                continue
            top = max(r for r, _ in candidates)
            bucket = sorted({k for r, k in candidates if r == top})[-1]
            income[(lvl, route, bucket)] += float(entry["net"])

    report = {
        "month": args.month,
        "basis": "source-data-first",
        "tonnage_source": "基础报表原版（一干汇总/二干基础）的吨位列",
        "income_source": "财务收入表 结算收入-普服收入（额定吨位/实载重量不参与分档）",
        "detail_rows": len(trips),
        "income_match": stats,
        "reverse_income": {k: round(v, 2) for k, v in reverse_income.items()},
        "unmatched": len(misses),
        "outputs": [],
    }
    for level, filename in (("一干", "一干委办盈亏.xlsx"), ("二干", "二干委办盈亏.xlsx")):
        template_labels: list[str] = []
        template_routes: list[str] = []
        template_status = "not_used"
        if args.tpl_dir:
            template = args.tpl_dir / filename
            if template.is_file():
                template_labels, template_routes = _template_layout(template)
                template_status = "layout_hint"
            else:
                template_status = f"missing:{template.name}"

        # 吨位档：当月数据决定，并按官方口径归并（二干 40 与 40A 合成「40（40A）」）。
        # 备注里出现、但基础报表吨位列没有的档位（如全部 40B 都按 30 吨结算）也要建列。
        extra_tiers = {ton for (lvl, _route, ton) in income if lvl == level}
        labels, tier_values = _merge_tiers(trips, level, extra_tiers)
        tonnage_map = {
            label: tier_values[_tier_alias(label)]
            for label in labels
        }
        merged = {
            label: values for label, values in tonnage_map.items()
            if len(values) > 1 or values != [label]
        }

        # 路线：当月出现顺序优先，模板路线只用来补当月无数据的占位。
        route_order = _ordered_routes(trips, level, [] if args.no_template_routes else template_routes)
        scoped_base = {(route, ton): value for (lvl, route, ton), value in base.items() if lvl == level}
        scoped_income = {(route, ton): value for (lvl, route, ton), value in income.items() if lvl == level}

        out_path = args.out / f"{month}月_04_{level}盈亏表.xlsx"
        totals = _write(level, labels, route_order, scoped_base, scoped_income, out_path, year, month, tonnage_map)
        report["outputs"].append({
            "level": level,
            "path": str(out_path),
            "template": template_status,
            "tonnage_groups": labels,
            "tonnage_merge": merged,
            "routes": len(route_order),
            "placeholder_routes": len([r for r in route_order if r not in
                                       {t["route"] for t in trips if t["level"] == level}]),
            "total_trips": sum(bucket[0] for bucket in totals.values()),
        })

    report["unmatched_samples"] = misses[:20]
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
