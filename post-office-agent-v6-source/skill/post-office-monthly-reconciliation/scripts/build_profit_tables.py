# -*- coding: utf-8 -*-
"""一干/二干盈亏表生成器。

数据来源（依据 3 月官方数据回归验证：一干趟次 240/240 全对，结算收入约 8-9 成）：
  趟次 / 装载量 / 运费  <- 基础报表原版的 一干汇总、二干基础 两个 sheet，
                         按 (标准邮路, 档位) 聚合。当月数据给了「吨位」列就用吨位
                         （备注优先）；没有吨位才按「容间」区间分档。不按一干/二干
                         写死，哪张表给什么就用什么；个别行没有吨位的，按容间推回。
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
  3. 基础报表的「吨位」列是车辆车型档，备注才是实际结算档：备注写「调度通知按30吨结算」
     「两辆12吨车合并一辆30吨车」的，整行按备注的吨位算；没备注的按该车该路线
     跑过的最低档算。当月表里没有吨位列时，按「容间」区间分档。
  4. 两趟合一趟、降吨位等特殊 Sheet 的记录已包含在一干汇总/二干基础内，不重复统计。
  5. 返程（银川-西安 等）收入不单独占行，并入对应正向路线的最高档。

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

from reconciliation.paths import publish_files

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


def year_month(value) -> tuple[int, int] | None:
    """封车时间 -> (年, 月)，用于校对基础报表是不是目标月份的数据。"""
    if isinstance(value, datetime):
        return (value.year, value.month)
    text = str(value or "").strip()
    match = re.match(r"(\d{4})\s*[-/.年]\s*(\d{1,2})", text)
    if match:
        return (int(match.group(1)), int(match.group(2)))
    return None


def _column(header, *names):
    """按表头别名找列号（不同导出格式同一列叫法不同）。"""
    for name in names:
        for i, value in enumerate(header):
            if str(value or "").strip() == name:
                return i
    return None


# 「容间 → 表头档位」：客户 2026-10 手写确认的区间，原样照用，不自己拟阈值。
#   小于55 / [55,100) / [100,130) / [130,150) / [150,180) / 大于180
# 客户原话：生成盈亏表时没有吨位就按容积来划分。
VOLUME_BINS = (
    (180, "大于180"),
    (150, "[150,180)"),
    (130, "[130,150)"),
    (100, "[100,130)"),
    (55, "[55,100)"),
    (0, "小于55"),
)
BIN_ORDER = ["小于55", "[55,100)", "[100,130)", "[130,150)", "[150,180)", "大于180"]


def _bin_of(volume) -> str:
    """容间（m³）-> 表头档位名。180 整按「大于180」处理。"""
    try:
        value = float(volume)
    except (TypeError, ValueError):
        return ""
    for lower, label in VOLUME_BINS:
        if value >= lower:
            return label
    return "小于55"


# 问题清单里给用户看的说法：不用 [55,100) 这种数学写法。
_PLAIN_BIN = {
    "小于55": "55方以下",
    "[55,100)": "55–99方",
    "[100,130)": "100–129方",
    "[130,150)": "130–149方",
    "[150,180)": "150–179方",
    "大于180": "180方以上",
}


def _plain_bin(label) -> str:
    return _PLAIN_BIN.get(str(label or "").strip(), str(label or ""))


# 「容间 → 吨位档」：当月表里没有吨位、或个别行吨位为空时用。
# 区间与二干价格表的容积下限一致，也与当月二干数据实测对得上
# （3 吨 16~19 方、12 吨 55~91 方、20 吨 127~128 方、30 吨 130~147 方、
#   40A 153~162 方、40B 183 方以上）。
_TIER_LINES = (
    (180, "40B"),
    (150, "40A"),
    (130, "30"),
    (100, "20"),
    (55, "12"),
    (45, "8"),
    (35, "5"),
    (16, "3"),
)


def _tier_from_volume(volume) -> str:
    """容间（m³）-> 吨位档。没有吨位列时用，拿不到容间返回空串。"""
    try:
        value = float(volume)
    except (TypeError, ValueError):
        return ""
    for lower, tier in _TIER_LINES:
        if value >= lower:
            return tier
    return "3"


# 客户确认的结算吨位：只对确认过的那一趟/那几趟生效，不是长期规则。
# 2026-11-02 客户确认：西安-陇南 苏E6W073，2026-09-08 这一趟（容间 35 方，
# 运费 2,250）按 3 吨结算，不是 5 吨。按「邮路关键字 + 车牌 + 日期」登记，
# 同车其他日期、其他邮路都不受影响；remark 里另有明确写法的，仍以 remark 为准。
_CONFIRMED_SETTLEMENTS = (
    {"route": "陇南", "plate": "苏E6W073", "mmdd": "0908", "tier": "3"},
)


def _confirmed_tier(trip: dict) -> str:
    """查客户确认过的结算吨位；没有就返回空串。"""
    route = trip.get("route") or ""
    plate = trip.get("plate") or ""
    mmdd = trip.get("mmdd") or ""
    for item in _CONFIRMED_SETTLEMENTS:
        if item["route"] not in route or item["plate"] != plate:
            continue
        if item.get("mmdd") and item["mmdd"] != mmdd:
            continue
        return item["tier"]
    return ""


def _read_summary_sheets(base_path: Path) -> tuple[dict, dict, dict]:
    """返回 ({(level,route,ton): [trips, transport, freight]}, trips 列表, 诊断信息)。

    level 先由来源 sheet 决定（一干汇总 -> 一干，二干基础 -> 二干），
    之后 _reclassify_by_finance 会按财务收入表的邮路级别纠正；
    档位在 _assign_settlement_tiers 里按「数据给了什么」再落一次：
      - 表里有「吨位」列且有值：用吨位列；个别行没有吨位的按「容间」推回。
        备注写了结算吨位的，以备注为准。
      - 表里没有「吨位」列：按客户 2026-10 确认的容间区间分档。

    不按一干/二干写死：哪张表带了吨位列，就以吨位为准；没带就用容间。

    注意：基础报表「吨位」列是车辆车型档，备注才是实际结算档。

    表头按别名匹配：官方导出是 吨位/运输量/封车时间，车管平台汇总导出是
    系数/邮运量/邮路执行日期。认不出来的 sheet 记进诊断信息，交给调用方报错，
    绝不静默跳过（跳过会让一干表整张变成空表，而退出码仍是 0）。
    """
    trips: list[dict] = []
    diag: dict[str, dict] = {}
    months: dict[tuple[int, int], int] = defaultdict(int)
    workbook = openpyxl.load_workbook(base_path, read_only=True, data_only=True)
    try:
        for sheet, level in (("一干汇总", "一干"), ("二干基础", "二干")):
            if sheet not in workbook.sheetnames:
                diag[sheet] = {"status": "missing"}
                continue
            rows = list(workbook[sheet].iter_rows(values_only=True))
            if not rows:
                diag[sheet] = {"status": "empty"}
                continue
            idx = _header_index(rows[0])
            c_ton = _column(rows[0], "吨位")
            c_vol = _column(rows[0], "容间")
            c_qty = _column(rows[0], "运输量", "邮运量")
            c_time = _column(rows[0], "封车时间", "邮路执行日期", "邮路执行时间")
            c_price = _column(rows[0], "运价")
            c_plate = _column(rows[0], "车牌号")
            missing = []
            if c_ton is None and c_vol is None:
                missing.append("吨位/容间")
            if c_qty is None:
                missing.append("运输量/邮运量")
            if c_time is None:
                missing.append("封车时间/邮路执行日期")
            if c_price is None:
                missing.append("运价")
            if c_plate is None:
                missing.append("车牌号")
            if missing:
                diag[sheet] = {
                    "status": "unreadable",
                    "missing_headers": missing,
                    "has_tonnage_column": c_ton is not None,
                    "headers": [str(v).strip() for v in rows[0] if str(v or "").strip()],
                }
                continue
            route_col = _route_column(rows[0], idx)
            remark_col = idx.get("备注")
            parsed = 0
            no_tier = 0
            skipped: list[dict] = []
            for row in rows[1:]:
                if not (row and isinstance(row[0], int)):
                    continue
                route = normalize_route(row[route_col])
                tonnage = str(row[c_ton] or "").strip() if c_ton is not None else ""
                volume = None
                if c_vol is not None:
                    try:
                        float(row[c_vol])
                    except (TypeError, ValueError):
                        pass
                    else:
                        volume = row[c_vol]
                if not route or (not tonnage and volume is None):
                    no_tier += 1 if route else 0
                    try:
                        qty = float(row[c_qty] or 0)
                        price = float(row[c_price] or 0)
                    except (TypeError, ValueError):
                        qty = price = 0.0
                    if route or qty or price:
                        skipped.append({
                            "route": route,
                            "plate": str(row[c_plate] or "").strip(),
                            "date": str(row[c_time] or ""),
                            "freight": price,
                        })
                    continue
                stamp = year_month(row[c_time])
                if stamp:
                    months[stamp] += 1
                trips.append({
                    "level": level,
                    "route": route,
                    # 先给个临时档位；_assign_settlement_tiers 会按最终归属重算。
                    "tonnage": tonnage or _bin_of(volume),
                    "tonnage_raw": tonnage,
                    "volume": volume,
                    "plate": str(row[c_plate] or "").strip(),
                    "mmdd": month_day(row[c_time]),
                    "remark": str(row[remark_col] or "").strip() if remark_col is not None else "",
                    "transport": float(row[c_qty] or 0),
                    "freight": float(row[c_price] or 0),
                })
                parsed += 1
            diag[sheet] = {
                "status": "ok", "rows": len(rows) - 1, "parsed": parsed,
                "tier_source": "吨位列" if c_ton is not None else "容间分档",
            }
            if no_tier:
                diag[sheet]["rows_without_tier"] = no_tier
            if skipped:
                diag[sheet]["skipped_rows"] = len(skipped)
                diag[sheet]["skipped_freight"] = round(
                    sum(item["freight"] for item in skipped), 2)
                diag[sheet]["skipped_samples"] = skipped[:10]
    finally:
        workbook.close()
    diag["months"] = {f"{y:04d}-{m:02d}": n for (y, m), n in sorted(months.items())}

    return _aggregate_base(trips), trips, diag


def _aggregate_base(trips: list[dict]) -> dict:
    """(level, route, 档位键) -> [趟次, 装载量, 运费]。"""
    base: dict[tuple[str, str, str], list] = defaultdict(lambda: [0, 0.0, 0.0])
    for trip in trips:
        slot = base[(trip["level"], trip["route"], _tier_alias(trip["tonnage"]))]
        slot[0] += 1
        slot[1] += trip["transport"]
        slot[2] += trip["freight"]
    return base


def _finance_route_levels(exact: dict) -> dict[str, str]:
    """财务收入表里每条标准路线的一/二干归属（按行数取多数）。"""
    votes: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
    for entries in exact.values():
        for entry in entries:
            level = "一干" if "一干" in str(entry.get("level", "")) else "二干"
            votes[normalize_route(entry["route"])][level] += 1
    return {route: max(counts, key=counts.get) for route, counts in votes.items()}


def _reclassify_by_finance(trips: list[dict], levels: dict[str, str]) -> list[dict]:
    """一/二干以客户结算口径（财务收入表）为准，基础报表的 sheet 名只是兜底。

    9 月就踩过一次：平凉/固原/陇南/天水/庆阳 220 趟被车管记进「一干汇总」，
    财务收入表全部记为二干；归回二干后两边委办趟次才和运输局对上（1453/1803）。
    """
    moved: dict[tuple[str, str, str], dict] = {}
    for trip in trips:
        target = levels.get(trip["route"])
        if not target or target == trip["level"]:
            continue
        slot = moved.setdefault(
            (trip["route"], trip["level"], target),
            {"trips": 0, "transport": 0.0, "freight": 0.0},
        )
        slot["trips"] += 1
        slot["transport"] += trip["transport"]
        slot["freight"] += trip["freight"]
        trip["level"] = target
    return [
        {"route": route, "from": src, "to": dst, "trips": v["trips"],
         "transport": round(v["transport"], 2), "freight": round(v["freight"], 2)}
        for (route, src, dst), v in sorted(moved.items())
    ]


def _adopt_base_month(diag: dict, year: int, month: int) -> str | None:
    """数据月份比目标月新时，按数据出表，返回该月份；否则返回 None。

    文件名定的是「要出哪个月的账」，数据才是「实际给到的月份」。两者不一致时不能一刀切：
      - 数据比目标月**新**：多半是文件名写错（把 9 月的数据存成了 8月的名字），按数据走；
      - 数据比目标月**旧**：数据没更新，必须停下（交给 _base_guard 报 month_mismatch），
        否则会出一张月份错的表，而且看起来一切正常。
    跨两个月的数据不猜，也不放行。
    """
    months = sorted(diag.get("months") or {})
    target = f"{year:04d}-{month:02d}"
    if len(months) != 1:
        return None
    return months[0] if months[0] > target else None


def _base_guard(diag: dict, year: int, month: int) -> dict | None:
    """核对基础报表原版：读不出来、整张为空、或月份对不上，都必须硬失败。

    这三种情况以前都会「照常出表」：一干汇总被静默跳过 -> 一干表只有合计 0 行；
    上传成上个月的文件 -> 出一张上个月的表，只是名字跟着数据变。都不该继续。
    """
    absent = [s for s in ("一干汇总", "二干基础") if diag.get(s, {}).get("status") == "missing"]
    if absent:
        return {
            "status": "missing_base_sheet",
            "message": f"基础报表原版里没有 {'、'.join(absent)} 工作表",
            "next_action": "请上传官方格式的基础报表原版（含 一干汇总/二干基础 两个 sheet）。",
        }

    unusable = {s: d for s, d in diag.items()
                if s in ("一干汇总", "二干基础") and d.get("status") == "unreadable"}
    if unusable:
        detail = "；".join(
            f"{sheet} 缺少列 {d['missing_headers']}（实际表头：{'、'.join(d['headers'])}）"
            for sheet, d in unusable.items()
        )
        return {
            "status": "unreadable_base",
            "message": f"基础报表原版的工作表读不出来：{detail}",
            "next_action": "请上传官方格式的基础报表原版（含 一干汇总/二干基础 两个 sheet），"
                           "或把该导出格式发给我们适配。",
        }

    empty = [s for s in ("一干汇总", "二干基础")
             if diag.get(s, {}).get("status") in ("ok", "empty") and not diag[s].get("parsed")]
    if empty:
        return {
            "status": "empty_base",
            "message": f"基础报表原版的 {'、'.join(empty)} 一行有效数据都没读到"
                       f"（该 sheet 共 {diag[empty[0]].get('rows', 0)} 行）",
            "next_action": "请确认该 sheet 有当月派车记录后重试。",
        }

    dropped = {
        sheet: info for sheet, info in diag.items()
        if sheet in ("一干汇总", "二干基础") and info.get("skipped_rows")
    }
    if dropped:
        detail = "；".join(
            f"{sheet} {info['skipped_rows']} 行（运费 {info.get('skipped_freight', 0):,.2f} 元）"
            for sheet, info in dropped.items()
        )
        return {
            "status": "base_rows_dropped",
            "message": f"基础报表里有趟次分不出车型/容间，不能静默丢掉：{detail}",
            "skipped_samples": {sheet: info.get("skipped_samples")
                                for sheet, info in dropped.items()},
            "next_action": "请核对这些行的吨位/容间，补齐后再出表；否则成本和趟次对不上。",
        }

    months = diag.get("months") or {}
    target = f"{year:04d}-{month:02d}"
    if months and target not in months:
        found = "、".join(f"{name}（{count} 行）" for name, count in months.items())
        return {
            "status": "month_mismatch",
            "message": f"基础报表原版的业务日期是 {found}，和目标月份 {target} 对不上",
            "next_action": f"请确认上传的是 {target} 的基础报表原版；"
                           "文件名写着 9 月但内容是 8 月的情况已经出过一次，务必核对数据日期。",
        }
    return None


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
    """档位 -> 排序数值。

    容间区间用区间序号（0~5）；吨位档用吨位数值（3 / 12 / 30 / 40 …）。
    两种档位不会出现在同一张表里，各自的数值只用来比大小。
    """
    text = str(label).strip()
    if text in BIN_ORDER:
        return float(BIN_ORDER.index(text))
    if text.startswith("40"):
        return 40.0
    match = re.search(r"\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else None


def _tier_sort_key(label: str) -> tuple[float, int]:
    """档位排序键：40A 在 40B 前；容间区间按区间顺序。"""
    text = str(label).strip()
    if text == "40B":
        return (40.0, 1)
    if text in ("40A", "40（40A）", "40"):
        return (40.0, 0)
    number = _tier_number(text)
    return (number if number is not None else 999.0, 0)


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
    """当月实际出现的档位，从低到高排序。"""
    seen: set[str] = set()
    for trip in trips:
        if trip["level"] == level:
            seen.add(trip["tonnage"])
    if not seen:
        return list(template_labels)
    return sorted(seen, key=_tier_sort_key)


def _tier_alias(label: str) -> str:
    """档位键：40 / 40A 合成一档（40A），40B 单独一档；容间区间原样保留。

    Excel 把吨位存成数值时会是 12.0，这里先归一成 12。
    """
    text = str(label).strip()
    whole = re.fullmatch(r"(\d+)\.0+", text)
    if whole:
        text = whole.group(1)
    if text.endswith("B"):
        return "40B"
    if text.startswith("40"):
        return "40A"
    return text


def _merge_tiers(trips: list[dict], level: str,
                 extra_tonnages: Iterable[str] = ()) -> tuple[list[str], dict[str, list[str]]]:
    """返回（当月出现的档位显示名列表, {显示名: [档位键]}）。

    - 有吨位档时按吨位分列；40 与 40A 在二干模板里合并显示为「40（40A）」，
      一干模板保持 40A 单列。
    - 没有吨位时按客户确认的容间区间分列，显示名就是区间名。
    extra_tonnages 保留参数位：收入里有、基础报表当列没有的档位也要建列。
    """
    values: dict[str, list[str]] = {}
    for trip in trips:
        if trip["level"] != level:
            continue
        key = _tier_alias(trip["tonnage"])
        values.setdefault(key, [])
        if key not in values[key]:
            values[key].append(key)
    for tonnage in extra_tonnages:
        key = _tier_alias(str(tonnage))
        values.setdefault(key, [])
        if key not in values[key]:
            values[key].append(key)

    keys = sorted(values, key=_tier_sort_key)
    labels = ["40（40A）" if key == "40A" and level == "二干" else key for key in keys]
    return labels, {label: values[key] for label, key in zip(labels, keys)}


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
    """表头显示名原样返回；40（40A）这类合并名在 _merge_tiers 里已经定好。"""
    return label


def _trip_level_income(trips: list[dict], exact: dict, by_plate: dict, year: int, max_gap_days: int = 7):
    """逐趟分配结算收入，返回 ({id(trip): 金额}, 没用上的财务收入记录)。

    与 allocate_income 同一套配对口径，但保留趟级结果，供按「车牌本路线最高档」归桶使用。
    没用上的记录不能丢：主表不算，交给问题清单让人核对。
    """
    result: dict[int, float] = {}
    unused: list[dict] = []
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
        for index, entry in enumerate(pool):
            if index not in used:
                unused.append(entry)
        for entry in by_plate.get(plate, []):
            if not entry.get("mmdd"):
                unused.append(entry)
    return result, unused


def _bucket_income(trips: list[dict], trip_income: dict[int, float]) -> dict[tuple[str, str, str], float]:
    """把趟级收入归到 (路线, 档位)。

    档位在 `_assign_settlement_tiers` 里已经按结算口径归一：
    二干备注优先，一干按容间区间。这里直接按归好的档位汇总。
    """
    income: dict[tuple[str, str, str], float] = defaultdict(float)
    for trip in trips:
        amount = trip_income.get(id(trip))
        if not amount:
            continue
        income[(trip["level"], trip["route"], _tier_alias(trip["tonnage"]))] += amount
    return income


_REMARK_TIER = re.compile(
    r"按\s*(\d+(?:\.\d+)?)\s*([AB]?)\s*(?:吨)?\s*(?=结算|$)"
)
_REMARK_MERGED = re.compile(
    r"合并(?:为)?(?:一)?辆\s*(\d+(?:\.\d+)?)\s*([AB]?)"
)


def _remark_tier(remark: str) -> str | None:
    """从备注提取结算吨位，返回「30」/「40A」这种档位名。

    支持这些写法：
      - 「调度通知按30吨结算」 -> 30
      - 「调度通知按40A结算」 -> 40A
      - 「两辆12吨的车合并为一辆30吨车」 -> 30
      - 「两辆12吨的车合并一辆40A的车」 -> 40A
    没有标记时返回 None。备注里的「按150方结算」是容积不是吨位，不处理。
    """
    text = remark or ""
    match = _REMARK_TIER.search(text) or _REMARK_MERGED.search(text)
    if not match:
        return None
    return f"{match.group(1)}{match.group(2)}"


def _assign_settlement_tiers(trips: list[dict]) -> dict[str, dict]:
    """按「数据给了什么」落档位：有吨位用吨位，没有吨位才用容间。

    不按一干/二干写死。哪张表带了「吨位」列（且有值），就以吨位为准；
    没有吨位（或个别行吨位为空）才回落到容间。返回每张表实际用的口径，
    供问题清单和 JSON 说明，避免用户拿到表却不知道是按什么分的。

    吨位口径下，备注优先：备注写了按多少吨结算、或几辆车合并成多少吨的，
    按备注；没写的，按这辆车在这条路线跑过的最低档（客户确认口径）。
    """
    by_level: dict[str, list[dict]] = defaultdict(list)
    for trip in trips:
        by_level[trip["level"]].append(trip)

    basis: dict[str, dict] = {}
    for level, group in by_level.items():
        has_tonnage = any(trip.get("tonnage_raw") for trip in group)
        derived = 0
        for trip in group:
            raw = trip.get("tonnage_raw") or ""
            volume = trip.get("volume")
            if has_tonnage:
                if raw:
                    tier = raw
                else:
                    tier = _tier_from_volume(volume)
                    derived += 1 if tier else 0
            else:
                tier = _bin_of(volume) if volume is not None else raw
            trip["tonnage"] = _tier_alias(tier) if tier else ""

        confirmed_count = 0
        confirmed_detail: dict[tuple[str, str, str], dict] = {}
        if has_tonnage:
            grouped: dict[tuple[str, str], list[dict]] = defaultdict(list)
            for trip in group:
                if trip["tonnage"]:
                    grouped[(trip["route"], trip["plate"])].append(trip)
            for plate_group in grouped.values():
                lowest = min(plate_group, key=lambda item: _tier_order(item["tonnage"]))["tonnage"]
                for trip in plate_group:
                    remark_tier = _remark_tier(trip["remark"])
                    base_tier = _tier_alias(remark_tier) if remark_tier else lowest
                    confirmed_tier = "" if remark_tier else _confirmed_tier(trip)
                    if confirmed_tier:
                        trip["tonnage"] = _tier_alias(confirmed_tier)
                        if trip["tonnage"] != base_tier:
                            confirmed_count += 1
                            key = (trip["route"], trip["plate"],
                                   trip.get("mmdd") or "", trip["tonnage"])
                            item = confirmed_detail.setdefault(
                                key, {"route": trip["route"], "plate": trip["plate"],
                                      "mmdd": trip.get("mmdd") or "",
                                      "volume": trip.get("volume"),
                                      "tier": trip["tonnage"], "trips": 0})
                            item["trips"] += 1
                    else:
                        trip["tonnage"] = base_tier

        basis[level] = {
            "basis": "吨位" if has_tonnage else "容间",
            "derived": derived,
            "rows": len(group),
            "confirmed": confirmed_count,
            "confirmed_detail": list(confirmed_detail.values()),
        }
    return basis


def _tier_note(level: str, info: dict) -> str:
    """用大白话说明这张表本月按什么分车型。"""
    if info.get("basis") == "吨位":
        text = (f"怎么分车型的：本月{level}表里填了「吨位」这一列，就按吨位分车型"
                "（3吨、12吨、30吨……）。备注里写了「按30吨结算」「两辆车合并成一辆12吨的车」"
                "这类话的，按备注算。")
        if info.get("derived"):
            text += (f"有 {info['derived']} 趟没填吨位，我按容间估了对应车型"
                     "（比如 65 方按 12 吨的车算）。")
        if info.get("confirmed"):
            items = []
            for item in info.get("confirmed_detail") or []:
                volume = item.get("volume")
                volume_text = f"容间{float(volume):g}方" if volume is not None else "没填容间"
                date_text = _mmdd_text(item.get("mmdd"))
                prefix = f"{item['route']} {date_text} {item['plate']}" if date_text \
                    else f"{item['route']} {item['plate']}"
                items.append(f"{prefix}（{volume_text}）按 {item['tier']} 吨的车算")
            text += (f"还有 {info['confirmed']} 趟是您确认过的："
                     + "；".join(items) + "。")
        return text
    return (f"怎么分类的：本月{level}表里没有填「吨位」，所以按容间分："
            "55方以下、55–99方、100–129方、130–149方、150–179方、180方以上"
            "（正好 180 方算最后一档）。")


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
    """档位排序键，用于「车牌本路线最高档」归桶。

    40A 与 40B 同属 40 吨档，靠第二位区分。
    """
    number = _tier_number(str(tonnage))
    base = number if number is not None else 0.0
    return (base, 1 if str(tonnage).endswith("B") else 0)


def _mmdd_text(mmdd) -> str:
    """0909 -> 9月9日。财务表/基础报表里的流水号是 MMDD。"""
    text = str(mmdd or "").strip()
    if len(text) == 4 and text.isdigit():
        return f"{int(text[:2])}月{int(text[2:])}日"
    return text


def _write_issue_sheet(workbook, level, year, month, issues: dict) -> None:
    """第二页「问题清单」：主表一个字不动，需要人工确认的地方集中放这里。

    这一页是给人看的：先按车型/分档汇总（趟次和运费要能和主表对上），
    再用大白话说明数字怎么来的、哪里需要确认。主表要跟官方模板逐格对照，
    任何多余的行列都会造成误判，所以问题一律放第二页。
    """
    sheet = workbook.create_sheet("问题清单")
    sheet.sheet_view.zoomScale = 85

    title = sheet.cell(1, 1, f"{year}年{month}月{level}盈亏表 · 问题清单")
    title.font = Font(name="宋体", bold=True, size=14)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    hint = sheet.cell(2, 1, "这一页用大白话说明数字是怎么来的、哪里需要您看一眼；主表的数字不受影响。")
    hint.font = Font(name="宋体", size=9, color="595959")
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)

    row = 4

    def section(caption: str) -> None:
        nonlocal row
        cell = sheet.cell(row, 1, caption)
        cell.font = Font(name="宋体", bold=True, size=11)
        row += 1

    def table(headers, body) -> None:
        nonlocal row
        for col, text in enumerate(headers, 1):
            cell = sheet.cell(row, col, text)
            cell.font = Font(name="宋体", bold=True, size=9)
            cell.border = BORDER
        row += 1
        if not body:
            cell = sheet.cell(row, 1, "无")
            cell.font = Font(name="宋体", size=9)
            row += 2
            return
        for item in body:
            for col, value in enumerate(item, 1):
                cell = sheet.cell(row, col, value)
                cell.font = Font(name="宋体", size=9)
                cell.border = BORDER
            row += 1
        row += 1

    def note(text: str) -> None:
        nonlocal row
        cell = sheet.cell(row, 1, text)
        cell.font = Font(name="宋体", size=9)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        sheet.row_dimensions[row].height = 30
        row += 1

    # 一、按吨位/容间汇总：让「成本和趟数对上」一眼能核。
    basis = issues.get("tier_basis", "容间")
    section("一、按吨位汇总" if basis == "吨位" else "一、按容间汇总")
    vehicle_rows = issues.get("vehicle_summary") or []
    total_trips, total_freight, total_income = issues.get("level_total") or (0, 0.0, 0.0)
    body = [
        [label if basis == "吨位" else _plain_bin(label),
         trips, round(freight, 2), round(income, 2), round(income - freight, 2)]
        for label, trips, freight, income in vehicle_rows
    ]
    if body:
        body.append(["合计", total_trips, round(total_freight, 2),
                     round(total_income, 2), round(total_income - total_freight, 2)])
    table(
        ["车型（吨位）" if basis == "吨位" else "容间（方）", "趟数",
         "运费（元）", "结算收入（元）", "盈亏（元）"],
        body,
    )
    if body:
        note("「盈亏」正数是赚，负数是亏。")

    # 二、需要您看一眼的地方：改判邮路只写在它归入的那一页。
    section("二、请您看一下这几处")
    moved = [item for item in issues.get("reclassified", []) if item.get("to") == level]
    unmatched = issues.get("unmatched") or []
    orphan = issues.get("orphan_income") or []
    unused = issues.get("unused_income") or []
    if moved:
        names = "、".join(item["route"] for item in moved)
        note(f"这几条线为什么在{level}表里：财务收入表把 {names} 这 {len(moved)} 条线算在{level}。"
             "算一干还是算二干，不影响总运费和总趟数。")
    if unmatched:
        note(f"有 {len(unmatched)} 趟的收入没找到，这几趟没有算进本表，麻烦您核对：")
        table(
            ["邮路", "车牌", "日期", "容间（方）", "运费（元）", "备注"],
            [[trip.get("route", ""), trip.get("plate", ""),
              _mmdd_text(trip.get("mmdd")), trip.get("volume", ""),
              round(float(trip.get("freight") or 0), 2), trip.get("remark", "")]
             for trip in unmatched],
        )
    if orphan:
        note(f"财务收入表里有 {len(orphan)} 笔收入，基础表里没有这台车，"
             "也没有算进本表，麻烦您核对：")
        table(
            ["邮路", "车牌", "日期", "结算收入（元）"],
            [[entry.get("route", ""), entry.get("plate", ""),
              _mmdd_text(entry.get("mmdd")), round(float(entry.get("net") or 0), 2)]
             for entry in orphan],
        )
    if unused:
        note(f"财务收入表里有 {len(unused)} 笔收入，基础表里有这台车，"
             "但日期/趟次没对上，也没有算进本表，麻烦您核对：")
        table(
            ["邮路", "车牌", "日期", "结算收入（元）"],
            [[entry.get("route", ""), entry.get("plate", ""),
              _mmdd_text(entry.get("mmdd")), round(float(entry.get("net") or 0), 2)]
             for entry in unused],
        )
    if not moved and not unmatched and not orphan and not unused:
        note("没有需要特别说明的。")

    # 三、这些数字是怎么来的：只讲人话，不写匹配算法名词。
    section("三、这些数字是怎么来的")
    stats = issues.get("stats") or {}
    if stats:
        note(f"收入怎么对上的：{stats.get('exact', 0)} 趟的车牌和日期都对上了；"
             f"{stats.get('near', 0)} 趟日期差了一两天，就按同一辆车前后 7 天里"
             "最接近的那条收入记录对上。")
    if issues.get("tier_source"):
        note(issues["tier_source"])
    reverse = issues.get("reverse_income") or {}
    if reverse:
        detail = "、".join(f"{route} {amount:,.2f} 元" for route, amount in sorted(reverse.items()))
        note(f"返程收入（银川-西安 这类）不单独占一行，已经加到对应的正向邮路里：{detail}。")
    if total_trips:
        unmatched_count = len(unmatched)
        all_trips = issues.get("all_trips") or (total_trips + unmatched_count)
        if unmatched_count:
            note(f"本表一共 {total_trips} 趟、运费 {total_freight:,.2f} 元；"
                 f"加上上面没找到收入的 {unmatched_count} 趟，"
                 f"正好是基础报表归在{level}的全部 {all_trips} 趟。"
                 "没找到收入的趟次没有算进本表，等核对后再补。")
        else:
            note(f"本表一共 {total_trips} 趟、运费 {total_freight:,.2f} 元。"
                 f"每一行的趟数和运费加起来，跟基础报表里归在{level}的记录完全一样，"
                 "成本和趟数都对得上。")
    if issues.get("month_note"):
        note(issues["month_note"])
    note("这次没给上个月的盈亏表，所以表底「跟上月比」那一块先空着，没有估数。")

    for col, width in zip(range(1, 7), (30, 12, 12, 12, 16, 30)):
        sheet.column_dimensions[get_column_letter(col)].width = width


def _verify_written(path, level, scoped_base, scoped_income) -> list[str]:
    """出表后复核：路线行逐列相加 == 合计行 == 源数据汇总。

    这是交付前的硬门槛。成本和趟次对不上时宁可停下，也不能把表交出去。
    """
    problems: list[str] = []
    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        total_row = next((r for r in range(4, sheet.max_row + 1)
                          if str(sheet.cell(r, 1).value or "").strip() == "合计"), None)
        if total_row is None:
            return [f"{level}：生成的表里找不到合计行"]
        columns: dict[str, int] = {}
        for column in range(1, sheet.max_column + 1):
            label = str(sheet.cell(2, column).value or "").strip()
            if label in SUMMARY_HEADERS:
                columns[label] = column
        missing = [header for header in SUMMARY_HEADERS if header not in columns]
        if missing:
            return [f"{level}：生成的表缺少汇总列 {missing}"]

        def number(row, column):
            try:
                return float(sheet.cell(row, column).value or 0)
            except (TypeError, ValueError):
                return 0.0

        route_sums = {
            header: sum(number(row, columns[header]) for row in range(4, total_row))
            for header in SUMMARY_HEADERS
        }
        table_totals = {
            header: number(total_row, columns[header]) for header in SUMMARY_HEADERS
        }
        base_trips = sum(int(value[0]) for value in scoped_base.values())
        base_freight = sum(float(value[2]) for value in scoped_base.values())
        base_income = sum(float(value) for value in scoped_income.values())
        checks = [
            ("路线行趟次相加", route_sums["趟次"], table_totals["趟次"], 0.5),
            ("路线行运费相加", route_sums["委办运费(元)"], table_totals["委办运费(元)"], 0.02),
            ("路线行收入相加", route_sums["结算收入(元)"], table_totals["结算收入(元)"], 0.02),
            ("合计趟次", table_totals["趟次"], base_trips, 0.5),
            ("合计运费", table_totals["委办运费(元)"], base_freight, 0.02),
            ("合计收入", table_totals["结算收入(元)"], base_income, 0.02),
        ]
        for label, got, want, tolerance in checks:
            if abs(got - want) > tolerance:
                problems.append(f"{level} {label}：表里 {got:,.2f}，应为 {want:,.2f}")
    finally:
        workbook.close()
    return problems


def _write(level, labels, route_order, base, income, out_path, year, month, tonnage_map,
           issues: dict | None = None):
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
        # 汇总列写具体数值而不是公式：openpyxl 写的公式没有缓存值，
        # Excel 会重算，但手机预览、微信预览、在线表格这类只读查看器
        # 直接显示 0——对方看到的表就成了空的。
        row_trips = sum(base.get((route, ton), [0, 0.0, 0.0])[0]
                        for label in labels for ton in tonnage_map.get(label, [label]))
        row_freight = sum(base.get((route, ton), [0, 0.0, 0.0])[2]
                          for label in labels for ton in tonnage_map.get(label, [label]))
        row_income = sum(income.get((route, ton), 0.0)
                         for label in labels for ton in tonnage_map.get(label, [label]))
        sheet.cell(row, summary_start, row_trips or None)
        sheet.cell(row, summary_start + 1, round(row_freight, 2))
        sheet.cell(row, summary_start + 2, round(row_income, 2))
        sheet.cell(row, summary_start + 3, round(row_income - row_freight, 2))
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
    total_trips = sum(totals[label][0] for label in labels)
    total_freight = sum(totals[label][2] for label in labels)
    total_income = sum(totals[label][3] for label in labels)
    for offset, value in enumerate([
        total_trips or None,
        round(total_freight, 2),
        round(total_income, 2),
        round(total_income - total_freight, 2),
    ]):
        sheet.cell(row, summary_start + offset, value)
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

    if issues:
        _write_issue_sheet(workbook, level, year, month, issues)

    workbook.save(out_path)
    workbook.close()
    return totals


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="生成一干/二干盈亏表")
    parser.add_argument("--base", type=Path, default=None, help="基础报表原版（含一干汇总/二干基础）")
    parser.add_argument("--finance", type=Path, default=None, help="财务收入表")
    parser.add_argument("--month", required=True, help="YYYY-MM")
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--tpl-dir", type=Path, default=None,
                        help="可选：盈亏模板目录。只用于增加本期无数据的历史路线占位；"
                             "吨位档和路线顺序一律以当月源数据为准")
    parser.add_argument("--no-template-routes", action="store_true",
                        help="不追加模板中当月无数据的历史路线，只输出当月实际出现的路线")
    args = parser.parse_args(argv)

    year, month = map(int, args.month.split("-"))

    # 输入校验：缺文件属于「请用户补充」，不是执行失败，返回 needs_input 且退出码 0。
    missing_inputs = []
    for role, label, path in (
        ("base_report", "基础报表原版", args.base),
        ("finance", "财务收入表", args.finance),
    ):
        if path is None or not path.is_file():
            item = {
                "kind": "missing",
                "role": role,
                "label": label,
                "required": 1,
                "found": 1 if path is not None and path.is_file() else 0,
                "action": f"请上传{label}" + (f"（没有找到 {path}）" if path is not None else ""),
            }
            if path is not None:
                parent = path.parent
                if parent.is_dir():
                    names = sorted(p.name for p in parent.iterdir() if p.is_file())
                    item["available_files"] = names[:30]
            missing_inputs.append(item)
    if missing_inputs:
        print(json.dumps({
            "ok": True,
            "status": "needs_input",
            "phase": "awaiting_inputs",
            "month": args.month,
            "missing_inputs": missing_inputs,
            "message": "当前上传的文件不完整，需要补充："
                       + "、".join(item["label"] for item in missing_inputs) + "。",
            "next_action": "请补充基础报表原版和/或财务收入表后重试。",
        }, ensure_ascii=False, indent=2))
        return 0

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

    base, trips, diag = _read_summary_sheets(args.base)
    # 数据月份比目标月新 -> 按数据出表；比目标月旧 -> 下面 _base_guard 会拦下来。
    month_note = ""
    adopted = _adopt_base_month(diag, year, month)
    if adopted:
        month_note = (f"文件/参数写的是 {args.month}，但基础报表原版的业务日期是 {adopted}，"
                      "已按数据出表（文件名疑似写错）")
        print(f"[WARN] {month_note}")
        year, month = map(int, adopted.split("-"))
        args.month = adopted
    guard = _base_guard(diag, year, month)
    if guard:
        print(f"[STOP] 基础报表原版不可用于 {args.month}：{guard.get('message', '')}")
        print(json.dumps({
            "ok": False,
            "phase": "base_report_check",
            "month": args.month,
            "base_sheets": diag,
            **guard,
        }, ensure_ascii=False, indent=2))
        return 2
    exact, by_plate, reverse_income = _read_finance(args.finance)
    # 一/二干归属以财务收入表为准（基础报表 sheet 名只是兜底）。
    reclassified = _reclassify_by_finance(trips, _finance_route_levels(exact))
    _income_flat, stats, misses = allocate_income(trips, exact, by_plate, year)
    # 财务表和基础报表不是同一个月时，逐趟几乎配不上：收入会大面积记 0，
    # 出的表看起来有数、其实是漏的。超过一半配不上就停下来问。
    if trips and len(misses) > len(trips) * 0.5:
        print(f"[STOP] 财务收入表和基础报表原版对不上：{len(misses)}/{len(trips)} 条派车记录配不上收入。")
        print(json.dumps({
            "ok": False,
            "phase": "finance_check",
            "month": args.month,
            "detail_rows": len(trips),
            "income_match": stats,
            "status": "finance_mismatch",
            "message": f"财务收入表 {len(misses)}/{len(trips)} 条派车记录配不上（车牌+日期），"
                       "两份文件很可能不是同一个月的",
            "next_action": "请确认财务收入表（运输局）和基础报表原版是同月的；"
                           "财务表通常是次月结算，但业务日期要覆盖当月。",
        }, ensure_ascii=False, indent=2))
        return 2
    trip_income, unused_income_raw = _trip_level_income(trips, exact, by_plate, year)
    # 收入没对上的趟次不进结果表：单独放问题清单，避免用 0 收入把盈亏带偏。
    matched_trips = [trip for trip in trips if id(trip) in trip_income]
    unmatched_trips = [trip for trip in trips if id(trip) not in trip_income]
    misses = unmatched_trips
    stats["miss"] = len(unmatched_trips)
    # 归属定下来、且剔除未匹配趟次后，再落档位和汇总。
    tier_basis = _assign_settlement_tiers(matched_trips)
    base = _aggregate_base(matched_trips)
    income = _bucket_income(matched_trips, trip_income)
    income = _apply_reverse_income(income, matched_trips, reverse_income)
    # 财务收入表里有、基础报表里找不到对应车辆/趟次的收入：不估算补足，
    # 单独放问题清单，等人核对。
    plates_in_base = {trip["plate"] for trip in trips}
    orphan_income: list[dict] = []
    for plate, entries in by_plate.items():
        if plate in plates_in_base:
            continue
        for entry in entries:
            if entry.get("route", "").endswith("-西安"):
                continue
            orphan_income.append({
                "level": "一干" if "一干" in str(entry.get("level", "")) else "二干",
                "route": normalize_route(entry["route"]),
                "plate": plate,
                "mmdd": entry.get("mmdd", ""),
                "net": float(entry.get("net") or 0),
            })
    for entry in orphan_income:
        entry["reason"] = "基础表里没有这台车"
    # 车牌在基础表里、但财务收入表的日期/趟次没对上：也不能丢，列问题清单。
    unused_income = [{
        "level": "一干" if "一干" in str(entry.get("level", "")) else "二干",
        "route": normalize_route(entry["route"]),
        "plate": entry.get("plate", ""),
        "mmdd": entry.get("mmdd", ""),
        "net": float(entry.get("net") or 0),
        "reason": "基础表里有这台车，但没对到具体趟次",
    } for entry in unused_income_raw]

    # 收入也要对平：主表收入 + 没对上的收入 = 财务收入表全部委办收入。
    finance_total = (sum(entry["net"] for entries in exact.values() for entry in entries)
                     + sum(reverse_income.values()))
    accounted = (sum(income.values())
                 + sum(entry["net"] for entry in unused_income)
                 + sum(entry["net"] for entry in orphan_income))
    if abs(accounted - finance_total) > 0.01:
        print("[STOP] 收入合计和财务收入表对不上，不能交付。")
        print(json.dumps({
            "ok": False,
            "status": "reconciliation_failed",
            "month": args.month,
            "finance_total": round(finance_total, 2),
            "accounted": round(accounted, 2),
            "unused_income": unused_income[:10],
            "orphan_income": orphan_income[:10],
            "next_action": "先核对财务收入表和未对上的收入明细，修好后重跑；不要直接交付。",
        }, ensure_ascii=False, indent=2))
        return 3

    tier_notes = {
        level: _tier_note(level, tier_basis.get(level, {"basis": "容间", "derived": 0, "rows": 0}))
        for level in ("一干", "二干")
    }
    ordered_basis = [(level, tier_basis[level]) for level in ("一干", "二干") if level in tier_basis]
    deliverables: list[dict] = []
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
        labels, tier_values = _merge_tiers(matched_trips, level, extra_tiers)
        tonnage_map = {label: tier_values[label] for label in labels}
        merged = {
            label: values for label, values in tonnage_map.items()
            if len(values) > 1 or values != [label]
        }

        # 路线：当月出现顺序优先，模板路线只用来补当月无数据的占位。
        route_order = _ordered_routes(matched_trips, level,
                                      [] if args.no_template_routes else template_routes)
        scoped_base = {(route, ton): value for (lvl, route, ton), value in base.items() if lvl == level}
        scoped_income = {(route, ton): value for (lvl, route, ton), value in income.items() if lvl == level}

        # 车型（二干）或容间档（一干）汇总：让「成本 + 趟次」在主表旁边直接能核。
        vehicle_rows = []
        for label in labels:
            keys = set(tonnage_map.get(label, [label]))
            group = [value for (route, ton), value in scoped_base.items() if ton in keys]
            group_income = sum(value for (route, ton), value in scoped_income.items() if ton in keys)
            vehicle_rows.append((
                label,
                sum(value[0] for value in group),
                sum(value[2] for value in group),
                group_income,
            ))
        level_total = (
            sum(row[1] for row in vehicle_rows),
            sum(row[2] for row in vehicle_rows),
            sum(row[3] for row in vehicle_rows),
        )

        out_path = args.out / f"{month}月_04_{level}盈亏表.xlsx"
        level_unmatched = [trip for trip in unmatched_trips if trip.get("level") == level]
        level_orphans = [entry for entry in orphan_income if entry.get("level") == level]
        level_unused = [entry for entry in unused_income if entry.get("level") == level]
        issues = {
            "unmatched": level_unmatched,
            "unmatched_freight": sum(float(trip.get("freight") or 0)
                                     for trip in level_unmatched),
            "all_trips": sum(1 for trip in trips if trip.get("level") == level),
            "orphan_income": level_orphans,
            "unused_income": level_unused,
            "reclassified": reclassified,
            "stats": stats,
            "tier_source": tier_notes[level],
            "tier_basis": (tier_basis.get(level) or {}).get("basis", "容间"),
            "vehicle_summary": vehicle_rows,
            "level_total": level_total,
            "reverse_income": reverse_income,
            "month_note": month_note,
        }
        totals = _write(level, labels, route_order, scoped_base, scoped_income, out_path,
                        year, month, tonnage_map, issues=issues)
        problems = _verify_written(out_path, level, scoped_base, scoped_income)
        if problems:
            print("[STOP] 出表后核对不平，不能交付：")
            for item in problems:
                print(f" - {item}")
            print(json.dumps({
                "ok": False,
                "status": "reconciliation_failed",
                "month": args.month,
                "level": level,
                "problems": problems,
                "next_action": "先核对基础报表和生成表，修好后重跑；不要直接交付。",
            }, ensure_ascii=False, indent=2))
            return 3
        deliverables.append({
            "level": level,
            "file": out_path.name,
            "total_trips": sum(bucket[0] for bucket in totals.values()),
            "total_freight": round(sum(bucket[2] for bucket in totals.values()), 2),
            "total_income": round(sum(bucket[3] for bucket in totals.values()), 2),
            "routes": len(route_order),
            "volume_groups": labels,
            "bin_merge": merged,
            "template": template_status,
            "tier_source": tier_notes[level],
        })

    published = [
        {"name": item["name"], "size": item["size"]}
        for item in publish_files(Path(args.out) / f"{month}月_04_{level}盈亏表.xlsx"
                                  for level in ("一干", "二干"))
    ]
    matched_freight = sum(trip["freight"] for trip in matched_trips)
    excluded_freight = sum(trip["freight"] for trip in unmatched_trips)
    base_freight = sum(trip["freight"] for trip in trips)
    if abs(matched_freight + excluded_freight - base_freight) > 0.01:
        print("[STOP] 剔除未匹配趟次后，运费合计和基础报表对不上，不能交付。")
        print(json.dumps({
            "ok": False,
            "status": "reconciliation_failed",
            "month": args.month,
            "base_freight": round(base_freight, 2),
            "matched_freight": round(matched_freight, 2),
            "excluded_freight": round(excluded_freight, 2),
            "next_action": "先核对基础报表和未匹配明细，修好后重跑；不要直接交付。",
        }, ensure_ascii=False, indent=2))
        return 3
    # 字段顺序有讲究：平台会截断过长的工具输出，模型只看得见前面一段。
    # 决定「要不要重跑」的 ok/status/published 排在前面，明细行次之，
    # 重分类清单和未匹配样例放最后，被截断也不影响出表判断。
    report = {
        "ok": True,
        "status": "profit_completed",
        "month": args.month,
        "month_note": month_note,
        "detail_rows": len(trips),
        "matched_rows": len(matched_trips),
        "excluded_rows": len(unmatched_trips),
        "income_match": stats,
        "unmatched": len(unmatched_trips),
        "orphan_income": len(orphan_income),
        "unused_income": len(unused_income),
        "reconciliation": {
            "base_rows": len(trips),
            "matched_rows": len(matched_trips),
            "excluded_rows": len(unmatched_trips),
            "base_freight": round(base_freight, 2),
            "matched_freight": round(matched_freight, 2),
            "excluded_freight": round(excluded_freight, 2),
            "freight_ok": True,
            "finance_total": round(finance_total, 2),
            "income_accounted": round(accounted, 2),
            "income_ok": True,
        },
        "published": published,
        "outputs": deliverables,
        "next_action": "已出表并发布，直接汇报，不要再执行本脚本或重新 detect。",
        "tier_source": "；".join(f"{level}按{info['basis']}"
                                for level, info in ordered_basis) or "按容间分档",
        "tier_basis": dict(ordered_basis),
        "level_source": "一/二干以财务收入表的邮路级别为准，基础报表 sheet 名兜底",
        "level_reclassified": [
            {"route": item["route"], "from": item["from"], "to": item["to"],
             "trips": item["trips"], "freight": round(item["freight"], 2)}
            for item in reclassified
        ],
        "reverse_income": {k: round(v, 2) for k, v in reverse_income.items()},
        "unmatched_samples": unmatched_trips[:5],
        "orphan_income_samples": orphan_income[:5],
        "unused_income_samples": unused_income[:5],
    }
    names = "、".join(item["file"] for item in deliverables)
    scale = "，".join(f"{item['level']} {item['total_trips']} 趟" for item in deliverables)
    print(f"[OK] 已出表并发布：{names}（{scale}）。直接汇报，不要重跑本脚本或 detect。")
    if month_note:
        print(f"[NOTE] {month_note}")
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
