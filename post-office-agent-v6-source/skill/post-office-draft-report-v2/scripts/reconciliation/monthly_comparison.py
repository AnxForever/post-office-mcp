from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from .normalize import normalize_text


def _find_total_row(ws, start_row: int) -> int | None:
    for row in range(start_row, ws.max_row + 1):
        if normalize_text(ws.cell(row, 1).value) == "合计":
            return row
    return None


def _num(ws, row, col):
    try:
        return float(ws.cell(row, col).value or 0)
    except Exception:
        return 0.0


def read_previous_month_totals(
    path: Path | str | None,
    kind: str,
    layout: dict[str, Any] | None = None,
) -> tuple[dict[str, float] | None, str | None]:
    """读取上月盈亏表合计行的吨位组汇总数据，不依赖固定列位置。"""
    if not path:
        return None, "上月盈亏表（暂无数据）"
    p = Path(path)
    if not p.is_file():
        return None, f"上月盈亏表不存在: {p.name}"
    try:
        wb = load_workbook(p, read_only=True, data_only=True, keep_links=False)
    except Exception as exc:
        return None, f"上月盈亏表无法读取: {exc}"
    try:
        if layout is None:
            return None, "上月盈亏表缺少动态布局"

        sheet_name = normalize_text(layout.get("sheet"))
        if not sheet_name or sheet_name not in wb.sheetnames:
            return None, f"上月盈亏表缺少{sheet_name or '有效月份'}Sheet"
        ws = wb[sheet_name]

        total_row = _find_total_row(ws, int(layout.get("total_row", 10)))
        if total_row is None:
            return None, "找不到合计行"

        # 逐组扫描：col 2 起每 7 列一组，趟次/运费/收入全0则跳过
        groups = []  # (start, trips, freight, income, load)
        col = 2
        while col + 6 <= ws.max_column:
            t = _num(ws, total_row, col)
            f = _num(ws, total_row, col + 3)
            i_ = _num(ws, total_row, col + 4)
            l = _num(ws, total_row, col + 1)
            if t == 0 and f == 0 and i_ == 0:
                col += 7
                continue
            groups.append((col, t, f, i_, l))
            col += 7

        if not groups:
            return None, "上月合计行无数据"

        # 最后一组如果是汇总（值远大于前端组），排除
        if len(groups) >= 2:
            prev_sum = sum(g[1] for g in groups[:-1])
            last_trips = groups[-1][1]
            if prev_sum > 0 and last_trips >= prev_sum * 0.8:
                groups = groups[:-1]

        trips = sum(g[1] for g in groups)
        freight = sum(g[2] for g in groups)
        income = sum(g[3] for g in groups)
        load = sum(g[4] for g in groups)

        return {
            "trips": trips,
            "load": load,
            "freight": freight,
            "income": income,
            "pnl": income - freight,
        }, None
    except Exception as exc:
        return None, f"上月盈亏表解析失败: {exc}"
    finally:
        wb.close()
