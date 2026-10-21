from __future__ import annotations

import json
import os
from copy import copy
from pathlib import Path
from typing import Any, Iterable, Mapping

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.views import Selection

from .normalize import destination_name, normalize_text, route_candidates, tonnage_key
from .template_discovery import discover_draft_template, discover_profit_template


REVIEW_HEADERS = [
    "序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", None, "车牌号", None, "容间",
    "吨位", "公里", "运输量", "路单号", "运价", "标准价", "备注", "是否跨档", "变动前容间", "变动前时间", "系数",
]

ISSUE_HEADERS = [
    "问题编号", "所在Sheet", "行号", "车牌号", "邮路名称",
    "路线类型", "当前吨位", "问题详情", "建议动作", "处理结果",
]

COLOR_PRIORITY = {None: 0, "yellow": 1, "orange": 2, "blue": 2.5, "deep_orange": 2.75, "red": 3}
SEVERITY_FILL = {
    "yellow": PatternFill("solid", fgColor="FFFFEB9C"),
    "orange": PatternFill("solid", fgColor="FFF4B183"),
    "red": PatternFill("solid", fgColor="FFFF0000"),
    "blue": PatternFill("solid", fgColor="FFBDD7EE"),
    "deep_orange": PatternFill("solid", fgColor="FFC65911"),
}

STRUCTURE_SOURCE_LABELS = {
    "model": "车型表",
    "online_caotan": "草滩在线测量表",
    "online_gangwu": "港务在线测量表",
    "vehicle": "车管基础数据",
    "finance": "财务收入表",
    "first": "一干",
    "second": "二干",
}


def _structure_source_label(role: Any, explicit_label: Any = None) -> str:
    label = str(explicit_label or "").strip()
    if label:
        return label
    return STRUCTURE_SOURCE_LABELS.get(str(role or "").strip(), "其他来源")


def _route_type_label(price_kind: Any) -> str | None:
    return {"first": "一干", "second": "二干"}.get(str(price_kind or "").strip().lower())


def strongest_color(*colors: str | None) -> str | None:
    return max(colors, key=lambda color: COLOR_PRIORITY.get(color, 0), default=None)


def _atomic_save(wb: Workbook, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix(output.suffix + ".tmp")
    wb.save(temporary)
    os.replace(temporary, output)


def _review_values(row: dict[str, Any], index: int) -> list[Any]:
    return [
        index,
        row.get("company"),
        row.get("dispatch_no"),
        row.get("route"),
        None,
        row.get("seal_time"),
        None,
        row.get("plate"),
        row.get("trailer"),
        row.get("volume"),
        row.get("tonnage"),
        row.get("distance"),
        row.get("transport_volume"),
        None,
        row.get("freight"),
        row.get("standard_price"),
        row.get("remark"),
        row.get("cross_tonnage", "否"),
        row.get("previous_volume"),
        row.get("previous_time"),
        row.get("coefficient"),
    ]


def _write_structure_discovery(wb: Workbook, structure: dict[str, Any]) -> None:
    ws = wb.create_sheet("结构发现")
    headers = ["分类", "来源/对象", "文件/路线", "工作表", "表头行/吨位", "字段定位/容间下限", "系数/备注"]
    ws.append(headers)
    raw_sources = structure.get("sources", [])
    sources = (
        [source for role_sources in raw_sources.values() for source in role_sources]
        if isinstance(raw_sources, Mapping)
        else raw_sources
    )
    for source in sources:
        sheets = source.get("target_sheets") or source.get("sheets") or [None]
        field_map = source.get("field_map") or {}
        detail = json.dumps(field_map, ensure_ascii=False, sort_keys=True) if field_map else ""
        for sheet in sheets:
            ws.append([
                "源文件结构",
                _structure_source_label(source.get("role"), source.get("role_label")),
                source.get("name") or Path(source.get("path", "")).name,
                sheet,
                source.get("header_row"),
                detail,
                None,
            ])
    catalogs = structure.get("catalogs", {})
    for kind, label in (("first", "一干吨位档"), ("second", "二干吨位档")):
        for tier in catalogs.get(kind, {}).get("tiers", []):
            ws.append([
                label,
                _structure_source_label(kind),
                None,
                None,
                tonnage_key(tier.get("key")),
                tier.get("min_volume"),
                tier.get("coefficient"),
            ])
    destinations = sorted({destination_name(route) for route in catalogs.get("second", {}).get("routes", {}) if destination_name(route)})
    for destination in destinations:
        ws.append(["二干目的地", _structure_source_label("second"), destination, None, None, None, None])

    header_fill = PatternFill("solid", fgColor="FF1F4E78")
    header_font = Font(color="FFFFFFFF", bold=True)
    thin = Side(style="thin", color="FF808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = copy(header_fill)
        cell.font = copy(header_font)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = copy(border)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = copy(border)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for index, width in enumerate([16, 16, 34, 36, 16, 48, 18], 1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False


def write_review_workbook(
    rows: Iterable[dict[str, Any]],
    template_path: Path | str,
    output: Path | str,
    month: tuple[int, int],
    *,
    structure: dict[str, Any] | None = None,
) -> Path:
    rows, template_path, output = list(rows), Path(template_path), Path(output)
    layout = discover_draft_template(template_path)
    wb = load_workbook(template_path, data_only=False, keep_links=False)
    base_ws = wb[layout["sheet"]]
    if base_ws.title != "基础报表":
        base_ws.title = "基础报表"
    for title in ("结构发现", "问题汇总"):
        if title in wb.sheetnames:
            wb.remove(wb[title])
    _write_structure_discovery(wb, structure or {})
    structure_ws = wb["结构发现"]
    wb.move_sheet(structure_ws, offset=-wb.index(structure_ws))
    issues_ws = wb.create_sheet("问题汇总", 1)
    issues_ws.append(ISSUE_HEADERS)
    header_fill = PatternFill("solid", fgColor="FF1F4E78")
    header_font = Font(color="FFFFFFFF", bold=True)
    thin = Side(style="thin", color="FF808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in issues_ws[1]:
        cell.fill = copy(header_fill)
        cell.font = copy(header_font)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = copy(border)
    issues_ws.freeze_panes = "A2"
    issues_ws.auto_filter.ref = issues_ws.dimensions

    data_start_row = int(layout["data_start_row"])
    trailer_column = int(layout["headers"]["车牌号"]) + 1
    layout_columns = {int(column) for column in layout["headers"].values()} | {trailer_column}
    layout_min_column = min(layout_columns)
    layout_max_column = min(max(layout_columns), 21)
    style_row = data_start_row if base_ws.max_row >= data_start_row else None
    data_styles = {
        column: copy(base_ws.cell(style_row, column)._style) if style_row else None
        for column in range(layout_min_column, layout_max_column + 1)
    }
    data_height = base_ws.row_dimensions[style_row].height if style_row else None
    if base_ws.max_row >= data_start_row:
        base_ws.delete_rows(data_start_row, base_ws.max_row - data_start_row + 1)

    issue_number = 1
    for index, row in enumerate(rows, 1):
        base_row_number = data_start_row + index - 1
        values = _review_values(row, index)
        target_values = {
            int(layout["headers"][header]): values[column - 1]
            for column, header in enumerate(REVIEW_HEADERS, 1)
            if header is not None
        }
        if trailer_column in set(layout["headers"].values()):
            raise ValueError("初版模板车牌号后缺少挂车号伴随空列")
        target_values[trailer_column] = values[8]
        for column, value in target_values.items():
            cell = base_ws.cell(base_row_number, column)
            if data_styles.get(column) is not None:
                cell._style = copy(data_styles[column])
            if value is not None:
                cell.value = value
        if data_height is not None:
            base_ws.row_dimensions[base_row_number].height = data_height
        issues = list(row.get("issues") or [])
        severity = strongest_color(*(issue.severity for issue in issues))
        for cell in base_ws.iter_cols(
            min_col=layout_min_column,
            max_col=layout_max_column,
            min_row=base_row_number,
            max_row=base_row_number,
        ):
            cell = cell[0]
            if severity:
                cell.fill = copy(SEVERITY_FILL[severity])
        for issue in issues:
            issues_ws.append([
                issue_number,
                "基础报表",
                base_row_number,
                row.get("plate"),
                row.get("route"),
                _route_type_label(row.get("price_kind")),
                row.get("tonnage"),
                issue.message,
                issue.action,
                None,
            ])
            issue_row_number = issues_ws.max_row
            for cell in issues_ws[issue_row_number]:
                cell.border = copy(border)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            issue_number += 1

    for index, width in enumerate([10, 14, 10, 14, 30, 12, 12, 32, 32, 24], 1):
        issues_ws.column_dimensions[issues_ws.cell(1, index).column_letter].width = width
    issues_ws.sheet_view.showGridLines = False
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    _atomic_save(wb, output)
    wb.close()
    check = load_workbook(output, read_only=True, data_only=False)
    try:
        valid_structure = check.sheetnames[:2] == ["结构发现", "问题汇总"] and "基础报表" in check.sheetnames
        if valid_structure:
            header_values = next(check["基础报表"].iter_rows(
                min_row=int(layout["header_row"]),
                max_row=int(layout["header_row"]),
                values_only=True,
            ), ())
            valid_structure = all(
                column <= len(header_values)
                and normalize_text(header_values[column - 1]) == normalize_text(header)
                for header, column in layout["headers"].items()
            )
        if not valid_structure:
            raise ValueError("初版工作簿结构校验失败")
    finally:
        check.close()
    return output


def _copy_cell_style(source, target) -> None:
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.number_format = source.number_format
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)


def _route_sheet(route: Any, split_layout: dict[str, Any]) -> str | None:
    targets = split_layout.get("route_targets", {})
    writable = set(split_layout.get("writable_sheets", ()))
    candidates = route_candidates(route)
    for candidate in candidates:
        target = targets.get(candidate)
        if target in writable:
            return target
    for candidate in candidates:
        target = targets.get(destination_name(candidate))
        if target in writable:
            return target
    return None


def _round_trip_sheet(route: Any, split_layout: dict[str, Any]) -> str | None:
    """往返行优先落到「<目的地>往返」页签（如 银川往返）。"""
    writable = set(split_layout.get("writable_sheets", ()))
    for candidate in route_candidates(route):
        target = f"{destination_name(candidate)}往返"
        if target in writable:
            return target
    return None


def _split_row_values(row: dict[str, Any], index: int, headers: list[Any]) -> list[Any]:
    direct = [
        index, row.get("company"), row.get("dispatch_no"), row.get("route"), None, row.get("seal_time"), None,
        row.get("plate"), row.get("trailer"), int(round(float(row["volume"]))) if row.get("volume") not in (None, "") else None,
        row.get("tonnage"), row.get("distance"), row.get("transport_volume"), None, row.get("freight"), row.get("standard_price"), row.get("remark"),
    ]
    result: list[Any] = []
    for column, header in enumerate(headers, 1):
        text = str(header or "")
        if column <= 16:
            result.append(direct[column - 1])
        elif "备注" in text:
            result.append(row.get("remark"))
        elif "运价" in text:
            result.append(row.get("freight"))
        elif column <= len(direct):
            result.append(direct[column - 1])
        else:
            result.append(None)
    return result


def _reset_split_sheet_display(ws: Any, last_row: int) -> None:
    """Remove template view/filter state while keeping the generated header frozen."""
    for index in list(ws.row_dimensions):
        if index > last_row:
            del ws.row_dimensions[index]
        else:
            ws.row_dimensions[index].hidden = False
    ws.sheet_properties.filterMode = False
    ws.auto_filter.filterColumn = []
    ws.auto_filter.sortState = None
    ws.freeze_panes = "A2"
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.zoomScale = 80
    ws.sheet_view.selection = [Selection(pane="bottomLeft", activeCell="A1", sqref="A1")]


DOWNGRADE_HEADERS = [
    "序号", "车辆承运单位", "派车单号", "邮路名称", "封车时间", "车牌号", "容间",
    "车型吨位", "车管结算吨位", "结算档位", "公里", "运输量", "运价", "标准价",
    "原运费", "差价", "备注",
]

SPECIAL_HEADERS = {
    "二干基础": [
        "序号", "车辆承运单位", "派车单号", "邮路名称", "邮路名称", "路单编号", "封车时间", None,
        "车牌号", None, "容间", "吨位", "公里", "运输量", "路单号", "运价", "标准价", "备注", None,
    ],
    "一干汇总": [
        "序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", None, "车牌号", None,
        "容间", "吨位", "公里", "运输量", "路单号", "运价（保留两位数）", "标准价", "运价", "备注",
    ],
    "两趟合一趟": [
        "序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", None, "车牌号", None,
        "容间", "吨位", "公里", "运输量", "路单号", "运价", "标准价", "备注", None,
    ],
}


def _special_row_values(row: dict[str, Any], index: int, headers: list[Any], *, two_trip: bool = False) -> list[Any]:
    """Map a corrected review row into an official special-sheet layout."""
    values: list[Any] = []
    seen_route = 0
    for column, header in enumerate(headers, 1):
        text = str(header or "")
        if text == "序号":
            value = index
        elif text == "车辆承运单位":
            value = row.get("company")
        elif text == "派车单号":
            value = row.get("dispatch_no")
        elif text == "邮路名称":
            seen_route += 1
            value = row.get("route") if seen_route == 1 else (row.get("standard_route") or row.get("route"))
        elif text == "封车时间":
            value = row.get("seal_time")
        elif text == "车牌号":
            value = row.get("plate")
        elif text == "容间":
            try:
                value = int(round(float(row["volume"]))) if row.get("volume") not in (None, "") else None
            except (TypeError, ValueError):
                value = None
        elif text == "吨位":
            value = row.get("tonnage")
        elif text == "公里":
            value = row.get("distance")
        elif text == "运输量":
            value = row.get("transport_volume")
        elif text == "运价（保留两位数）":
            value = round(float(row["freight"]), 2) if row.get("freight") not in (None, "") else None
        elif text == "运价":
            value = row.get("freight")
        elif text == "标准价":
            value = row.get("standard_price")
        elif text == "备注":
            value = row.get("remark")
        elif two_trip and column == len(headers):
            # The official two-trip tab uses the trailing blank column for the
            # per-vehicle unit price used by the manual merge calculation.
            value = row.get("standard_price")
        else:
            value = None
        values.append(value)
    return values


def _write_special_sheet(
    workbook: Workbook,
    name: str,
    rows: list[dict[str, Any]],
    *,
    style_source: Any,
    two_trip: bool = False,
) -> int:
    if name not in workbook.sheetnames:
        return 0
    ws = workbook[name]
    headers = SPECIAL_HEADERS[name]
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for column, header in enumerate(headers, 1):
        target = ws.cell(1, column, header)
        if style_source is not None and column <= style_source.max_column:
            _copy_cell_style(style_source.cell(1, column), target)
    row_style = []
    for column in range(1, len(headers) + 1):
        if style_source is not None and style_source.max_row >= 2 and column <= style_source.max_column:
            row_style.append(copy(style_source.cell(2, column)._style))
        else:
            row_style.append(None)
    for index, row in enumerate(rows, 1):
        values = _special_row_values(row, index, headers, two_trip=two_trip)
        for column, value in enumerate(values, 1):
            cell = ws.cell(index + 1, column, value)
            if row_style[column - 1] is not None:
                cell._style = copy(row_style[column - 1])
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
    else:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    return len(rows)


def _write_downgrade_sheet(workbook: Workbook, rows: list[dict[str, Any]]) -> int:
    downgrade_rows = [row for row in rows if row.get("downgraded")]
    if not downgrade_rows:
        return 0
    ws = workbook["降吨位"] if "降吨位" in workbook.sheetnames else workbook.create_sheet("降吨位")
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)

    header_fill = PatternFill("solid", fgColor="FF4472C4")
    header_font = Font(color="FFFFFFFF", bold=True, size=10)
    deep_orange = PatternFill("solid", fgColor="FFC65911")
    thin = Side(style="thin", color="FF808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(DOWNGRADE_HEADERS)
    for cell in ws[1]:
        cell.fill = copy(header_fill)
        cell.font = copy(header_font)
        cell.border = copy(border)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for index, row in enumerate(downgrade_rows, 1):
        ws.append([
            index,
            row.get("company"),
            row.get("dispatch_no"),
            row.get("route"),
            row.get("seal_time"),
            row.get("plate"),
            row.get("volume"),
            row.get("model_tonnage"),
            row.get("settlement_tonnage"),
            row.get("tonnage"),
            row.get("distance"),
            row.get("transport_volume"),
            row.get("freight"),
            row.get("standard_price"),
            row.get("original_freight"),
            row.get("price_difference"),
            row.get("remark"),
        ])
        for cell in ws[ws.max_row]:
            cell.fill = copy(deep_orange)
            cell.border = copy(border)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for column, width in enumerate([7, 24, 16, 30, 14, 14, 11, 11, 14, 11, 9, 11, 12, 12, 12, 12, 42], 1):
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(DOWNGRADE_HEADERS))}{ws.max_row}"
    ws.sheet_view.showGridLines = False
    return len(downgrade_rows)


def write_split_workbook(
    rows: Iterable[dict[str, Any]],
    template_path: Path | str,
    split_layout: dict[str, Any],
    output: Path | str,
) -> tuple[Path, dict[str, int]]:
    rows, output = list(rows), Path(output)
    wb = load_workbook(template_path, data_only=False, keep_links=False)
    original_sheets = list(wb.sheetnames)
    if split_layout.get("sheets") != original_sheets:
        wb.close()
        raise ValueError("拆分母版布局与实际Sheet顺序不一致")
    writable_sheets = [name for name in split_layout.get("writable_sheets", ()) if name in wb.sheetnames]
    if not writable_sheets:
        wb.close()
        raise ValueError("拆分母版布局没有可写入线路页签")

    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        sheet = None
        if "往返邮路" in str(row.get("remark") or ""):
            for route in (row.get("standard_route"), row.get("route")):
                if route and (sheet := _round_trip_sheet(route, split_layout)):
                    break
        if not sheet:
            sheet = next(
                (
                    target
                    for route in (row.get("standard_route"), row.get("route"))
                    if route and (target := _route_sheet(route, split_layout))
                ),
                "※未匹配目的地",
            )
        grouped.setdefault(sheet, []).append(row)
    if not grouped:
        wb.close()
        raise ValueError("修订初版没有可拆分的有效记录")

    if "※未匹配目的地" in grouped and "※未匹配目的地" not in wb.sheetnames:
        unmatched = wb.copy_worksheet(wb[writable_sheets[0]])
        unmatched.title = "※未匹配目的地"
        writable_sheets.append(unmatched.title)

    for name in writable_sheets:
        ws = wb[name]
        sheet_rows = grouped.get(ws.title, [])
        headers = [ws.cell(1, column).value for column in range(1, ws.max_column + 1)]
        max_column = max(17, max((index for index, value in enumerate(headers, 1) if value is not None), default=17))
        headers = headers[:max_column]
        if len(headers) < max_column:
            headers.extend([None] * (max_column - len(headers)))
        style_row = 2 if ws.max_row >= 2 else 1
        styles = [copy(ws.cell(style_row, column)._style) for column in range(1, max_column + 1)]
        style_height = ws.row_dimensions[style_row].height
        for merged in list(ws.merged_cells.ranges):
            if merged.max_row > 1:
                ws.unmerge_cells(str(merged))
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for column in range(1, max_column + 1):
            ws.cell(2, column)._style = copy(styles[column - 1])
        ws.row_dimensions[2].height = style_height
        if not sheet_rows:
            ws.cell(2, 1, "")
            ws.auto_filter.ref = f"A1:{get_column_letter(max_column)}1"
            _reset_split_sheet_display(ws, 2)
            continue
        for index, row in enumerate(sheet_rows, 1):
            target_row = index + 1
            values = _split_row_values(row, index, headers)
            for column, value in enumerate(values, 1):
                cell = ws.cell(target_row, column, value)
                cell._style = copy(styles[column - 1])
            ws.row_dimensions[target_row].height = style_height
        total_row = len(sheet_rows) + 2
        ws.cell(total_row, 1, "合计")
        for column in range(1, max_column + 1):
            ws.cell(total_row, column)._style = copy(styles[column - 1])
            header = str(headers[column - 1] or "")
            if "运输量" in header or ("运价" in header and "标准" not in header):
                letter = get_column_letter(column)
                ws.cell(total_row, column, f"=SUM({letter}2:{letter}{total_row-1})")
        ws.auto_filter.ref = f"A1:{get_column_letter(max_column)}{total_row}"
        _reset_split_sheet_display(ws, total_row)
    style_source = wb[writable_sheets[0]] if writable_sheets else None
    first_rows = [row for row in rows if row.get("price_kind") == "first"]
    second_rows = [row for row in rows if row.get("price_kind") == "second"]
    two_trip_rows = [
        row for row in rows
        if any(token in str(row.get("remark") or "") for token in ("两辆", "两趟合一趟", "合并为一辆"))
    ]
    _write_special_sheet(wb, "二干基础", second_rows, style_source=style_source)
    _write_special_sheet(wb, "一干汇总", first_rows, style_source=style_source)
    _write_special_sheet(wb, "两趟合一趟", two_trip_rows, style_source=style_source, two_trip=True)
    _write_downgrade_sheet(wb, rows)
    for ws in wb.worksheets:
        ws.sheet_view.topLeftCell = "A1"
        ws.sheet_view.zoomScale = 80
        ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]
    wb.active = 0
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    _atomic_save(wb, output)
    wb.close()
    return output, {name: len(items) for name, items in grouped.items()}


PROFIT_SUBHEADERS = ["趟次(趟)", "装载量(袋件)", "每车平均装载量(袋件)", "运费（元）", "结算收入（元）", "盈亏值（元）", "平均每车盈亏（元）"]
PROFIT_SUMMARY_HEADERS = ["趟次", "委办运费(元)", "结算收入(元)", "总盈亏(元)"]


def _tonnage_order(value: str) -> tuple[float, int]:
    key = tonnage_key(value)
    if key == "40A":
        return 40.0, 1
    if key == "40B":
        return 40.0, 2
    try:
        return float(key), 0
    except ValueError:
        return 9999.0, 0


def _display_tonnage(value: str) -> Any:
    key = tonnage_key(value)
    try:
        number = float(key)
        return int(number) if number.is_integer() else number
    except ValueError:
        return key


def write_profit_workbook(
    template_path: Path | str,
    layout: dict[str, Any],
    metrics: dict[tuple[str, str], Any] | Path | str,
    output: Path | str | tuple[int, int],
    month: tuple[int, int] | str,
    kind: str | None = None,
    previous_totals: dict[tuple[str, str], dict[str, float]] | None = None,
    warnings: list[str] | None = None,
) -> Path:
    # Keep old callers working while the public interface transitions to an explicit layout.
    if not isinstance(layout, dict) or not {"sheet", "groups", "route_rows"} <= set(layout):
        old_metrics = layout
        old_output = metrics
        old_month = output
        old_kind = month
        layout = discover_profit_template(template_path)
        metrics = old_metrics
        output = old_output
        month = old_month
        kind = str(old_kind)
    if not isinstance(metrics, dict) or not isinstance(month, tuple) or kind not in {"first", "second"}:
        raise ValueError("盈亏表写入参数无效")

    source_wb = load_workbook(template_path, data_only=False, keep_links=False)
    try:
        source = source_wb[layout["sheet"]]
        ws = source_wb.copy_worksheet(source)
        group_row, header_row = [int(value) for value in layout["header_rows"]]
        title_row = int(layout["title_row"])
        source_groups = {
            tonnage_key(item.get("tonnage") or item.get("label")): item
            for item in layout["groups"]
            if tonnage_key(item.get("tonnage") or item.get("label"))
        }
        if not source_groups:
            raise ValueError("盈亏模板动态布局没有有效7列吨位分组")
        required_groups = {
            tonnage_key(tonnage)
            for _, tonnage in metrics
            if tonnage_key(tonnage)
        }
        groups = sorted(set(source_groups) | required_groups, key=_tonnage_order)
        standard_group = next(iter(source_groups.values()))
        source_summary_start = max(int(item["end_column"]) for item in source_groups.values()) + 1
        template_routes = [
            normalize_text(item.get("route"))
            for item in layout.get("route_rows", [])
            if normalize_text(item.get("route"))
        ]
        source_route_rows = {
            normalize_text(item.get("route")): int(item["row"])
            for item in layout.get("route_rows", [])
            if normalize_text(item.get("route")) and item.get("row")
        }
        metric_routes = {normalize_text(route) for route, _ in metrics if normalize_text(route)}
        routes = template_routes + sorted(metric_routes - set(template_routes))
        data_start_row = header_row + 1
        standard_data_row = next(
            iter(source_route_rows.values()),
            int(layout.get("total_row") or data_start_row),
        )
        source_total_row = int(layout.get("total_row") or standard_data_row)

        final_column = 1 + 7 * len(groups) + 4
        summary_start = 2 + 7 * len(groups)
        source_final_column = source_summary_start + 3
        target_total_row = data_start_row + len(routes)
        rewrite_max_row = max(source_total_row + 3, target_total_row + 3)
        rewrite_max_column = max(source_final_column, final_column)
        for merged in list(ws.merged_cells.ranges):
            intersects_rewritten_table = (
                merged.max_row >= title_row
                and merged.min_row <= rewrite_max_row
                and merged.max_col >= 1
                and merged.min_col <= rewrite_max_column
            )
            if intersects_rewritten_table:
                ws.unmerge_cells(str(merged))
        for row in range(title_row, rewrite_max_row + 1):
            for column in range(1, rewrite_max_column + 1):
                ws.cell(row, column).value = None

        ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=final_column)
        ws.cell(title_row, 1, f"{month[0]}年{month[1]}月{'一干' if kind == 'first' else '二干'}委办运输盈亏情况")
        _copy_cell_style(source.cell(title_row, 1), ws.cell(title_row, 1))
        ws.merge_cells(start_row=group_row, start_column=1, end_row=header_row, end_column=1)
        ws.cell(group_row, 1, source.cell(group_row, 1).value or "邮路名称")
        _copy_cell_style(source.cell(group_row, 1), ws.cell(group_row, 1))
        ws.column_dimensions["A"].width = source.column_dimensions["A"].width

        group_starts: dict[str, int] = {}
        group_sources: dict[str, dict[str, Any]] = {}
        for index, group in enumerate(groups):
            start = 2 + index * 7
            group_starts[group] = start
            source_group = source_groups.get(group, standard_group)
            group_sources[group] = source_group
            source_start = int(source_group["start_column"])
            ws.merge_cells(start_row=group_row, start_column=start, end_row=group_row, end_column=start + 6)
            ws.cell(group_row, start, source.cell(group_row, source_start).value if group in source_groups else _display_tonnage(group))
            _copy_cell_style(source.cell(group_row, source_start), ws.cell(group_row, start))
            for offset, header in enumerate(PROFIT_SUBHEADERS):
                target = ws.cell(header_row, start + offset, header)
                _copy_cell_style(source.cell(header_row, source_start + offset), target)
                source_letter = get_column_letter(source_start + offset)
                ws.column_dimensions[get_column_letter(start + offset)].width = source.column_dimensions[source_letter].width

        for offset, header in enumerate(PROFIT_SUMMARY_HEADERS):
            column = summary_start + offset
            source_column = min(source_summary_start + offset, source.max_column)
            if source.cell(group_row, source_column).style_id == 0:
                source_column = int(standard_group["start_column"]) + min(offset, 6)
            ws.merge_cells(start_row=group_row, start_column=column, end_row=header_row, end_column=column)
            ws.cell(group_row, column, header)
            _copy_cell_style(source.cell(group_row, source_column), ws.cell(group_row, column))
            ws.column_dimensions[get_column_letter(column)].width = source.column_dimensions[get_column_letter(source_column)].width

        for target_row, route in enumerate(routes, data_start_row):
            source_row = source_route_rows.get(route, standard_data_row)
            ws.row_dimensions[target_row].height = source.row_dimensions[source_row].height
            ws.cell(target_row, 1, route)
            _copy_cell_style(source.cell(source_row, 1), ws.cell(target_row, 1))
            for group, start in group_starts.items():
                source_start = int(group_sources[group]["start_column"])
                for offset in range(7):
                    _copy_cell_style(source.cell(source_row, source_start + offset), ws.cell(target_row, start + offset))
            for offset in range(4):
                source_column = min(source_summary_start + offset, source.max_column)
                if source.cell(source_row, source_column).style_id == 0:
                    source_column = int(standard_group["start_column"]) + min(offset, 6)
                _copy_cell_style(source.cell(source_row, source_column), ws.cell(target_row, summary_start + offset))
            if route not in metric_routes:
                continue
            for group, start in group_starts.items():
                metric = metrics.get((route, group))
                trips = metric.trips if metric else 0
                load = metric.load if metric else 0
                freight = metric.freight if metric else 0
                income = metric.income if metric else 0
                ws.cell(target_row, start, trips)
                ws.cell(target_row, start + 1, load)
                ws.cell(target_row, start + 2, f"=IF({get_column_letter(start)}{target_row}=0,0,{get_column_letter(start+1)}{target_row}/{get_column_letter(start)}{target_row})")
                ws.cell(target_row, start + 3, freight)
                ws.cell(target_row, start + 4, income)
                ws.cell(target_row, start + 5, f"={get_column_letter(start+4)}{target_row}-{get_column_letter(start+3)}{target_row}")
                ws.cell(target_row, start + 6, f"=IF({get_column_letter(start)}{target_row}=0,0,{get_column_letter(start+5)}{target_row}/{get_column_letter(start)}{target_row})")
            trip_cells = "+".join(f"{get_column_letter(start)}{target_row}" for start in group_starts.values())
            freight_cells = "+".join(f"{get_column_letter(start+3)}{target_row}" for start in group_starts.values())
            income_cells = "+".join(f"{get_column_letter(start+4)}{target_row}" for start in group_starts.values())
            ws.cell(target_row, summary_start, f"={trip_cells}")
            ws.cell(target_row, summary_start + 1, f"={freight_cells}")
            ws.cell(target_row, summary_start + 2, f"={income_cells}")
            ws.cell(target_row, summary_start + 3, f"={get_column_letter(summary_start+2)}{target_row}-{get_column_letter(summary_start+1)}{target_row}")

        total_row = data_start_row + len(routes)
        ws.row_dimensions[total_row].height = source.row_dimensions[source_total_row].height
        ws.cell(total_row, 1, "合计")
        for column in range(1, final_column + 1):
            if column == 1:
                source_column = 1
            elif column < summary_start:
                group_index, offset = divmod(column - 2, 7)
                source_column = int(group_sources[groups[group_index]]["start_column"]) + offset
            else:
                source_column = min(source_summary_start + column - summary_start, source.max_column)
                if source.cell(source_total_row, source_column).style_id == 0:
                    source_column = int(standard_group["start_column"]) + min(column - summary_start, 6)
            _copy_cell_style(source.cell(source_total_row, source_column), ws.cell(total_row, column))
            data_source_row = source_route_rows.get(routes[0], standard_data_row) if routes else standard_data_row
            data_font = copy(source.cell(data_source_row, source_column).font)
            data_font.bold = True
            ws.cell(total_row, column).font = data_font
            if column > 1:
                if routes:
                    letter = get_column_letter(column)
                    ws.cell(total_row, column, f"=SUM({letter}{data_start_row}:{letter}{total_row-1})")
                else:
                    ws.cell(total_row, column, 0)

        prev_row, current_row, change_row = total_row + 1, total_row + 2, total_row + 3
        for comparison_row in (prev_row, current_row, change_row):
            ws.row_dimensions[comparison_row].height = source.row_dimensions[source_total_row].height
            for column in range(1, final_column + 1):
                _copy_cell_style(ws.cell(total_row, column), ws.cell(comparison_row, column))
        ws.cell(prev_row, 1, "上月" if previous_totals else "上月（暂无数据）")
        ws.cell(current_row, 1, "本月")
        ws.cell(change_row, 1, "环比%")
        for column in range(2, final_column + 1):
            letter = get_column_letter(column)
            ws.cell(current_row, column, f"={letter}{total_row}")
            ws.cell(prev_row, column, "--")
            ws.cell(change_row, column, "--")

        if previous_totals:
            comparable_columns: set[int] = set()
            previous_groups = {group for _, group in previous_totals}
            all_groups_comparable = all(group in previous_groups for group in group_starts)
            for group, start in group_starts.items():
                if group not in previous_groups:
                    continue
                values = [previous_totals.get((route, group), {}) for route in routes]
                for offset, field in ((0, "trips"), (3, "freight"), (4, "income")):
                    value = sum(float(item.get(field, 0) or 0) for item in values)
                    ws.cell(prev_row, start + offset, value)
                    comparable_columns.add(start + offset)
                ws.cell(prev_row, start + 5, f"={get_column_letter(start+4)}{prev_row}-{get_column_letter(start+3)}{prev_row}")
                comparable_columns.add(start + 5)
            if all_groups_comparable:
                trip_cells = "+".join(f"{get_column_letter(start)}{prev_row}" for start in group_starts.values())
                freight_cells = "+".join(f"{get_column_letter(start+3)}{prev_row}" for start in group_starts.values())
                income_cells = "+".join(f"{get_column_letter(start+4)}{prev_row}" for start in group_starts.values())
                ws.cell(prev_row, summary_start, f"={trip_cells}")
                ws.cell(prev_row, summary_start + 1, f"={freight_cells}")
                ws.cell(prev_row, summary_start + 2, f"={income_cells}")
                ws.cell(prev_row, summary_start + 3, f"={get_column_letter(summary_start+2)}{prev_row}-{get_column_letter(summary_start+1)}{prev_row}")
                comparable_columns.update(range(summary_start, summary_start + 4))
            for column in comparable_columns:
                letter = get_column_letter(column)
                ws.cell(change_row, column, f'=IF({letter}{prev_row}=0,"--",({letter}{current_row}-{letter}{prev_row})/{letter}{prev_row})')
                ws.cell(change_row, column).number_format = "0.00%"

        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 80
        ws.sheet_view.zoomScaleNormal = 80
        target_title = f"{str(month[0])[2:]}年{month[1]}月"
        for sheet in list(source_wb.worksheets):
            if sheet is not ws:
                source_wb.remove(sheet)
        ws.title = target_title
        source_wb.active = 0
        source_wb.calculation.calcMode = "auto"
        source_wb.calculation.fullCalcOnLoad = True
        source_wb.calculation.forceFullCalc = True
        _atomic_save(source_wb, Path(output))
        return Path(output)
    finally:
        source_wb.close()
