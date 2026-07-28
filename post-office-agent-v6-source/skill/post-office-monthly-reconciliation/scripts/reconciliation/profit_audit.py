from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .normalize import tonnage_key
from .template_discovery import discover_profit_template


SUBHEADERS = ["趟次(趟)", "装载量(袋件)", "每车平均装载量(袋件)", "运费（元）", "结算收入（元）", "盈亏值（元）", "平均每车盈亏（元）"]
SUMMARY_HEADERS = ["趟次", "委办运费(元)", "结算收入(元)", "总盈亏(元)"]


def _number_equal(actual: Any, expected: Any, tolerance: float = 1e-6) -> bool:
    try:
        return abs(float(actual or 0) - float(expected or 0)) <= tolerance
    except Exception:
        return actual == expected


def audit_profit_workbook(
    path: Path | str,
    template_path: Path | str,
    metrics: Mapping[tuple[str, str], Any],
    month: tuple[int, int],
    kind: str,
) -> dict[str, Any]:
    path, template_path = Path(path), Path(template_path)
    errors: list[str] = []
    spot_candidates: list[dict[str, Any]] = []
    try:
        output_layout = discover_profit_template(path)
    except ValueError as exc:
        return {"errors": [f"盈亏表动态布局无效: {exc}"], "spot_candidates": []}
    try:
        template_layout = discover_profit_template(template_path)
    except ValueError as exc:
        return {"errors": [f"盈亏模板动态布局无效: {exc}"], "spot_candidates": []}
    wb = load_workbook(path, read_only=False, data_only=False, keep_links=False)
    template_wb = load_workbook(template_path, read_only=False, data_only=False, keep_links=False)
    try:
        expected_sheet = f"{str(month[0])[2:]}年{month[1]}月"
        if wb.sheetnames != [expected_sheet]:
            errors.append(f"盈亏表必须且只能保留当月Sheet {expected_sheet}")
        ws = wb[output_layout["sheet"]]
        template = template_wb[template_layout["sheet"]]
        title_row = int(output_layout["title_row"])
        group_row, header_row = [int(value) for value in output_layout["header_rows"]]
        output_groups = output_layout.get("groups", [])
        complete_groups = bool(output_groups) and all(
            int(group.get("width", 0)) == 7
            and int(group["end_column"]) == int(group["start_column"]) + 6
            for group in output_groups
        )
        if not complete_groups:
            return {"errors": errors + ["吨位区不是完整7列分组"], "spot_candidates": []}
        group_starts = {
            tonnage_key(group.get("tonnage") or group.get("label")): int(group["start_column"])
            for group in output_groups
            if tonnage_key(group.get("tonnage") or group.get("label"))
        }
        summary_start = max(int(group["end_column"]) for group in output_groups) + 1
        for group, start in group_starts.items():
            actual_headers = [ws.cell(header_row, start + offset).value for offset in range(7)]
            if actual_headers != SUBHEADERS:
                errors.append(f"吨位{group}的7列子表头不正确")
        if [ws.cell(group_row, summary_start + offset).value for offset in range(4)] != SUMMARY_HEADERS:
            errors.append("盈亏表汇总表头不正确")

        final_column = summary_start + 3
        required_merges = {
            f"A{title_row}:{get_column_letter(final_column)}{title_row}",
            f"A{group_row}:A{header_row}",
        }
        required_merges.update(
            f"{get_column_letter(start)}{group_row}:{get_column_letter(start+6)}{group_row}"
            for start in group_starts.values()
        )
        required_merges.update(
            f"{get_column_letter(summary_start+offset)}{group_row}:{get_column_letter(summary_start+offset)}{header_row}"
            for offset in range(4)
        )
        actual_merges = {str(item) for item in ws.merged_cells.ranges}
        missing_merges = sorted(required_merges - actual_merges)
        if missing_merges:
            errors.append("盈亏表缺少合并单元格: " + ", ".join(missing_merges))
        template_title_row = int(template_layout["title_row"])
        if (
            ws.cell(title_row, 1).font.name != template.cell(template_title_row, 1).font.name
            or ws.cell(title_row, 1).font.sz != template.cell(template_title_row, 1).font.sz
        ):
            errors.append("盈亏表标题字体与母版不一致")

        template_routes = [
            str(item["route"])
            for item in template_layout.get("route_rows", [])
            if item.get("route")
        ]
        metric_routes = {route for route, _ in metrics}
        expected_routes = template_routes + sorted(metric_routes - set(template_routes))
        route_rows: dict[str, int] = {}
        total_row = None
        data_start_row = header_row + 1
        for row in range(data_start_row, ws.max_row + 1):
            route = ws.cell(row, 1).value
            if route == "合计":
                total_row = row
                break
            if route not in (None, ""):
                route_rows[str(route)] = row
        if list(route_rows) != expected_routes:
            errors.append("盈亏表路线清单或顺序与模板及源数据重算结果不一致")

        template_sample_row = next(
            (int(item["row"]) for item in template_layout.get("route_rows", []) if item.get("row")),
            int(template_layout.get("total_row") or template_layout["header_rows"][1]),
        )
        template_groups = {
            tonnage_key(group.get("tonnage") or group.get("label")): group
            for group in template_layout.get("groups", [])
            if tonnage_key(group.get("tonnage") or group.get("label"))
        }
        standard_template_group = next(iter(template_groups.values()))
        for route in expected_routes:
            row = route_rows.get(route)
            if row is None:
                continue
            if route in template_routes and route not in metric_routes:
                populated = [
                    ws.cell(row, column).coordinate
                    for column in range(2, final_column + 1)
                    if ws.cell(row, column).value not in (None, "")
                ]
                if populated:
                    errors.append(f"无数据固定路线{route}应留空，发现值: {', '.join(populated[:10])}")
                continue
            for group, start in group_starts.items():
                metric = metrics.get((route, group))
                expected_values = {
                    0: metric.trips if metric else 0,
                    1: metric.load if metric else 0,
                    3: metric.freight if metric else 0,
                    4: metric.income if metric else 0,
                }
                labels = {0: "趟次", 1: "装载量", 3: "运费", 4: "结算收入"}
                for offset, expected in expected_values.items():
                    actual = ws.cell(row, start + offset).value
                    if not _number_equal(actual, expected):
                        errors.append(f"{route}/{group}吨{labels[offset]}与源数据重算结果不一致")
                expected_formulas = {
                    2: f"=IF({get_column_letter(start)}{row}=0,0,{get_column_letter(start+1)}{row}/{get_column_letter(start)}{row})",
                    5: f"={get_column_letter(start+4)}{row}-{get_column_letter(start+3)}{row}",
                    6: f"=IF({get_column_letter(start)}{row}=0,0,{get_column_letter(start+5)}{row}/{get_column_letter(start)}{row})",
                }
                for offset, expected in expected_formulas.items():
                    if ws.cell(row, start + offset).value != expected:
                        errors.append(f"{route}/{group}吨派生单元格公式不正确")
                template_group = template_groups.get(group, standard_template_group)
                template_start = int(template_group["start_column"])
                for offset in range(7):
                    if ws.cell(row, start + offset).number_format != template.cell(template_sample_row, template_start + offset).number_format:
                        errors.append(f"{route}/{group}吨第{offset+1}子列数字格式与母版不一致")
                if metric:
                    spot_candidates.append({
                        "kind": kind,
                        "route": route,
                        "tonnage": group,
                        "source": {"trips": metric.trips, "load": metric.load, "freight": metric.freight, "income": metric.income},
                        "output": {
                            "trips": ws.cell(row, start).value,
                            "load": ws.cell(row, start + 1).value,
                            "freight": ws.cell(row, start + 3).value,
                            "income": ws.cell(row, start + 4).value,
                        },
                    })
            trip_cells = "+".join(f"{get_column_letter(start)}{row}" for start in group_starts.values())
            freight_cells = "+".join(f"{get_column_letter(start+3)}{row}" for start in group_starts.values())
            income_cells = "+".join(f"{get_column_letter(start+4)}{row}" for start in group_starts.values())
            expected_summary = [
                f"={trip_cells}", f"={freight_cells}", f"={income_cells}",
                f"={get_column_letter(summary_start+2)}{row}-{get_column_letter(summary_start+1)}{row}",
            ]
            if [ws.cell(row, summary_start + offset).value for offset in range(4)] != expected_summary:
                errors.append(f"{route}汇总公式不正确")
        if total_row is None:
            errors.append("盈亏表缺少合计行")
        elif expected_routes:
            for column in range(2, final_column + 1):
                letter = get_column_letter(column)
                expected = f"=SUM({letter}{data_start_row}:{letter}{total_row-1})"
                if ws.cell(total_row, column).value != expected:
                    errors.append(f"合计行{letter}{total_row}公式不正确")
        if total_row is not None:
            comparison_labels = [ws.cell(total_row + offset, 1).value for offset in (1, 2, 3)]
            if comparison_labels[0] not in {"上月", "上月（暂无数据）"} or comparison_labels[1:] != ["本月", "环比%"]:
                errors.append("盈亏表底部缺少上月/本月/环比%三行")
    finally:
        wb.close()
        template_wb.close()
    return {"errors": errors, "spot_candidates": spot_candidates}
