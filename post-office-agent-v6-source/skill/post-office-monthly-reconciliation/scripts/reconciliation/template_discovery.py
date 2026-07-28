from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Mapping

from openpyxl import load_workbook

from .excel_io import PROFIT_GROUP_HEADERS, REVIEW_HEADERS
from .normalize import destination_name, normalize_text, parse_month_token, route_candidates, tonnage_key


_DRAFT_FAILURE = "初版模板未找到完整业务表头"
_PROFIT_FAILURE = "盈亏模板未找到有效月份页签"
_SPLIT_FAILURE = "拆分母版未发现可写入线路页签"
_RESERVED_SHEET_MARKERS = (
    "汇总", "基础", "降吨位", "明细", "统计", "占比", "两趟合一趟",
    "说明", "封面", "目录", "须知", "摘要", "历史", "保留",
)
_CHINESE_TEXT = re.compile(r"[\u3400-\u9fff]")


def _failure(path: Path, reason: str, detail: Any = None) -> ValueError:
    suffix = f": {detail}" if detail else ""
    return ValueError(f"{path.name}: {reason}{suffix}")


def _normalized(value: Any) -> str:
    return normalize_text(value).casefold()


def _load(path: Path, reason: str):
    try:
        return load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except Exception as exc:
        raise _failure(path, reason, exc) from exc


def discover_draft_template(path: Path | str) -> dict[str, Any]:
    template = Path(path)
    workbook = _load(template, _DRAFT_FAILURE)
    expected = [(header, _normalized(header)) for header in REVIEW_HEADERS if normalize_text(header)]
    candidates: list[dict[str, Any]] = []
    try:
        for worksheet in workbook.worksheets:
            for row_number, values in enumerate(worksheet.iter_rows(values_only=True), 1):
                columns: dict[str, int] = {}
                for column, value in enumerate(values, 1):
                    key = _normalized(value)
                    if key and key not in columns:
                        columns[key] = column
                if all(key in columns for _, key in expected):
                    candidates.append({
                        "sheet": worksheet.title,
                        "header_row": row_number,
                        "headers": {header: columns[key] for header, key in expected},
                        "data_start_row": row_number + 1,
                    })
    except Exception as exc:
        raise _failure(template, _DRAFT_FAILURE, exc) from exc
    finally:
        workbook.close()
    if len(candidates) != 1:
        raise _failure(template, _DRAFT_FAILURE)
    return candidates[0]


def _profit_groups(worksheet) -> tuple[int, int, int, list[dict[str, Any]]] | None:
    expected = [_normalized(header) for header in PROFIT_GROUP_HEADERS]
    row_limit = min(worksheet.max_row or 0, 30)
    column_limit = worksheet.max_column or 0
    for header_row in range(2, row_limit + 1):
        group_row = header_row - 1
        groups: list[dict[str, Any]] = []
        for start in range(1, max(1, column_limit - len(expected) + 2)):
            actual = [_normalized(worksheet.cell(header_row, start + offset).value) for offset in range(len(expected))]
            label = normalize_text(worksheet.cell(group_row, start).value)
            if actual == expected and label:
                groups.append({
                    "label": label,
                    "tonnage": tonnage_key(label),
                    "start_column": start,
                    "end_column": start + len(expected) - 1,
                    "width": len(expected),
                })
        if not groups:
            continue
        title_row = next(
            (
                row
                for row in range(group_row - 1, 0, -1)
                if any(normalize_text(worksheet.cell(row, column).value) for column in range(1, column_limit + 1))
            ),
            0,
        )
        if title_row:
            return title_row, group_row, header_row, groups
    return None


def _sheet_year(worksheet, token: tuple[int | None, int], title_row: int) -> int | None:
    if token[0] is not None:
        return token[0]
    match = re.search(r"(?<!\d)(\d{2})\s*年", normalize_text(worksheet.title))
    if match:
        return 2000 + int(match.group(1))
    for cell in worksheet[title_row]:
        title_token = parse_month_token(cell.value)
        if title_token and title_token[0] is not None and title_token[1] == token[1]:
            return title_token[0]
    return None


def _profit_rows(worksheet, header_row: int) -> tuple[list[dict[str, Any]], int | None]:
    routes: list[dict[str, Any]] = []
    total_row = None
    for row in range(header_row + 1, (worksheet.max_row or header_row) + 1):
        value = normalize_text(worksheet.cell(row, 1).value)
        if value == "合计":
            total_row = row
            break
        if value and "-" in value:
            routes.append({"route": value, "row": row})
    return routes, total_row


def _resolve_candidate_years(candidates: list[dict[str, Any]]) -> None:
    current_year = None
    previous_month = None
    for candidate in candidates:
        if candidate["year"] is not None:
            current_year = candidate["year"]
        elif current_year is not None:
            if previous_month is not None and candidate["month"] < previous_month:
                current_year += 1
            candidate["year"] = current_year
        previous_month = candidate["month"]

    current_year = None
    next_month = None
    for candidate in reversed(candidates):
        if candidate["year"] is not None:
            current_year = candidate["year"]
        elif current_year is not None:
            if next_month is not None and candidate["month"] > next_month:
                current_year -= 1
            candidate["year"] = current_year
        next_month = candidate["month"]


def discover_profit_template(path: Path | str) -> dict[str, Any]:
    template = Path(path)
    workbook = _load(template, _PROFIT_FAILURE)
    candidates: list[dict[str, Any]] = []
    try:
        for sheet_order, worksheet in enumerate(workbook.worksheets):
            token = parse_month_token(worksheet.title)
            if token is None:
                continue
            structure = _profit_groups(worksheet)
            if structure is None:
                continue
            title_row, group_row, header_row, groups = structure
            routes, total_row = _profit_rows(worksheet, header_row)
            year = _sheet_year(worksheet, token, title_row)
            layout = {
                "sheet": worksheet.title,
                "title_row": title_row,
                "header_rows": [group_row, header_row],
                "route_rows": routes,
                "groups": groups,
                "total_row": total_row,
            }
            candidates.append({
                "year": year,
                "month": token[1],
                "sheet_order": sheet_order,
                "layout": layout,
            })
    except Exception as exc:
        raise _failure(template, _PROFIT_FAILURE, exc) from exc
    finally:
        workbook.close()
    if not candidates:
        raise _failure(template, _PROFIT_FAILURE)
    _resolve_candidate_years(candidates)
    selected = max(
        candidates,
        key=lambda item: (item["year"] or 0, item["month"], item["sheet_order"]),
    )
    return selected["layout"]


def _alias_mapping(aliases: Mapping[str, Any]) -> Mapping[str, Any]:
    nested = aliases.get("aliases") if isinstance(aliases, Mapping) else None
    return nested if isinstance(nested, Mapping) else aliases


def _reserved_sheet(name: str) -> bool:
    text = normalize_text(name)
    return bool(re.fullmatch(r"sheet\d*", text, flags=re.IGNORECASE)) or any(marker in text for marker in _RESERVED_SHEET_MARKERS)


def _sheet_destination(name: str) -> str:
    text = re.sub(r"[\(（].*$", "", normalize_text(name)).strip()
    text = re.sub(r"(?:(?:一干|二干|业务|往返))+$", "", text).strip()
    return text


def _semantic_route_sheet(name: str) -> bool:
    """Recognize route-only sheet names whose intentionally blank bodies offer no clues."""
    text = normalize_text(name)
    if _reserved_sheet(text):
        return False
    destination = _sheet_destination(text)
    if not re.fullmatch(r"[\u3400-\u9fff]{2,12}", destination):
        return False
    carrier_qualified = bool(re.fullmatch(r"[\u3400-\u9fff]{2,12}[（(][^()（）]{1,20}[）)]", text))
    round_trip = bool(re.fullmatch(r"[\u3400-\u9fff]{2,12}往返", text))
    return carrier_qualified or round_trip


def _is_route_text(value: str) -> bool:
    origin, separator, destination = normalize_text(value).partition("-")
    return bool(separator and _CHINESE_TEXT.search(origin) and _CHINESE_TEXT.search(destination))


def _route_evidence(worksheet, aliases: Mapping[str, Any]) -> tuple[list[str], bool]:
    clues: list[str] = []
    has_route_header = False
    max_row = min(worksheet.max_row or 0, 30)
    max_column = min(worksheet.max_column or 0, 64)
    for row in worksheet.iter_rows(min_row=1, max_row=max_row, max_col=max_column, values_only=True):
        for value in row:
            text = normalize_text(value).replace("－", "-").replace("—", "-")
            if _normalized(text) in {"邮路", "邮路名称", "线路", "线路名称"}:
                has_route_header = True
            if "-" not in text:
                continue
            for candidate in route_candidates(text, aliases):
                if _is_route_text(candidate) and candidate not in clues:
                    clues.append(candidate)
    return clues, has_route_header


def discover_split_template(path: Path | str, aliases: Mapping[str, Any]) -> dict[str, Any]:
    template = Path(path)
    workbook = _load(template, _SPLIT_FAILURE)
    sheet_names = list(workbook.sheetnames)
    writable: list[str] = []
    targets: dict[str, str] = {}
    alias_map = _alias_mapping(aliases)
    try:
        evidence: list[tuple[Any, list[str], bool, str]] = []
        clue_destinations: set[str] = set()
        for worksheet in workbook.worksheets:
            if _reserved_sheet(worksheet.title):
                continue
            clues, has_route_header = _route_evidence(worksheet, alias_map)
            sheet_target = _sheet_destination(worksheet.title)
            evidence.append((worksheet, clues, has_route_header, sheet_target))
            clue_destinations.update(destination_name(route) for route in clues if destination_name(route))

        for worksheet, clues, has_route_header, sheet_target in evidence:
            if (
                not clues
                and not has_route_header
                and sheet_target not in clue_destinations
                and not _semantic_route_sheet(worksheet.title)
            ):
                continue
            writable.append(worksheet.title)
            for route in clues:
                targets.setdefault(route, worksheet.title)
                destination = destination_name(route)
                if destination:
                    targets.setdefault(destination, worksheet.title)
            if sheet_target:
                targets.setdefault(sheet_target, worksheet.title)
            if _semantic_route_sheet(worksheet.title):
                targets.setdefault(normalize_text(worksheet.title), worksheet.title)
    except Exception as exc:
        raise _failure(template, _SPLIT_FAILURE, exc) from exc
    finally:
        workbook.close()
    if not writable:
        raise _failure(template, _SPLIT_FAILURE)
    return {"sheets": sheet_names, "writable_sheets": writable, "route_targets": targets}
