from __future__ import annotations

import hashlib
import json
import re
import warnings
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

from .normalize import (
    normalize_dispatch_no,
    normalize_plate,
    normalize_text,
    normalize_trailer,
    parse_daily_sheet_date,
    parse_month_token,
    tonnage_key,
)

warnings.filterwarnings("ignore", message="Workbook contains no default style.*", module="openpyxl.styles.stylesheet")


MODEL_FIELDS = {
    "sequence": ["序号"],
    "route": ["邮路", "邮路名称"],
    "plate": ["车号", "车牌号"],
    "trailer": ["挂车号", "车厢车牌号", "车厢码"],
    "length": ["长(m)", "长（m）", "长"],
    "width": ["宽(m)", "宽（m）", "宽"],
    "height": ["高(m)", "高（m）", "高"],
    "volume": ["测量容积(方)", "测量容积（方）", "测量容积", "容积"],
    "problem": ["车型问题"],
    "remark": ["备注"],
}

VEHICLE_FIELDS = {
    "route": ["邮路名称"],
    "date": ["邮路执行日期"],
    "dispatch_no": ["派车单号"],
    "plate": ["车牌号"],
    "trailer": ["车厢码", "挂车号"],
    "transport_volume": ["邮运量"],
    "settlement_tonnage": ["结算吨位"],
}

FINANCE_FIELDS = {
    "route": ["邮路名称", "邮路", "线路"],
    "tonnage": ["额定吨位", "结算吨位", "要求派车吨位", "实际派车吨位"],
    "settlement_income": ["结算收入"],
    "universal_income": ["普服收入"],
}

PRICE_ROUTE_HEADERS = {"邮路名称", "路向"}
PRICE_BLANK_ROUTE_RUN = 10
PROFIT_GROUP_HEADERS = ["趟次(趟)", "装载量(袋件)", "每车平均装载量(袋件)", "运费（元）", "结算收入（元）", "盈亏值（元）", "平均每车盈亏（元）"]
REVIEW_HEADERS = [
    "序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", None, "车牌号", None, "容间",
    "吨位", "公里", "运输量", "路单号", "运价", "标准价", "备注", "是否跨档", "变动前容间", "变动前时间", "系数",
]

DEFAULT_COMPANY_MAP_PATH = Path(__file__).resolve().parents[2] / "assets" / "company-map.json"


@dataclass
class FileProfile:
    path: Path
    role: str
    site: str | None = None
    sheet_names: list[str] = field(default_factory=list)
    header_row: int | None = None
    field_map: dict[str, int] = field(default_factory=dict)
    months: list[tuple[int | None, int]] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def file_sha256(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def discover_workbooks(input_dir: Path | str) -> list[Path]:
    root = Path(input_dir)
    if not root.exists():
        return []

    def is_agentos_system_file(path: Path) -> bool:
        relative_parts = path.relative_to(root).parts
        return bool(relative_parts) and relative_parts[0].lower() in {"skills", "output"}

    return sorted(
        path
        for path in root.rglob("*.xlsx")
        if not is_agentos_system_file(path)
        and not path.name.startswith("~$")
        and "_02_基础报表初版" not in path.name
        and "_03_基础报表拆分版" not in path.name
        and "_04_" not in path.name
    )


def _header_key(value: Any) -> str:
    return re.sub(r"[\s:：]", "", normalize_text(value)).lower()


def _matches_header(actual: str, candidate: str) -> bool:
    left, right = _header_key(actual), _header_key(candidate)
    return left == right or (right and right in left)


def _find_header_column(values: Sequence[Any], aliases: Sequence[str]) -> int | None:
    alias_keys = {_header_key(alias) for alias in aliases}
    for column, value in enumerate(values, 1):
        if value is not None and _header_key(value) in alias_keys:
            return column
    for column, value in enumerate(values, 1):
        if value is not None and any(_matches_header(str(value), alias) for alias in aliases):
            return column
    return None


def find_header_row(ws, required_groups: Mapping[str, Sequence[str]], max_rows: int = 30) -> tuple[int, dict[str, int]]:
    for row_number, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, values_only=True), 1):
        values = list(row)
        mapping: dict[str, int] = {}
        for standard, aliases in required_groups.items():
            column = _find_header_column(values, aliases)
            if column:
                mapping[standard] = column
        if len(mapping) == len(required_groups):
            return row_number, mapping
    raise ValueError(f"工作表 {ws.title!r} 未找到必需表头: {list(required_groups)}")


def _try_header(ws, groups: Mapping[str, Sequence[str]]) -> tuple[int | None, dict[str, int]]:
    try:
        return find_header_row(ws, groups)
    except ValueError:
        return None, {}


def _sheet_months(sheet_names: Iterable[str]) -> list[tuple[int | None, int]]:
    result: list[tuple[int | None, int]] = []
    for name in sheet_names:
        token = parse_month_token(name)
        if token and token not in result:
            result.append(token)
    return result


def _paired_role(path: Path, first_role: str, second_role: str) -> str | None:
    filename = normalize_text(path.name).casefold()
    has_first = "一干" in filename or "first" in filename
    has_second = "二干" in filename or "second" in filename
    if has_first == has_second:
        return None
    return first_role if has_first else second_role


def _price_header(ws) -> tuple[int | None, dict[str, int]]:
    route_column = distance_column = None
    header_row = None
    has_tier_or_coefficient = False
    for row, values in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        for column, value in enumerate(values, 1):
            text = normalize_text(value)
            if text in PRICE_ROUTE_HEADERS and route_column is None:
                route_column, header_row = column, row
            if ("公里" in text or "里程" in text) and distance_column is None:
                distance_column = column
            if "系数" in text or "吨" in text:
                has_tier_or_coefficient = True
    if route_column and distance_column and has_tier_or_coefficient:
        return header_row, {"route": route_column, "distance": distance_column}
    return None, {}


def _profit_group_header(ws) -> int | None:
    expected = [normalize_text(value) for value in PROFIT_GROUP_HEADERS]
    for row, values in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        for column in range(0, len(values) - len(expected) + 1):
            actual = [normalize_text(value) for value in values[column:column + len(expected)]]
            if actual == expected:
                return row
    return None


def _review_header(ws) -> tuple[int | None, dict[str, int]]:
    expected = [normalize_text(value) for value in REVIEW_HEADERS]
    for row, values in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        actual = [normalize_text(value) for value in values[:len(expected)]]
        if actual == expected:
            return row, {f"column_{column}": column for column in range(1, len(expected) + 1)}
    return None, {}


def classify_workbook(path: Path | str) -> FileProfile:
    path = Path(path)
    try:
        wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except Exception as exc:
        return FileProfile(path=path, role="damaged", errors=[str(exc)])
    profile = FileProfile(path=path, role="unknown", sheet_names=list(wb.sheetnames))
    profile.months = _sheet_months(profile.sheet_names)
    file_month = parse_month_token(path.name)
    if file_month and file_month not in profile.months:
        profile.months.append(file_month)
    try:
        for ws in wb.worksheets:
            if hasattr(ws, "reset_dimensions") and ws.calculate_dimension() == "A1:A1":
                ws.reset_dimensions()
            row, fields = _try_header(ws, VEHICLE_FIELDS)
            if row:
                profile.role, profile.header_row, profile.field_map = "vehicle", row, fields
                date_column = fields["date"]
                for values in ws.iter_rows(min_row=row + 1, values_only=True):
                    value = values[date_column - 1] if date_column <= len(values) else None
                    parsed = parse_date(value)
                    token = (parsed.year, parsed.month) if parsed else None
                    if token and token not in profile.months:
                        profile.months.append(token)
                return profile
            finance_required = {key: FINANCE_FIELDS[key] for key in ("route", "tonnage", "settlement_income", "universal_income")}
            row, fields = _try_header(ws, finance_required)
            if row:
                profile.role, profile.header_row, profile.field_map = "finance", row, fields
                return profile
            model_required = {key: MODEL_FIELDS[key] for key in ("route", "plate", "trailer", "volume")}
            row, fields = _try_header(ws, model_required)
            if row:
                lowered = normalize_text(path.name).lower()
                daily_sheets = sum(1 for name in wb.sheetnames if parse_month_token(name))
                if "车型" in lowered and "在线" not in lowered:
                    role = "model"
                elif "港务" in lowered:
                    role = "online_gangwu"
                elif "草滩" in lowered or "在线" in lowered or daily_sheets >= 2:
                    role = "online_caotan"
                else:
                    role = "model"
                profile.role, profile.header_row, profile.field_map = role, row, fields
                profile.site = "港务" if role == "online_gangwu" else ("草滩" if role == "online_caotan" else None)
                return profile
            row, fields = _price_header(ws)
            price_role = _paired_role(path, "first_price", "second_price")
            if row and price_role:
                profile.role, profile.header_row, profile.field_map = price_role, row, fields
                return profile
            row = _profit_group_header(ws)
            profit_role = _paired_role(path, "first_profit_template", "second_profit_template")
            if row and profit_role:
                profile.role, profile.header_row, profile.field_map = profit_role, row, {"profit_group": 2}
                return profile
            row, fields = _review_header(ws)
            if row:
                profile.role, profile.header_row, profile.field_map = "draft_template", row, fields
                return profile
    finally:
        wb.close()
    return profile


def _to_number(value: Any) -> float | int | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return value
    try:
        return float(normalize_text(value).replace(",", ""))
    except ValueError:
        return None


def _value(row: Sequence[Any], field_map: Mapping[str, int], key: str) -> Any:
    column = field_map.get(key)
    return row[column - 1] if column and column <= len(row) else None


def _record_from_model_row(row: Sequence[Any], field_map: Mapping[str, int], source_row: int, source_sheet: str) -> dict[str, Any]:
    length, width, height = (_to_number(_value(row, field_map, key)) for key in ("length", "width", "height"))
    volume = _to_number(_value(row, field_map, "volume"))
    if volume is None and all(item is not None for item in (length, width, height)):
        volume = float(length) * float(width) * float(height)
    plate = normalize_plate(_value(row, field_map, "plate"))
    trailer = normalize_trailer(_value(row, field_map, "trailer"))
    return {
        "source_sheet": source_sheet,
        "source_row": source_row,
        "sequence": _value(row, field_map, "sequence"),
        "route": _value(row, field_map, "route"),
        "plate": _value(row, field_map, "plate"),
        "trailer": _value(row, field_map, "trailer") or "--",
        "plate_key": plate,
        "trailer_key": trailer,
        "length": length,
        "width": width,
        "height": height,
        "volume": volume,
        "measured_volume": volume,
        "problem": _value(row, field_map, "problem"),
        "remark": _value(row, field_map, "remark"),
    }


def read_model_rows(path: Path | str) -> list[dict[str, Any]]:
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    result: list[dict[str, Any]] = []
    required = {key: MODEL_FIELDS[key] for key in ("route", "plate", "trailer", "volume")}
    try:
        for ws in wb.worksheets:
            header_row, required_map = _try_header(ws, required)
            if not header_row:
                continue
            _, all_map = _try_header(ws, {key: aliases for key, aliases in MODEL_FIELDS.items() if key in required_map or key not in required})
            field_map = dict(required_map)
            # Add optional fields from the same header row without requiring them all.
            header_values = [cell.value for cell in next(ws.iter_rows(min_row=header_row, max_row=header_row))]
            for key, aliases in MODEL_FIELDS.items():
                column = _find_header_column(header_values, aliases)
                if column:
                    field_map[key] = column
            for source_row, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
                record = _record_from_model_row(row, field_map, source_row, ws.title)
                if record["plate_key"] and normalize_text(record["route"]):
                    record["source_id"] = f"M{len(result)+1:06d}"
                    result.append(record)
    finally:
        wb.close()
    return result


def read_online_rows(path: Path | str, target_month: tuple[int, int], site: str) -> list[dict[str, Any]]:
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    result: list[dict[str, Any]] = []
    required = {key: MODEL_FIELDS[key] for key in ("route", "plate", "trailer", "volume")}
    try:
        for source_sheet_order, ws in enumerate(wb.worksheets):
            token = parse_month_token(ws.title)
            if not token or token[1] != target_month[1]:
                continue
            header_row, field_map = find_header_row(ws, required)
            header_values = [cell.value for cell in next(ws.iter_rows(min_row=header_row, max_row=header_row))]
            for key, aliases in MODEL_FIELDS.items():
                column = _find_header_column(header_values, aliases)
                if column:
                    field_map[key] = column
            for source_row, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
                record = _record_from_model_row(row, field_map, source_row, ws.title)
                if not record["plate_key"]:
                    continue
                record["site"] = site
                record["measurement_date"] = parse_daily_sheet_date(ws.title, target_month[0])
                record["source_file"] = path.name
                record["source_path"] = str(path.resolve())
                record["source_sheet_order"] = source_sheet_order
                record["source_order"] = len(result)
                result.append(record)
    finally:
        wb.close()
    return result


def load_company_map(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    companies = data.get("companies") if isinstance(data, dict) else None
    if not isinstance(companies, list) or not companies:
        raise ValueError("company-map.json 必须定义至少一家公司")
    tokens: set[str] = set()
    company_names: set[str] = set()
    for entry in companies:
        if not isinstance(entry, dict):
            raise ValueError("company-map.json 公司定义无效")
        token = normalize_text(entry.get("source_token"))
        company = normalize_text(entry.get("company"))
        if not token or not company or token in tokens or company in company_names:
            raise ValueError("company-map.json 公司定义无效")
        tokens.add(token)
        company_names.add(company)
    return data


def _company_from_path(path: Path, company_map: Mapping[str, Any]) -> str:
    stem = normalize_text(path.stem)
    for entry in company_map["companies"]:
        if normalize_text(entry["source_token"]) in stem:
            return normalize_text(entry["company"])
    raise ValueError(f"车管文件未在 company-map.json 中映射: {path.name}")


def read_vehicle_rows(
    paths: Iterable[Path | str],
    company_map: Mapping[str, Any] | None = None,
) -> list[dict[str, Any]]:
    company_map = company_map or load_company_map(DEFAULT_COMPANY_MAP_PATH)
    result: list[dict[str, Any]] = []
    for raw_path in paths:
        path = Path(raw_path)
        company_name = _company_from_path(path, company_map)
        wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
        try:
            for ws in wb.worksheets:
                header_row, field_map = _try_header(ws, VEHICLE_FIELDS)
                if not header_row:
                    continue
                for source_row, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
                    plate = normalize_plate(_value(row, field_map, "plate"))
                    if not plate:
                        continue
                    trailer = normalize_trailer(_value(row, field_map, "trailer"))
                    execution_datetime = parse_date(_value(row, field_map, "date"))
                    result.append({
                        "source_file": path.name,
                        "source_path": str(path.resolve()),
                        "source_sheet": ws.title,
                        "source_row": source_row,
                        "company": company_name,
                        "route": _value(row, field_map, "route"),
                        "date": _value(row, field_map, "date"),
                        "execution_date": execution_datetime.date() if execution_datetime else None,
                        "dispatch_no": normalize_dispatch_no(_value(row, field_map, "dispatch_no")),
                        "plate": plate,
                        "trailer": trailer,
                        "plate_key": plate,
                        "trailer_key": trailer,
                        "transport_volume": _to_number(_value(row, field_map, "transport_volume")),
                        "settlement_tonnage": _to_number(_value(row, field_map, "settlement_tonnage")),
                    })
        finally:
            wb.close()
    return result


def read_finance_rows(path: Path | str) -> list[dict[str, Any]]:
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    result: list[dict[str, Any]] = []
    try:
        for ws in wb.worksheets:
            if hasattr(ws, "reset_dimensions"):
                ws.reset_dimensions()
            header_row, field_map = _try_header(ws, FINANCE_FIELDS)
            if not header_row:
                continue
            for source_row, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
                route = _value(row, field_map, "route")
                tonnage = _to_number(_value(row, field_map, "tonnage"))
                if not normalize_text(route) or tonnage in (None, 0):
                    continue
                settlement = _to_number(_value(row, field_map, "settlement_income")) or 0
                universal = _to_number(_value(row, field_map, "universal_income")) or 0
                result.append({
                    "source_sheet": ws.title,
                    "source_row": source_row,
                    "route": route,
                    "tonnage": tonnage,
                    "settlement_income": settlement,
                    "universal_income": universal,
                    "net_income": settlement - universal,
                })
    finally:
        wb.close()
    return result


def _number_from_cell(ws, row: int, column: int) -> float | int | None:
    value = ws.cell(row, column).value
    number = _to_number(value)
    if number is not None:
        return number
    if isinstance(value, str) and value.startswith("="):
        match = re.fullmatch(r"=\$?([A-Z]+)\$?(\d+)\s*\*\s*([0-9.]+)", value.replace(" ", ""), re.I)
        if match:
            ref_column = column_index_from_string(match.group(1).upper())
            ref_row = int(match.group(2))
            base = _to_number(ws.cell(ref_row, ref_column).value)
            if base is not None:
                calculated = Decimal(str(base)) * Decimal(match.group(3))
                return float(calculated) if calculated % 1 else int(calculated)
    return None


def _extract_first_number(value: Any) -> float | None:
    text = normalize_text(value)
    match = re.search(r"(?:≥|>=)\s*(\d+(?:\.\d+)?)", text)
    if not match:
        return None
    return float(match.group(1))


def _extract_coefficient(value: Any, allow_unlabeled_number: bool = False) -> float | None:
    if allow_unlabeled_number:
        number = _to_number(value)
        if number is not None and not isinstance(value, bool) and 0 < float(number) <= 2:
            return float(number)
    match = re.search(r"系数\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)", normalize_text(value))
    if not match:
        return None
    coefficient = float(match.group(1))
    return coefficient if 0 < coefficient <= 2 else None


def _price_header_positions(ws) -> tuple[dict[str, int] | None, dict[str, int] | None]:
    for row in range(1, min(ws.max_row, 30) + 1):
        values = [ws.cell(row, column).value for column in range(1, ws.max_column + 1)]
        route_column = _find_header_column(values, tuple(PRICE_ROUTE_HEADERS))
        distance_column = _find_header_column(values, ("路向里程", "公里（KM）", "公里", "里程"))
        if route_column and distance_column:
            return {"row": row, "column": route_column}, {"row": row, "column": distance_column}
    return None, None


def _price_tier_header(
    ws,
    column: int,
    anchor_row: int,
    numeric_coefficient_rows: set[int],
) -> dict[str, Any] | None:
    first_window = max(1, anchor_row - 2)
    last_window = min(anchor_row, max(1, ws.max_row - 2))
    for start_row in range(first_window, last_window + 1):
        tonnage = threshold = coefficient = None
        tonnage_row = threshold_row = coefficient_row = None
        for row in range(start_row, min(ws.max_row, start_row + 2) + 1):
            raw = ws.cell(row, column).value
            text = normalize_text(raw)
            if tonnage is None and "吨" in text:
                tonnage = tonnage_key(raw)
                tonnage_row = row
            if threshold is None:
                threshold = _extract_first_number(raw)
                if threshold is not None:
                    threshold_row = row
            if coefficient is None:
                coefficient = _extract_coefficient(raw, row in numeric_coefficient_rows)
                if coefficient is not None:
                    coefficient_row = row
        if tonnage and threshold is not None and coefficient is not None:
            return {
                "key": tonnage,
                "column": column,
                "tonnage_row": tonnage_row,
                "threshold_row": threshold_row,
                "coefficient_row": coefficient_row,
                "min_volume": threshold,
                "coefficient": coefficient,
            }
    return None


def read_price_catalog(path: Path | str, kind: str) -> dict[str, Any]:
    wb = load_workbook(path, read_only=False, data_only=False, keep_links=False)
    catalog: dict[str, Any] = {"kind": kind, "tiers": [], "routes": {}, "sheets": []}
    tier_definitions: dict[str, dict[str, Any]] = {}
    try:
        for ws in wb.worksheets:
            if ws.max_row < 4 or ws.max_column < 4:
                continue
            route_header, distance_header = _price_header_positions(ws)
            if not route_header or not distance_header:
                continue
            header_row = route_header["row"]
            route_column = route_header["column"]
            distance_column = distance_header["column"]
            numeric_coefficient_rows = {
                row
                for row in range(1, min(ws.max_row, 30) + 1)
                if any(
                    _header_key(ws.cell(row, column).value) == "系数"
                    for column in range(1, distance_column + 1)
                )
            }
            tier_headers: list[dict[str, Any]] = []
            seen_keys: set[str] = set()
            for column in range(distance_column + 1, ws.max_column + 1):
                tier_header = _price_tier_header(ws, column, header_row, numeric_coefficient_rows)
                if not tier_header:
                    continue
                key = tier_header["key"]
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                tier_headers.append(tier_header)
                tier_definitions.setdefault(key, {
                    "key": key,
                    "min_volume": tier_header["min_volume"],
                    "coefficient": tier_header["coefficient"],
                })
            if not tier_headers:
                continue
            data_start = max(
                max(item[row_key] for item in tier_headers)
                for row_key in ("tonnage_row", "threshold_row", "coefficient_row")
            ) + 1
            catalog["sheets"].append({
                "sheet": ws.title,
                "route_header": route_header,
                "distance_header": distance_header,
                "tier_headers": [
                    {key: item[key] for key in ("key", "column", "tonnage_row", "threshold_row", "coefficient_row")}
                    for item in tier_headers
                ],
                "data_start_row": data_start,
            })
            populated_route_seen = False
            blank_route_count = 0
            for row in range(data_start, ws.max_row + 1):
                route = normalize_text(ws.cell(row, route_column).value)
                if not route:
                    if populated_route_seen:
                        blank_route_count += 1
                        if blank_route_count >= PRICE_BLANK_ROUTE_RUN:
                            break
                    continue
                populated_route_seen = True
                blank_route_count = 0
                distance = _number_from_cell(ws, row, distance_column)
                prices = {
                    item["key"]: _number_from_cell(ws, row, item["column"])
                    for item in tier_headers
                }
                prices = {key: value for key, value in prices.items() if value is not None}
                if not prices:
                    continue
                catalog["routes"].setdefault(route, {
                    "distance": distance,
                    "prices": prices,
                    "sheet": ws.title,
                    "kind": kind,
                })
    finally:
        wb.close()
    catalog["tiers"] = sorted(tier_definitions.values(), key=lambda item: (item["min_volume"], item["key"]))
    return catalog


def parse_date(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    text = normalize_text(value)
    for pattern in (r"(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})", r"(20\d{2})(\d{2})(\d{2})"):
        match = re.search(pattern, text)
        if match:
            try:
                return datetime(*map(int, match.groups()))
            except ValueError:
                return None
    return None
