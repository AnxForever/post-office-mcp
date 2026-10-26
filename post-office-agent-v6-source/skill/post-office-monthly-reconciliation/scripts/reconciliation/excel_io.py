from __future__ import annotations

import hashlib
import json
import re
import warnings
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, get_column_letter

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
    # 往返属性=往返、邮路往返属性=去程 的行，按价格表「往返」价结算；
    # 同一派车单号的返程行不单独计价。
    "round_trip": ["往返属性"],
    "route_direction": ["邮路往返属性"],
}

FINANCE_FIELDS = {
    "route": ["邮路名称", "邮路", "线路"],
    "tonnage": ["额定吨位", "结算吨位", "要求派车吨位", "实际派车吨位"],
    "settlement_income": ["结算收入"],
    "universal_income": ["普服收入"],
}

# 车管平台「汇总导出」的最低必需列。平台导出没有 结算吨位/往返属性/邮路往返属性，
# 所以这几个字段只在存在时读取，缺了不报错——价格一律按价格表重算，
# 平台自带的「运价/结算金额」不参与计价（口径与官方报表不一致，且有重复计）。
VEHICLE_PLATFORM_FIELDS = {
    "route": ["邮路名称"],
    "date": ["邮路执行日期"],
    "dispatch_no": ["派车单号"],
    "plate": ["车牌号"],
}
# 车管平台汇总里的可选列：有就用，没有就留空。
VEHICLE_PLATFORM_OPTIONAL = {
    "trailer": ["车厢码", "挂车号"],
    "transport_volume": ["邮运量"],
    "capacity": ["容间"],
    "company": ["班次单号", "签订局名称"],
    "level": ["邮路级别"],
}
# 容积测量表（草滩+港务合并导出）的必需列：没有日期表头，日期单独识别。
MEASURE_FIELDS = {
    "route": ["邮路名称"],
    "plate": ["车牌号"],
    "volume": ["测量容积(方)", "测量容积（方）", "测量容积", "容积"],
}
MEASURE_OPTIONAL = {
    "trailer": ["车厢车牌号", "挂车号", "车厢码"],
    "length": ["长(m)", "长（m）", "长"],
    "width": ["宽(m)", "宽（m）", "宽"],
    "height": ["高(m)", "高（m）", "高"],
    "problem": ["车型问题"],
    "remark": ["备注", "未装容积(米)", "未装容积（米）"],
}
# 容积测量表的日期列没有表头，值形如「9.1」。
MEASURE_DATE_PATTERN = re.compile(r"^\d{1,2}\.\d{1,2}$")

PRICE_ROUTE_HEADERS = {"邮路名称", "路向"}
# 基础报表原版的标志性页签；这两个名字出现在工作簿里就是基础报表，不再往下猜。
BASE_REPORT_SHEETS = {"一干汇总", "二干基础"}
PRICE_BLANK_ROUTE_RUN = 10
PROFIT_GROUP_HEADERS = ["趟次(趟)", "装载量(袋件)", "每车平均装载量(袋件)", "运费（元）", "结算收入（元）", "盈亏值（元）", "平均每车盈亏（元）"]

# 在线测量表的两种日 Sheet 命名：草滩「4月1」或「4月1日」，港务「4.1」。
# 草滩实际导出不带「日」字，两种都要认。
DAILY_CAOTAN_SHEET = re.compile(r"^\d+月\d+日?$")
DAILY_GANGWU_SHEET = re.compile(r"^\d+\.\d+$")
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


def _row_width(ws, header_row: int, sample_rows: int = 5) -> int:
    """只读模式下 max_column 会跟着已读内容收缩，表头行末尾的空列会被截掉。

    以表头行和随后几行数据里最长的那条为准，保证日期列这类「无表头」列不被丢掉。
    """
    widest = 0
    for row in ws.iter_rows(min_row=header_row, max_row=header_row + sample_rows, values_only=True):
        widest = max(widest, len(row))
    return widest


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


def _pair_feature(ws) -> str | None:
    """用表内标题判断一干/二干。

    档位名不能用来区分：一干 2025 版模板也有「7.5米 45方 8吨」档，
    二干 2026 版价格表也有 40B 档。标题里的「一干 / 二干」才是可靠信号。
    """
    texts: list[str] = []
    for values in ws.iter_rows(min_row=1, max_row=3, values_only=True):
        for value in values:
            text = normalize_text(value)
            if text:
                texts.append(text)
    joined = " ".join(texts)
    has_first = "一干" in joined
    has_second = "二干" in joined
    if has_first and not has_second:
        return "first"
    if has_second and not has_first:
        return "second"
    return None


def _paired_role(path: Path, first_role: str, second_role: str, feature: str | None = None) -> str | None:
    """先按内容特征判断，判不出来再看文件名；都没有则返回 None，交给调用方报歧义。"""
    if feature == "first":
        return first_role
    if feature == "second":
        return second_role
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
    has_vehicle_columns = False
    for row, values in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        for column, value in enumerate(values, 1):
            text = normalize_text(value)
            if text in PRICE_ROUTE_HEADERS and route_column is None:
                route_column, header_row = column, row
            if ("公里" in text or "里程" in text) and distance_column is None:
                distance_column = column
            if "系数" in text or "吨" in text:
                has_tier_or_coefficient = True
            if text in {"车牌号", "车号", "派车单号", "封车时间", "运单号"}:
                has_vehicle_columns = True
    # 价格表是按路线报价，不出现单车字段；出现说明是报表，不是价格表。
    if route_column and distance_column and has_tier_or_coefficient and not has_vehicle_columns:
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
        values = tuple(values)
        if len(values) < len(expected):
            # 只读模式常把表头行末尾的空单元格截掉，补齐后再比，否则第 18~21 列永远对不上。
            values = values + (None,) * (len(expected) - len(values))
        actual = [normalize_text(value) for value in values[:len(expected)]]
        if actual == expected:
            return row, {f"column_{column}": column for column in range(1, len(expected) + 1)}
    return None, {}


def _platform_header(ws) -> tuple[int | None, dict[str, int]]:
    """车管平台汇总导出的表头。

    只在开头 5 行里找：平台导出的表头必然在表头行，而初版模板/报表的数据区
    深处也可能出现同样的列名组合，放宽行数会把它们误判成车管汇总。
    """
    try:
        return find_header_row(ws, VEHICLE_PLATFORM_FIELDS, max_rows=5)
    except ValueError:
        return None, {}


def _has_unheaded_date_column(ws, header_row: int, probe_rows: int = 20) -> bool:
    """表头有空列、且该列的值形如「9.1」——合并容积测量表的特征。"""
    header = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    empty_columns = [i for i, value in enumerate(header) if value in (None, "")]
    if not empty_columns:
        return False
    for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + probe_rows, values_only=True):
        for index in empty_columns:
            if index < len(row) and MEASURE_DATE_PATTERN.match(normalize_text(row[index])):
                return True
    return False


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
    # 基础报表原版的「一干汇总/二干基础」页签也含车管必需列，必须先按 Sheet 名判定，
    # 否则会被认成车管汇总导出，整份报表只剩两个页签可用。
    if any(normalize_text(name) in BASE_REPORT_SHEETS for name in wb.sheetnames):
        profile.role = "base_report"
        wb.close()
        return profile
    try:
        for ws in wb.worksheets:
            if hasattr(ws, "reset_dimensions") and ws.calculate_dimension() == "A1:A1":
                ws.reset_dimensions()
            # 21 列初版模板必须最先判：它的数据区里也可能出现一段
            # 「邮路名称/车牌号/日期」表头，晚判就会把模板认成车管汇总导出。
            row, fields = _review_header(ws)
            if row:
                profile.role, profile.header_row, profile.field_map = "draft_template", row, fields
                return profile
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
            # 车管平台「汇总导出」：必备列比逐家公司导出少三个
            # （结算吨位 / 往返属性 / 邮路往返属性），必须单独识别。
            row, fields = _platform_header(ws)
            if row:
                profile.role, profile.header_row, profile.field_map = "vehicle_platform", row, fields
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
                daily_caotan = sum(1 for name in wb.sheetnames if DAILY_CAOTAN_SHEET.match(normalize_text(name)))
                daily_gangwu = sum(1 for name in wb.sheetnames if DAILY_GANGWU_SHEET.match(normalize_text(name)))
                # 合并后的容积测量表：没有「日」页签，日期放在一列无表头的单元格里。
                # 先判它，否则会被当成「车型表」忽略掉。
                if not daily_caotan and not daily_gangwu and _has_unheaded_date_column(ws, row):
                    profile.role, profile.header_row, profile.field_map = "online_measure", row, fields
                    profile.site = "合并"
                    return profile
                # 先看日 Sheet 命名：草滩「8月1日」，港务「8.1」。
                # 文件名只在日 Sheet 无法区分时兜底。
                if daily_gangwu > daily_caotan:
                    role = "online_gangwu"
                elif daily_caotan > daily_gangwu:
                    role = "online_caotan"
                elif "车型" in lowered and "在线" not in lowered:
                    role = "model"
                elif "港务" in lowered:
                    role = "online_gangwu"
                elif "草滩" in lowered or "在线" in lowered:
                    role = "online_caotan"
                elif daily_caotan + daily_gangwu:
                    role = "online_caotan"
                else:
                    role = "model"
                profile.role, profile.header_row, profile.field_map = role, row, fields
                profile.site = "港务" if role == "online_gangwu" else ("草滩" if role == "online_caotan" else None)
                return profile
            row, fields = _price_header(ws)
            if row:
                price_role = _paired_role(path, "first_price", "second_price", _pair_feature(ws))
                profile.role = price_role or "price_candidate"
                profile.header_row, profile.field_map = row, fields
                return profile
            row = _profit_group_header(ws)
            if row:
                profit_role = _paired_role(
                    path, "first_profit_template", "second_profit_template", _pair_feature(ws)
                )
                profile.role = profit_role or "profit_candidate"
                profile.header_row, profile.field_map = row, {"profit_group": 2}
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


def _rounded_volume(value: Any) -> int | None:
    try:
        number = Decimal(str(value).replace(",", ""))
    except (InvalidOperation, ValueError):
        return None
    return int(number.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _record_from_model_row(row: Sequence[Any], field_map: Mapping[str, int], source_row: int, source_sheet: str) -> dict[str, Any]:
    length, width, height = (_to_number(_value(row, field_map, key)) for key in ("length", "width", "height"))
    raw_volume = _value(row, field_map, "volume")
    numeric_volume = _to_number(raw_volume)
    volume_formula = raw_volume if isinstance(raw_volume, str) and raw_volume.lstrip().startswith("=") else None
    unrounded_volume: Decimal | None = None
    volume_method = "missing"
    if numeric_volume is not None:
        unrounded_volume = Decimal(str(numeric_volume))
        volume_method = "source_value_rounded"
    elif all(item is not None for item in (length, width, height)):
        unrounded_volume = Decimal(str(length)) * Decimal(str(width)) * Decimal(str(height))
        volume_method = "dimensions_formula_fallback" if volume_formula else "dimensions_fallback"
    volume = _rounded_volume(unrounded_volume)
    volume_column = field_map.get("volume")
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
        "volume_unrounded": float(unrounded_volume) if unrounded_volume is not None else None,
        "volume_method": volume_method,
        "volume_formula": volume_formula,
        "volume_cell": f"{get_column_letter(volume_column)}{source_row}" if volume_column else None,
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
                        "round_trip_attribute": normalize_text(_value(row, field_map, "round_trip")),
                        "route_direction": normalize_text(_value(row, field_map, "route_direction")),
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
    if isinstance(value, str):
        # 2026 版一干表把里程写成「1088公里」，前面是数就是里程数。
        match = re.match(r"\s*(-?\d+(?:\.\d+)?)\s*(?:公里|km|KM)?\s*$", value)
        if match:
            return float(match.group(1))
    return None


def _extract_first_number(value: Any) -> float | None:
    """容间下限。两种写法都要认：

    2024 版价格表写「≥55立方米」，2026 版二干合同写「55m³≤R」，
    含义相同，都是「容积下限 55」。
    """
    text = normalize_text(value)
    match = re.search(r"(?:≥|>=)\s*(\d+(?:\.\d+)?)", text)
    if match:
        return float(match.group(1))
    match = re.search(r"(\d+(?:\.\d+)?)\s*(?:m3|m³|立方米|方)?\s*(?:≤|<=)", text)
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


def _price_header_positions(
    ws,
) -> tuple[dict[str, int] | None, dict[str, int] | None, list[int]]:
    """返回 (路线列, 里程列, 其余路线列)。

    2026 版一干表同时给了全称「西安市-北京市」和简称「西安-北京」两列，
    两者都要能当路线名用，否则车管里的「西安-北京」匹配不上价格表。
    """
    for row in range(1, min(ws.max_row, 30) + 1):
        values = [ws.cell(row, column).value for column in range(1, ws.max_column + 1)]
        route_columns = [
            column
            for column, value in enumerate(values, 1)
            if value is not None and _header_key(value) in {_header_key(item) for item in PRICE_ROUTE_HEADERS}
        ]
        distance_column = _find_header_column(values, ("路向里程", "公里（KM）", "公里", "里程"))
        if route_columns and distance_column:
            return (
                {"row": row, "column": route_columns[0]},
                {"row": row, "column": distance_column},
                {"row": row, "columns": route_columns[1:]},
            )
    return None, None, {"row": 0, "columns": []}


# 2026 版一干表把档位改成每 5m³ 一档：「[55,60)」「[60.65)」（表里逗号有时候写成点）。
_PRICE_BAND_RE = re.compile(r"[\[\(]\s*(\d+(?:\.\d+)?)\s*[,.]\s*(\d+(?:\.\d+)?)\s*[\)\]]")


def _band_range(value: Any) -> tuple[float, float] | None:
    text = normalize_text(value)
    match = _PRICE_BAND_RE.search(text)
    if not match:
        return None
    low, high = float(match.group(1)), float(match.group(2))
    return (low, high) if high > low else None


def _is_banded_sheet(ws, header_row: int, distance_column: int) -> bool:
    """档位头写成 [55,60) 这种容积区间，就是 2026 版「每 5m³ 一档」价目表。"""
    found = 0
    for column in range(distance_column + 1, ws.max_column + 1):
        for row in range(header_row, min(ws.max_row, header_row + 4) + 1):
            if _band_range(ws.cell(row, column).value):
                found += 1
                break
    return found >= 2


def _row_with_label(ws, header_row: int, label: str) -> int | None:
    """找写着「吨位」这一类标题的行（它可能在里程列右边，如 2026 版一干表）。"""
    for row in range(header_row, min(ws.max_row, header_row + 5) + 1):
        for column in range(1, ws.max_column + 1):
            if _header_key(ws.cell(row, column).value) == label:
                return row
    return None


def _first_data_row(ws, header_row: int, route_column: int) -> int:
    for row in range(header_row + 1, ws.max_row + 1):
        text = normalize_text(ws.cell(row, route_column).value)
        if text and text not in PRICE_ROUTE_HEADERS:
            return row
    return header_row + 1


def _banded_tier_headers(
    ws,
    header_row: int,
    route_column: int,
    distance_column: int,
    cached_rows: dict[int, tuple] | None = None,
) -> tuple[list[dict[str, Any]], int] | None:
    """解析「每 5m³ 一档」的价目表。

    档位信息分散在四行：吨位区间（如 12-13）、方量区间（如 [60,65)）、
    档位下限数字、系数。系数是数据行上方最后一个数字，表里不写「系数」二字。
    吨位取该档区间的上限（[60,65) → 13），与表头「12吨（≥55m³）」的写法一致。
    """
    if not _is_banded_sheet(ws, header_row, distance_column):
        return None
    tonnage_row = _row_with_label(ws, header_row, "吨位")
    data_start = _first_data_row(ws, header_row, route_column)
    tiers: list[dict[str, Any]] = []
    last_row = min(ws.max_row, header_row + 5)
    for column in range(distance_column + 1, ws.max_column + 1):
        band = None
        for row in range(header_row, last_row + 1):
            band = _band_range(ws.cell(row, column).value)
            if band:
                break
        if band:
            min_volume, max_volume = band
        else:
            # 往返页前三档只写了「5吨（≥35m³）」，没写方量区间。
            min_volume = _extract_first_number(ws.cell(header_row, column).value)
            max_volume = None
            if min_volume is None:
                continue
        label = ""
        if tonnage_row:
            label = normalize_text(ws.cell(tonnage_row, column).value)
        match = re.search(r"(\d+(?:\.\d+)?)\s*(?:吨)?\s*$", label) or re.search(
            r"(\d+(?:\.\d+)?)", label
        )
        key = tonnage_key(match.group(1)) if match else ""
        if not key:
            key = tonnage_key(min_volume)
        coefficient = None
        coefficient_row = None
        for row in range(data_start - 1, header_row, -1):
            # 系数列常写成公式（=0.65+(M5-55)/5*0.024），公式本身取不到数值，
            # 所以还要看 data_only 读出来的缓存值。
            number = _number_from_cell(ws, row, column)
            if number is None and cached_rows:
                cached = cached_rows.get(row) or ()
                value = cached[column - 1] if column - 1 < len(cached) else None
                number = _to_number(value)
            if number is not None and 0 < float(number) <= 2:
                coefficient = float(number)
                coefficient_row = row
                break
        if coefficient is None:
            continue
        tiers.append({
            "key": key,
            "column": column,
            "tonnage_row": tonnage_row,
            "threshold_row": None,
            "coefficient_row": coefficient_row,
            "min_volume": min_volume,
            "max_volume": max_volume,
            "coefficient": coefficient,
        })
    if not tiers:
        return None
    # 没写上限的档位（往返页前三档）用下一档的下限补齐；最后一档是开区间。
    for index, tier in enumerate(tiers):
        if tier.get("max_volume") is not None:
            continue
        following = tiers[index + 1] if index + 1 < len(tiers) else None
        tier["max_volume"] = following["min_volume"] if following else None
    return tiers, data_start


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


def _special_price_columns(ws, header_row: int, distance_column: int) -> dict[int, tuple[str, str]]:
    """Find 返程 / 往返 price columns that sit outside the normal tier grid.

    二干价目表把返程、往返各放一列（都只写 12 吨价），列头上方标着
    「返程」「往返」，和常规的「12吨/20吨…」档位列不在一起。
    返回 {列号: (标签, 吨位键)}。
    """
    result: dict[int, tuple[str, str]] = {}
    for column in range(distance_column + 1, ws.max_column + 1):
        label = None
        for row in range(max(1, header_row - 2), header_row + 1):
            text = normalize_text(ws.cell(row, column).value)
            if text in {"往返", "返程"}:
                label = text
                break
        if not label:
            continue
        key = None
        for row in range(header_row, min(ws.max_row, header_row + 2) + 1):
            if "吨" in normalize_text(ws.cell(row, column).value):
                key = tonnage_key(ws.cell(row, column).value)
                break
        if key:
            result[column] = (label, key)
    return result


def read_price_catalog(path: Path | str, kind: str) -> dict[str, Any]:
    wb = load_workbook(path, read_only=False, data_only=False, keep_links=False)
    # 档位系数常是公式，再读一份缓存值用于还原系数。
    # 用只读模式，避免把整张价目表再占一份内存（客户价目表最大近 200 列 × 上百行）。
    try:
        cached_wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    except Exception:
        cached_wb = None
    catalog: dict[str, Any] = {"kind": kind, "tiers": [], "routes": {}, "sheets": []}
    tier_definitions: dict[str, dict[str, Any]] = {}
    try:
        for ws in wb.worksheets:
            if ws.max_row < 4 or ws.max_column < 4:
                continue
            route_header, distance_header, alt_route_header = _price_header_positions(ws)
            if not route_header or not distance_header:
                continue
            header_row = route_header["row"]
            route_column = route_header["column"]
            alt_route_columns = [
                column for column in alt_route_header.get("columns", []) if column
            ]
            distance_column = distance_header["column"]
            # 「往返」是独立 Sheet（一干价目表）：整页价格都是往返价，
            # 不能覆盖同路线在「单程」页的价格。
            sheet_round_trip = "往返" in normalize_text(ws.title)
            special_columns = _special_price_columns(ws, header_row, distance_column)
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
            cached_rows: dict[int, tuple] = {}
            if cached_wb is not None and ws.title in cached_wb.sheetnames:
                cached_sheet = cached_wb[ws.title]
                for row in cached_sheet.iter_rows(
                    min_row=header_row, max_row=header_row + 6, values_only=True
                ):
                    cached_rows[header_row + len(cached_rows)] = row
            banded = _banded_tier_headers(
                ws, header_row, route_column, distance_column, cached_rows
            )
            if banded:
                tier_headers, data_start = banded
                for tier_header in tier_headers:
                    key = tier_header["key"]
                    if key in seen_keys:
                        continue
                    seen_keys.add(key)
                    tier_definitions.setdefault(key, {
                        "key": key,
                        "min_volume": tier_header["min_volume"],
                        "max_volume": tier_header["max_volume"],
                        "coefficient": tier_header["coefficient"],
                    })
            else:
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
                "round_trip": sheet_round_trip,
                "route_header": route_header,
                "distance_header": distance_header,
                "tier_headers": [
                    {key: item.get(key) for key in ("key", "column", "tonnage_row", "threshold_row", "coefficient_row")}
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
                # 全称列（西安市-北京市）和简称列（西安-北京）都是合法路线名。
                alt_routes = [
                    normalize_text(ws.cell(row, column).value) for column in alt_route_columns
                ]
                distance = _number_from_cell(ws, row, distance_column)
                prices = {
                    item["key"]: _number_from_cell(ws, row, item["column"])
                    for item in tier_headers
                }
                prices = {key: value for key, value in prices.items() if value is not None}
                round_trip_prices: dict[str, Any] = {}
                return_prices: dict[str, Any] = {}
                for column, (label, special_key) in special_columns.items():
                    value = _number_from_cell(ws, row, column)
                    if value is None:
                        continue
                    if label == "往返":
                        round_trip_prices[special_key] = value
                    else:
                        return_prices[special_key] = value
                if not prices and not round_trip_prices:
                    continue
                if sheet_round_trip:
                    entry = catalog["routes"].get(route)
                    if entry is None:
                        entry = catalog["routes"][route] = {
                            "distance": distance,
                            "prices": {},
                            "sheet": ws.title,
                            "kind": kind,
                        }
                    entry.setdefault("round_trip_prices", {}).update(prices)
                    for name in alt_routes:
                        if name and name != route:
                            catalog["routes"].setdefault(name, entry)
                    continue
                entry = catalog["routes"].get(route)
                if entry is None:
                    entry = catalog["routes"][route] = {
                        "distance": distance,
                        "prices": prices,
                        "sheet": ws.title,
                        "kind": kind,
                    }
                elif not entry.get("prices"):
                    entry["prices"] = prices
                if round_trip_prices:
                    entry.setdefault("round_trip_prices", {}).update(round_trip_prices)
                if return_prices:
                    entry.setdefault("return_prices", {}).update(return_prices)
                for name in alt_routes:
                    if name and name != route:
                        catalog["routes"].setdefault(name, entry)
    finally:
        wb.close()
        if cached_wb is not None:
            cached_wb.close()
    catalog["tiers"] = sorted(tier_definitions.values(), key=lambda item: (item["min_volume"], item["key"]))
    _merge_route_supplement(catalog)
    return catalog


SUPPLEMENT_ASSET = "route-supplement.json"


def _merge_route_supplement(catalog: dict[str, Any], assets_dir: Path | None = None) -> None:
    """合并合同外专线的补充价目。

    兰州、平凉、陇南、洛川港、兴平港、泾阳港、庆阳港这几条不在两张合同价目表里，
    但每月都要结算。价格取自官方基础报表逐单反查（见 assets/route-supplement.json），
    每个档位在当月内取值唯一，可以直接当固定价用。

    只补「价格表里完全不存在」的路线；价格表里已有的路线一律以合同价为准，
    避免补充值覆盖合同调价后的新价。
    """
    root = assets_dir or Path(__file__).resolve().parents[2] / "assets"
    path = root / SUPPLEMENT_ASSET
    if not path.is_file():
        return
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return
    routes = data.get("routes") if isinstance(data, dict) else None
    if not isinstance(routes, dict):
        return
    tiers = catalog.get("tiers") or []
    for route, entry in routes.items():
        name = normalize_text(route)
        if not name or name in catalog["routes"] or not isinstance(entry, dict):
            continue
        prices = entry.get("prices")
        if not isinstance(prices, dict) or not prices:
            continue
        cleaned = {}
        for key, value in prices.items():
            number = _to_number(value)
            if number is not None:
                cleaned[tonnage_key(key)] = number
        if not cleaned:
            continue
        # 官方专线按固定容间区间定档，区间下限直接带进目录，供 _supplement_tier 使用。
        thresholds = {
            tonnage_key(key): _to_number(value)
            for key, value in (entry.get("volume_thresholds") or {}).items()
            if _to_number(value) is not None
        }
        catalog["routes"][name] = {
            "distance": _to_number(entry.get("distance")),
            "prices": cleaned,
            "standard_price": _to_number(entry.get("standard_price")),
            "volume_thresholds": thresholds,
            "sheet": "补充价目",
            "kind": catalog.get("kind"),
            "supplement": True,
        }
    # 补充价目的档位（3/5/12/20/30/40A）只用于它自己的路线查价，
    # 绝不能插进主档位表：一干主表是每 5 方一档，插进一个 min_volume=3 的档
    # 会打乱整个阶梯，让所有路线都选错档。


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


def read_vehicle_platform_rows(path: Path | str) -> list[dict[str, Any]]:
    """读取车管平台「汇总导出」（总/一干/二干 页签）。

    平台导出和逐家公司导出是同一种业务表，但少了三个字段：
    结算吨位、往返属性、邮路往返属性。价格一律按价格表重算，
    平台自带的「运价/结算金额/系数」不参与计价——它的口径与官方报表不一致，
    而且同一邮路同一天多趟时结算金额会重复记在每一行上。

    优先取「总」页签（含一干+二干全集）；没有「总」时合并一干、二干，
    避免同一趟在两个页签里各出现一次。
    """
    path = Path(path)
    wb = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    result: list[dict[str, Any]] = []
    try:
        sheets = _platform_sheet_order(wb.sheetnames)
        for sheet_name in sheets:
            ws = wb[sheet_name]
            if hasattr(ws, "reset_dimensions"):
                ws.reset_dimensions()
            header_row, field_map = _try_header(ws, VEHICLE_PLATFORM_FIELDS)
            if not header_row:
                continue
            header_values = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
            for key, aliases in VEHICLE_PLATFORM_OPTIONAL.items():
                column = _find_header_column(header_values, aliases)
                if column:
                    field_map[key] = column
            for source_row, row in enumerate(
                ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1
            ):
                plate = normalize_plate(_value(row, field_map, "plate"))
                if not plate:
                    continue
                execution_datetime = parse_date(_value(row, field_map, "date"))
                if execution_datetime is None:
                    continue
                company = normalize_text(_value(row, field_map, "company"))
                result.append({
                    "source_file": path.name,
                    "source_path": str(path.resolve()),
                    "source_sheet": ws.title,
                    "source_row": source_row,
                    "company": company,
                    "route": _value(row, field_map, "route"),
                    "date": _value(row, field_map, "date"),
                    "execution_date": execution_datetime.date(),
                    "dispatch_no": normalize_dispatch_no(_value(row, field_map, "dispatch_no")),
                    "plate": plate,
                    "trailer": normalize_trailer(_value(row, field_map, "trailer")),
                    "plate_key": plate,
                    "trailer_key": normalize_trailer(_value(row, field_map, "trailer")),
                    "transport_volume": _to_number(_value(row, field_map, "transport_volume")),
                    # 平台不提供结算吨位；留空让计价完全按价格表走。
                    "settlement_tonnage": None,
                    "round_trip_attribute": "",
                    "route_direction": "",
                    "platform_level": normalize_text(_value(row, field_map, "level")),
                    "platform_capacity": _to_number(_value(row, field_map, "capacity")),
                })
    finally:
        wb.close()
    return result


def _platform_sheet_order(sheet_names: Sequence[str]) -> list[str]:
    """「总」优先；没有「总」才用一干/二干，防止同一趟被两个页签重复读入。"""
    normalized = {normalize_text(name): name for name in sheet_names}
    if "总" in normalized:
        return [normalized["总"]]
    selected = [normalized[name] for name in ("一干", "二干") if name in normalized]
    return selected or list(sheet_names)


def read_measure_rows(
    path: Path | str,
    target_month: tuple[int, int],
) -> list[dict[str, Any]]:
    """读取合并后的车辆容积测量表（草滩 + 港务一并导出）。

    与逐场地在线表的两点差异：
      1. 没有「日」页签，日期是**无表头的一列**，值形如 9.1；
      2. 同一张表里混着两个场地、上百条邮路，所以场地记为「合并」。

    日期列有一个必须处理的陷阱：Excel 把日期存成数字，9.10 与 9.1 的值完全相同，
    所以 10/20/30 号会退化成 9.1/9.2/9.3，和 1/2/3 号混在一起。这里按
    「1~3 号同时也是 10/20/30」的歧义口径处理，把两条日期都产出，
    由下游「日期+车牌」精确匹配后再按容间接近程度择优——直接把 9.1 当成 1 号
    会让 10 号那天的趟次永远配不上测量值。
    """
    path = Path(path)
    wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    result: list[dict[str, Any]] = []
    year = target_month[0]
    try:
        for source_sheet_order, ws in enumerate(wb.worksheets):
            if hasattr(ws, "reset_dimensions"):
                ws.reset_dimensions()
            header_row, field_map = _try_header(ws, MEASURE_FIELDS)
            if not header_row:
                continue
            # 只读模式下 ws.max_column 会按已读内容收缩，_try_header 读完之后
            # 表头末尾的 None 列（测量表的日期列）就看不到了。
            # 用表头行 + 首行数据里较长的那条来展开列数，否则日期列会被截掉。
            header_values = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
            probe = next(ws.iter_rows(min_row=header_row + 1, max_row=header_row + 1, values_only=True), ())
            if len(probe) > len(header_values):
                header_values = tuple(header_values) + (None,) * (len(probe) - len(header_values))
            for key, aliases in MEASURE_OPTIONAL.items():
                column = _find_header_column(header_values, aliases)
                if column:
                    field_map[key] = column
            date_column = _find_measure_date_column(ws, header_row, header_values, target_month)
            for source_row, row in enumerate(
                ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1
            ):
                record = _record_from_model_row(row, field_map, source_row, ws.title)
                if not record["plate_key"]:
                    continue
                raw_date = row[date_column - 1] if date_column and date_column <= len(row) else None
                for measurement_date in _measure_dates(raw_date, year, target_month[1]):
                    item = dict(record)
                    item["site"] = "合并"
                    item["measurement_date"] = measurement_date
                    item["source_file"] = path.name
                    item["source_path"] = str(path.resolve())
                    item["source_sheet_order"] = source_sheet_order
                    item["source_order"] = len(result)
                    result.append(item)
    finally:
        wb.close()
    return result


def _find_measure_date_column(ws, header_row: int, header_values: Sequence[Any], target_month: tuple[int, int]) -> int | None:
    """容积测量表的日期列没有表头，靠「值形如 9.1」来定位。"""
    month = target_month[1]
    best: tuple[int, int] | None = None
    for column, value in enumerate(header_values, 1):
        if value not in (None, ""):
            continue
        hits = 0
        for row in ws.iter_rows(min_row=header_row + 1, max_row=header_row + 40, values_only=True):
            cell = row[column - 1] if column <= len(row) else None
            if _measure_dates(cell, target_month[0], month):
                hits += 1
        if hits and (best is None or hits > best[1]):
            best = (column, hits)
    return best[0] if best else None


def _measure_dates(value: Any, year: int, month: int) -> list[date]:
    """把「9.1」这类无表头日期换算成日期；1/2/3 号同时产出 10/20/30 号的候选。

    必须走 Decimal 的字符串形式取小数位，不能用 fraction * 100：
    0.1 * 100 == 10.000000000000002，四舍五入会把 9.1 算成 10 号，
    于是 1 号的测量记录整体丢掉、10 号的反而翻倍。
    """
    if value is None or value == "":
        return []
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError):
        return []
    if not number.is_finite():
        return []
    whole = int(number)
    text = format(number, "f")
    _, _, fraction = text.partition(".")
    fraction = fraction.rstrip("0")
    if not fraction:
        return []
    day = int(fraction)
    if whole != month or not 1 <= day <= 31:
        return []
    days = {day}
    if day in (1, 2, 3):
        # 9.10 / 9.20 / 9.30 在浮点里等于 9.1 / 9.2 / 9.3，无法区分，两个都保留。
        days.add(day * 10)
    result: list[date] = []
    for candidate in sorted(days):
        try:
            result.append(date(year, month, candidate))
        except ValueError:
            continue
    return result
