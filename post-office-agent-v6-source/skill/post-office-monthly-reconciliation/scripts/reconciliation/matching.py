from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Iterable, Mapping

from .normalize import normalize_plate, normalize_text, route_candidates, tonnage_key


@dataclass(frozen=True)
class Issue:
    code: str
    severity: str
    message: str
    action: str


@dataclass
class MatchResult:
    record: dict[str, Any] | None
    issues: list[Issue] = field(default_factory=list)
    candidates: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class PriceMatch:
    kind: str | None = None
    catalog_kind: str | None = None
    route: str | None = None
    tonnage: str | None = None
    distance: float | int | None = None
    coefficient: float | int | None = None
    standard_price: float | int | None = None
    freight: float | int | None = None
    model_tonnage: str | None = None
    settlement_tonnage: float | int | None = None
    original_coefficient: float | int | None = None
    original_freight: float | int | None = None
    price_difference: float | int | None = None
    downgraded: bool = False
    issues: list[Issue] = field(default_factory=list)


def select_online(candidates: Iterable[dict[str, Any]]) -> MatchResult:
    candidates = list(candidates)
    if not candidates:
        return MatchResult(None, [Issue("online_missing", "red", "在线测量表未找到车辆", "核实车牌和挂车并人工补充容间")])
    order = lambda item: (item.get("measurement_date") or datetime.min, int(item.get("source_order") or 0))
    nonzero = [item for item in candidates if isinstance(item.get("volume"), (int, float)) and item["volume"] > 0]
    if nonzero:
        return MatchResult(max(nonzero, key=order), candidates=candidates)
    selected = max(candidates, key=order)
    all_zero = all(item.get("volume") in (0, 0.0, None, "") for item in candidates)
    dimensions_empty = all(all(item.get(key) in (None, "", 0, 0.0) for key in ("length", "width", "height")) for item in candidates)
    if all_zero and dimensions_empty:
        return MatchResult(selected, [Issue("online_unmeasured", "orange", "在线记录确认未测量，容间为0且长宽高为空", "人工确认是否保持0")], candidates)
    return MatchResult(selected, [Issue("online_zero_inconsistent", "red", "在线容间为0但尺寸信息不一致", "回查原始测量表")], candidates)


def _date_key(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    return value if isinstance(value, date) else None


def index_model_summary(rows: Iterable[dict[str, Any]]) -> dict[tuple[date, str], dict[str, Any]]:
    """Index selected online rows by their only permitted match key."""
    result: dict[tuple[date, str], dict[str, Any]] = {}
    for row in rows:
        measurement_date = _date_key(row.get("measurement_date"))
        plate = normalize_plate(row.get("plate"))
        if measurement_date is not None and plate:
            result[(measurement_date, plate)] = row
    return result


def match_vehicle_row(
    vehicle: Mapping[str, Any],
    index: Mapping[tuple[date, str], dict[str, Any]],
) -> dict[str, Any] | None:
    """Match by date and plate, then use the latest non-zero same-plate record."""
    execution_date = _date_key(vehicle.get("execution_date"))
    plate = normalize_plate(vehicle.get("plate"))
    if execution_date is None or not plate:
        return None
    result = index.get((execution_date, plate))
    if result is not None:
        return result
    plate_rows = [(d, row) for (d, p), row in index.items() if p == plate]
    if plate_rows:
        nonzero = [
            (measurement_date, row)
            for measurement_date, row in plate_rows
            if isinstance(row.get("measured_volume"), (int, float)) and row["measured_volume"] > 0
        ]
        _, selected = max(nonzero or plate_rows, key=lambda item: item[0])
        # Never annotate the shared summary row in-place. The same row can be an
        # exact match for a later vehicle record and must not inherit this flag.
        sel = dict(selected)
        sel["_cross_station"] = True
        return sel
    return None


def _alias_entry(route: str, aliases: Mapping[str, Any]) -> Any:
    normalized = normalize_text(route)
    for key, entry in aliases.items():
        if normalize_text(key) == normalized:
            return entry
    return None


def _alias_kind(route: str, aliases: Mapping[str, Any]) -> str | None:
    entry = _alias_entry(route, aliases)
    if isinstance(entry, Mapping):
        value = normalize_text(entry.get("price_kind"))
        return value or None
    return None


def _alias_route(route: str, aliases: Mapping[str, Any]) -> str | None:
    entry = _alias_entry(route, aliases)
    if isinstance(entry, str):
        return normalize_text(entry) or None
    if isinstance(entry, Mapping):
        return normalize_text(entry.get("standard_route")) or None
    return None


def _alias_price_source(route: str, aliases: Mapping[str, Any]) -> str | None:
    entry = _alias_entry(route, aliases)
    if isinstance(entry, Mapping):
        value = normalize_text(entry.get("price_source"))
        return value or None
    return None


def _baseline_price(
    route_info: Mapping[str, Any],
    catalog: Mapping[str, Any],
    fallback: Any,
    key: str = "",
) -> Any:
    """Return the contract base price (the tier whose coefficient is 1.0).

    Price-table columns already carry the coefficient: e.g. "系数0.63 /
    12吨" holds 30吨价×0.63. The official report therefore prints that base
    price in 标准价 and the selected tier's cell value in 运价.
    """
    prices = route_info.get("prices", {})
    for tier in catalog.get("tiers", []):
        try:
            if Decimal(str(tier.get("coefficient"))) != Decimal(1):
                continue
        except Exception:
            continue
        value = prices.get(tonnage_key(tier.get("key")))
        if value not in (None, ""):
            return value
        return fallback
    # 2026 版一干表每 5m³ 一档，没有系数正好等于 1 的档位。
    # 官方「标准价」是系数 1 对应的合同价，用「运价 ÷ 系数」还原。
    return _base_from_coefficient(fallback, _tier_coefficient(catalog, key))


def _base_from_coefficient(value: Any, coefficient: Decimal) -> Any:
    if value in (None, "") or not coefficient:
        return value
    try:
        base = (Decimal(str(value)) / coefficient).quantize(Decimal("0.01"))
    except Exception:
        return value
    return float(base) if base % 1 else int(base)


def _round_trip_price(
    prices: Mapping[str, Any],
    catalog: Mapping[str, Any],
    key: str,
    coefficient: Any,
) -> tuple[Any, Decimal] | None:
    """Return (标准价, 运价) from a 往返 price column.

    一干价目表的「往返」Sheet 写全了各档位，直接取所在档位单元格。
    二干价目表只写 12 吨往返价，其余档位按「12 吨价 × 系数 / 12 吨系数」
    折算，标准价仍是 12 吨往返价（与官方表一致）。
    """
    base_key: str | None = None
    base_coefficient = Decimal("1")
    for tier in catalog.get("tiers", []):
        try:
            if Decimal(str(tier.get("coefficient"))) != Decimal(1):
                continue
        except Exception:
            continue
        base_key = tonnage_key(tier.get("key"))
        break
    base_value = prices.get(base_key) if base_key else None
    if base_value in (None, ""):
        for candidate_key, candidate_value in prices.items():
            if candidate_value not in (None, ""):
                base_key, base_value = candidate_key, candidate_value
                base_coefficient = _tier_coefficient(catalog, base_key)
                break
    if base_value in (None, ""):
        return None
    try:
        base_decimal = Decimal(str(base_value))
    except Exception:
        return None
    direct = prices.get(key)
    if direct not in (None, ""):
        freight = Decimal(str(direct))
        # 2026 版一干表没有系数 1 的档位，标准价要用本档「运价 ÷ 系数」还原，
        # 不能拿第一档（2.75 吨）的价当基准。
        if base_key != key:
            standard = _base_from_coefficient(direct, _tier_coefficient(catalog, key))
        else:
            standard = direct
        return standard, freight
    try:
        # 二干只给 12 吨往返价，其余档位由系数折算；官方保留两位小数。
        freight = (base_decimal * Decimal(str(coefficient)) / base_coefficient).quantize(Decimal("0.01"))
    except Exception:
        return None
    return base_value, freight


def _tier_coefficient(catalog: Mapping[str, Any], key: str) -> Decimal:
    for tier in catalog.get("tiers", []):
        if tonnage_key(tier.get("key")) == key:
            try:
                return Decimal(str(tier.get("coefficient")))
            except Exception:
                break
    return Decimal("1")


def _select_tier(volume_number: Decimal, catalog: Mapping[str, Any]) -> Mapping[str, Any] | None:
    eligible: list[tuple[Decimal, Mapping[str, Any]]] = []
    for tier in catalog.get("tiers", []):
        try:
            minimum = Decimal(str(tier["min_volume"]))
        except Exception:
            continue
        # 2026 版一干表按「[55,60)」这种容积区间分档，上限是开区间；
        # 老表只有下限（≥55），没有 max_volume。
        maximum = tier.get("max_volume")
        if maximum is not None:
            try:
                if volume_number >= Decimal(str(maximum)):
                    continue
            except Exception:
                pass
        if volume_number >= minimum:
            eligible.append((minimum, tier))
    return max(eligible, key=lambda item: item[0])[1] if eligible else None


def match_price(
    route: Any,
    volume: Any,
    catalogs: Mapping[str, Mapping[str, Any]],
    aliases: Mapping[str, Any],
    *,
    round_trip: bool = False,
) -> PriceMatch:
    try:
        volume_number = Decimal(str(volume))
    except Exception:
        # Keep resolving route kind and distance even when the model summary
        # has no volume. The official special sheets still classify such rows;
        # only tonnage, standard price and freight remain blank.
        volume_number = None
    candidates = route_candidates(route, aliases)
    resolved_kind: str | None = None
    catalog_kind: str | None = None
    resolved_route: str | None = None
    for candidate in candidates:
        preferred = _alias_kind(candidate, aliases)
        if preferred in catalogs:
            resolved_kind = preferred
            catalog_kind = _alias_price_source(candidate, aliases) or preferred
            if catalog_kind not in catalogs:
                catalog_kind = preferred
            resolved_route = _alias_route(candidate, aliases) or candidate
            break
        matching_kinds = [kind for kind, catalog in catalogs.items() if candidate in catalog.get("routes", {})]
        if matching_kinds:
            resolved_kind = matching_kinds[0]
            catalog_kind = resolved_kind
            resolved_route = candidate
            break
    if not resolved_kind:
        return PriceMatch(issues=[Issue("price_route_uncovered", "red", "路线未被一干或二干价格表覆盖且无法识别干线类型", "维护路线别名或价格资产")])

    catalog_kind = catalog_kind or resolved_kind
    catalog = catalogs[catalog_kind]
    tier = _select_tier(volume_number, catalog) if volume_number is not None else None
    route_info = None
    for candidate in ([resolved_route] if resolved_route else []) + candidates:
        route_info = catalog.get("routes", {}).get(candidate)
        if route_info:
            resolved_route = candidate
            break
    if volume_number is None:
        issues = [Issue("volume_missing", "red", "无法用容间判定吨位", "人工补充有效容间")]
        if not route_info:
            issues.append(Issue("price_route_uncovered", "red", "路线未被对应干线价格表覆盖", "维护路线别名或价格资产"))
        return PriceMatch(
            kind=resolved_kind,
            catalog_kind=catalog_kind,
            route=resolved_route,
            distance=route_info.get("distance") if route_info else None,
            issues=issues,
        )
    if not tier:
        issues = [Issue("tonnage_below_minimum", "red", "容间低于价格表最低吨位档", "人工确认吨位和价格")]
        if not route_info:
            issues.append(Issue("price_route_uncovered", "red", "路线未被对应干线价格表覆盖", "维护路线别名或价格资产"))
        return PriceMatch(
            kind=resolved_kind,
            catalog_kind=catalog_kind,
            route=resolved_route,
            distance=route_info.get("distance") if route_info else None,
            issues=issues,
        )

    key = tonnage_key(tier.get("key"))
    coefficient = tier.get("coefficient")
    if not route_info:
        return PriceMatch(
            kind=resolved_kind,
            catalog_kind=catalog_kind,
            route=resolved_route,
            tonnage=key,
            model_tonnage=key,
            coefficient=coefficient,
            issues=[Issue("price_route_uncovered", "red", "路线未被对应干线价格表覆盖", "维护路线别名或价格资产")],
        )
    round_trip_prices = route_info.get("round_trip_prices") or {}
    if round_trip and round_trip_prices:
        priced = _round_trip_price(round_trip_prices, catalog, key, coefficient)
        if priced is None:
            return PriceMatch(
                kind=resolved_kind,
                catalog_kind=catalog_kind,
                route=resolved_route,
                tonnage=key,
                model_tonnage=key,
                distance=route_info.get("distance"),
                coefficient=coefficient,
                issues=[Issue("price_missing", "red", "往返价目缺失", "维护价格表往返价或人工补充")],
            )
        standard, freight = priced
        return PriceMatch(
            kind=resolved_kind,
            catalog_kind=catalog_kind,
            route=resolved_route,
            tonnage=key,
            model_tonnage=key,
            distance=route_info.get("distance"),
            coefficient=coefficient,
            standard_price=standard,
            freight=float(freight) if freight % 1 else int(freight),
        )
    tier_price = route_info.get("prices", {}).get(key)
    if tier_price in (None, ""):
        return PriceMatch(
            kind=resolved_kind,
            catalog_kind=catalog_kind,
            route=resolved_route,
            tonnage=key,
            model_tonnage=key,
            distance=route_info.get("distance"),
            coefficient=coefficient,
            issues=[Issue("price_missing", "red", "路线已匹配但吨位价格为空", "维护价格资产或人工补充")],
        )
    # 价格表单元格已经是"标准价×系数"，直接作为运价；再乘一次系数会把
    # 12吨一干价压低 37%、40吨二干价抬高 72%。
    standard = _baseline_price(route_info, catalog, tier_price, key)
    freight = float(tier_price)
    return PriceMatch(
        kind=resolved_kind,
        catalog_kind=catalog_kind,
        route=resolved_route,
        tonnage=key,
        model_tonnage=key,
        distance=route_info.get("distance"),
        coefficient=coefficient,
        standard_price=standard,
        freight=freight if freight % 1 else int(freight),
    )


def _tonnage_number(value: Any) -> Decimal | None:
    key = tonnage_key(value)
    if key in {"40A", "40B"}:
        return Decimal("40")
    try:
        return Decimal(key)
    except Exception:
        return None


def apply_settlement_tonnage(
    price: PriceMatch,
    settlement_tonnage: Any,
    catalog: Mapping[str, Any],
    *,
    confirmed: bool = False,
) -> PriceMatch:
    """Apply an explicitly confirmed settlement-tier override.

    The raw vehicle-management ``结算吨位`` is a candidate signal, not proof of a
    dispatch instruction. Official workbooks contain manually confirmed notes
    such as ``调度通知按30吨结算``; applying every numeric difference would
    incorrectly classify ordinary decimal settlement weights as downgrades.
    Therefore phase one only reports a candidate. The caller must pass
    ``confirmed=True`` after the corrected initial workbook records the business
    decision.
    """
    model_tonnage = price.model_tonnage or price.tonnage
    model_number = _tonnage_number(model_tonnage)
    settlement_number = _tonnage_number(settlement_tonnage)
    if model_number is None or settlement_number is None or model_number < Decimal("30"):
        return replace(price, model_tonnage=model_tonnage)
    if settlement_number >= model_number:
        return replace(price, model_tonnage=model_tonnage, settlement_tonnage=float(settlement_number))

    if not confirmed:
        issue = Issue(
            "settlement_downgrade_candidate",
            "yellow",
            f"降吨位候选：车型={model_tonnage}，车管结算吨位={settlement_tonnage}；未自动改价",
            "请结合调度通知或人工备注确认结算档位",
        )
        return replace(
            price,
            model_tonnage=model_tonnage,
            settlement_tonnage=float(settlement_number),
            issues=[*price.issues, issue],
        )

    eligible: list[tuple[Decimal, Mapping[str, Any]]] = []
    for tier in catalog.get("tiers", []):
        number = _tonnage_number(tier.get("key"))
        if number is not None and number <= settlement_number:
            eligible.append((number, tier))
    if not eligible:
        issue = Issue(
            "settlement_tonnage_unmapped",
            "red",
            f"车管结算吨位{settlement_tonnage}无可用价格档，未自动降吨位",
            "核对车管结算吨位和价格表吨位档",
        )
        return replace(price, model_tonnage=model_tonnage, issues=[*price.issues, issue])

    _, settlement_tier = max(eligible, key=lambda item: (item[0], str(item[1].get("key"))))
    settlement_key = tonnage_key(settlement_tier.get("key"))
    coefficient = settlement_tier.get("coefficient")
    route_info = catalog.get("routes", {}).get(price.route or "")
    # 价格表单元格就是该档位的最终运价；标准价保持合同基准档（系数1.0）价格。
    tier_price = route_info.get("prices", {}).get(settlement_key) if route_info else None
    standard_price = price.standard_price
    if standard_price in (None, "") and route_info:
        standard_price = _baseline_price(route_info, catalog, tier_price)
    freight = original_freight = difference = None
    if tier_price not in (None, ""):
        freight_value = Decimal(str(tier_price))
        freight = float(freight_value) if freight_value % 1 else int(freight_value)
        if price.freight not in (None, ""):
            original_value = Decimal(str(price.freight))
            original_freight = float(original_value) if original_value % 1 else int(original_value)
            difference_value = freight_value - original_value
            difference = float(difference_value) if difference_value % 1 else int(difference_value)

    details = f"车型={model_tonnage}→结算={settlement_key}"
    if original_freight is not None and difference is not None:
        details += f"，原运费={original_freight}，差价={difference:+}"
    issue = Issue(
        "settlement_downgrade",
        "deep_orange",
        f"降吨位：{details}（结算吨位来自车管数据）",
        "核对车管结算吨位与降档价格",
    )
    issues = [issue for issue in price.issues if issue.code != "price_missing"]
    if tier_price in (None, ""):
        issues.append(Issue("price_missing", "red", "降吨位后对应吨位价格为空", "维护价格表或人工补充"))
    issues.append(issue)
    return replace(
        price,
        tonnage=settlement_key,
        model_tonnage=model_tonnage,
        settlement_tonnage=float(settlement_number),
        coefficient=coefficient,
        standard_price=standard_price,
        freight=freight,
        original_coefficient=price.coefficient,
        original_freight=original_freight,
        price_difference=difference,
        downgraded=True,
        issues=issues,
    )
