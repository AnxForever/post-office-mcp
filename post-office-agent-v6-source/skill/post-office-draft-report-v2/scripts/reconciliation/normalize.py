from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Any, Mapping


_DASHES = {"", "-", "--", "—", "–", "－", "无", "无挂车", "无挂", "NONE", "N/A"}
_BUSINESS_SUFFIX_RE = re.compile(
    r"(?:港航|旺临|特快|陕串|港串|提速|陕临|车次|陕T|港T|港|航|旺)"
    r"\s*[TＴ]?[一二三四五六七八九十A-Za-z0-9]*$",
    re.I,
)
_PAREN_RE = re.compile(r"[（(][^（）()]*[）)]")
_WASTE_RE = re.compile(r"废\s*\d*$")


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).replace("\u00a0", " ")
    return re.sub(r"\s+", " ", text).strip()


def normalize_plate(value: Any) -> str:
    return re.sub(r"\s+", "", normalize_text(value)).upper()


def normalize_trailer(value: Any) -> str:
    text = normalize_plate(value)
    if text in _DASHES or not text:
        return "--"
    return text


def normalize_dispatch_no(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return normalize_text(value)


def parse_month_token(value: Any) -> tuple[int | None, int] | None:
    text = normalize_text(value)
    if not text:
        return None
    match = re.search(r"(?<!\d)(20\d{2})[年./_-]\s*(1[0-2]|0?[1-9])(?:月|[./_-])?", text)
    if match:
        return int(match.group(1)), int(match.group(2))
    match = re.search(r"(?<!\d)(1[0-2]|0?[1-9])\s*月", text)
    if match:
        return None, int(match.group(1))
    match = re.search(r"(?<!\d)(1[0-2]|0?[1-9])\.(?:[0-3]?\d)(?!\d)", text)
    if match:
        return None, int(match.group(1))
    return None


def parse_daily_sheet_date(value: Any, default_year: int) -> datetime | None:
    text = normalize_text(value)
    if not text:
        return None
    patterns = (
        r"(?:(20\d{2})年)?\s*(1[0-2]|0?[1-9])月\s*([0-3]?\d)日?",
        r"(?:(20\d{2})[./_-])?\s*(1[0-2]|0?[1-9])\.([0-3]?\d)(?!\d)",
    )
    for pattern in patterns:
        match = re.fullmatch(pattern, text)
        if not match:
            continue
        year = int(match.group(1)) if match.group(1) else int(default_year)
        try:
            return datetime(year, int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None


def _append_unique(items: list[str], value: str) -> None:
    value = normalize_text(value)
    if value and value not in items:
        items.append(value)


def route_candidates(value: Any, aliases: Mapping[str, Any] | None = None) -> list[str]:
    original = normalize_text(value)
    if not original:
        return []
    candidates: list[str] = []
    _append_unique(candidates, original)
    no_paren = _PAREN_RE.sub("", original).strip()
    _append_unique(candidates, no_paren)
    no_suffix = _BUSINESS_SUFFIX_RE.sub("", no_paren).strip()
    _append_unique(candidates, no_suffix)
    no_transfer = re.sub(r"(?:转|本)$", "", no_suffix).strip()
    _append_unique(candidates, no_transfer)
    no_waste = _WASTE_RE.sub("", no_transfer).strip()
    _append_unique(candidates, no_waste)
    for item in tuple(candidates):
        normalized = item.replace("西安草滩-", "西安-")
        normalized = normalized.replace("西安航站-", "西安-").replace("航站-", "西安-")
        normalized = normalized.replace("西安-上海东", "西安-上海").replace("西安-上海西", "西安-上海")
        _append_unique(candidates, normalized)
    if aliases:
        lookup = {normalize_text(k): v for k, v in aliases.items()}
        for item in tuple(candidates):
            entry = lookup.get(item)
            if isinstance(entry, str):
                _append_unique(candidates, entry)
            elif isinstance(entry, Mapping):
                _append_unique(candidates, str(entry.get("standard_route", "")))
    return candidates


def tonnage_key(value: Any) -> str:
    text = normalize_text(value).upper().replace("吨", "")
    if not text:
        return ""
    if "40A" in text or "40(A" in text or "40（A" in text:
        return "40A"
    if "40B" in text or "40(B" in text or "40（B" in text:
        return "40B"
    match = re.search(r"(\d+(?:\.\d+)?)", text)
    if not match:
        return text
    try:
        number = Decimal(match.group(1))
    except InvalidOperation:
        return match.group(1)
    return format(number.normalize(), "f")


def destination_name(route: Any) -> str:
    text = route_candidates(route)[-1] if route_candidates(route) else normalize_text(route)
    return text.split("-")[-1].strip()
