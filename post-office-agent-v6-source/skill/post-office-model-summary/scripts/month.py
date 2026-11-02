from __future__ import annotations

from datetime import datetime
from pathlib import Path

from reconciliation.excel_io import discover_workbooks
from reconciliation.normalize import parse_month_token


def infer_month_from_inputs(input_dir: Path | str, roles: set[str] | None = None) -> dict:
    """从上传文件推断目标月份，让用户不必记 YYYY-MM。

    roles 不为空时只看这些角色的文件。车型汇总阶段只该看在线测量表，
    否则同目录里的价格表年份（如「2024年10月新一干价格表」）会把月份带偏。

    关键点：财务收入表常常是「上个月」的（业务上允许），不能拿它当目标月。
    所以按角色加权——车管/容积测量/基础报表这些「业务月」证据权重大，
    财务表权重小；只有在没有其他证据时才用财务表。

    依据优先级：
      1. 文件名里的月份（9月.xlsx、2026.08运输局.xlsx）
      2. 工作表名里的月份（草滩的 9月1日、港务的 9.1）

    返回结构（不再「猜一个就完事」）：
        {
          "month": "2026-09" | None,
          "month_only": 9,              # 只推断了月，没有年份
          "year_confirmed": bool,       # 年份是否来自文件而非默认
          "confidence": "high"|"medium"|"low",
          "evidence": [{"file":..., "month":..., "year":...}],
          "conflicts": [...],           # 不同月份的证据，供反问用户
          "questions": [...]            # 需要 agent 主动反问用户的点
        }
    只要 confidence != "high"，调用方就必须先反问用户，不能直接往下跑。
    """
    from collections import Counter

    from reconciliation.excel_io import discover_workbooks
    from reconciliation.normalize import parse_month_token

    root = Path(input_dir)
    result = {
        "month": None, "month_only": None, "year_confirmed": False,
        "confidence": "low", "evidence": [], "conflicts": [], "questions": [], "notes": [],
    }
    if not root.is_dir():
        result["questions"].append("没有找到上传文件，请先上传当月报表")
        return result

    # 财务表月份可能滞后，权重低；其余业务表权重高
    LOW_WEIGHT_HINTS = ("运输局", "财务", "收入")

    paths = discover_workbooks(root)
    if roles:
        from reconciliation.detect import detect_file

        paths = [path for path in paths if detect_file(path).role in roles]

    counter: Counter[tuple[int | None, int]] = Counter()
    for path in paths:
        token = parse_month_token(path.name)
        if not token:
            continue
        weight = 1 if any(hint in path.name for hint in LOW_WEIGHT_HINTS) else 3
        counter[token] += weight
        result["evidence"].append({
            "file": path.name, "month": token[1], "year": token[0], "weight": weight,
        })

    if not counter:
        # 文件名没有月份，退回读工作表名（只读 sheet 名，不加载数据）
        try:
            import openpyxl

            for path in paths:
                try:
                    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                except Exception:
                    continue
                try:
                    for name in wb.sheetnames:
                        token = parse_month_token(name)
                        if token:
                            counter[token] += 1
                            result["evidence"].append({
                                "file": f"{path.name}#{name}", "month": token[1],
                                "year": token[0], "weight": 1,
                            })
                finally:
                    wb.close()
        except Exception:
            pass

    if not counter:
        result["questions"].append(
            "文件名和工作表名里都没有月份信息，请告诉我这是哪个月的报表"
        )
        return result

    # 按月聚合权重，取最高分
    month_score: Counter[int] = Counter()
    for (year, mon), weight in counter.items():
        month_score[mon] += weight
    best = max(month_score.values())
    winners = [mon for mon, score in month_score.items() if score == best]

    if len(winners) != 1:
        result["conflicts"] = [
            {"month": mon, "score": score} for mon, score in month_score.most_common()
        ]
        result["questions"].append(
            "上传的文件里出现了多个不同的月份（"
            + "、".join(f"{mon}月" for mon in sorted(month_score))
            + "），请确认要核对的是哪个月"
        )
        return result

    month = winners[0]
    # 次高分只差一点点时也要问，避免两个月份证据接近却硬选一个
    second = sorted(month_score.values(), reverse=True)
    if len(second) > 1 and second[1] > 0 and (best - second[1]) <= 1:
        result["conflicts"] = [
            {"month": mon, "score": score} for mon, score in month_score.most_common()
        ]
        result["questions"].append(
            f"{month}月和其他月份的线索接近，请确认要核对的是不是 {month} 月"
        )

    # 年份判定：优先取「同一月份」的证据；该月没有年份时，用其他文件的年份兜底
    # （财务表常写上月的完整日期，年份同样有效）。全都没有才需要反问。
    years = {year for (year, mon) in counter if year is not None and mon == month}
    result["month_only"] = month
    if len(years) == 1:
        year = next(iter(years))
        result["year_confirmed"] = True
    elif not years:
        other_years = {year for (year, _mon) in counter if year is not None}
        if len(other_years) == 1:
            # 从别的文件借到唯一确定的年份：能定下来，只需在汇报里说明依据
            year = next(iter(other_years))
            result["year_confirmed"] = True
            result.setdefault("notes", []).append(
                f"{month} 月的文件本身没写年份，参照同批文件的 {year} 年"
            )
        else:
            year = datetime.now().year
            result["year_confirmed"] = False
            result["questions"].append(
                f"文件里只有「{month}月」没有年份，我暂按 {year} 年处理，"
                f"如果不对请告诉我正确年份"
            )
    else:
        year = datetime.now().year
        result["year_confirmed"] = False
        result["questions"].append(
            f"{month} 月的文件出现了多个年份（{'、'.join(str(y) for y in sorted(years))}），"
            "请确认用哪一年"
        )

    result["month"] = f"{year:04d}-{month:02d}"
    if result["conflicts"]:
        result["confidence"] = "low"
    elif result["year_confirmed"]:
        result["confidence"] = "high"
    else:
        result["confidence"] = "medium"
    return result
