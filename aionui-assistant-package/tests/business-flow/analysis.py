#!/usr/bin/env python3
"""审计邮局报表助手的真实业务阶段与模板结构。"""

from __future__ import annotations

import json
import os
import re
from collections import Counter
from pathlib import Path

import openpyxl


PROJECT = Path(os.environ.get("POST_OFFICE_PROJECT_ROOT", "/mnt/d/项目/邮局"))
PACKAGE = PROJECT / "aionui-assistant-package"
OUT = Path(__file__).with_name("validation_report.md")

DRAFT_HEADERS = [
    "序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
    "车牌号", "", "容间", "吨位", "公里", "运输量", "路单号", "运价", "标准价",
    "备注", "是否跨档", "变动前容间", "变动前时间", "系数",
]
FORMAL_HEADERS = DRAFT_HEADERS[:17]
FORMAL_HEADERS_WITH_ROUNDED_FREIGHT = [
    "序号", "车辆承运单位", "派车单号", "邮路名称", "路单编号", "封车时间", "",
    "车牌号", "", "容间", "吨位", "公里", "运输量", "路单号", "运价（保留两位数）",
    "标准价", "运价", "备注",
]


def header_values(ws, row=1):
    values = next(ws.iter_rows(min_row=row, max_row=row, values_only=True), ())
    last = 0
    for col, value in enumerate(values, 1):
        if value is not None:
            last = col
    return [str(value or "").strip() for value in values[:last]]


def workbook_matrix(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
    ws = wb[sheet]
    matrix = tuple(tuple(row) for row in ws.iter_rows(values_only=True))
    wb.close()
    return matrix


def audit_drafts():
    files = {
        "correct": PROJECT / "交付物/4月_02_基础报表初版.xlsx",
        "expected": PACKAGE / "tests/expected-draft.xlsx",
        "current": PACKAGE / "tests/output/4月_02_基础报表初版.xlsx",
    }
    result = {}
    for label, path in files.items():
        assert path.exists(), f"缺少初版文件: {path}"
        wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
        assert "基础报表" in wb.sheetnames, f"{label} 缺少基础报表 Sheet"
        ws = wb["基础报表"]
        headers = header_values(ws)
        result[label] = {
            "path": str(path),
            "sheets": wb.sheetnames,
            "rows": ws.max_row - 1,
            "columns": len(headers),
            "headers": headers,
        }
        wb.close()
        assert headers == DRAFT_HEADERS, f"{label} 初版不是 21 列 v1.16: {headers}"

    correct = workbook_matrix(files["correct"], "基础报表")
    expected = workbook_matrix(files["expected"], "基础报表")
    current = workbook_matrix(files["current"], "基础报表")
    result["correct_equals_expected"] = correct == expected
    result["current_equals_expected"] = current == expected
    assert result["correct_equals_expected"], "回归 expected 与已确认正确初版不一致"
    assert result["current_equals_expected"], "当前回归输出与 expected 不一致"
    return result


def audit_formal_templates():
    result = {}
    special_names = {"Sheet1", "Sheet2", "Sheet3", "Sheet4", "降吨位", "二干基础", "两趟合一趟", "一干汇总"}
    for month in (3, 4, 5):
        path = PROJECT / f"数据/委办报表开发/{month}月/{month}月基础报表原版.xlsx"
        assert path.exists(), f"缺少拆分模板: {path}"
        wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
        business = {}
        for sn in wb.sheetnames:
            if sn in special_names:
                continue
            ws = wb[sn]
            headers = header_values(ws)
            if "序号" in headers and "邮路名称" in headers:
                business[sn] = len(headers)
        dist = Counter(business.values())
        header_patterns = Counter()
        for sn in business:
            header_patterns[tuple(header_values(wb[sn]))] += 1
        result[str(month)] = {
            "path": str(path),
            "sheet_count": len(wb.sheetnames),
            "business_sheet_count": len(business),
            "column_distribution": dict(sorted(dist.items())),
            "special_sheets": [
                name for name in ("二干基础", "一干汇总", "两趟合一趟", "降吨位")
                if name in wb.sheetnames
            ],
            "non_17_sheets": sorted(name for name, cols in business.items() if cols != 17),
            "header_patterns": {" | ".join(k): v for k, v in header_patterns.items()},
        }
        wb.close()
        assert business, f"{month}月拆分模板没有识别到业务 Sheet"
        allowed = {tuple(FORMAL_HEADERS), tuple(FORMAL_HEADERS_WITH_ROUNDED_FREIGHT)}
        unexpected = [list(pattern) for pattern in header_patterns if pattern not in allowed]
        assert not unexpected, f"{month}月常规拆分 Sheet 出现未知表头结构: {unexpected}"
    return result


def find_pnl_header_sheet(wb):
    candidates = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        if (
            str(ws.cell(row=2, column=1).value or "").strip() == "邮路名称"
            and re.match(r"^(?:\d{2}年)?\d+月$", sn)
        ):
            candidates.append(sn)
    if candidates:
        return candidates[-1]
    raise AssertionError("盈亏模板中未找到邮路名称表头")


def audit_one_pnl(path):
    wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
    sn = find_pnl_header_sheet(wb)
    ws = wb[sn]
    summary_labels = {"趟次", "委办运费(元)", "结算收入(元)", "总盈亏(元)"}
    starts = []
    summary_start = None
    for col in range(2, ws.max_column + 1):
        value = str(ws.cell(row=2, column=col).value or "").strip()
        if value in summary_labels:
            summary_start = summary_start or col
        elif value:
            starts.append((col, value))
    assert starts, f"{path} 未识别到车型组"
    assert summary_start is not None, f"{path} 未识别到汇总区"

    widths = []
    for i, (start, name) in enumerate(starts):
        end = starts[i + 1][0] if i + 1 < len(starts) else summary_start
        widths.append((name, end - start))
    summary = [str(ws.cell(row=2, column=c).value or "").strip() for c in range(summary_start, summary_start + 4)]
    labels_col1 = [str(ws.cell(row=r, column=1).value or "").strip() for r in range(1, ws.max_row + 1)]
    result = {
        "path": str(path),
        "sheet": sn,
        "vehicle_groups": [name for _, name in starts],
        "group_widths": widths,
        "summary_headers": summary,
        "has_total": "合计" in labels_col1,
        "has_previous": "上月" in labels_col1,
        "has_current": "本月" in labels_col1,
        "has_mom": "环比" in labels_col1,
    }
    wb.close()
    assert all(width == 7 for _, width in widths), f"车型组不是每组 7 列: {widths}"
    assert summary == ["趟次", "委办运费(元)", "结算收入(元)", "总盈亏(元)"], f"汇总区不符: {summary}"
    assert result["has_total"], f"{path} 缺合计行"
    return result


def audit_pnl_templates():
    result = {}
    for month in (3, 4, 5):
        for level in ("一干", "二干"):
            path = PROJECT / f"数据/委办报表开发/{month}月/{level}委办盈亏.xlsx"
            assert path.exists(), f"缺少盈亏模板: {path}"
            result[f"{month}月{level}"] = audit_one_pnl(path)
    assert any(v["has_mom"] for v in result.values()), "所有官方盈亏样本都未发现环比区"
    return result


def render_report(data):
    drafts = data["drafts"]
    formal = data["formal_templates"]
    pnl = data["pnl_templates"]
    lines = [
        "# Validation Report — 邮局报表业务链路",
        "",
        "## 阶段结构",
        "",
        "| 阶段 | 权威样本 | 结构 | Assertions | 状态 |",
        "|---|---|---:|---|---|",
        f"| 基础报表初版 | 交付物正确版 | {drafts['correct']['columns']} 列 / {drafts['correct']['rows']} 条 | 21列v1.16、与expected一致 | PASS |",
    ]
    for month, item in formal.items():
        dist = ", ".join(f"{k}列×{v}" for k, v in item["column_distribution"].items())
        lines.append(f"| {month}月正式拆分模板 | 邮局基础报表原版 | {dist} | 逐Sheet照搬17/18列结构 | PASS |")
    for key, item in pnl.items():
        lines.append(
            f"| {key}盈亏模板 | 邮局官方样本 `{item['sheet']}` | {len(item['vehicle_groups'])} 个车型组 | 每组7列、4汇总列、合计行 | PASS |"
        )
    lines += [
        "",
        "## 流程结论",
        "",
        "1. Step 1—6 生成 **21 列基础报表初版**，包含结构发现、问题汇总和带核对标记的基础报表。",
        "2. 初版必须先由人工核对修正；未经确认不得继续下游。",
        "3. Step 7 读取核对后初版，按当月官方基础报表原版的 Sheet、顺序与表头，逐 Sheet 生成 **17/18 列模板结构**的拆分正式版。",
        "4. Step 8 从核对后初版聚合运输量/运价，从财务收入表聚合趟次/结算收入，套用官方盈亏模板；车型组必须动态读取。",
        "",
        "## 断言结果",
        "",
        f"- 正确初版与 expected 逐格一致：{'PASS' if drafts['correct_equals_expected'] else 'FAIL'}",
        f"- 当前输出与 expected 逐格一致：{'PASS' if drafts['current_equals_expected'] else 'FAIL'}",
        f"- 盈亏样本中发现环比区：{'PASS' if any(v['has_mom'] for v in pnl.values()) else 'FAIL'}",
        "",
        "所有数字均由 `analysis.py` 读取 Excel 后生成。",
    ]
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    data = {
        "drafts": audit_drafts(),
        "formal_templates": audit_formal_templates(),
        "pnl_templates": audit_pnl_templates(),
    }
    render_report(data)
    print(json.dumps(data, ensure_ascii=False, indent=2))
    print(f"validation_report={OUT}")


if __name__ == "__main__":
    main()
