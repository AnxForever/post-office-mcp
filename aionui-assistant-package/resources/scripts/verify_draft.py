#!/usr/bin/env python3
"""
邮局报表助手 · 初版输出验证工具

读 build_draft.py 生成的初版，自检：
- Sheet 结构（结构发现/问题汇总/基础报表）
- 21 列字段完整性
- 关键指标异常
- 可疑项标注是否齐全（不应有悄悄留空）
- 运价公式抽样验算

用法:
  python3 verify_draft.py <初版.xlsx>
"""
import sys
import openpyxl
from pathlib import Path


def main():
    if len(sys.argv) < 2:
        print("用法: python3 verify_draft.py <初版.xlsx>")
        sys.exit(1)

    path = Path(sys.argv[1])
    if not path.exists():
        print(f"❌ 文件不存在: {path}")
        sys.exit(2)

    wb = openpyxl.load_workbook(path, data_only=True)

    issues = []

    # 1. Sheet 结构
    required_sheets = ['结构发现', '问题汇总', '基础报表']
    for sn in required_sheets:
        if sn not in wb.sheetnames:
            issues.append(f"缺少 Sheet: {sn}")

    print("=" * 60)
    print(f"验证: {path.name}")
    print("=" * 60)
    print(f"Sheets: {wb.sheetnames}")

    if '基础报表' not in wb.sheetnames:
        print("❌ 没有基础报表 Sheet，无法验证")
        sys.exit(3)

    ws = wb['基础报表']
    print(f"\n[基础报表] 行={ws.max_row} 列={ws.max_column}")

    # 2. 21 列字段
    expected_cols = ['序号', '车辆承运单位', '派车单号', '邮路名称', '路单编号', '封车时间',
                     None, '车牌号', None, '容间', '吨位', '公里', '运输量', '路单号',
                     '运价', '标准价', '备注', '是否跨档', '变动前容间', '变动前时间', '系数']
    if ws.max_column != 21:
        issues.append(f"列数不对: {ws.max_column}, 应为 21")
    for i, name in enumerate(expected_cols, 1):
        actual = ws.cell(row=1, column=i).value
        if name and actual != name:
            issues.append(f"Col{i} 表头错: 实际 '{actual}', 应为 '{name}'")
    print(f"列数: {ws.max_column} {'✅' if ws.max_column == 21 else '❌'}")

    # 3. 关键指标
    total = ws.max_row - 1
    no_ton = sum(1 for r in range(2, ws.max_row + 1) if not ws.cell(row=r, column=11).value)
    no_price = sum(1 for r in range(2, ws.max_row + 1) if not ws.cell(row=r, column=16).value)
    no_km = sum(1 for r in range(2, ws.max_row + 1) if not ws.cell(row=r, column=12).value)
    cross = sum(1 for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=18).value == '是')
    print(f"\n[指标] 总={total} 吨位空={no_ton} 价格空={no_price} 公里空={no_km} 跨档={cross}")

    # 4. 可疑项标注检查（不应有悄悄留空）
    silent_empty = 0
    for r in range(2, ws.max_row + 1):
        ton = ws.cell(row=r, column=11).value
        vol = ws.cell(row=r, column=10).value
        remark = str(ws.cell(row=r, column=17).value or '')
        # 容间>0 但吨位空 = 应该有备注说明
        if vol and vol > 0 and not ton:
            if '吨位' not in remark and '容间' not in remark:
                silent_empty += 1
                if silent_empty <= 3:
                    print(f"  ❌ Row{r}: 容间={vol} 吨位空 但无说明备注")
    if silent_empty:
        issues.append(f"{silent_empty} 条容间>0但吨位空且无说明备注（悄悄留空）")

    # 5. 运价公式抽样验算
    formula_errors = 0
    samples = 0
    for r in range(2, min(ws.max_row + 1, 102)):
        freight = ws.cell(row=r, column=15).value
        price = ws.cell(row=r, column=16).value
        coef = ws.cell(row=r, column=21).value
        if freight and price and coef:
            samples += 1
            expected = round(float(coef) * float(price), 2)
            actual = round(float(freight), 2)
            if abs(actual - expected) > 0.05:
                formula_errors += 1
                if formula_errors <= 3:
                    print(f"  ❌ Row{r}: {coef}×{price}={expected} != 运价{actual}")
    print(f"\n[运价公式] 抽样 {samples} 条，错误 {formula_errors} 条")
    if formula_errors:
        issues.append(f"{formula_errors} 条运价公式错误")

    # 6. 问题汇总覆盖率
    if '问题汇总' in wb.sheetnames:
        ws_p = wb['问题汇总']
        prob_count = ws_p.max_row - 3
        print(f"\n[问题汇总] {prob_count} 条")

    # === 总结 ===
    print(f"\n{'='*60}")
    if not issues:
        print("✅ 全部通过")
        sys.exit(0)
    else:
        print(f"❌ 发现 {len(issues)} 个问题:")
        for i, iss in enumerate(issues, 1):
            print(f"  {i}. {iss}")
        sys.exit(1)


if __name__ == '__main__':
    main()
