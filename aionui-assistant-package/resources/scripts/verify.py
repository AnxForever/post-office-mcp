#!/usr/bin/env python3
"""
邮局报表助手 · 输出验证脚本
对比本次输出与基准版的指标，确认结果正确。

用法:
  python3 verify.py <本次输出.xlsx> <基准版.xlsx>

退出码 0 = 关键指标一致, 非 0 = 有差异
"""
import sys
import openpyxl


def load_report(path):
    """加载基础报表，返回 (records, problem_count, cross_count)"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['基础报表']

    total = ws.max_row - 1
    no_ton = sum(1 for r in range(2, ws.max_row + 1) if not ws.cell(row=r, column=11).value)
    no_price = sum(1 for r in range(2, ws.max_row + 1) if not ws.cell(row=r, column=16).value)
    no_km = sum(1 for r in range(2, ws.max_row + 1) if not ws.cell(row=r, column=12).value)
    no_freight = sum(1 for r in range(2, ws.max_row + 1) if not ws.cell(row=r, column=15).value)
    cross = sum(1 for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=18).value == '是')

    has_struct = '结构发现' in wb.sheetnames
    has_prob = '问题汇总' in wb.sheetnames
    prob_count = 0
    if has_prob:
        ws_p = wb['问题汇总']
        # 找最后一行有效问题
        for r in range(ws_p.max_row, 3, -1):
            if ws_p.cell(row=r, column=1).value is not None:
                prob_count = r - 3
                break

    wb.close()
    return {
        'total': total, 'no_ton': no_ton, 'no_price': no_price,
        'no_km': no_km, 'no_freight': no_freight, 'cross': cross,
        'has_struct': has_struct, 'has_prob': has_prob, 'prob_count': prob_count,
    }


def main():
    if len(sys.argv) < 3:
        print("用法: python3 verify.py <本次输出.xlsx> <基准版.xlsx>")
        sys.exit(1)

    cur_path, base_path = sys.argv[1], sys.argv[2]
    cur = load_report(cur_path)
    base = load_report(base_path)

    print(f"{'指标':<16} {'本次':>8} {'基准':>8} {'差异':>8} {'判定':>6}")
    print("-" * 50)

    all_ok = True
    for key, label in [
        ('total', '总记录数'),
        ('no_ton', '吨位为空'),
        ('no_price', '价格为空'),
        ('no_km', '公里为空'),
        ('no_freight', '运价为空'),
        ('cross', '跨档检出'),
        ('prob_count', '问题汇总数'),
    ]:
        cv, bv = cur[key], base[key]
        diff = cv - bv
        ok = diff == 0
        if not ok:
            all_ok = False
        print(f"{label:<16} {cv:>8} {bv:>8} {diff:>+8} {'✅' if ok else '❌':>6}")

    print("-" * 50)
    print(f"结构发现Sheet: {'✅' if cur['has_struct'] else '❌'}")
    print(f"问题汇总Sheet: {'✅' if cur['has_prob'] else '❌'}")

    print(f"\n{'✅ 关键指标全部一致' if all_ok else '❌ 有指标差异，请检查'}")
    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()
