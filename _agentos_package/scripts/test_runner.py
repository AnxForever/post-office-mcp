#!/usr/bin/env python3
r"""
邮局报表助手 · 回归测试 runner

跑全套脚本，对比 expected 输出，报告差异。
用于：脚本/提示词改动后验证没退化。

用法:
  python3 test_runner.py

依赖：测试输入数据放在 D:\项目\邮局\数据\ 下（见 test-inputs.md）
"""
import sys
import os
import subprocess
import openpyxl
from pathlib import Path

# === 配置 ===
SCRIPT_DIR = Path(__file__).parent
TEST_DIR = SCRIPT_DIR.parent.parent / 'tests'
EXPECTED_DRAFT = TEST_DIR / 'expected-draft.xlsx'

# 输入数据路径（可被环境变量覆盖）
DATA_ROOT = Path(os.environ.get('POST_OFFICE_DATA_ROOT', '/mnt/d/项目/邮局/数据'))

INPUTS = {
    'caotan':   DATA_ROOT / '草滩在线表.xlsx',
    'gangwu':   DATA_ROOT / '港务在线表.xlsx',
    'baseline': DATA_ROOT / '委办报表开发/4月/4月车型.xlsx',
    'yigan':    DATA_ROOT / '委办报表开发/4月/2024年10月新一干价格表(1).xlsx',
    'ergan':    DATA_ROOT / '委办报表开发/4月/二干运费价.xlsx',
    'carrier':  DATA_ROOT / '每月车管基础数据下载/4月',
    'finance':  DATA_ROOT / '委办报表开发/4月/2026.04财务收入表.xlsx',
    'split_template': DATA_ROOT / '委办报表开发/4月/4月基础报表原版.xlsx',
    'pnl_dir': DATA_ROOT / '委办报表开发/4月',
}


def check_inputs():
    """检查输入文件齐全"""
    missing = []
    for name, path in INPUTS.items():
        if not path.exists():
            missing.append(f"  {name}: {path}")
    if missing:
        print("❌ 缺少测试输入文件:")
        for m in missing:
            print(m)
        print(f"\n设置环境变量 POST_OFFICE_DATA_ROOT 指向数据根目录")
        print(f"当前: {DATA_ROOT}")
        return False
    return True


def run_draft(output_dir):
    """跑初版生成"""
    cmd = [
        sys.executable, str(SCRIPT_DIR / 'build_draft.py'), '4', str(output_dir),
        str(INPUTS['caotan']), str(INPUTS['gangwu']), str(INPUTS['baseline']),
        str(INPUTS['yigan']), str(INPUTS['ergan']), str(INPUTS['carrier']),
    ]
    print(f"\n▶ 跑 build_draft.py...")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        print(f"❌ 脚本失败 (exit {result.returncode})")
        print(result.stderr[-500:])
        return None
    # 打印关键指标
    for line in result.stdout.split('\n'):
        if any(k in line for k in ['[汇总]', '[匹配]', '[问题汇总]', '红(缺失)']):
            print(f"  {line.strip()}")
    return output_dir / '4月_02_基础报表初版.xlsx'


def compare_drafts(actual_path, expected_path):
    """对比两个初版，返回差异列表"""
    wba = openpyxl.load_workbook(actual_path, data_only=True, read_only=True)
    wbe = openpyxl.load_workbook(expected_path, data_only=True, read_only=True)
    wa = wba['基础报表']; we = wbe['基础报表']
    rows_a = [tuple(row) for row in wa.iter_rows(values_only=True)]
    rows_e = [tuple(row) for row in we.iter_rows(values_only=True)]

    diffs = []

    # 1. 行数
    if len(rows_a) != len(rows_e):
        diffs.append(f"行数: actual={len(rows_a)} expected={len(rows_e)}")

    # 2. 列数
    cols_a = max((len(row) for row in rows_a), default=0)
    cols_e = max((len(row) for row in rows_e), default=0)
    if cols_a != cols_e:
        diffs.append(f"列数: actual={cols_a} expected={cols_e}")

    # 3. 关键指标
    def metrics(rows):
        data = rows[1:]
        return {
            'rows': len(data),
            'no_ton': sum(1 for row in data if len(row) < 11 or not row[10]),
            'no_price': sum(1 for row in data if len(row) < 16 or not row[15]),
            'no_km': sum(1 for row in data if len(row) < 12 or not row[11]),
            'cross': sum(1 for row in data if len(row) >= 18 and row[17] == '是'),
        }
    ma = metrics(rows_a)
    me = metrics(rows_e)
    for key in ['rows', 'no_ton', 'no_price', 'no_km', 'cross']:
        if ma[key] != me[key]:
            diffs.append(f"指标{key}: actual={ma[key]} expected={me[key]} (diff={ma[key]-me[key]:+d})")

    # 4. 逐格内容对比（按 行号+车牌 匹配，避免顺序差异）
    # 简化：按行对比（前提是脚本顺序稳定）
    cell_diffs = 0
    for row_a, row_e in zip(rows_a[1:], rows_e[1:]):
        for va, ve in zip(row_a, row_e):
            if va != ve:
                try:
                    if abs(float(va) - float(ve)) < 0.001:
                        continue
                except (TypeError, ValueError):
                    pass
                cell_diffs += 1
    if cell_diffs:
        diffs.append(f"单元格内容差异: {cell_diffs} 处")

    wba.close(); wbe.close()
    return diffs, ma, me


def run_split(output_dir, draft_path):
    cmd = [
        sys.executable, str(SCRIPT_DIR / 'build_split.py'), '4', str(output_dir),
        str(draft_path), str(INPUTS['split_template']),
    ]
    print("\n▶ 跑 build_split.py...")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        print(result.stdout[-1000:]); print(result.stderr[-1000:])
        return None
    print("  ✅ 拆分脚本执行成功")
    return output_dir / '4月_03_基础报表拆分版.xlsx'


def effective_headers(ws):
    values = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    last = max((i for i, value in enumerate(values, 1) if value is not None), default=0)
    return tuple(str(value or '').strip() for value in values[:last])


def verify_split(path, draft_rows):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    allowed = {
        ('序号','车辆承运单位','派车单号','邮路名称','路单编号','封车时间','','车牌号','','容间','吨位','公里','运输量','路单号','运价','标准价','备注'),
        ('序号','车辆承运单位','派车单号','邮路名称','路单编号','封车时间','','车牌号','','容间','吨位','公里','运输量','路单号','运价（保留两位数）','标准价','运价','备注'),
    }
    data_rows = 0
    for ws in wb.worksheets:
        headers = effective_headers(ws)
        assert headers in allowed, f"拆分 Sheet {ws.title} 表头未照搬模板: {headers}"
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and isinstance(row[0], int):
                data_rows += 1
    wb.close()
    assert data_rows == draft_rows, f"拆分行数不守恒: {data_rows} != {draft_rows}"
    return {'data_rows': data_rows}


def run_pnl(output_dir, draft_path):
    cmd = [
        sys.executable, str(SCRIPT_DIR / 'build_pnl.py'), '4', str(output_dir),
        str(draft_path), str(INPUTS['finance']), str(INPUTS['pnl_dir']),
    ]
    print("\n▶ 跑 build_pnl.py...")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        print(result.stdout[-1500:]); print(result.stderr[-1500:])
        return None
    for line in result.stdout.splitlines():
        if '[模板]' in line or '未归类路线=' in line or '抽样通过' in line:
            print(f"  {line.strip()}")
    return [output_dir / f'4月_04_{level}盈亏表.xlsx' for level in ('一干', '二干')]


def latest_template_sheet(wb):
    import re
    candidates = []
    for sn in wb.sheetnames:
        if re.match(r'^(?:\d{2}年)?\d+月$', sn):
            ws = wb[sn]
            if str(ws.cell(2, 1).value or '').strip() == '邮路名称':
                candidates.append(sn)
    assert candidates, '盈亏模板无有效月份 Sheet'
    return candidates[-1]


def pnl_structure(ws):
    group_starts = []
    summary_start = None
    for col in range(2, ws.max_column + 1):
        value = str(ws.cell(2, col).value or '').strip()
        if value in ('趟次', '委办运费(元)', '结算收入(元)', '总盈亏(元)'):
            summary_start = summary_start or col
        elif value:
            group_starts.append((col, value))
    total_row = next(r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or '').strip() == '合计')
    route_order = [str(ws.cell(r, 1).value or '').strip() for r in range(4, total_row) if str(ws.cell(r, 1).value or '').strip()]
    return group_starts, summary_start, total_row, route_order


def verify_pnl(paths):
    metrics = {}
    for path, level in zip(paths, ('一干', '二干')):
        template_wb = openpyxl.load_workbook(INPUTS['pnl_dir'] / f'{level}委办盈亏.xlsx', data_only=False, read_only=True)
        template_ws = template_wb[latest_template_sheet(template_wb)]
        expected_groups, expected_summary, _, expected_routes = pnl_structure(template_ws)

        wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
        ws = wb[wb.sheetnames[0]]
        groups, summary_start, total_row, routes = pnl_structure(ws)
        assert [v for _, v in groups] == [v for _, v in expected_groups], f'{level}车型组未使用最新模板'
        assert summary_start == expected_summary, f'{level}汇总区列号不符'
        assert routes == expected_routes, f'{level}路线顺序未照搬最新模板'
        labels = [str(ws.cell(r, 1).value or '').strip() for r in range(total_row, ws.max_row + 1)]
        assert '环比' in labels and '4月' in labels, f'{level}缺本月/环比区'

        current_row = next(r for r in range(total_row + 1, ws.max_row + 1) if str(ws.cell(r, 1).value or '').strip() == '4月')
        group_trip_cols = [2 + i * 7 for i in range(len(groups))]
        group_transport_cols = [c + 1 for c in group_trip_cols]
        group_freight_cols = [c + 3 for c in group_trip_cols]
        group_settle_cols = [c + 4 for c in group_trip_cols]
        trip_total = sum(float(ws.cell(total_row, c).value or 0) for c in group_trip_cols)
        transport_total = sum(float(ws.cell(total_row, c).value or 0) for c in group_transport_cols)
        freight_total = sum(float(ws.cell(total_row, c).value or 0) for c in group_freight_cols)
        settle_total = sum(float(ws.cell(total_row, c).value or 0) for c in group_settle_cols)
        assert float(ws.cell(current_row, 2).value or 0) == trip_total
        assert abs(float(ws.cell(current_row, 3).value or 0) - transport_total) < 1
        assert abs(float(ws.cell(current_row, 4).value or 0) - freight_total) < 0.01
        assert abs(float(ws.cell(current_row, 5).value or 0) - settle_total) < 0.01
        assert abs(float(ws.cell(current_row, 6).value or 0) - (settle_total - freight_total)) < 0.01
        metrics[level] = {'groups': len(groups), 'routes': len(routes), 'trips': trip_total}
        wb.close(); template_wb.close()
    return metrics


def main():
    print("=" * 60)
    print("邮局报表助手 · 回归测试")
    print("=" * 60)

    if not check_inputs():
        sys.exit(1)

    if not EXPECTED_DRAFT.exists():
        print(f"❌ 缺少 expected: {EXPECTED_DRAFT}")
        sys.exit(2)

    output_dir = TEST_DIR / 'output'
    output_dir.mkdir(exist_ok=True)

    actual = run_draft(output_dir)
    if actual is None or not actual.exists():
        print("❌ 初版未生成")
        sys.exit(3)

    print(f"\n▶ 对比 expected...")
    diffs, ma, me = compare_drafts(actual, EXPECTED_DRAFT)

    print(f"\n{'='*60}")
    print(f"{'指标':<16} {'actual':>8} {'expected':>8} {'diff':>6}")
    print("-" * 42)
    for k in ['rows', 'no_ton', 'no_price', 'no_km', 'cross']:
        d = ma[k] - me[k]
        flag = '✅' if d == 0 else '❌'
        print(f"{k:<16} {ma[k]:>8} {me[k]:>8} {d:>+6} {flag}")

    split_path = run_split(output_dir, actual)
    if split_path is None or not split_path.exists():
        diffs.append('拆分版未生成')
    else:
        try:
            split_metrics = verify_split(split_path, ma['rows'])
            print(f"  ✅ 拆分行数守恒: {split_metrics['data_rows']}")
        except AssertionError as exc:
            diffs.append(str(exc))

    pnl_paths = run_pnl(output_dir, actual)
    if pnl_paths is None or not all(path.exists() for path in pnl_paths):
        diffs.append('盈亏表未生成')
    else:
        try:
            pnl_metrics = verify_pnl(pnl_paths)
            for level, item in pnl_metrics.items():
                print(f"  ✅ {level}: 车型组={item['groups']} 路线={item['routes']} 趟次={item['trips']:.0f}")
        except AssertionError as exc:
            diffs.append(str(exc))

    if diffs:
        print(f"\n❌ 发现 {len(diffs)} 处差异:")
        for d in diffs:
            print(f"  - {d}")
        sys.exit(1)
    else:
        print(f"\n✅ 全部通过，输出与 expected 一致")
        sys.exit(0)


if __name__ == '__main__':
    main()
