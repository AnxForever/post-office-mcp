#!/usr/bin/env python3
"""
邮局报表助手 · 源文件结构探查工具

只读不写。打开所有源文件，打印本次动态识别的结构。
助手调 build_*.py 前先调本工具，确认结构识别正确。

用法:
  python3 inspect_sources.py <草滩表> <港务表> <车型表> <一干价格表> <二干价格表> <车管目录>
"""
import sys
import re
import os
from pathlib import Path
import openpyxl


def find_header_row(ws, keywords, scan=8):
    """前 scan 行中命中 keywords 最多的行"""
    best_r, best_hits = 1, -1
    for r in range(1, min(scan, ws.max_row) + 1):
        txt = [str(ws.cell(r, c).value or '') for c in range(1, ws.max_column + 1)]
        hits = sum(1 for kw in keywords if any(kw in t for t in txt))
        if hits > best_hits:
            best_hits, best_r = hits, r
    return best_r


def locate_ordered(ws, hrow, specs):
    hdr = {c: str(ws.cell(hrow, c).value or '') for c in range(1, ws.max_column + 1)}
    used, res = set(), {}
    for field, kws in specs:
        col = None
        for kw in kws:
            for c, h in hdr.items():
                if c in used:
                    continue
                if kw in h:
                    col = c
                    break
            if col:
                break
        res[field] = col
        if col:
            used.add(col)
    return res


def parse_coef(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r'([\d.]+)', str(v))
    return float(m.group(1)) if m else None


def parse_limit(v):
    if v is None:
        return 0.0
    m = re.search(r'≥\s*(\d+)', str(v)) or re.search(r'(\d+)\s*(?:立方米|m3|方)', str(v))
    return float(m.group(1)) if m else 0.0


def discover_price_sheet(ws):
    ton_row = coef_row = limit_row = None
    for r in range(1, min(8, ws.max_row) + 1):
        txt = [str(ws.cell(r, c).value or '') for c in range(1, ws.max_column + 1)]
        joined = ' '.join(txt)
        if ton_row is None and sum(1 for t in txt if re.search(r'\d+\.?\d*吨', t)) >= 2:
            ton_row = r
        if limit_row is None and ('≥' in joined or 'm3' in joined or '立方米' in joined):
            limit_row = r
        if coef_row is None and '系数' in joined:
            coef_row = r
    if ton_row is None:
        return None
    if coef_row is None:
        coef_row = ton_row - 1
    if limit_row is None:
        limit_row = ton_row + 1
    tiers = []
    for c in range(1, ws.max_column + 1):
        m = re.search(r'(\d+\.?\d*吨[A-B]?)', str(ws.cell(ton_row, c).value or ''))
        if not m:
            continue
        coef = parse_coef(ws.cell(coef_row, c).value)
        if coef is None:
            continue
        tiers.append((parse_limit(ws.cell(limit_row, c).value), m.group(1), coef, c))
    if not tiers:
        return None
    route_col = km_col = None
    for hr in [coef_row, ton_row, limit_row]:
        if not (1 <= hr <= ws.max_row):
            continue
        for c in range(1, ws.max_column + 1):
            h = str(ws.cell(hr, c).value or '')
            if route_col is None and any(k in h for k in ['路向', '邮路', '路线']):
                route_col = c
            if km_col is None and any(k in h for k in ['里程', '公里', 'KM', 'km']):
                km_col = c
    data_start = max(r for r in [ton_row, coef_row, limit_row] if r) + 1
    tiers.sort(key=lambda t: -t[0])
    return {'route_col': route_col, 'km_col': km_col, 'data_start': data_start, 'tiers': tiers}


def inspect_online(path, src_name, pat_str):
    """探查在线表"""
    print(f"\n[{src_name}] {path}")
    if not Path(path).exists():
        print(f"  ❌ 文件不存在")
        return
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    pat = re.compile(pat_str)
    matched = [sn for sn in wb.sheetnames if pat.match(sn.strip())]
    print(f"  Sheet总数: {len(wb.sheetnames)}, 匹配月份格式的: {len(matched)}")
    if matched:
        ws = wb[matched[0]]
        hr = find_header_row(ws, ['序号', '邮路', '车'])
        print(f"  表头行(首Sheet): {hr}")
        cols = locate_ordered(ws, hr, [
            ('挂车', ['车厢车牌号', '车厢码', '挂车号', '挂车']),
            ('车牌', ['车牌号', '车号']),
            ('邮路', ['邮路名称', '邮路', '路线']),
            ('容间', ['测量容积', '容积']),
        ])
        print(f"  列定位: {cols}")
    wb.close()


def inspect_baseline(path):
    print(f"\n[基准表] {path}")
    if not Path(path).exists():
        print(f"  ❌ 文件不存在")
        return
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
    hr = find_header_row(ws, ['序号', '邮路', '车牌', '车号'])
    print(f"  表头行: {hr} (数据从 {hr+1} 起)")
    cols = locate_ordered(ws, hr, [
        ('挂车', ['车厢车牌号', '车厢码', '挂车号', '挂车']),
        ('车牌', ['车牌号', '车号']),
        ('邮路', ['邮路名称', '邮路', '路线']),
        ('容间', ['测量容积', '容积']),
    ])
    print(f"  列定位: {cols}")
    wb.close()


def inspect_price(path, label):
    print(f"\n[{label}] {path}")
    if not Path(path).exists():
        print(f"  ❌ 文件不存在")
        return
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets_used = []
    first_tiers = None
    for sn in wb.sheetnames:
        ws = wb[sn]
        disc = discover_price_sheet(ws)
        if not disc:
            continue
        n_routes = sum(1 for r in range(disc['data_start'], ws.max_row + 1)
                       if str(ws.cell(r, disc['route_col']).value or '').strip()) if disc['route_col'] else 0
        if n_routes == 0:
            continue
        sheets_used.append(sn)
        if first_tiers is None:
            first_tiers = disc['tiers']
    print(f"  价格Sheet: {sheets_used}")
    if first_tiers:
        print(f"  吨位档 (容积下限, 吨位名, 系数):")
        for lim, ton, coef, _ in first_tiers:
            print(f"    ≥{lim:>5} → {ton} (系数 {coef})")
    wb.close()


def inspect_carrier(path):
    print(f"\n[车管目录] {path}")
    if not Path(path).exists():
        print(f"  ❌ 目录不存在")
        return
    files = sorted([f for f in os.listdir(path) if f.endswith('.xlsx') and not f.startswith('~$')])
    print(f"  文件数: {len(files)}")
    for f in files[:3]:
        wb = openpyxl.load_workbook(os.path.join(path, f), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        hr = find_header_row(ws, ['车牌号', '派车单号', '邮运量'])
        print(f"  示例 {f}: 表头行={hr}")
        wb.close()
    print(f"  公司列表: {[os.path.splitext(f)[0].split(' (')[0] for f in files]}")


def main():
    if len(sys.argv) < 6:
        print("用法: python3 inspect_sources.py <草滩表> <港务表> <车型表> <一干价格表> <二干价格表> <车管目录>")
        sys.exit(1)
    caotan, gangwu, baseline, yigan, ergan, carrier = sys.argv[1:7]

    print("=" * 60)
    print("源文件结构探查（动态识别结果）")
    print("=" * 60)

    inspect_online(caotan, '草滩在线表', r'^(\d+)月(\d+)日?$')
    inspect_online(gangwu, '港务在线表', r'^(\d+)\.(\d+)$')
    inspect_baseline(baseline)
    inspect_price(yigan, '一干价格表')
    inspect_price(ergan, '二干价格表')
    inspect_carrier(carrier)

    print(f"\n{'='*60}")
    print("探查完成。如果以上识别和实际不符，说明源文件改版，先核对再处理。")
    print("=" * 60)


if __name__ == '__main__':
    main()
