#!/usr/bin/env python3
"""
子skill·车型汇总表 —— 邮局月度报表 Step 1，合并草滩+港务两个在线表当月Sheet，生成车型汇总表。
车型汇总表是交付产物①，也是初版的容积来源。
跨场地补容间：草滩记录→港务测了→取港务容积（串行车辆）。

用法:
  python3 build_summary.py <月份> <输出目录> <草滩在线表> <港务在线表>

输出: <输出目录>/<月>月_01_车型汇总表.xlsx
"""
import sys, os, re
import openpyxl
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Border, Side



HDR = ['序号', '邮路', '车号', '挂车号', '长（m）', '宽（m）', '高（m）', '测量容积（方）', '预报载重（吨）', '未装容积（米）', '车型问题']


def extract(path, pat_str, month, src):
    """从在线表提取目标月所有记录, 返回records列表。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    pat = re.compile(pat_str)
    records = []
    for wks in wb.worksheets:
        sn = str(wks.title).strip()
        m = pat.match(sn)
        if not m or int(m.group(1)) != month:
            continue
        rows = []
        for i, row in enumerate(wks.iter_rows(values_only=True)):
            rows.append(row)
            if i >= 3:
                break
        hr = None
        for ri in range(min(4, len(rows))):
            joined = ' '.join(str(x or '') for x in rows[ri])
            if '车' in joined and ('邮路' in joined or '序号' in joined):
                hr = ri
                break
        if hr is None or hr >= len(rows):
            continue
        col = {}
        for c, h in enumerate(rows[hr]):
            h = str(h or '')
            if ('车厢' in h or '挂车' in h) and 'trailer' not in col:
                col['trailer'] = c
            if ('车号' in h or '车牌' in h) and 'plate' not in col:
                col['plate'] = c
            if '邮路' in h and 'route' not in col:
                col['route'] = c
            if h == '长' and 'len' not in col:
                col['len'] = c
            if h == '宽' and 'wid' not in col:
                col['wid'] = c
            if h == '高' and 'hgt' not in col:
                col['hgt'] = c
            if '容积' in h and 'vol' not in col:
                col['vol'] = c
            if ('车型问题' in h or h == '问题') and 'issue' not in col:
                col['issue'] = c
            if '预报载重' in h and 'forecast' not in col:
                col['forecast'] = c
            if ('未装容积' in h or h == '未装') and 'unused' not in col:
                col['unused'] = c
        if 'plate' not in col:
            continue
        date_str = '2026-%02d-%02d' % (month, int(m.group(2)))
        for row in wks.iter_rows(min_row=hr + 2, values_only=True):
            if not row:
                continue
            gv = lambda k: row[col[k]] if k in col and col[k] < len(row) else None
            pl = str(gv('plate') or '').strip()
            rt = str(gv('route') or '').strip()
            if not pl and not rt:
                continue
            vol = gv('vol')
            try:
                vol = float(vol) if vol is not None else None
            except Exception:
                vol = None
            # 提取容积为空时,用长*宽*高计算(草滩公式列data_only取不到缓存值)
            if (vol is None or vol == 0):
                lv, wv, hv = gv('len'), gv('wid'), gv('hgt')
                try:
                    if lv and wv and hv: vol = float(lv) * float(wv) * float(hv)
                except: pass
            records.append({
                'route': gv('route'), 'plate': pl.upper(),
                'trailer': gv('trailer'), 'len': gv('len'),
                'wid': gv('wid'), 'hgt': gv('hgt'),
                'vol': vol, 'issue': gv('issue'),
                'forecast': gv('forecast'), 'unused': gv('unused'),
                'cross_filled': False,
            })
    wb.close()
    return records


def main():
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)
    M = int(sys.argv[1]); OUT = sys.argv[2]
    CAOTAN, GANGWU = sys.argv[3], sys.argv[4]
    os.makedirs(OUT, exist_ok=True)

    print('Step 1: 提取在线表...')
    ct = extract(CAOTAN, r'^(\d+)月(\d+)日?$', M, '草滩')
    gw = extract(GANGWU, r'^(\d+)\.(\d+)$', M, '港务')
    print('  草滩=%d 港务=%d' % (len(ct), len(gw)))

    # 跨场地补容间(串行): 草滩记录vol为None→同车牌在港务有值的补上
    print('Step 2: 跨场地补容间（串行：草滩记录→港务容积）...')
    # 建港务 plate→vol索引(跳过None/0)
    gw_vol = {}
    for r in gw:
        if r['vol'] and r['vol'] > 0:
            p = r['plate']
            if p not in gw_vol:
                gw_vol[p] = r['vol']
    filled = 0
    for r in ct:
        if (r['vol'] is None or r['vol'] == 0) and r['plate'] in gw_vol:
            r['vol'] = gw_vol[r['plate']]
            r['cross_filled'] = True
            filled += 1
    if filled:
        print('  草滩记录从港务补容积: %d条' % filled)

    # 合并写输出
    print('Step 3: 生成车型汇总表...')
    B = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    FH = PatternFill('solid', fgColor='4472C4')
    FONT_H = Font(color='FFFFFF', bold=True)

    wb_o = openpyxl.Workbook(); ws = wb_o.active; ws.title = '车型汇总'
    for c, h in enumerate(HDR, 1):
        cl = ws.cell(1, c, h); cl.fill = FH; cl.font = FONT_H; cl.border = B

    seq = 1
    for r in ct + gw:
        ws.cell(seq + 1, 1, seq).border = B
        ws.cell(seq + 1, 2, r['route']).border = B
        ws.cell(seq + 1, 3, r['plate']).border = B
        ws.cell(seq + 1, 4, r['trailer']).border = B
        ws.cell(seq + 1, 5, r['len']).border = B
        ws.cell(seq + 1, 6, r['wid']).border = B
        ws.cell(seq + 1, 7, r['hgt']).border = B
        ws.cell(seq + 1, 8, r['vol']).border = B
        ws.cell(seq + 1, 9, r['forecast']).border = B
        ws.cell(seq + 1, 10, r['unused']).border = B
        ws.cell(seq + 1, 11, r['issue']).border = B
        seq += 1

    for c, w in enumerate([6, 28, 12, 14, 12, 12, 12, 14, 12, 12, 20], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    F = os.path.join(OUT, '%d月_01_车型汇总表.xlsx' % M)
    wb_o.save(F); wb_o.close()
    total = len(ct) + len(gw)
    print('  草滩=%d 港务=%d 合并=%d' % (len(ct), len(gw), total))
    if filled:
        print('  其中草滩%d条容积从港务补(串行)' % filled)
    print('\n[OK] 车型汇总表已生成: %s' % F)
    print('[STOP] 车型汇总表是初版的容积来源，下一步用 build_draft 生成初版。')


if __name__ == '__main__':
    main()
