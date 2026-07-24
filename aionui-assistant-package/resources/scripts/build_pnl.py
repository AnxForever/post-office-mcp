#!/usr/bin/env python3
"""
子skill·盈亏 —— 邮局月度报表 Step 8，从「人工修正后的初版」+ 财务收入表生成两干线盈亏表。
- 装载量/运费 来自初版「基础报表」sheet 的 运输量/运价（按 路线+吨位 聚合）
- 趟次/结算收入 来自财务收入表（结算收入 = Col23 - Col31 普服收入）
- 盈亏 = 结算收入 - 运费
格式：照搬盈亏模板（列数/表头/路线顺序/万元数字格式/SUM公式）。

用法:
  python3 build_pnl.py <月份> <输出目录> <修正后初版.xlsx> <财务收入表.xlsx> <盈亏模板目录>
  盈亏模板目录须含: 一干委办盈亏.xlsx  二干委办盈亏.xlsx

输出: <输出目录>/<月>月_04_一干盈亏表.xlsx  +  <月>月_04_二干盈亏表.xlsx
末尾自动做 3 条路线抽样验证（趟次/结算收入/运费）。
"""
import sys, os, re
from copy import copy
from collections import defaultdict
import openpyxl


from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

WANYUAN = '0"."0,"万元"'
REPORT_YEAR = int(os.environ.get('POST_OFFICE_REPORT_YEAR', '2026'))


def std(rt, to_dest=False):
    r = rt.strip(); r = re.sub(r'\([^)]*\)', '', r).strip()
    r = re.sub(r'(港|航|旺临|特快|港航|港T串?|陕T?|陕串?|车次|港串?|提速|陕临)\d*$', '', r).strip()
    r = re.sub(r'(转|本)$', '', r).strip(); r = re.sub(r'废\d+$', '', r).strip()
    r = r.replace('航站-', '-').replace('西安草滩-', '西安-')
    r = r.replace('-上海东','-上海').replace('-上海西','-上海')
    if to_dest and r.startswith('西安-'): r = r[3:]
    return r


def _col_by_header(ws, hrow=1):
    """按表头名 → 列号映射（不写死列序）。"""
    return {str(ws.cell(hrow, c).value or '').strip(): c for c in range(1, ws.max_column+1)}

def _find_col(ws, hrow, keywords):
    """在 hrow 行按关键词找列号（首个命中）。"""
    for c in range(1, ws.max_column+1):
        h = str(ws.cell(hrow, c).value or '')
        if any(kw in h for kw in keywords): return c
    return None

def month_template_sheet(wb):
    """选工作簿中最后一个有效月份 Sheet（例如 26年3月），不取 Sheet1 等辅助页。"""
    candidates = []
    for sn in wb.sheetnames:
        if not re.match(r'^(?:\d{2}年)?\d+月$', sn):
            continue
        ws = wb[sn]
        if str(ws.cell(row=2, column=1).value or '').strip() == '邮路名称':
            candidates.append(sn)
    if not candidates:
        raise ValueError('盈亏模板中未找到有效月份 Sheet（Row2 Col1 应为“邮路名称”）')
    return candidates[-1]


def parse_template(path):
    """读取最近月份模板的车型组、路线顺序、汇总区和环比区。"""
    wb = openpyxl.load_workbook(path, data_only=False)
    wb_values = openpyxl.load_workbook(path, data_only=True)
    sn = month_template_sheet(wb)
    ws = wb[sn]; ws_values = wb_values[sn]
    group_starts = []
    summary_start = None
    for col in range(2, ws.max_column + 1):
        value = str(ws.cell(row=2, column=col).value or '').strip()
        if value in ('趟次', '委办运费(元)', '结算收入(元)', '总盈亏(元)'):
            summary_start = summary_start or col
        elif value:
            group_starts.append((col, value))
    if not group_starts or summary_start is None:
        raise ValueError(f'{path} / {sn} 未识别到车型组或汇总区')
    for i, (start, name) in enumerate(group_starts):
        end = group_starts[i+1][0] if i+1 < len(group_starts) else summary_start
        if end - start != 7:
            raise ValueError(f'{path} / {sn} 车型组“{name}”不是 7 列')
    ton_keys = []
    for _, name in group_starts:
        m = re.search(r'(\d+\.?\d*)', name)
        if not m:
            raise ValueError(f'{path} / {sn} 无法从车型组“{name}”解析吨位')
        num = m.group(1)
        suffix_match = re.search(r'([A-B])', name)
        suffix = suffix_match.group(1) if suffix_match else ''
        ton_keys.append(f'{num}吨{suffix}'.rstrip())
    total_row = next((r for r in range(4, ws.max_row + 1) if str(ws.cell(r, 1).value or '').strip() == '合计'), None)
    if total_row is None:
        raise ValueError(f'{path} / {sn} 缺少合计行')
    routes = [str(ws.cell(r, 1).value or '').strip() for r in range(4, total_row) if str(ws.cell(r, 1).value or '').strip()]
    mom_row = next((r for r in range(total_row + 1, ws.max_row + 1) if str(ws.cell(r, 1).value or '').strip() == '环比'), None)
    compare = None
    if mom_row:
        header_row, previous_row, current_row = mom_row - 3, mom_row - 2, mom_row - 1
        previous_values = [ws_values.cell(current_row, c).value for c in range(2, 7)]
        if any(value is None for value in previous_values):
            trip_cols = [start for start, _ in group_starts]
            transport_cols = [start + 1 for start, _ in group_starts]
            freight_cols = [start + 3 for start, _ in group_starts]
            settle_cols = [start + 4 for start, _ in group_starts]
            trips = sum(float(ws_values.cell(total_row, c).value or 0) for c in trip_cols)
            transport = sum(float(ws_values.cell(total_row, c).value or 0) for c in transport_cols)
            freight = sum(float(ws_values.cell(total_row, c).value or 0) for c in freight_cols)
            settle = sum(float(ws_values.cell(total_row, c).value or 0) for c in settle_cols)
            previous_values = [trips, transport, freight, settle, settle - freight]
        compare = {
            'offset': header_row - total_row,
            'headers': [ws.cell(header_row, c).value for c in range(1, 7)],
            'previous_label': ws.cell(current_row, 1).value,
            'previous_values': previous_values,
            'style_rows': (header_row, previous_row, current_row, mom_row),
        }
    return {
        'workbook': wb, 'values_workbook': wb_values, 'sheet': sn, 'ws': ws,
        'group_starts': group_starts, 'ton_keys': ton_keys, 'summary_start': summary_start,
        'ncols': summary_start + 3, 'routes': routes, 'total_row': total_row, 'compare': compare,
    }


def route_aliases(route):
    full = std(str(route or ''))
    dest = full.split('-')[-1] if '-' in full else full
    return {x for x in (full, dest) if x}


def template_route_aliases(routes):
    aliases = set()
    for route in routes:
        aliases.update(route_aliases(route))
    return aliases


def classify_route(route, yg_aliases, er_aliases):
    aliases = route_aliases(route)
    yg_hits = aliases & yg_aliases
    er_hits = aliases & er_aliases
    full = std(str(route or ''))
    if full in yg_aliases and full not in er_aliases:
        return True
    if full in er_aliases and full not in yg_aliases:
        return False
    if yg_hits and not er_hits:
        return True
    if er_hits and not yg_hits:
        return False
    return None


def load_basic_sum(path, yg_routes, er_routes):
    """从修正后初版聚合；一/二干按最新模板路线动态判定，歧义项不静默灌入。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['基础报表'] if '基础报表' in wb.sheetnames else wb[wb.sheetnames[-1]]
    H = _col_by_header(ws)
    c_route, c_ton = H.get('邮路名称'), H.get('吨位')
    c_tran, c_frt = H.get('运输量'), H.get('运价')
    basic = defaultdict(lambda: {'transport':0.0,'freight':0.0})
    unresolved = []
    yg_aliases = template_route_aliases(yg_routes)
    er_aliases = template_route_aliases(er_routes)
    for row in range(2, ws.max_row+1):
        route = str(ws.cell(row, c_route).value or "").strip() if c_route else ""
        ton = ws.cell(row, c_ton).value if c_ton else None
        if not route or not ton: continue
        transport = (ws.cell(row, c_tran).value if c_tran else 0) or 0
        freight = (ws.cell(row, c_frt).value if c_frt else 0) or 0
        is_yg = classify_route(route, yg_aliases, er_aliases)
        if is_yg is None:
            unresolved.append(route)
            continue
        try: transport = float(transport)
        except: transport = 0.0
        try: freight = float(freight)
        except: freight = 0.0
        for alias in route_aliases(route):
            basic[(is_yg, alias, str(ton).strip())]['transport'] += transport
            basic[(is_yg, alias, str(ton).strip())]['freight'] += freight
    wb.close(); return basic, sorted(set(unresolved))


def rated_ton(rated, ton_keys):
    """按模板车型组动态映射额定吨位；同数值 A/B 时 A 为上限、B 为兜底。"""
    try:
        value = float(rated)
    except (TypeError, ValueError):
        return None
    if value <= 0:
        return None
    parsed = []
    for key in ton_keys:
        m = re.search(r'(\d+\.?\d*)', key)
        if m:
            parsed.append((float(m.group(1)), key))
    for limit, key in parsed[:-1]:
        if value <= limit:
            return key
    return parsed[-1][1] if parsed else None


def load_fin(path, yg_tons, er_tons):
    """财务收入表 → (is_yg, 路线, 吨位) → {trips, settle, pufu}。列按表头动态定位。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hr = 1
    c_lv = _find_col(ws, hr, ['邮路级别'])
    c_rt = _find_col(ws, hr, ['邮路名称'])
    c_rated = _find_col(ws, hr, ['额定吨位'])
    c_settle = _find_col(ws, hr, ['结算收入'])
    c_pufu = _find_col(ws, hr, ['普服收入'])
    print(f"  [发现] 财务表列: 邮路级别=c{c_lv} 邮路名称=c{c_rt} 额定吨位=c{c_rated} 结算收入=c{c_settle} 普服收入=c{c_pufu}")
    fin = defaultdict(lambda: {'trips':0,'settle':0.0,'pufu':0.0})
    for row in range(hr+1, ws.max_row+1):
        lv = str(ws.cell(row, c_lv).value or '') if c_lv else ''; is_yg = '一干' in lv
        rt = str(ws.cell(row, c_rt).value or '').strip() if c_rt else ''
        rated = ws.cell(row, c_rated).value if c_rated else None
        settle = (ws.cell(row, c_settle).value if c_settle else 0) or 0
        pufu = (ws.cell(row, c_pufu).value if c_pufu else 0) or 0
        ton = rated_ton(rated, yg_tons if is_yg else er_tons)
        if ton is None: continue
        r = std(rt)
        if is_yg and not r.startswith('西安'): continue
        for alias in route_aliases(r):
            k = (is_yg, alias, ton)
            fin[k]['trips'] += 1; fin[k]['settle'] += float(settle); fin[k]['pufu'] += float(pufu)
    wb.close(); return fin


def copy_cell_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format


def main():
    if len(sys.argv) < 6:
        print(__doc__); sys.exit(1)
    M = int(sys.argv[1]); OUT = sys.argv[2]
    DRAFT = sys.argv[3]; FIN_REVENUE = sys.argv[4]; PNL_DIR = sys.argv[5]
    os.makedirs(OUT, exist_ok=True)

    B = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

    print("读取修正后初版 + 财务表...")
    templates = {
        '一干': parse_template(f"{PNL_DIR}/一干委办盈亏.xlsx"),
        '二干': parse_template(f"{PNL_DIR}/二干委办盈亏.xlsx"),
    }
    print(f"  [模板] 一干={templates['一干']['sheet']} 车型组={templates['一干']['ton_keys']}")
    print(f"  [模板] 二干={templates['二干']['sheet']} 车型组={templates['二干']['ton_keys']}")
    basic_sum, unresolved = load_basic_sum(DRAFT, templates['一干']['routes'], templates['二干']['routes'])
    fin = load_fin(FIN_REVENUE, templates['一干']['ton_keys'], templates['二干']['ton_keys'])
    print(f"  初版聚合键={len(basic_sum)} 财务聚合键={len(fin)} 未归类路线={len(unresolved)}")
    if unresolved:
        print("  [!]️ 未在一/二干最新模板中唯一归类的路线（未灌入盈亏表）:")
        for route in unresolved[:20]:
            print(f"    - {route}")

    filled = []  # 模板内实际落数的行，供末尾抽样验证

    for level, is_yg in [('一干', True), ('二干', False)]:
        print(f"Step 8: {level}盈亏...")
        tpl = templates[level]
        ws_std = tpl['ws']; ncols = tpl['ncols']
        veh_groups = tpl['group_starts']; n_groups = len(veh_groups)
        summary_start = tpl['summary_start']; route_order = tpl['routes']
        group_ton_keys = tpl['ton_keys']

        wb_o = openpyxl.Workbook(); ws_o = wb_o.active; ws_o.title = f"{REPORT_YEAR % 100}年{M}月"
        ws_o.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        c = ws_o.cell(row=1, column=1, value=f'{REPORT_YEAR}年{M}月{level}委办运输盈亏情况')
        copy_cell_style(ws_std.cell(1, 1), c)
        if not c.has_style:
            c.font = Font(name='宋体', bold=True, size=48); c.alignment = ALIGN

        ws_o.merge_cells(start_row=2, start_column=1, end_row=3, end_column=1)
        c = ws_o.cell(row=2, column=1, value='邮路名称'); copy_cell_style(ws_std.cell(2, 1), c)
        for gi, (sc, name) in enumerate(veh_groups):
            ws_o.merge_cells(start_row=2, start_column=2+gi*7, end_row=2, end_column=2+gi*7+6)
            c = ws_o.cell(row=2, column=2+gi*7, value=name); copy_cell_style(ws_std.cell(2, sc), c)
        summary_labels = [ws_std.cell(row=2, column=summary_start+j).value for j in range(4)]
        for j, label in enumerate(summary_labels):
            ws_o.merge_cells(start_row=2, start_column=summary_start+j, end_row=3, end_column=summary_start+j)
            c = ws_o.cell(row=2, column=summary_start+j, value=label); copy_cell_style(ws_std.cell(2, summary_start+j), c)

        for gi in range(n_groups):
            src_start = veh_groups[gi][0]
            for j in range(7):
                s = ws_std.cell(row=3, column=src_start+j).value
                c = ws_o.cell(row=3, column=2+gi*7+j, value=s); copy_cell_style(ws_std.cell(3, src_start+j), c)

        row = 4
        totals = {k:{'trips':0,'transport':0,'freight':0,'settle':0,'pnl':0} for k in group_ton_keys}
        for route_name in route_order:
            c = ws_o.cell(row=row, column=1, value=route_name); copy_cell_style(ws_std.cell(4, 1), c)
            for gi, ton_key in enumerate(group_ton_keys):
                route_key = std(route_name)
                fk = (is_yg, route_key, ton_key)
                trips = fin[fk]['trips']; net_settle = fin[fk]['settle'] - fin[fk]['pufu']
                transport = basic_sum[fk]['transport']; freight = basic_sum[fk]['freight']
                pnl = net_settle - freight; avg_load = transport/trips if trips>0 else 0; avg_pnl = pnl/trips if trips>0 else 0
                if trips > 0 and freight > 0:
                    filled.append((level, route_name, ton_key, trips, round(net_settle,2), round(freight,2)))
                sc2 = 2+gi*7
                for j, v in enumerate([trips, round(transport), round(avg_load,1), round(freight,2), round(net_settle,2), round(pnl,2), round(avg_pnl,2)]):
                    cc = ws_o.cell(row=row, column=sc2+j, value=v)
                    copy_cell_style(ws_std.cell(4, veh_groups[gi][0]+j), cc)
                    if j == 0: cc.number_format = '0_ '
                for k in totals[ton_key]: totals[ton_key][k] += {'trips':trips,'transport':transport,'freight':freight,'settle':net_settle,'pnl':pnl}[k]

            vt = [2+gi*7 for gi in range(n_groups)]; vf = [2+gi*7+3 for gi in range(n_groups)]; vs = [2+gi*7+4 for gi in range(n_groups)]
            for j, (cols, fmt) in enumerate([(vt, None), (vf, WANYUAN), (vs, WANYUAN)]):
                col = summary_start+j; formula = '=' + '+'.join([f'{get_column_letter(cc)}{row}' for cc in cols])
                cc = ws_o.cell(row=row, column=col, value=formula); copy_cell_style(ws_std.cell(4, col), cc)
                if fmt: cc.number_format = fmt
            cc = ws_o.cell(row=row, column=summary_start+3, value=f'={get_column_letter(summary_start+2)}{row}-{get_column_letter(summary_start+1)}{row}')
            copy_cell_style(ws_std.cell(4, summary_start+3), cc); cc.number_format = WANYUAN
            row += 1

        c = ws_o.cell(row=row, column=1, value='合计'); copy_cell_style(ws_std.cell(tpl['total_row'], 1), c)
        for gi, ton_key in enumerate(group_ton_keys):
            t = totals[ton_key]; sc2 = 2+gi*7
            al = t['transport']/t['trips'] if t['trips']>0 else 0; ap = t['pnl']/t['trips'] if t['trips']>0 else 0
            for j, v in enumerate([t['trips'], round(t['transport']), round(al,1), round(t['freight'],2), round(t['settle'],2), round(t['pnl'],2), round(ap,2)]):
                cc = ws_o.cell(row=row, column=sc2+j, value=v)
                copy_cell_style(ws_std.cell(tpl['total_row'], veh_groups[gi][0]+j), cc)
                if j == 0: cc.number_format = '0_ '
        for j, label in enumerate(['trips','freight','settle','pnl']):
            col = summary_start+j
            if label == 'pnl':
                formula = f'={get_column_letter(summary_start+2)}{row}-{get_column_letter(summary_start+1)}{row}'
            else:
                formula = f'=SUM({get_column_letter(col)}{4}:{get_column_letter(col)}{row-1})'
            cc = ws_o.cell(row=row, column=col, value=formula); copy_cell_style(ws_std.cell(tpl['total_row'], col), cc)
            if j >= 1: cc.number_format = WANYUAN

        # 底部环比区：上月取最近模板“本月”行，本月取本次汇总，公式按官方格式。
        compare = tpl['compare']
        if compare:
            header_row = row + compare['offset']
            prev_row, curr_row, mom_row = header_row + 1, header_row + 2, header_row + 3
            src_header, src_prev, src_curr, src_mom = compare['style_rows']
            for cidx in range(1, 7):
                c = ws_o.cell(header_row, cidx, value=compare['headers'][cidx-1]); copy_cell_style(ws_std.cell(src_header, cidx), c)
            ws_o.cell(prev_row, 1, value=compare['previous_label'])
            current_values = [
                sum(t['trips'] for t in totals.values()),
                round(sum(t['transport'] for t in totals.values())),
                round(sum(t['freight'] for t in totals.values()), 2),
                round(sum(t['settle'] for t in totals.values()), 2),
                round(sum(t['pnl'] for t in totals.values()), 2),
            ]
            for cidx in range(1, 7):
                copy_cell_style(ws_std.cell(src_prev, cidx), ws_o.cell(prev_row, cidx))
                copy_cell_style(ws_std.cell(src_curr, cidx), ws_o.cell(curr_row, cidx))
                copy_cell_style(ws_std.cell(src_mom, cidx), ws_o.cell(mom_row, cidx))
            for cidx, value in enumerate(compare['previous_values'], 2):
                ws_o.cell(prev_row, cidx, value=value)
            ws_o.cell(curr_row, 1, value=f'{M}月')
            for cidx, value in enumerate(current_values, 2):
                ws_o.cell(curr_row, cidx, value=value)
            ws_o.cell(mom_row, 1, value='环比')
            for cidx in range(2, 7):
                letter = get_column_letter(cidx)
                ws_o.cell(mom_row, cidx, value=f'=({letter}{curr_row}-{letter}{prev_row})/{letter}{prev_row}')

        for col in range(1, ncols+1):
            letter = get_column_letter(col); src_w = ws_std.column_dimensions[letter].width
            if src_w: ws_o.column_dimensions[letter].width = src_w
        last_row = mom_row if compare else row
        for r in range(2, last_row+1):
            for c in range(1, ncols+1): ws_o.cell(row=r, column=c).border = B
        fname = f'{OUT}/{M}月_04_{level}盈亏表.xlsx'
        wb_o.save(fname); wb_o.close()
        print(f'  [OK] {fname}')

    for tpl in templates.values():
        tpl['workbook'].close(); tpl['values_workbook'].close()

    # ===== 抽样验证：从「模板内实际落数」的行里抽 3 条对账 =====
    print("\n[验证] 从盈亏表实际落数行抽 3 条对账（趟次/结算收入/运费均应>0）...")
    if not filled:
        print("  [!]️ 盈亏表零落数！请查：路线标准化 / 吨位映射 / 一二干判定")
    else:
        idxs = sorted(set([0, len(filled)//2, len(filled)-1]))
        all_ok = True
        for i in idxs:
            level, rt, ton, trips, settle, freight = filled[i]
            ok = trips > 0 and settle != 0 and freight > 0
            all_ok = all_ok and ok
            print(f"  [{level}] {rt} {ton}: 趟次={trips} 结算收入={settle} 运费={freight} => {'PASS' if ok else 'FAIL'}")
        print(f"  覆盖：{len(filled)} 个(路线×吨位)落数 | " + ("[OK] 抽样通过" if all_ok else "[!]️ 有空值需核对"))


if __name__ == "__main__":
    main()
