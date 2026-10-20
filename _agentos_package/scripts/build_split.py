#!/usr/bin/env python3
"""
子skill·拆分 —— 邮局月度报表 Step 7，从「人工修正后的初版」生成拆分版。
数据来源：修正后初版的「基础报表」sheet（21列 v1.16 契约），不再重算价格/车管。

用法:
  python3 build_split.py <月份> <输出目录> <修正后初版.xlsx> <拆分模板(基础报表原版).xlsx>

输出: <输出目录>/<月>月_03_基础报表拆分版.xlsx
"""
import sys, os, re
from copy import copy
from collections import defaultdict
import openpyxl


from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

from _paths import require_file, resolve_out, save_workbook


def std(rt, to_dest=False):
    r = rt.strip(); r = re.sub(r'\([^)]*\)', '', r).strip()
    r = re.sub(r'(港|航|旺临|特快|港航|港T串?|陕T?|陕串?|车次|港串?|提速|陕临)\d*$', '', r).strip()
    r = re.sub(r'(转|本)$', '', r).strip(); r = re.sub(r'废\d+$', '', r).strip()
    r = r.replace('航站-', '-').replace('西安草滩-', '西安-')
    r = r.replace('-上海东','-上海').replace('-上海西','-上海')
    if to_dest and r.startswith('西安-'): r = r[3:]
    return r


def _col_by_header(ws, hrow=1):
    """基础报表 sheet：按表头名 → 列号映射（不写死列序）。"""
    return {str(ws.cell(hrow, c).value or '').strip(): c for c in range(1, ws.max_column+1)}


SPECIAL_SHEETS = {'降吨位', '二干基础', '两趟合一趟', '一干汇总'}
BLANK_SHEETS = {'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4'}


def effective_headers(ws):
    """返回模板第一行直到最后一个非空表头（保留中间占位空列）。"""
    last = 0
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value is not None:
            last = col
    return [str(ws.cell(row=1, column=col).value or '').strip() for col in range(1, last + 1)]


def copy_header_layout(src, dst, headers):
    """复制模板表头样式、列宽和页面设置，不复制历史数据。"""
    for col, value in enumerate(headers, 1):
        sc = src.cell(row=1, column=col)
        dc = dst.cell(row=1, column=col, value=value or None)
        if sc.has_style:
            dc.font = copy(sc.font)
            dc.fill = copy(sc.fill)
            dc.border = copy(sc.border)
            dc.alignment = copy(sc.alignment)
            dc.protection = copy(sc.protection)
        if sc.number_format:
            dc.number_format = sc.number_format
        letter = get_column_letter(col)
        if src.column_dimensions[letter].width:
            dst.column_dimensions[letter].width = src.column_dimensions[letter].width
    dst.freeze_panes = src.freeze_panes
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines
    dst.page_margins = copy(src.page_margins)
    dst.page_setup = copy(src.page_setup)


def template_hints(wb, sheet_names):
    """从模板已有数据学习 (标准路线, 公司) → Sheet，避免固原/天水等双 Sheet 误分。"""
    exact = {}
    route_only = defaultdict(set)
    for sname in sheet_names:
        ws = wb[sname]
        headers = effective_headers(ws)
        route_cols = [i + 1 for i, h in enumerate(headers) if h == '邮路名称']
        company_col = next((i + 1 for i, h in enumerate(headers) if h == '车辆承运单位'), None)
        if not route_cols:
            continue
        route_col = route_cols[0]
        blank_run = 0
        # 有些模板因格式污染 max_row 超过百万；连续空 50 行即停止。
        for row in ws.iter_rows(min_row=2, values_only=True):
            route = str(row[route_col - 1] or '').strip() if len(row) >= route_col else ''
            company = str(row[company_col - 1] or '').strip() if company_col and len(row) >= company_col else ''
            if not route:
                blank_run += 1
                if blank_run >= 50:
                    break
                continue
            blank_run = 0
            key = std(route)
            exact[(key, company)] = sname
            route_only[key].add(sname)
    unique_route = {route: next(iter(names)) for route, names in route_only.items() if len(names) == 1}
    return exact, unique_route


def record_values(headers, rec, seq):
    """按模板表头语义生成一行；支持 17 列和带双运价的 18 列模板。"""
    result = []
    seen = defaultdict(int)
    vol = rec.get('final_vol')
    if isinstance(vol, (int, float)) and vol > 0:
        vol = int(round(vol))
    else:
        vol = None
    freight = rec.get('freight')
    rounded_freight = round(freight, 2) if isinstance(freight, (int, float)) else freight
    for header in headers:
        seen[header] += 1
        if header == '序号': value = seq
        elif header == '车辆承运单位': value = rec.get('company')
        elif header == '派车单号': value = rec.get('dispatch')
        elif header == '邮路名称': value = rec.get('route') if seen[header] == 1 else std(rec.get('route') or '')
        elif header == '路单编号': value = rec.get('waybill_no')
        elif header == '封车时间': value = rec.get('seal_date_str')
        elif header == '车牌号': value = rec.get('plate')
        elif header == '容间': value = vol
        elif header == '吨位': value = rec.get('ton')
        elif header == '公里': value = rec.get('km')
        elif header == '运输量': value = rec.get('transport_qty')
        elif header == '路单号': value = rec.get('waybill_id')
        elif header == '运价（保留两位数）': value = rounded_freight
        elif header == '运价': value = rounded_freight
        elif header == '标准价': value = rec.get('price')
        elif header == '备注': value = rec.get('issue')
        else: value = None
        result.append(value)
    return result

def load_draft(path):
    """从修正后初版的「基础报表」sheet 按【表头名】读回 records（不依赖固定列序）。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['基础报表'] if '基础报表' in wb.sheetnames else wb[wb.sheetnames[-1]]
    H = _col_by_header(ws)
    def g(row, name):
        c = H.get(name)
        return ws.cell(row, c).value if c else None
    recs = []
    for row in range(2, ws.max_row+1):
        route = str(g(row, '邮路名称') or "").strip()
        plate = str(g(row, '车牌号') or "").strip()
        if not route and not plate: continue
        recs.append({
            'route': route, 'plate': plate,
            'company': g(row, '车辆承运单位'), 'dispatch': g(row, '派车单号'),
            'seal_date_str': g(row, '封车时间'), 'final_vol': g(row, '容间'),
            'ton': g(row, '吨位'), 'km': g(row, '公里'),
            'transport_qty': g(row, '运输量'), 'freight': g(row, '运价'),
            'price': g(row, '标准价'), 'issue': g(row, '备注'),
            'waybill_no': g(row, '路单编号'), 'waybill_id': g(row, '路单号'),
        })
    wb.close(); return recs


def main():
    if len(sys.argv) < 5:
        print(__doc__); sys.exit(1)
    M = int(sys.argv[1]); OUT = str(resolve_out(sys.argv[2]))
    DRAFT = str(require_file(sys.argv[3], '核对后初版'))
    SPLIT_TEMPLATE = str(require_file(sys.argv[4], '拆分模板'))

    FILL_H = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    FONT_H = Font(color="FFFFFF", bold=True, size=10)
    B = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    print("读取修正后初版...")
    records = load_draft(DRAFT)
    print(f"  记录={len(records)}")

    print("Step 7: 拆分...")
    wb_t = openpyxl.load_workbook(SPLIT_TEMPLATE, data_only=False)
    STD = [s for s in wb_t.sheetnames if s not in SPECIAL_SHEETS | BLANK_SHEETS and effective_headers(wb_t[s])]
    STD.sort(key=lambda s: (len(s), s))  # 短名优先: 银川→银川往返
    tmpl_info = {}
    for sname in STD:
        ws = wb_t[sname]
        hdrs = effective_headers(ws)
        tmpl_info[sname] = {'hdrs': hdrs, 'ncols': len(hdrs)}
    hint_exact, hint_route = template_hints(wb_t, STD)

    def dc(rt):
        d, f = std(rt), std(rt)
        return (d.split('-')[-1] if '-' in d else d), f

    sr = defaultdict(list); ex = defaultdict(list); ambiguous = []; excluded_reverse = 0
    for rec in records:
        d, f = dc(rec['route']); m = None
        is_ambiguous = False
        route_key = std(rec['route'])
        company = str(rec.get('company') or '')
        # 1. 模板hints优先 (处理 平凉→平凉二干 等映射)
        m_exact = hint_exact.get((route_key, company))
        # exact hint指向往返Sheet时不信任(模板写入顺序导致覆盖)
        if m_exact and '往返' in m_exact:
            m_exact = None
        m_route = hint_route.get(route_key) if not m_exact else None
        m = m_exact or m_route
        # route级hint指向往返Sheet时改到基础Sheet
        if m and '往返' in m:
            base = m.replace('往返', '')
            if base in STD: m = base
        # 3. 无hints时用目的地匹配
        if not m:
            exact_dest = [s for s in STD if s.split('（')[0] == d]
            if len(exact_dest) == 1:
                m = exact_dest[0]
            elif len(exact_dest) > 1:
                ambiguous.append((rec['route'], company, exact_dest))
                is_ambiguous = True
        # 4. 最后兜底字符串匹配
        if not m and not is_ambiguous:
            for s in STD:
                if s in rec['route'] or s.split('（')[0] in f:
                    m = s; break
        # 官方规则：排除反向路线（如 宝鸡-西安 不放入宝鸡Sheet）
        if m:
            sheet_city = m.split('（')[0].split('(')[0]
            clean_route = std(rec['route'])
            if clean_route.startswith(sheet_city + '-') and not clean_route.startswith('西安-'):
                excluded_reverse += 1
                continue
            sr[m].append(rec)
        else: ex[f"*{d}"].append(rec)

    sh = STD + sorted(ex.keys())
    SPLIT = f"{OUT}/{M}月_03_基础报表拆分版.xlsx"
    wb_o = openpyxl.Workbook(); wb_o.remove(wb_o.active)
    for sname in sh:
        ws = wb_o.create_sheet(title=sname[:31])
        info = tmpl_info.get(sname, {}); ncols = info.get('ncols', 17)
        hdrs = info.get('hdrs', ['序号','车辆承运单位','派车单号','邮路名称','路单编号','封车时间','','车牌号','','容间','吨位','公里','运输量','路单号','运价','标准价','备注'])
        if sname in wb_t.sheetnames:
            copy_header_layout(wb_t[sname], ws, hdrs)
        else:
            for c, h in enumerate(hdrs, 1):
                cl = ws.cell(row=1, column=c, value=h or None); cl.fill = FILL_H; cl.font = FONT_H; cl.border = B
        rs = sr.get(sname, []) if sname in STD else ex.get(sname, [])
        rw = 2; tt = 0; tf = 0
        for i, rec in enumerate(rs):
            out = record_values(hdrs, rec, i + 1)
            for c, val in enumerate(out, 1):
                if c <= ncols:
                    cell = ws.cell(row=rw, column=c, value=val)
                    if sname in wb_t.sheetnames and wb_t[sname].max_row >= 2:
                        src = wb_t[sname].cell(row=2, column=c)
                        if src.has_style:
                            cell.font = copy(src.font)
                            cell.fill = copy(src.fill)
                            cell.border = copy(src.border)
                            cell.alignment = copy(src.alignment)
                            cell.protection = copy(src.protection)
                            cell.number_format = src.number_format
                    else:
                        cell.border = B
            if isinstance(rec.get('transport_qty'), (int, float)): tt += rec['transport_qty']
            if isinstance(rec.get('freight'), (int, float)): tf += rec['freight']
            rw += 1
        if rs:
            qty_col = next((i + 1 for i, h in enumerate(hdrs) if h == '运输量'), None)
            freight_col = next((i + 1 for i, h in enumerate(hdrs) if h == '运价（保留两位数）'), None)
            if freight_col is None:
                freight_col = next((i + 1 for i, h in enumerate(hdrs) if h == '运价'), None)
            for c, v in [(qty_col, tt), (freight_col, round(tf, 2))]:
                if c is not None and c <= ncols:
                    cl = ws.cell(row=rw, column=c, value=v); cl.border = B; cl.font = Font(bold=True)
        for c, w in enumerate([6,24,16,28,12,14,2,12,2,14,8,8,10,14,10,10,20,20], 1):
            if c <= ncols: ws.column_dimensions[get_column_letter(c)].width = w
    # ====== 降吨位 Sheet ======
    dg_recs = []
    for rec in records:
        issue = str(rec.get('issue') or '')
        if '降吨位' in issue:
            dg_recs.append(rec)
    if dg_recs:
        dg_hdrs = ['序号','车辆承运单位','派车单号','邮路名称','路单编号','封车时间','车牌号','容间','吨位','公里','运输量','运价（保留两位数）','标准价','运价','备注','原运费','差价']
        ws_dg = wb_o.create_sheet(title='降吨位')
        for c, h in enumerate(dg_hdrs, 1):
            cl = ws_dg.cell(row=1, column=c, value=h); cl.fill = FILL_H; cl.font = FONT_H; cl.border = B
        for i, rec in enumerate(dg_recs):
            rw = i + 2
            # 标准列
            vals = record_values(dg_hdrs[:15], rec, i + 1)
            for c, v in enumerate(vals, 1):
                ws_dg.cell(row=rw, column=c, value=v).border = B
            # 吨位列: 显示原车型吨位
            remark = str(rec.get('issue') or '')
            m = re.search(r'车型=(.+?)→', remark)
            ws_dg.cell(row=rw, column=9, value=m.group(1) if m else rec.get('ton')).border = B
            # 原运费 + 差价
            m2 = re.search(r'原运费=([\d.]+).*?差价=([\d.+-]+)', remark)
            if m2:
                ws_dg.cell(row=rw, column=16, value=float(m2.group(1))).border = B
                ws_dg.cell(row=rw, column=17, value=float(m2.group(2))).border = B
        for c, w in enumerate([6,24,16,28,12,14,12,14,8,8,10,14,10,10,20,14,14], 1):
            ws_dg.column_dimensions[get_column_letter(c)].width = w
        print(f"  [降吨位] 生成降吨位Sheet: {len(dg_recs)} 条")

    wb_t.close()
    SPLIT = save_workbook(wb_o, SPLIT)
    print(f"\n[OK] 拆分版已生成: {SPLIT}  ({len(sh) + (1 if dg_recs else 0)} 个 Sheet，其中新建*Sheet {len(ex)} 个)")
    print(f"  [反向] 排除反向路线(如宝鸡-西安): {excluded_reverse} 条")
    if ex:
        print(f"  [!] 模板无对应Sheet、已新建*前缀: {', '.join(sorted(ex.keys()))}")
    if ambiguous:
        print(f"  [!] {len(ambiguous)} 条路线存在多个同城Sheet且模板无唯一历史提示，已放入*Sheet，需人工确认")


if __name__ == "__main__":
    main()
