#!/usr/bin/env python3
"""
子skill·初版 —— 邮局月度报表 Step 1-6，生成基础报表初版。【动态发现版】
一切字段位置、Sheet、吨位档/系数都从实际文件表头运行时发现，不写死列号/档位。
铁律：拿不准就留空并标注原因，绝不编造。生成初版后停止，等人工核对。

用法:
  python3 build_draft.py <月份> <输出目录> <草滩表> <港务表> <一干价格表> <二干价格表> <车管目录>

输出: <输出目录>/<月>月_02_基础报表初版.xlsx  (Sheet: 结构发现 + 问题汇总 + 基础报表)
"""
import sys, os, re
from datetime import datetime, timedelta
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

from _paths import require_dir, require_file, resolve_out, save_workbook



V16_HDR = ['序号','车辆承运单位','派车单号','邮路名称','路单编号','封车时间','','车牌号','','容间','吨位','公里','运输量','路单号','运价','标准价','备注','是否跨档','变动前容间','变动前时间','系数']

# 公司简称→全称映射：从外部 JSON 加载（不写死在代码里）
# JSON 路径：本脚本同级目录的 ../company-map.json
# 新增公司时改 JSON 即可，不用改代码
import json as _json
from pathlib import Path as _Path
_COMPANY_MAP_PATH = _Path(__file__).parent.parent / 'company-map.json'
try:
    _raw = _json.loads(_COMPANY_MAP_PATH.read_text(encoding='utf-8'))
    COMPANY_MAP = _raw.get('mappings', {})
except Exception as _e:
    # JSON 读不到时回退到内置默认（保证脚本能跑）
    print(f"  [警告] 加载 company-map.json 失败 ({_e})，用内置默认")
    COMPANY_MAP = {'则一':'马鞍山则一物流有限公司','富海源':'西安富海源运输有限公司','拓达':'合肥拓达货运有限公司','盛丰':'盛丰物流集团有限公司','胜途':'西安胜途汽车服务有限公司','荣欣':'武汉市荣欣物流有限公司'}

DISCOVERY_LOG = []  # 收集结构发现记录，写入「结构发现」Sheet 供人工核验
def logd(msg): DISCOVERY_LOG.append(msg); print("  [发现]", msg)


# ============ 通用动态发现工具 ============
def find_header_row(ws, keywords, scan=8):
    """返回表头行号：前 scan 行中命中 keywords 最多的行。"""
    best_r, best_hits = 1, -1
    for r in range(1, min(scan, ws.max_row)+1):
        txt = [str(ws.cell(r, c).value or '') for c in range(1, ws.max_column+1)]
        hits = sum(1 for kw in keywords if any(kw in t for t in txt))
        if hits > best_hits: best_hits, best_r = hits, r
    return best_r

def locate_ordered(ws, hrow, ordered_specs):
    """按 (字段,[关键词]) 顺序解析列号；已占用的列不复用（避免'车牌号'误命中'车厢车牌号'）。"""
    hdr = {c: str(ws.cell(hrow, c).value or '') for c in range(1, ws.max_column+1)}
    used, res = set(), {}
    for field, kws in ordered_specs:
        col = None
        for kw in kws:
            for c, h in hdr.items():
                if c in used: continue
                if kw in h: col = c; break
            if col: break
        res[field] = col
        if col: used.add(col)
    return res

def parse_coef(v):
    """系数解析：数字直接用；'系数0.63'/'系数 1' 抽数字；无数字(如'返程')→None。"""
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    m = re.search(r'([\d.]+)', str(v))
    return float(m.group(1)) if m else None

def parse_limit(v):
    """容积下限解析：'(≥55立方米）'/'4.2米(≥16m3)' → 55/16。"""
    if v is None: return 0.0
    m = re.search(r'≥\s*(\d+)', str(v)) or re.search(r'(\d+)\s*(?:立方米|m3|方)', str(v))
    return float(m.group(1)) if m else 0.0

def discover_price_sheet(ws):
    """从价格表某 Sheet 动态发现：(route_col, km_col, data_start, tiers)。
    tiers=[(容积下限, 吨位名, 系数, 列号)]，按下限降序。识别不出返回 None。"""
    ton_row = coef_row = limit_row = None
    for r in range(1, min(8, ws.max_row)+1):
        txt = [str(ws.cell(r, c).value or '') for c in range(1, ws.max_column+1)]
        joined = ' '.join(txt)
        if ton_row is None and sum(1 for t in txt if re.search(r'\d+\.?\d*吨', t)) >= 2: ton_row = r
        if limit_row is None and ('≥' in joined or 'm3' in joined or '立方米' in joined): limit_row = r
        if coef_row is None and '系数' in joined: coef_row = r
    if ton_row is None: return None
    if coef_row is None: coef_row = ton_row - 1
    if limit_row is None: limit_row = ton_row + 1
    tiers = []
    for c in range(1, ws.max_column+1):
        m = re.search(r'(\d+\.?\d*吨[A-B]?)', str(ws.cell(ton_row, c).value or ''))
        if not m: continue
        coef = parse_coef(ws.cell(coef_row, c).value)
        if coef is None: continue  # 排除'返程'等无系数列
        tiers.append((parse_limit(ws.cell(limit_row, c).value), m.group(1), coef, c))
    if not tiers: return None
    route_col = km_col = None
    for hr in [coef_row, ton_row, limit_row]:
        if not (1 <= hr <= ws.max_row): continue
        for c in range(1, ws.max_column+1):
            h = str(ws.cell(hr, c).value or '')
            if route_col is None and any(k in h for k in ['路向', '邮路', '路线']): route_col = c
            if km_col is None and any(k in h for k in ['里程', '公里', 'KM', 'km']): km_col = c
    data_start = max(r for r in [ton_row, coef_row, limit_row] if r) + 1
    tiers.sort(key=lambda t: -t[0])
    return route_col, km_col, data_start, tiers


def std(rt, to_dest=False):
    """路线标准化 L0-L6（见 references/route_normalize.md）"""
    r = rt.strip(); r = re.sub(r'\([^)]*\)', '', r).strip()
    r = re.sub(r'(港|航|旺临|特快|港航|港T串?|陕T?|陕串?|车次|港串?|提速|陕临)\d*$', '', r).strip()
    r = re.sub(r'(转|本)$', '', r).strip(); r = re.sub(r'废\d+$', '', r).strip()
    r = r.replace('航站-', '-').replace('西安草滩-', '西安-')
    r = r.replace('-上海东', '-上海').replace('-上海西', '-上海')
    if to_dest and r.startswith('西安-'): r = r[3:]
    return r


def main():
    if len(sys.argv) < 8:
        print(__doc__); sys.exit(1)
    M = int(sys.argv[1]); OUT = str(resolve_out(sys.argv[2]))
    CAOTAN = str(require_file(sys.argv[3], '草滩在线表'))
    GANGWU = str(require_file(sys.argv[4], '港务在线表'))
    YIGAN_PRICE = str(require_file(sys.argv[5], '一干价格表'))
    ERGAN_PRICE = str(require_file(sys.argv[6], '二干价格表'))
    CARRIER_DIR = str(require_dir(sys.argv[7], '车管基础数据目录'))
    REPORT_YEAR = int(os.environ.get('POST_OFFICE_REPORT_YEAR', '2026'))

    FILL_H = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    FILL_R = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
    FILL_O = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    FILL_Y = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    FILL_WARN = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    FILL_B = PatternFill(start_color="BBD7EE", end_color="BBD7EE", fill_type="solid")  # 浅蓝：串行
    FILL_DG = PatternFill(start_color="C65911", end_color="C65911", fill_type="solid")  # 深橙：降吨位
    FONT_H = Font(color="FFFFFF", bold=True, size=10)
    B = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # ====== Step 1: 提取合并在线表（表头/列动态定位）======
    print("Step 1: 提取合并在线表...")
    ONLINE_SPEC = [('trailer', ['车厢车牌号', '车厢码', '挂车号', '挂车']),
                   ('plate', ['车牌号', '车号']),
                   ('route', ['邮路名称', '邮路', '路线']),
                   ('len', ['长']), ('wid', ['宽']), ('hgt', ['高']),
                   ('vol', ['测量容积', '容积']), ('issue', ['车型问题', '问题'])]
    def extract(fp, src, pat_str):
        wb = openpyxl.load_workbook(fp, data_only=True); recs = []
        pat = re.compile(pat_str); sheets_hit = 0
        for sn in wb.sheetnames:
            m = pat.match(sn.strip())
            if not m or int(m.group(1)) != M: continue
            ws = wb[sn]; sheets_hit += 1
            hr = find_header_row(ws, ['序号', '邮路', '车'])
            col = locate_ordered(ws, hr, ONLINE_SPEC)
            if sheets_hit == 1:
                logd(f"{src}在线表: 表头行={hr}, 列定位={{ {', '.join(f'{k}=c{v}' for k,v in col.items() if v)} }}")
            for r in range(hr+1, ws.max_row+1):
                rt = str(ws.cell(r, col['route']).value or "").strip() if col['route'] else ""
                pl = str(ws.cell(r, col['plate']).value or "").strip() if col['plate'] else ""
                if not rt and not pl: continue
                gv = lambda k: ws.cell(r, col[k]).value if col.get(k) else None
                vol = gv('vol')
                if (vol is None or vol == 0):
                    lv, wv, hv = gv('len'), gv('wid'), gv('hgt')
                    try: vol = float(lv) * float(wv) * float(hv) if lv and wv and hv else vol
                    except: pass
                recs.append({'date': datetime(REPORT_YEAR, M, int(m.group(2))), 'route': rt, 'plate': pl,
                    'trailer': str(gv('trailer') or "").strip(), 'len': gv('len'), 'wid': gv('wid'),
                    'hgt': gv('hgt'), 'vol': vol, 'issue': gv('issue'), 'src': src})
        wb.close(); return recs

    ct = extract(CAOTAN, '草滩', r'^(\d+)月(\d+)日?$')
    gw = extract(GANGWU, '港务', r'^(\d+)\.(\d+)$')
    online = ct + gw
    print(f"  草滩={len(ct)} 港务={len(gw)} 合并={len(online)}")

    # ====== Step 2: 在线表索引（容积+车型问题来源）+ 车管数据（行来源）======
    print("Step 2a: 建立在线表索引（容积+车型问题来源=草滩+港务合并）...")
    ol_by_date_plate = defaultdict(list)  # (日期, 车牌) -> [记录]
    ol_by_plate = defaultdict(list)       # 车牌 -> [记录]
    for r in online:
        if r['plate']:
            p = r['plate'].upper()
            ol_by_plate[p].append(r)
            if r['date']:
                ol_by_date_plate[(r['date'].strftime('%Y-%m-%d'), p)].append(r)
    # 跨场地补容间: 同(日期,车牌)草滩港务可能都有一条,取非零容积(串行: 港务发车→草滩测)
    filled = 0
    for key, recs in ol_by_date_plate.items():
        for rec in recs:
            if rec.get('vol') and rec['vol'] > 0:
                # 对该(日期,车牌)的其他空容积记录,补上非零值
                for other in recs:
                    if other is not rec and (not other.get('vol') or other['vol'] == 0):
                        other['vol'] = rec['vol']
                        other['issue'] = rec.get('issue') or other.get('issue')
                        filled += 1
    logd(f"在线表索引: 唯一车牌={len(ol_by_plate)}, 唯一(日期+车牌)={len(ol_by_date_plate)}"
         + (f", 跨场地补容间={filled}" if filled else ""))

    print("Step 2b: 车管数据（行来源）...")
    CAR_CORE = [('trailer', ['车厢码', '车厢车牌号', '挂车']),
                ('plate', ['车牌号', '车号']), ('route', ['邮路名称', '邮路', '路线']),
                ('dispatch', ['派车单号']), ('seal', ['邮路执行日期', '封车']),
                ('qty', ['邮运量', '运输量']),
                ('settle_ton', ['结算吨位'])]
    records = []; car_logged = False
    for f in sorted(os.listdir(CARRIER_DIR)):
        if not f.endswith('.xlsx') or f.startswith('~$'): continue
        wb_car = openpyxl.load_workbook(os.path.join(CARRIER_DIR, f), data_only=True)
        ws_car = wb_car[wb_car.sheetnames[0]]
        hr_car = find_header_row(ws_car, ['车牌号', '车号', '派车单号', '邮运量'])
        cc = locate_ordered(ws_car, hr_car, CAR_CORE)
        if not car_logged:
            logd(f"车管文件(示例{f}): 表头行={hr_car}, 列定位={{ {', '.join(f'{k}=c{v}' for k,v in cc.items() if v)} }}")
            car_logged = True
        comp_key = re.sub(r'\s*\(\d+\)\s*$', '', f.replace('.xlsx', '')).strip()
        company = COMPANY_MAP.get(comp_key, comp_key)
        for row in range(hr_car+1, ws_car.max_row+1):
            pl = str(ws_car.cell(row, cc['plate']).value or "").strip().upper() if cc.get('plate') else ""
            if not pl: continue
            rt = str(ws_car.cell(row, cc['route']).value or "").strip() if cc.get('route') else ""
            tr = str(ws_car.cell(row, cc['trailer']).value or "").strip().upper() if cc.get('trailer') else ""
            dp = str(ws_car.cell(row, cc['dispatch']).value or "").strip() if cc.get('dispatch') else ""
            dv = ws_car.cell(row, cc['seal']).value if cc.get('seal') else None
            dt = None; sd = None
            if isinstance(dv, datetime): dt = dv; sd = dv.strftime('%Y-%m-%d')
            elif isinstance(dv, str):
                try: dt_v = datetime.strptime(dv[:10], '%Y-%m-%d'); dt = dt_v; sd = dt_v.strftime('%Y-%m-%d')
                except: pass
            records.append({'route': rt, 'plate': pl, 'trailer': tr, 'date': dt,
                'vol': None, 'issue': None,
                'company': company,
                'dispatch': dp.replace('.0', '') if dp.endswith('.0') else dp,
                'seal_date_str': sd,
                'transport_qty': ws_car.cell(row, cc['qty']).value if cc.get('qty') else None,
                'settle_ton': ws_car.cell(row, cc['settle_ton']).value if cc.get('settle_ton') else None})
        wb_car.close()
    logd(f"车管记录共 {len(records)} 条（行来源）")

    # ====== Step 2c: 匹配在线表（容积+车型问题）。串行:同车牌跨站点取非零容积 ====
    print("Step 2c: 匹配在线表（容积+车型问题）...")
    vol_ok=0; vol_online=0; vol_none=0
    for r in records:
        dk = r['date'].strftime('%Y-%m-%d') if r['date'] else None
        exacts = ol_by_date_plate.get((dk, r['plate']), []) if dk else []
        r['measured_other_date'] = False
        r['unmeasured'] = False
        r['vol_cross_station'] = False  # 串行: 容积从同车跨站点取得
        # 优先: (日期+车牌)精确命中且容积>0
        found = False
        if exacts:
            best = next((x for x in exacts if x.get('vol') and x['vol'] > 0), None)
            if best:
                r['vol'] = best['vol']; r['issue'] = best.get('issue')
                found = True
        # 兜底: 同车牌跨站点/跨日期取非零容积(串行车:草滩记录→港务测了)
        if not found:
            others = ol_by_plate.get(r['plate'], [])
            best_other = next((x for x in others if x.get('vol') and x['vol'] > 0), None)
            if best_other:
                r['vol'] = best_other['vol']; r['issue'] = best_other.get('issue')
                r['vol_cross_station'] = True
                found = True
        if not found:
            if ol_by_plate.get(r['plate']):
                r['measured_other_date'] = True
                if not r.get('other_dates'):
                    r['other_dates'] = sorted({rec['date'].strftime('%m/%d') for rec in ol_by_plate[r['plate']] if rec['date']})
            else:
                r['unmeasured'] = True
    print(f"  在线表匹配完成（容积源=草滩+港务合并，串行同车牌跨站点取容积）")

    # 价格表动态发现
    def load_price(fp):
        """返回 {sheet顺序保留}: data={route:{km,ps}}, tiers(取首个价格Sheet)。"""
        wb = openpyxl.load_workbook(fp, data_only=True)
        data = {}; tiers = None; sheets_used = []
        for sn in wb.sheetnames:
            ws = wb[sn]
            disc = discover_price_sheet(ws)
            if not disc: continue
            rc, kc, ds, tr = disc
            if rc is None: continue
            n_routes = 0
            for row in range(ds, ws.max_row+1):
                rt = str(ws.cell(row, rc).value or "").strip()
                if not rt: continue
                n_routes += 1
                ps = {ton: ws.cell(row, c).value for (lim, ton, cf, c) in tr if ws.cell(row, c).value is not None}
                if rt not in data:
                    data[rt] = {'km': ws.cell(row, kc).value if kc else None, 'ps': ps}
            if n_routes == 0: continue  # 空 Sheet（有档位表头但无路线）不计
            sheets_used.append(sn)
            if tiers is None: tiers = tr
        wb.close(); return data, tiers, sheets_used

    yigan_data, yigan_tiers, yg_sheets = load_price(YIGAN_PRICE)
    ergan_data, ergan_tiers, eg_sheets = load_price(ERGAN_PRICE)
    logd(f"一干价格表: 价格Sheet={yg_sheets}, 吨位档={[(t[1],t[2],t[0]) for t in (yigan_tiers or [])]}")
    logd(f"二干价格表: 价格Sheet={eg_sheets}, 吨位档={[(t[1],t[2],t[0]) for t in (ergan_tiers or [])]}")

    # 二干判定：从二干价格表目的地【动态派生】，不写死城市名
    ergan_dests = {std(rt, to_dest=True) for rt in ergan_data.keys()}
    def is_er(rt): return std(rt, to_dest=True) in ergan_dests
    logd(f"二干目的地集(动态派生, {len(ergan_dests)}个): {sorted(ergan_dests)}")

    # ====== 吨位判定（独立于价格匹配）======
    for r in records:
        r['missing'] = False  # 行来源=车管数据，不存在"缺失"
        r['final_vol'] = r['vol']
        ie = is_er(r['route'])
        tiers = ergan_tiers if ie else yigan_tiers
        r['ton'] = None; r['coeff'] = None
        if r['final_vol'] and r['final_vol'] > 0 and tiers:
            for limit, ton, coef, _c in tiers:  # 已按下限降序
                if r['final_vol'] >= limit: r['ton'] = ton; r['coeff'] = coef; break
        r['is_ergan'] = ie

    # ====== 降吨位：结算吨位(C35)覆盖车型吨位 ======
    def _ton_num(tier_name):
        """从吨位名提取数值: '40吨B'→40, '30吨'→30, '2.75吨'→2.75"""
        m = re.search(r'(\d+\.?\d*)', str(tier_name))
        return float(m.group(1)) if m else 0

    downgrade_count = 0
    for r in records:
        st = r.get('settle_ton')
        if st is None: continue
        try: st_val = float(st)
        except (ValueError, TypeError): continue
        if st_val <= 0: continue
        r['settle_ton'] = st_val

        tiers = ergan_tiers if r['is_ergan'] else yigan_tiers
        if not tiers: continue
        # 只有车型吨位≥30吨才考虑降吨位（12吨车的结算9吨是载重记录，非降吨位）
        if not r['ton']: continue
        vol_ton_num = _ton_num(r['ton'])
        if vol_ton_num < 30: continue

        # 找结算吨位对应的价格档: 取吨位数值≤st_val的最大档(档位已按容积下限降序)
        settle_tier = settle_coef = None
        settle_ton_num = 0
        for limit, ton, coef, _c in tiers:
            tn = _ton_num(ton)
            if tn <= st_val:
                settle_tier = ton; settle_coef = coef; settle_ton_num = tn; break

        # 降吨位: 结算档 < 车型档 且 车型≥30吨
        if settle_tier and settle_ton_num < vol_ton_num:
            r['orig_ton'] = r['ton']    # 车型吨位（仅标记，不用于价格）
            r['orig_coeff'] = r['coeff']
            r['ton'] = settle_tier      # 结算吨位 → 用于价格匹配
            r['coeff'] = settle_coef
            r['downgrade'] = True
            downgrade_count += 1

    if downgrade_count:
        logd(f"降吨位: 结算吨位(C35)与车型吨位不一致, 共 {downgrade_count} 条按结算吨位定档")

    # ====== Step 3: 跨档 + 变动前 ======
    pt = defaultdict(list)
    for r in records:
        if r['ton']: pt[r['plate']].append(r['ton'])
    pn = {p: Counter(t).most_common(1)[0][0] for p, t in pt.items()}
    records.sort(key=lambda x: (x['plate'], x['date'] or datetime(2000, 1, 1)))
    pv_rec = {}
    for r in records:
        r['cross'] = bool(r['ton'] and r['ton'] != pn.get(r['plate']) and not r.get('downgrade'))
        if r['plate'] in pv_rec:
            r['prev_vol'] = pv_rec[r['plate']]['final_vol']; r['prev_date'] = pv_rec[r['plate']]['date']
        pv_rec[r['plate']] = r

    miss = sum(1 for r in records if r['missing']); unm = sum(1 for r in records if r['unmeasured'])
    other = sum(1 for r in records if r.get('measured_other_date'))
    cross = sum(1 for r in records if r['cross']); dg = sum(1 for r in records if r.get('downgrade')); total = len(records)
    print(f"  红(缺失)={miss} 橙(未测量)={unm} 蓝(派车日未测·别处测过)={other} 黄(跨档)={cross} 降吨位={dg} 总={total}")

    # ====== Step 4: 价格匹配 ======
    print("Step 4: 价格匹配...")
    km_ok = price_ok = 0
    for r in records:
        rt, ton = r['route'], r['ton']; r['km'] = None; r['price'] = None; r['freight'] = None
        data = ergan_data if r['is_ergan'] else yigan_data
        for lv in range(7):
            n = std(rt, lv)
            if n and n in data:
                r['km'] = data[n]['km']
                if ton and ton in data[n]['ps']:
                    try: r['price'] = float(data[n]['ps'][ton])
                    except: r['price'] = data[n]['ps'][ton]
                break
        if r['price'] and r['coeff']: r['freight'] = round(r['coeff'] * r['price'], 2)
        if r['km']: km_ok += 1
        if r['price']: price_ok += 1
    print(f"  公里={km_ok}({100*km_ok/total:.1f}%) 标准价={price_ok}({100*price_ok/total:.1f}%)")

    # ====== 降吨位：计算原运费 ======
    # 官方公式: 原运费 = 当前路线标准价 × 原车型系数
    # 差价 = 当前运费(结算吨位) - 原运费(车型吨位)
    for r in records:
        if not r.get('downgrade'): continue
        if r.get('price') and r.get('orig_coeff'):
            try:
                r['orig_freight'] = round(r['orig_coeff'] * float(r['price']), 2)
                if r.get('freight') is not None:
                    r['price_diff'] = round(r['freight'] - r['orig_freight'], 2)
            except: pass

    # Step 5: 车管数据已在 Step 2b 读入 records，无需单独匹配
    # 承运单位/派车单号/封车时间/运输量 已随行来源填入

    # ====== Step 6: 生成初版 ======
    print("Step 6: 生成初版...")
    er_min = min((t[0] for t in ergan_tiers), default=None) if ergan_tiers else None
    yg_min = min((t[0] for t in yigan_tiers), default=None) if yigan_tiers else None
    for r in records:
        issues = []
        vol_str = f"{r['final_vol']}方" if r['final_vol'] else "0方(未测量)"
        if r.get('vol_cross_station'):
            issues.append("串行：容积取自同车跨站点/跨日期测量（草滩记录→港务测量或反之），请人工核对")
        if r.get('measured_other_date'):
            dates = r.get('other_dates', [])
            dates_str = '/'.join(dates[:3]) + ('等' if len(dates) > 3 else '')
            issues.append(f"派车当日在线表未测，同车在{dates_str}测过。容积未跨日借用（留空），请人工核对取哪天的容积")
        if r.get('unmeasured'):
            issues.append("在线表未测量该车（当月无任何记录），容间缺失，待人工补")
        if not r['ton']:
            need = er_min if r['is_ergan'] else yg_min
            issues.append(f"吨位无法判定：容间={vol_str}，{'二干' if r['is_ergan'] else '一干'}路线最低需≥{need}方。可能为串行未测量或容积异常")
        if r['km'] is None:
            dest = re.sub(r'\([^)]*\)', '', r['route']).split('-')[-1].strip()
            issues.append(f"价格表未覆盖：路线「{r['route']}」在{'二干' if r['is_ergan'] else '一干'}价格表中未找到（目的地={dest}）。需人工确认标准价和公里")
        if r['price'] is None and r['km'] is not None:
            issues.append(f"吨位无价格：路线在价格表中已找到（公里={r['km']}），但吨位「{r['ton']}」无对应价格列。运价无法计算")
        if r.get('downgrade'):
            orig_tn = r.get('orig_ton', '?')
            orig_f = r.get('orig_freight')
            diff = r.get('price_diff')
            if diff is not None:
                issues.append(f"降吨位：车型={orig_tn}→结算={r['ton']}，原运费={orig_f} 差价={diff:+}（结算吨位来自车管C35）")
            else:
                issues.append(f"降吨位：车型={orig_tn}→结算={r['ton']}（结算吨位来自车管C35，价格待补）")
        r['issues'] = issues

    DRAFT = f"{OUT}/{M}月_02_基础报表初版.xlsx"
    wb_o = openpyxl.Workbook()

    # 结构发现 Sheet（放最前，供人工核验本月各表结构是否被正确识别）
    ws_s = wb_o.active; ws_s.title = "结构发现"
    ws_s.cell(row=1, column=1, value="[*] 本次运行动态识别的各源文件结构（请核对是否正确，若识别错请检查源文件表头）").font = Font(bold=True, size=11, color="0070C0")
    for i, line in enumerate(DISCOVERY_LOG, start=3):
        ws_s.cell(row=i, column=1, value=line).font = Font(size=10)
    ws_s.column_dimensions['A'].width = 160

    ws_i = wb_o.create_sheet(title="问题汇总")
    ws_i.merge_cells('A1:K1')
    ws_i.cell(row=1, column=1, value="[!] 需人工核对修正 —— 按「问题类型」列筛选可批量处理；核对后在「处理结果」列标注，全部确认无误后方可拆分/盈亏").font = Font(bold=True, size=11, color="FF0000")
    HDR_I = ['问题类型', 'Sheet', '行号', '车牌号', '路线', '路线类型', '当前吨位', '容间', '问题详情（请逐条核对）', '建议动作', '处理结果（人工填写）']
    for c, h in enumerate(HDR_I, 1):
        cl = ws_i.cell(row=3, column=c, value=h); cl.fill = FILL_H; cl.font = FONT_H; cl.border = B

    # 问题分类（按优先级）+ 建议动作，方便人工按类型批量审核
    def categorize(r):
        det = ' '.join(r['issues'])
        if '派车当日在线表未测' in det: return (1, '派车日未测', '查同车在线表其他日期(见详情)的容积，确认取哪天')
        if '串行' in det: return (2, '串行(跨站点取容积)', '已从同车牌跨站点取得容积，请核对是否正确，特别关注波动大(跨档)的车')
        if '未测量该车' in det: return (3, '完全未测量', '该车当月无任何测量，需人工补容间或标缺')
        if '价格表未覆盖' in det: return (4, '价格表未覆盖', '维护价格表路线别名，或人工补标准价/公里')
        if '吨位无价格' in det: return (5, '吨位无价格', '路线已匹配到价格表，补该吨位对应的标准价')
        if '跨档' in det: return (6, '跨档', '核对本次吨位（偏离同车牌众数），确认是否真跨档')
        if '吨位无法判定' in det: return (7, '吨位无法判定', '容间异常，人工确认吨位')
        if '降吨位' in det: return (0, '降吨位', '结算吨位低于车型吨位，已自动按结算吨位定档。请核对车管C35结算吨位是否准确')
        return (9, '其他', '人工核对')
    TYPE_COLOR = {'派车日未测': FILL_B, '串行(跨站点取容积)': FILL_B, '完全未测量': FILL_O, '跨档': FILL_Y, '降吨位': FILL_DG}

    rows = []
    for i, r in enumerate(records):
        if r['issues']:
            prio, typ, action = categorize(r)
            rows.append((prio, typ, i + 1, r, action))
    rows.sort(key=lambda x: x[0])  # 按类型分组排序，同类问题集中，便于批量处理

    ir = 4
    for prio, typ, seq, r, action in rows:
        ws_i.cell(row=ir, column=1, value=typ).border = B
        ws_i.cell(row=ir, column=2, value="基础报表").border = B
        ws_i.cell(row=ir, column=3, value=seq).border = B
        ws_i.cell(row=ir, column=4, value=r['plate']).border = B
        ws_i.cell(row=ir, column=5, value=r['route']).border = B
        ws_i.cell(row=ir, column=6, value='二干' if r['is_ergan'] else '一干').border = B
        ws_i.cell(row=ir, column=7, value=r['ton'] if r['ton'] else '未判定').border = B
        ws_i.cell(row=ir, column=8, value=r['final_vol']).border = B
        ws_i.cell(row=ir, column=9, value='\n'.join(f"• {iss}" for iss in r['issues'])).border = B
        ws_i.cell(row=ir, column=10, value=action).border = B
        ws_i.cell(row=ir, column=11, value="").border = B
        fill = TYPE_COLOR.get(typ, FILL_WARN)
        for c in range(1, 12): ws_i.cell(row=ir, column=c).fill = fill
        ir += 1
    ic = len(rows)
    for c, w in enumerate([14, 12, 6, 12, 36, 8, 8, 10, 52, 38, 15], 1): ws_i.column_dimensions[get_column_letter(c)].width = w
    ws_i.row_dimensions[1].height = 22
    for rw in range(4, ir): ws_i.row_dimensions[rw].height = 55

    ws_d = wb_o.create_sheet(title="基础报表")
    for c, h in enumerate(V16_HDR, 1):
        cl = ws_d.cell(row=1, column=c, value=h); cl.fill = FILL_H; cl.font = FONT_H; cl.border = B
    for i, r in enumerate(records):
        row = i + 2
        rp = []
        if r['issue'] and str(r['issue']).strip() not in ('', 'None'):
            raw = str(r['issue']).strip()
            # 纯数字 → 标注含义不明确
            if re.match(r'^\d+\.?\d*$', raw):
                rp.append(f'{raw}（车型问题原文为纯数字，含义不明确，请确认是否为车型代号或挂车号错填）')
            else:
                rp.append(raw)
        if r['issues']: rp.append('【[!]需核对】' + '；'.join(r['issues']))
        remark = ' | '.join(rp) if rp else None
        vals = {1: i+1, 2: r.get('company'), 3: r.get('dispatch'), 4: r['route'],
            6: r.get('seal_date_str'), 8: r['plate'],
            10: r['final_vol'] if r['final_vol'] and r['final_vol'] > 0 else None,
            11: r['ton'], 12: r['km'], 13: r.get('transport_qty'),
            15: r['freight'], 16: r['price'], 17: remark,
            18: '是' if r['cross'] else '否',
            19: r.get('prev_vol'), 20: r['prev_date'].strftime('%Y-%m-%d') if r.get('prev_date') else None, 21: r['coeff']}
        color = FILL_R if r['missing'] else (FILL_O if r['unmeasured'] else (FILL_DG if r.get('downgrade') else (FILL_Y if r['cross'] else (FILL_B if r.get('vol_cross_station') else None))))
        for c in range(1, 22):
            cl = ws_d.cell(row=row, column=c, value=vals.get(c)); cl.border = B
            if color: cl.fill = color
    for c, w in enumerate([6, 24, 16, 28, 12, 14, 2, 12, 2, 14, 8, 8, 10, 14, 10, 10, 20, 8, 10, 14, 8], 1):
        ws_d.column_dimensions[get_column_letter(c)].width = w
    DRAFT = save_workbook(wb_o, DRAFT)

    print(f"\n[OK] 初版已生成: {DRAFT}")
    print(f"[汇总] 总={total} | 红(缺失)={miss} 橙(未测量)={unm} 黄(跨档)={cross} 降吨位={dg}")
    print(f"[匹配] 公里={100*km_ok/total:.1f}% 标准价={100*price_ok/total:.1f}%")
    print(f"[问题汇总] {ic} 条待人工核对")
    print("\n[STOP] 铁律：初版生成后停止。拆分/盈亏须在人工核对修正初版之后再执行。")


if __name__ == "__main__":
    main()
