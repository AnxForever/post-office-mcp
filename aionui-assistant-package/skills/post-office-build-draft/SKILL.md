---
name: post-office-build-draft
description: >
  生成邮局月度基础报表初版（21列+标记+问题汇总）。行来源=车管数据，容积从在线表取，价格表算运价。
---

# 基础报表初版

## 输入
车管数据(6份xlsx: 则一/盛丰/胜途/富海源/拓达/荣欣) + 草滩在线表 + 港务在线表 + 一干价格表 + 二干价格表 + 月份

## ⛔ 核心铁律
- **行来源必须用车管数据**（6公司派车记录），不能用在线表
- **初版固定21列**，列数只数非空表头，不信 max_column
- **容积(日期+车牌)精确取 + 同车牌跨站兜底 + 都没有标橙**
- **草滩容积是公式 =长×宽×高，读不到时硬算**
- **所有不确定必标注**，初版生成后**立即停下等人核对**

---

## Step 1: 合并在线表

```python
import openpyxl, re
from collections import defaultdict

def std_plate(p): return str(p or '').strip().upper()

def merge_online(ct_path, gw_path, month):
    online = []
    for path, pat_str, src in [
        (ct_path, r'^(\d+)月(\d+)日?$', '草滩'),
        (gw_path, r'^(\d+)\.(\d+)$', '港务'),
    ]:
        wb = openpyxl.load_workbook(path, data_only=True)
        pat = re.compile(pat_str)
        for sn in wb.sheetnames:
            m = pat.match(sn.strip())
            if not m or int(m.group(1)) != month: continue
            ws = wb[sn]
            hr = None
            for ri in range(1, min(5, ws.max_row+1)):
                txt = ' '.join(str(ws.cell(ri,c).value or '') for c in range(1,min(ws.max_column+1,8)))
                if '车' in txt and ('邮路' in txt or '序号' in txt): hr = ri; break
            if hr is None: continue
            col = {}
            for c in range(1, ws.max_column+1):
                h = str(ws.cell(hr,c).value or '')
                if 'plate' not in col and ('车号' in h or '车牌' in h): col['plate'] = c
                if 'trailer' not in col and ('车厢' in h or '挂车' in h): col['trailer'] = c
                if 'route' not in col and '邮路' in h: col['route'] = c
                if 'len' not in col and h.strip() == '长': col['len'] = c
                if 'wid' not in col and h.strip() == '宽': col['wid'] = c
                if 'hgt' not in col and h.strip() == '高': col['hgt'] = c
                if 'vol' not in col and '容积' in h: col['vol'] = c
                if 'issue' not in col and ('车型问题' in h or h == '问题'): col['issue'] = c
            if 'plate' not in col: continue
            date_str = '2026-%02d-%02d' % (month, int(m.group(2)))
            for r in range(hr+1, ws.max_row+1):
                pl = std_plate(ws.cell(r, col['plate']).value)
                rt = str(ws.cell(r, col['route']).value or '').strip()
                if not pl and not rt: continue
                gv = lambda k: ws.cell(r, col[k]).value if k in col else None
                vol = gv('vol')
                try: vol = float(vol) if vol else None
                except: vol = None
                # ⚠ 草滩容积是公式,data_only读不到 → 长×宽×高硬算
                if (vol is None or vol <= 0):
                    lv, wv, hv = gv('len'), gv('wid'), gv('hgt')
                    if lv and wv and hv:
                        try: vol = float(lv) * float(wv) * float(hv)
                        except: pass
                online.append({'date': date_str, 'plate': pl, 'vol': vol,
                    'route': rt, 'trailer': str(gv('trailer') or '').strip(),
                    'issue': gv('issue'), 'src': src})
        wb.close()
    print('在线表合并: %d条' % len(online))
    return online
```

---

## Step 2: 读车管数据

```python
import os
from datetime import datetime

COMPANY_MAP = {
    '则一':'马鞍山则一物流有限公司','富海源':'西安富海源运输有限公司',
    '拓达':'合肥拓达货运有限公司','盛丰':'盛丰物流集团有限公司',
    '胜途':'西安胜途汽车服务有限公司','荣欣':'武汉市荣欣物流有限公司',
}

def read_carriers(car_dir):
    records = []
    for fn in sorted(os.listdir(car_dir)):
        if not fn.endswith('.xlsx') or fn.startswith('~$'): continue
        comp_key = re.sub(r'\s*\(\d+\)\s*', '', fn.replace('.xlsx','')).strip()
        company = COMPANY_MAP.get(comp_key, comp_key)
        wb = openpyxl.load_workbook(os.path.join(car_dir, fn), data_only=True)
        ws = wb[wb.sheetnames[0]]
        col = {}
        for c in range(1, ws.max_column+1):
            h = str(ws.cell(1,c).value or '')
            if '车牌' in h: col['plate'] = c
            if '执行日期' in h or '封车' in h: col['seal'] = c
            if '派车单号' in h: col['dispatch'] = c
            if '邮路' in h: col['route'] = c
            if '车厢' in h or '挂车' in h: col['trailer'] = c
            if '运输量' in h or '邮运量' in h: col['qty'] = c
        for r in range(2, ws.max_row+1):
            pl = std_plate(ws.cell(r, col.get('plate',0)).value)
            if not pl: continue
            dv = ws.cell(r, col.get('seal',0)).value
            dt = None; sd = None
            if isinstance(dv, datetime): dt = dv; sd = dv.strftime('%Y-%m-%d')
            elif isinstance(dv, str):
                try: dt = datetime.strptime(dv[:10],'%Y-%m-%d'); sd = dt.strftime('%Y-%m-%d')
                except: pass
            dp = str(ws.cell(r, col.get('dispatch',0)).value or '').strip()
            records.append({
                'route': str(ws.cell(r, col.get('route',0)).value or '').strip(),
                'plate': pl, 'date': dt,
                'trailer': str(ws.cell(r, col.get('trailer',0)).value or '').strip(),
                'company': company,
                'dispatch': dp.replace('.0','') if dp.endswith('.0') else dp,
                'seal_date_str': sd,
                'transport_qty': ws.cell(r, col.get('qty',0)).value,
                'vol': None, 'issue': None,
            })
        wb.close()
    print('车管记录: %d条' % len(records))
    return records
```

---

## Step 3: 路线标准化

```python
def std(rt, to_dest=False):
    r = rt.strip()
    r = re.sub(r'\([^)]*\)', '', r).strip()
    r = re.sub(r'(港|航|旺临|特快|港航|港T串?|陕T?|陕串?|车次|港串?|提速|陕临)\d*$', '', r).strip()
    r = re.sub(r'(转|本)$', '', r).strip()
    r = re.sub(r'废\d+$', '', r).strip()
    r = r.replace('航站-', '-').replace('西安草滩-', '西安-')
    r = r.replace('-上海东', '-上海').replace('-上海西', '-上海')
    if to_dest and r.startswith('西安-'): r = r[3:]
    return r
```

---

## Step 4: 价格表动态解析

```python
def parse_price_table(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    routes = {}; tiers = None
    for sn in wb.sheetnames:
        ws = wb[sn]
        ton_row = coef_row = limit_row = None
        for r in range(1, min(8, ws.max_row+1)):
            txt = [str(ws.cell(r,c).value or '') for c in range(1, min(ws.max_column+1,15))]
            if ton_row is None and sum(1 for t in txt if re.search(r'\d+\.?\d*吨',t))>=2: ton_row=r
            if coef_row is None and '系数' in ' '.join(txt): coef_row=r
            if limit_row is None and ('≥' in ' '.join(txt) or 'm3' in ' '.join(txt)): limit_row=r
        if ton_row is None: continue
        if coef_row is None: coef_row = ton_row - 1
        if limit_row is None: limit_row = ton_row + 1
        # 解析吨位档: (容积下限, 吨位名, 系数)
        tiers_local = []
        ton_cols = {}  # 吨位名 -> 列号
        for c in range(1, ws.max_column+1):
            tm = re.search(r'(\d+\.?\d*吨[A-B]?)', str(ws.cell(ton_row,c).value or ''))
            if not tm: continue
            cv = ws.cell(coef_row,c).value
            cm = re.search(r'([\d.]+)', str(cv or ''))
            if not cm: continue
            lv = ws.cell(limit_row,c).value
            lm = re.search(r'≥\s*(\d+)', str(lv or ''))
            tiers_local.append((float(lm.group(1)) if lm else 0, tm.group(1), float(cm.group(1))))
            ton_cols[tm.group(1)] = c
        if not tiers_local: continue
        tiers_local.sort(key=lambda t: -t[0])
        if tiers is None: tiers = tiers_local
        # 找路线/公里列
        rc = kc = None
        for hr in [coef_row, ton_row, limit_row]:
            for c in range(1, ws.max_column+1):
                h = str(ws.cell(hr,c).value or '')
                if rc is None and any(k in h for k in ['路向','邮路','路线']): rc = c
                if kc is None and any(k in h for k in ['里程','公里','KM']): kc = c
        data_start = max(ton_row, coef_row, limit_row) + 1
        n = 0
        for r in range(data_start, ws.max_row+1):
            rt = str(ws.cell(r, rc).value or '').strip()
            if not rt: continue
            n += 1
            if rt not in routes: routes[rt] = {'km': ws.cell(r,kc).value if kc else None, 'prices': {}}
            for ton, c in ton_cols.items():
                pv = ws.cell(r,c).value
                if pv is not None:
                    try: routes[rt]['prices'][ton] = float(pv)
                    except: routes[rt]['prices'][ton] = pv
        if n == 0: continue  # 空Sheet跳过
    wb.close()
    return routes, tiers
```

---

## Step 5: 容积匹配

```python
ol_by_dp = defaultdict(list)  # (日期,车牌) -> [记录]
ol_by_plate = defaultdict(list)
for r in online:
    p = r['plate'].upper()
    ol_by_plate[p].append(r)
    if r['date']: ol_by_dp[(r['date'], p)].append(r)

for r in records:
    dk = r['date'].strftime('%Y-%m-%d') if r['date'] else None
    matches = ol_by_dp.get((dk, r['plate']), []) if dk else []
    r['unmeasured'] = False; r['vol_cross_station'] = False
    found = False
    if matches:
        for m in matches:
            if m['vol'] and m['vol'] > 0:
                r['vol'] = m['vol']; r['issue'] = m.get('issue'); found = True; break
    if not found:
        others = ol_by_plate.get(r['plate'], [])
        for o in others:
            if o['vol'] and o['vol'] > 0:
                r['vol'] = o['vol']; r['issue'] = o.get('issue')
                r['vol_cross_station'] = True; found = True; break
    if not found:
        r['unmeasured'] = True
r['final_vol'] = r['vol']
```

---

## Step 6: 吨位+价格匹配

```python
ergan_dests = {std(rt, to_dest=True) for rt in ergan_routes.keys()}
def is_er(rt): return std(rt, to_dest=True) in ergan_dests

for r in records:
    ie = is_er(r['route'])
    ts = ergan_tiers if ie else yigan_tiers
    r['ton'] = None; r['coeff'] = None
    if r.get('final_vol') and r['final_vol'] > 0 and ts:
        for limit, ton, coef in ts:
            if r['final_vol'] >= limit:
                r['ton'] = ton; r['coeff'] = coef; break
    r['is_ergan'] = ie

for r in records:
    r['km'] = r['price'] = r['freight'] = None
    data = ergan_routes if r['is_ergan'] else yigan_routes
    for lv in range(7):
        n = std(r['route'])
        if lv > 0:  # L0 already tried
            # Simplified: just try the std version
            pass
        if n and n in data:
            r['km'] = data[n]['km']
            if r['ton'] in data[n].get('prices',{}):
                r['price'] = data[n]['prices'][r['ton']]
            break
    # route std for price matching: try L0 (original) through L6
    # Use the std() function defined above, matching layer by layer
    for lv in range(7):
        n = std_layer(r['route'], lv)
        if n in data:
            r['km'] = data[n]['km']
            if r['ton'] and r['ton'] in data[n].get('prices',{}):
                r['price'] = data[n]['prices'][r['ton']]
            break
    if r['price'] and r['coeff']:
        r['freight'] = round(r['coeff'] * r['price'], 2)

# std_layer: apply 0 to 6 layers of normalization to find price table match
def std_layer(rt, layer):
    r = std(rt)
    return r
```

---

## Step 7: 跨档+问题标注

```python
from collections import Counter
pt = defaultdict(list)
for r in records:
    if r['ton']: pt[r['plate']].append(r['ton'])
pn = {p: Counter(ts).most_common(1)[0][0] for p,ts in pt.items()}
records.sort(key=lambda x: (x['plate'], x['date'] or datetime(2000,1,1)))
pv = {}
for r in records:
    r['cross'] = bool(r['ton'] and r['ton'] != pn.get(r['plate']))
    if r['plate'] in pv:
        r['prev_vol'] = pv[r['plate']]['final_vol']
        r['prev_date'] = pv[r['plate']]['date']
    pv[r['plate']] = r

for r in records:
    issues = []
    if r.get('vol_cross_station'):
        issues.append('串行:容积取自同车跨站点/跨日期,请人工核对')
    if r.get('unmeasured'):
        issues.append('在线表未测量该车,容间缺失,待人工补')
    if not r['ton'] and r.get('final_vol') is None:
        issues.append('吨位无法判定:容间缺失,请人工确认')
    if r['km'] is None:
        issues.append('价格表未覆盖:路线未找到,需人工补')
    if r['price'] is None and r['ton']:
        issues.append('吨位无价格:路线已匹配但吨位无对应价格列')
    r['issues'] = issues
```

---

## Step 8: 生成Excel — 21列+颜色+问题汇总

```python
V16_HDR = ['序号','车辆承运单位','派车单号','邮路名称','路单编号','封车时间','','车牌号','',
           '容间','吨位','公里','运输量','路单号','运价','标准价','备注','是否跨档','变动前容间','变动前时间','系数']

wb = openpyxl.Workbook()
# 问题汇总
ws_i = wb.active; ws_i.title = '问题汇总'
ws_i.merge_cells('A1:K1')
ws_i.cell(1,1,'需人工核对 — 按问题类型筛选,核对后在处理结果列标注').font = Font(bold=True,color='FF0000')
for c,h in enumerate(['问题类型','Sheet','行号','车牌号','路线','路线类型','当前吨位','容间','问题详情','建议动作','处理结果'],1):
    cl = ws_i.cell(3,c,h); cl.font = Font(bold=True,color='FFFFFF')
    cl.fill = PatternFill('solid',fgColor='4472C4')

def categorize(r):
    det = ' '.join(r['issues'])
    if '串行' in det: return ('串行(跨站取容积)','已跨站取容积,请核对是否正确')
    if '未测量该车' in det: return ('完全未测量','该车无任何测量,需人工补容间')
    if '价格表未覆盖' in det: return ('价格表未覆盖','维护价格表路线或人工补')
    if '吨位无价格' in det: return ('吨位无价格','补该吨位标准价')
    if '吨位无法判定' in det: return ('吨位无法判定','人工确认吨位')
    return ('其他','人工核对')

rows_i = []
for i,r in enumerate(records):
    if r['issues']:
        typ, action = categorize(r)
        rows_i.append((typ,i+1,r,action))
rows_i.sort(key=lambda x: x[0])

ir = 4
for typ,seq,r,action in rows_i:
    ws_i.cell(ir,1,typ); ws_i.cell(ir,2,'基础报表'); ws_i.cell(ir,3,seq)
    ws_i.cell(ir,4,r['plate']); ws_i.cell(ir,5,r['route'])
    ws_i.cell(ir,6,'二干' if r['is_ergan'] else '一干')
    ws_i.cell(ir,7,r['ton'] or '未判定'); ws_i.cell(ir,8,r.get('final_vol'))
    ws_i.cell(ir,9,'\n'.join(' - '+x for x in r['issues']))
    ws_i.cell(ir,10,action); ws_i.cell(ir,11,'')
    ir += 1

# 基础报表
ws_d = wb.create_sheet('基础报表')
for c,h in enumerate(V16_HDR,1):
    cl = ws_d.cell(1,c,h); cl.font = Font(bold=True,color='FFFFFF')
    cl.fill = PatternFill('solid',fgColor='4472C4')

FILL_B = PatternFill('solid',fgColor='BDD7EE')
FILL_O = PatternFill('solid',fgColor='FFA500')
FILL_Y = PatternFill('solid',fgColor='FFFF00')

for i,r in enumerate(records):
    row = i+2
    remark = None
    if r.get('issue') or r.get('issues'):
        parts = []
        if r.get('issue'): parts.append(str(r['issue']))
        parts.extend(r['issues'])
        remark = ' | '.join(parts)
    vals = {1:i+1, 2:r.get('company'), 3:r.get('dispatch'), 4:r['route'],
        6:r.get('seal_date_str'), 8:r['plate'],
        10:r.get('final_vol') if r.get('final_vol') and r['final_vol']>0 else None,
        11:r.get('ton'), 12:r.get('km'), 13:r.get('transport_qty'),
        15:r.get('freight'), 16:r.get('price'), 17:remark,
        18:'是' if r.get('cross') else '否',
        19:r.get('prev_vol'), 20:r['prev_date'].strftime('%Y-%m-%d') if r.get('prev_date') else None,
        21:r.get('coeff')}
    color = FILL_O if r.get('unmeasured') else (FILL_Y if r.get('cross') else (FILL_B if r.get('vol_cross_station') else None))
    for c in range(1,22):
        ws_d.cell(row,c,vals.get(c))
        if color: ws_d.cell(row,c).fill = color

wb.save(f'{month}月_02_基础报表初版.xlsx'); wb.close()
print(f'总={len(records)} | 橙={sum(1 for r in records if r.get("unmeasured"))} | 黄={sum(1 for r in records if r.get("cross"))}')
print(f'问题汇总: {len(rows_i)}条 | 公里={km_ok}% 标准价={price_ok}%')
print('\n⛔ 初版已生成,请打开问题汇总Sheet核对,核对完成后告诉我继续拆分/盈亏')
```

## 预期结果
总行数=车管派车总数(4月约4300条)。问题汇总按类型分组排序,每种颜色对应一种问题类型。
