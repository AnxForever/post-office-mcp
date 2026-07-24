#!/usr/bin/env python3
"""
Step 2: 在线合并表 vs 基准表对比 + 容积跨档检查 → 生成基础报表初版
用法: python3 compare_check.py <月份> <合并表路径> <基准表路径> <二干价格表路径> <输出路径>
"""
import sys, os, re
from datetime import datetime, timedelta
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

# v1.16 输出表头
OUTPUT_HEADERS = [
    '序号', '车辆承运单位', '派车单号', '邮路名称', '路单编号',
    '封车时间', '', '车牌号', '', '容间', '吨位', '公里', '运输量',
    '路单号', '运价', '标准价', '备注', '是否跨档', '变动前容间', '变动前时间', '系数'
]

# 吨位档位规则
YIGAN_BRACKETS = [
    (55, '12吨', 0.63), (100, '20吨', 0.83), (130, '30吨', 1.0),
    (150, '40吨A', 1.12), (180, '40吨B', 1.23)
]
ERGAN_BRACKETS = [
    (16, '3吨', 0.58), (35, '5吨', 0.71), (45, '8吨', 0.81),
    (55, '12吨', 1.0), (100, '20吨', 1.32), (130, '30吨', 1.59),
    (150, '40吨A', 1.78)
]


def load_merged_table(path):
    """加载在线合并表 → {(plate, trailer): [records]}"""
    wb = openpyxl.load_workbook(path, data_only=True)
    index = defaultdict(list)
    for sname in wb.sheetnames:
        m = re.match(r'(\d+)月(\d+)日', sname)
        if not m:
            continue
        date = datetime(2026, int(m.group(1)), int(m.group(2)))
        ws = wb[sname]
        for row in range(2, ws.max_row + 1):
            plate = str(ws.cell(row=row, column=3).value or "").strip().upper()
            route = str(ws.cell(row=row, column=2).value or "").strip()
            trailer = str(ws.cell(row=row, column=4).value or "").strip().upper()
            if not plate:
                continue
            index[(plate, trailer)].append({
                'date': date, 'route': route,
                'volume': ws.cell(row=row, column=8).value,
                'length': ws.cell(row=row, column=5).value,
                'width': ws.cell(row=row, column=6).value,
                'height': ws.cell(row=row, column=7).value,
                'source': str(ws.cell(row=row, column=13).value or ""),
            })
    wb.close()
    return index


def load_baseline(path):
    """加载基准表（N月车型）→ [records]"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Sheet1']
    records = []
    for row in range(4, ws.max_row + 1):
        plate = str(ws.cell(row=row, column=3).value or "").strip().upper()
        route = str(ws.cell(row=row, column=2).value or "").strip()
        trailer = str(ws.cell(row=row, column=4).value or "").strip().upper()
        if not plate and not route:
            continue
        date_serial = ws.cell(row=row, column=12).value
        if isinstance(date_serial, (int, float)):
            dt = datetime(1899, 12, 30) + timedelta(days=int(date_serial))
        else:
            dt = None
        records.append({
            'route': route, 'plate': plate, 'trailer': trailer,
            'volume': ws.cell(row=row, column=8).value,
            'vehicle_issue': ws.cell(row=row, column=11).value,
            'date': dt,
        })
    wb.close()
    return records


def load_ergan_routes(path):
    """加载二干路线列表"""
    wb = openpyxl.load_workbook(path, data_only=True)
    routes = set()
    for sname in ['单程', '往返']:
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        for row in range(5, ws.max_row + 1):
            r = str(ws.cell(row=row, column=3).value or "").strip()
            if r:
                routes.add(r)
    wb.close()
    return routes


def get_bracket(volume, brackets):
    """根据容间返回(吨位, 系数)"""
    if volume is None:
        return (None, None)
    result = (None, None)
    for th, ton, coeff in brackets:
        if volume >= th:
            result = (ton, coeff)
    return result


def is_vehicle_measured(all_online_records):
    """判断一辆车是否被实际测量过"""
    for r in all_online_records:
        v = r['volume']
        if v is not None and float(v) > 0:
            return True
        if r['length'] is not None or r['width'] is not None or r['height'] is not None:
            return True
    return False


def get_best_volume(baseline_volume, all_online_records):
    """
    获取最佳容间值:
    - 优先用基准表中的值（如果>0）
    - 如果基准表为0，从在线表中找非零值
    - 都没有则返回0
    """
    if baseline_volume is not None and float(baseline_volume) > 0:
        return baseline_volume
    for r in all_online_records:
        v = r['volume']
        if v is not None and float(v) > 0:
            return v
    return baseline_volume  # 保持原值（可能是0）


def main():
    if len(sys.argv) < 6:
        print("用法: python3 compare_check.py <月份> <合并表> <基准表> <二干价格表> <输出>")
        sys.exit(1)

    month = int(sys.argv[1])
    merged_path = sys.argv[2]
    baseline_path = sys.argv[3]
    ergan_price_path = sys.argv[4]
    output_path = sys.argv[5]

    print("📦 加载数据...")
    online_index = load_merged_table(merged_path)
    baseline_records = load_baseline(baseline_path)
    ergan_routes = load_ergan_routes(ergan_price_path)

    def is_ergan(route):
        for er in ergan_routes:
            if er in route or route in er:
                return True
        return False

    # 逐条处理
    print(f"🔍 对比 {len(baseline_records)} 条记录...")
    results = []

    for r in baseline_records:
        key = (r['plate'], r['trailer'])
        online_recs = online_index.get(key, [])
        # 如果按(plate+trailer)找不到，只按plate找
        if not online_recs:
            for k, v in online_index.items():
                if k[0] == r['plate']:
                    online_recs.extend(v)

        # 规则①: 在线表缺失
        is_missing = len(online_recs) == 0

        # 规则②: 未测量
        volume = get_best_volume(r['volume'], online_recs)
        is_unmeasured = False
        if (volume is None or float(volume) == 0) and not is_missing:
            is_unmeasured = not is_vehicle_measured(online_recs)

        # 路线类型 + 吨位档
        route_type = '二干' if is_ergan(r['route']) else '一干'
        brackets = ERGAN_BRACKETS if route_type == '二干' else YIGAN_BRACKETS
        ton, coeff = get_bracket(volume, brackets)

        results.append({
            'route': r['route'], 'plate': r['plate'], 'trailer': r['trailer'],
            'volume': volume if not is_unmeasured else 0,
            'ton': ton, 'coeff': coeff,
            'vehicle_issue': r['vehicle_issue'],
            'date': r['date'],
            'is_missing': is_missing,
            'is_unmeasured': is_unmeasured,
            'route_type': route_type,
        })

    # 规则③: 跨档检查（按车牌排序）
    results.sort(key=lambda x: (x['plate'], x['date'] or datetime(2000, 1, 1)))
    prev = {}
    for r in results:
        plate = r['plate']
        if plate in prev and r['ton'] and prev[plate]['ton']:
            if r['ton'] != prev[plate]['ton']:
                r['is_cross'] = True
                r['prev_volume'] = prev[plate]['volume']
                r['prev_date'] = prev[plate]['date']
        prev[plate] = r

    # 统计
    missing = sum(1 for r in results if r.get('is_missing'))
    unmeasured = sum(1 for r in results if r.get('is_unmeasured'))
    cross = sum(1 for r in results if r.get('is_cross'))
    print(f"  🔴 缺失: {missing}  🟠 未测量: {unmeasured}  🟡 跨档: {cross}")

    # 写入
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "基础报表"

    FILL_H = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    FILL_RED = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")
    FILL_ORANGE = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    FILL_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    FONT_H = Font(color="FFFFFF", bold=True, size=10)
    BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    for c, h in enumerate(OUTPUT_HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = FILL_H; cell.font = FONT_H; cell.border = BORDER

    for i, r in enumerate(results):
        row = i + 2
        vals = {
            1: i + 1,                        # 序号
            4: r['route'],                   # 邮路名称
            8: r['plate'],                   # 车牌号
            10: r['volume'] if r['volume'] > 0 else None,  # 容间
            11: r['ton'],                    # 吨位
            17: str(r['vehicle_issue']) if r['vehicle_issue'] else None,  # 备注
            18: '是' if r.get('is_cross') else '否',  # 是否跨档
            19: r.get('prev_volume'),        # 变动前容间
            20: r['prev_date'].strftime('%Y-%m-%d') if r.get('prev_date') else None,
            21: r['coeff'],                  # 系数
        }

        for c in range(1, len(OUTPUT_HEADERS) + 1):
            cell = ws.cell(row=row, column=c, value=vals.get(c))
            cell.border = BORDER

        # 标色
        if r.get('is_missing'):
            for c in range(1, len(OUTPUT_HEADERS) + 1):
                ws.cell(row=row, column=c).fill = FILL_RED
        elif r.get('is_unmeasured'):
            for c in range(1, len(OUTPUT_HEADERS) + 1):
                ws.cell(row=row, column=c).fill = FILL_ORANGE
        elif r.get('is_cross'):
            for c in range(1, len(OUTPUT_HEADERS) + 1):
                ws.cell(row=row, column=c).fill = FILL_YELLOW

    widths = [6,18,16,26,12,14,2,12,2,14,8,8,10,14,10,10,20,8,10,14,8]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    wb_out.save(output_path)
    print(f"✅ {output_path} ({os.path.getsize(output_path)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
