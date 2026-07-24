#!/usr/bin/env python3
"""
Step 1: 提取在线表数据并合并
用法: python3 extract_merge.py <月份> <草滩表路径> <港务表路径> <输出路径>
示例: python3 extract_merge.py 4 数据/草滩.xlsx 数据/港务.xlsx 输出/4月合并表.xlsx
"""
import sys, os, re
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter

HEADERS = ['序号', '邮路', '车号', '挂车号', '长(m)', '宽(m)', '高(m)',
           '测量容积(方)', '预报载重(吨)', '未装容积(米)', '车型问题', '备注', '来源']


def extract_month(filepath, month, sheet_pattern, source_label):
    """从文件中提取指定月份的所有数据"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    records = []

    for sname in wb.sheetnames:
        m = sheet_pattern.match(sname.strip())
        if not m:
            continue
        sheet_month = int(m.group(1))
        if sheet_month != month:
            continue

        day = int(m.group(2))
        date = datetime(2026, sheet_month, day)
        ws = wb[sname]

        # 找表头行
        header_row = 3
        for row in range(1, min(8, ws.max_row + 1)):
            c1 = str(ws.cell(row=row, column=1).value or "")
            if '序号' in c1:
                header_row = row
                break

        for row in range(header_row + 1, ws.max_row + 1):
            route = str(ws.cell(row=row, column=2).value or "").strip()
            plate = str(ws.cell(row=row, column=3).value or "").strip()
            if not route and not plate:
                continue
            records.append({
                'date': date, 'sheet': sname,
                'route': route, 'plate': plate,
                'trailer': str(ws.cell(row=row, column=4).value or "").strip(),
                'length': ws.cell(row=row, column=5).value,
                'width': ws.cell(row=row, column=6).value,
                'height': ws.cell(row=row, column=7).value,
                'volume': ws.cell(row=row, column=8).value,
                'weight': ws.cell(row=row, column=9).value,
                'unused_volume': ws.cell(row=row, column=10).value,
                'vehicle_issue': ws.cell(row=row, column=11).value,
                'remark': ws.cell(row=row, column=12).value,
                'source': source_label,
            })

    wb.close()
    return records


def main():
    if len(sys.argv) < 5:
        print("用法: python3 extract_merge.py <月份> <草滩表> <港务表> <输出路径>")
        sys.exit(1)

    month = int(sys.argv[1])
    caotan_path = sys.argv[2]
    gangwu_path = sys.argv[3]
    output_path = sys.argv[4]

    # 提取
    # 草滩: "4月1日" 格式
    ct = extract_month(caotan_path, month,
                       re.compile(r'^(\d+)月(\d+)日?$'), '草滩')
    # 港务: "4.1" 格式
    gw = extract_month(gangwu_path, month,
                       re.compile(r'^(\d+)\.(\d+)$'), '港务')

    print(f"草滩: {len(ct)} 条, 港务: {len(gw)} 条")

    # 合并（不去重）
    all_records = ct + gw
    print(f"合并: {len(all_records)} 条")

    # 写入
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    by_date = defaultdict(list)
    for r in all_records:
        by_date[r['date']].append(r)

    FILL_H = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    FONT_H = Font(color="FFFFFF", bold=True, size=10)
    BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    for date in sorted(by_date.keys()):
        sheet_name = f"{date.month}月{date.day}日"
        ws = wb_out.create_sheet(title=sheet_name)

        for c, h in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill = FILL_H; cell.font = FONT_H; cell.border = BORDER

        for i, r in enumerate(by_date[date]):
            row = i + 2
            vals = [i+1, r['route'], r['plate'], r['trailer'],
                    r['length'], r['width'], r['height'], r['volume'],
                    r['weight'], r['unused_volume'], r['vehicle_issue'],
                    r['remark'], r['source']]
            for c, v in enumerate(vals, 1):
                ws.cell(row=row, column=c, value=v).border = BORDER

        widths = [6, 28, 12, 14, 8, 8, 8, 14, 14, 10, 12, 20, 8]
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w

    wb_out.save(output_path)
    print(f"✅ {output_path} ({os.path.getsize(output_path)/1024:.0f} KB)")


if __name__ == "__main__":
    main()
