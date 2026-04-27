import pandas as pd
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill

# --- 1. 設定とフィルタリング条件 ---
FILE_PATH = 'Fill_26_06cy_rev01.xlsm'
OUTPUT_FILE = "Production_Master_Report.xlsx"
DAY_START_TIME = datetime.time(3, 30)

# メンテナンス除外キーワード
MAINT_KEYWORDS = [
    'P/C', 'CLN', 'SETUP', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC',
    '原価改定', '段取', 'メンテナンス', '点検', '清掃', '切替', '予備',
    'WAIT', 'SAMPLE', 'サンプル', 'P/C洗浄', '段取り'
]

# --- 2. データ抽出ロジック ---
def to_time(val):
    if isinstance(val, datetime.time): return val
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60)
    return None

df_raw = pd.read_excel(FILE_PATH, sheet_name='Fill', header=None)

line_start_cols = {
    'Pump': 2, 'Ref1': 11, 'Flexible': 20, 'Ref3': 29, 
    'Ref4': 38, 'Ref5': 47, 'Awa': 56
}

# 各ラインの列構成を特定
line_config = {}
for line, start_idx in line_start_cols.items():
    found_ton_col = None
    for c in range(start_idx, start_idx + 10):
        h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))]
        if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals):
            found_ton_col = c
            break
    line_config[line] = {
        'prod': start_idx, 
        'start': start_idx + 2, 
        'finish': start_idx + 3, 
        'ton': found_ton_col or (start_idx + 4)
    }

date_col = 63
tasks = []
for i in range(3, len(df_raw)):
    date_val = df_raw.iloc[i, date_col]
    if not isinstance(date_val, (datetime.datetime, pd.Timestamp)): continue
    
    for line, cols in line_config.items():
        prod_raw = df_raw.iloc[i, cols['prod']]
        st_raw = df_raw.iloc[i, cols['start']]
        fn_raw = df_raw.iloc[i, cols['finish']]
        tn_raw = df_raw.iloc[i, cols['ton']]
        
        if pd.isna(st_raw) or pd.isna(fn_raw) or pd.isna(prod_raw): continue
        
        product = str(prod_raw).strip()
        if product.lower() in ['nan', '連操なし', '']: continue
        
        try: ton = float(tn_raw) if pd.notna(tn_raw) else 0.0
        except: ton = 0.0
        
        s_time, f_time = to_time(st_raw), to_time(fn_raw)
        if s_time and f_time:
            dt_s = datetime.datetime.combine(date_val.date(), s_time)
            dt_f = datetime.datetime.combine(date_val.date(), f_time)
            if s_time < DAY_START_TIME: dt_s += datetime.timedelta(days=1)
            if f_time < DAY_START_TIME: dt_f += datetime.timedelta(days=1)
            if dt_f <= dt_s: dt_f += datetime.timedelta(days=1)
            
            # メンテナンス判定
            is_maint = (ton <= 0) or any(kw.upper() in product.upper() for kw in MAINT_KEYWORDS)
            tasks.append({'Line': line, 'Product': product, 'Start': dt_s, 'Finish': dt_f, 'Ton': ton, 'is_maint': is_maint})

# --- 3. EXCELシート生成 (Hourly Volume) ---
wb = openpyxl.Workbook()
ws_vol = wb.active
ws_vol.title = "Hourly_Production_Volume"

# 1週間の時間軸作成
start_week = datetime.datetime(2026, 6, 1, 0, 0) # 基準日
hour_list = [start_week + datetime.timedelta(hours=h) for h in range(168)]

# ヘッダー作成
thick_black = Side(style='thick', color='000000')
ws_vol.cell(1, 1, "Line").font = Font(bold=True)
ws_vol.cell(1, 2, "Product").font = Font(bold=True)

for i, h_dt in enumerate(hour_list):
    cell = ws_vol.cell(1, 3 + i, h_dt.strftime('%m/%d %H:00'))
    cell.font = Font(bold=True)
    cell.alignment = Alignment(text_rotation=90, horizontal='center')

# データの書き込み (メンテナンス除外)
clean_tasks = [t for t in tasks if not t['is_maint']]
unique_items = sorted(list(set([(t['Line'], t['Product']) for t in clean_tasks])))

for r_idx, (line, product) in enumerate(unique_items):
    row_num = r_idx + 2
    ws_vol.cell(row_num, 1, line)
    ws_vol.cell(row_num, 2, product)
    for c_idx, h_dt in enumerate(hour_list):
        h_end = h_dt + datetime.timedelta(hours=1)
        ton_sum = 0
        for t in clean_tasks:
            if t['Line'] == line and t['Product'] == product:
                overlap = min(t['Finish'], h_end) - max(t['Start'], h_dt)
                if overlap.total_seconds() > 0:
                    dur = (t['Finish'] - t['Start']).total_seconds()
                    ton_sum += t['Ton'] * (overlap.total_seconds() / dur)
        if ton_sum > 0:
            cell = ws_vol.cell(row_num, 3 + c_idx, round(ton_sum, 2))
            cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

# 0:00に太い黒線を引く (修正済みロジック)
for col_idx in range(3, ws_vol.max_column + 1):
    header = str(ws_vol.cell(1, col_idx).value)
    if header.endswith("00:00") and not (header.endswith("10:00") or header.endswith("20:00")):
        for r in range(1, ws_vol.max_row + 1):
            cell = ws_vol.cell(r, col_idx)
            cell.border = Border(left=thick_black, right=cell.border.right, top=cell.border.top, bottom=cell.border.bottom)

# --- 4. ビジュアル生成と挿入 ---
name_map = {'Pump': 'Pump', 'Ref1': 'Ref-1', 'Flexible': 'Flexible', 'Ref3': 'Ref-3', 'Ref4': 'Ref-4', 'Ref5': 'Ref-5', 'Awa': 'Awa'}
requested_order = ['Pump', 'Ref1', 'Flexible', 'Ref3', 'Ref4', 'Ref5', 'Awa']
plot_order_names = [name_map[n] for n in requested_order[::-1]]
line_to_y = {name_map[n]: i for i, n in enumerate(requested_order[::-1])}

plt.figure(figsize=(24, 12))
for t in tasks:
    y = line_to_y[name_map[t['Line']]]
    s, e = mdates.date2num(t['Start']), mdates.date2num(t['Finish'])
    color = '#7F7F7F' if t['is_maint'] else '#1F4E78'
    plt.hlines(y, s, e, colors=color, linewidth=6 if not t['is_maint'] else 2)

plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
plt.yticks(range(len(plot_order_names)), plot_order_names)
plt.savefig("temp_plan.png", bbox_inches='tight')

ws_vis = wb.create_sheet("Visual_Schedule")
ws_vis.add_image(openpyxl.drawing.image.Image("temp_plan.png"), 'B2')

wb.save(OUTPUT_FILE)
print(f"完了: {OUTPUT_FILE}")
