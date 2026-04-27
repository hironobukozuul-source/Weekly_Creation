import streamlit as st
import pandas as pd
import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from io import BytesIO
import os
import requests

# --- フォント設定 (Streamlit Cloudでの日本語文字化け対策) ---
def setup_japanese_font():
    font_path = "NotoSansJP-Regular.ttf"
    if not os.path.exists(font_path):
        # サーバー上にフォントがない場合、GitHubからダウンロード
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf"
        r = requests.get(url)
        with open(font_path, 'wb') as f:
            f.write(r.content) # 
    
    font_prop = font_manager.FontProperties(fname=font_path)
    # デフォルトのフォントとして登録
    font_manager.fontManager.addfont(font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
    return font_prop

# フォントの初期化
jp_font = setup_japanese_font() # 

# --- 定数設定 ---
NAME_MAP = {'Pump': 'Pump', 'Ref1': 'Ref-1', 'Flexible': 'Flexible', 'Ref3': 'Ref-3', 'Ref4': 'Ref-4', 'Ref5': 'Ref-5', 'Awa': 'Awa'} # 
LINE_START_COLS = {'Pump': 2, 'Ref1': 11, 'Flexible': 20, 'Ref3': 29, 'Ref4': 38, 'Ref5': 47, 'Awa': 56} # 
MON_START = datetime.time(3, 30) # 
DATE_COL = 63 # 
MAINT_KEYWORDS = ['P/C', 'CLN', 'SETUP', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC', '原価改定', '段取', 'メンテナンス', '点検', '清掃', '切替', '予備', 'WAIT', 'SAMPLE', 'サンプル'] # 

def to_time(val):
    if isinstance(val, datetime.time): return val # [cite: 3]
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60) # [cite: 3]
    return None

@st.cache_data
def get_available_weeks(df_raw):
    dates = pd.to_datetime(df_raw.iloc[3:, DATE_COL], errors='coerce').dropna()
    mondays = dates[dates.dt.weekday == 0].dt.date.unique() # [cite: 3]
    return sorted(mondays)

def process_tasks(df_raw):
    tasks = []
    line_config = {}
    for line, start_idx in LINE_START_COLS.items():
        found_ton_col = None # [cite: 4]
        for c in range(start_idx, start_idx + 10):
            h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))]
            if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals):
                found_ton_col = c
                break
        line_config[line] = {'prod': start_idx, 'start': start_idx+2, 'finish': start_idx+3, 'ton': found_ton_col or (start_idx+4)} # [cite: 5]

    for i in range(3, len(df_raw)):
        date_val = df_raw.iloc[i, DATE_COL]
        if not isinstance(date_val, (datetime.datetime, pd.Timestamp)): continue
        for line, cols in line_config.items():
            prod_raw, st_raw, fn_raw, tn_raw = df_raw.iloc[i, cols['prod']], df_raw.iloc[i, cols['start']], df_raw.iloc[i, cols['finish']], df_raw.iloc[i, cols['ton']]
            if pd.isna(st_raw) or pd.isna(fn_raw) or pd.isna(prod_raw): continue # [cite: 5, 6]
            product = str(prod_raw).strip()
            if product.lower() in ['nan', '連操なし', '']: continue # [cite: 6]
            try: ton = float(tn_raw) if pd.notna(tn_raw) else 0.0
            except: ton = 0.0
            s_t, f_t = to_time(st_raw), to_time(fn_raw)
            if s_t and f_t: # [cite: 6, 7]
                dt_s, dt_f = datetime.datetime.combine(date_val.date(), s_t), datetime.datetime.combine(date_val.date(), f_t)
                if s_t < MON_START: dt_s += datetime.timedelta(days=1)
                if f_t < MON_START: dt_f += datetime.timedelta(days=1)
                if dt_f <= dt_s: dt_f += datetime.timedelta(days=1) # [cite: 7]
            
                is_m = (ton <= 0) or any(kw.upper() in product.upper() for kw in MAINT_KEYWORDS) # [cite: 8]
                tasks.append({'Line': line, 'Product': product, 'Start': dt_s, 'Finish': dt_f, 'Ton': ton, 'is_maint': is_m})
    return pd.DataFrame(tasks)

# --- ガントチャート生成 ---
def generate_plot(df_tasks, start_date):
    plot_start, plot_end = datetime.datetime.combine(start_date, MON_START), datetime.datetime.combine(start_date, MON_START) + datetime.timedelta(days=7) # [cite: 8]
    requested_order = ['Pump', 'Ref1', 'Flexible', 'Ref3', 'Ref4', 'Ref5', 'Awa']
    plot_order = [NAME_MAP[n] for n in requested_order[::-1]]
    line_to_y = {NAME_MAP[n]: i for i, n in enumerate(requested_order[::-1])} # [cite: 8, 9]

    fig, ax = plt.subplots(figsize=(28, 14), facecolor='white')
    line_offset_state = {line: 30 for line in plot_order}

    merged = []
    for line_key in requested_order:
        line_df = df_tasks[df_tasks['Line'] == line_key].sort_values('Start')
        if line_df.empty: continue
        curr = line_df.iloc[0].to_dict()
        curr['Segments'], curr['TotalTon'] = [(curr['Start'], curr['Finish'])], curr['Ton'] # [cite: 9]
        for idx in range(1, len(line_df)):
            nxt = line_df.iloc[idx].to_dict() # [cite: 10]
            if nxt['Product'] == curr['Product'] and nxt['is_maint'] == curr['is_maint'] and (nxt['Start'] - curr['Finish']) <= datetime.timedelta(hours=4):
                curr['Finish'], curr['TotalTon'] = nxt['Finish'], curr['TotalTon'] + nxt['Ton']
                curr['Segments'].append((nxt['Start'], nxt['Finish'])) # [cite: 10]
            else:
                merged.append(curr); curr = nxt; curr['Segments'], curr['TotalTon'] = [(curr['Start'], curr['Finish'])], curr['Ton'] # [cite: 10, 11]
        merged.append(curr)

    for camp in merged:
        if camp['Finish'] < plot_start or camp['Start'] > plot_end: continue
        line_name = NAME_MAP[camp['Line']]
        y, is_m, color = line_to_y[line_name], camp['is_maint'], ('#7F7F7F' if camp['is_maint'] else '#1F4E78') # [cite: 11]
        for s_dt, f_dt in camp['Segments']:
            s, e = max(mdates.date2num(s_dt), mdates.date2num(plot_start)), min(mdates.date2num(f_dt), mdates.date2num(plot_end))
            if e > s: # [cite: 11, 12]
                if is_m: ax.hlines(y, s, e, colors=color, linestyles='dotted', linewidth=2.5, zorder=3)
                else: ax.hlines(y, s, e, colors=color, linewidth=5, capstyle='butt', zorder=3)
        mid = mdates.date2num(max(camp['Start'], plot_start) + (min(camp['Finish'], plot_end) - max(camp['Start'], plot_start))/2)
        if is_m: ax.text(mid, y + 0.1, camp['Product'], ha='center', va='bottom', fontsize=9, color='#555555', fontweight='bold', fontproperties=jp_font) # [cite: 12]
        else:
            y_off = line_offset_state[line_name]
            line_offset_state[line_name] = -30 if y_off > 0 else 30
            ax.annotate(f"{camp['Product']}\n{camp['TotalTon']:.1f}t", xy=(mid, y), xytext=(0, y_off), textcoords='offset points', ha='center', va=('bottom' if y_off > 0 else 'top'), bbox=dict(boxstyle='square,pad=0.3', fc='white', ec=color, lw=1, alpha=0.9), arrowprops=dict(arrowstyle='->', color=color, connectionstyle='arc3'), fontsize=10, fontweight='bold', fontproperties=jp_font) # [cite: 13]

    ax.set_xlim(mdates.date2num(plot_start), mdates.date2num(plot_end))
    for i in range(9): ax.axvline(mdates.date2num(plot_start + datetime.timedelta(days=i)), color='red', alpha=0.4, linewidth=2, zorder=5)
    curr_h = plot_start # [cite: 13, 14]
    while curr_h <= plot_end:
        ax.axvline(mdates.date2num(curr_h), color='#EEEEEE', linewidth=0.7, zorder=1)
        curr_h += datetime.timedelta(hours=1)
    for i in range(len(plot_order) - 1): ax.axhline(i + 0.5, color='#CCCCCC', linewidth=1.2, alpha=0.8, zorder=2)
    ax.xaxis.set_major_locator(mdates.DayLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter('\n%m/%d (%a)')) # [cite: 14, 15]
    ax.xaxis.set_minor_locator(plt.FixedLocator([mdates.date2num(plot_start + datetime.timedelta(hours=3*i)) for i in range(57)]))
    ax.xaxis.set_minor_formatter(FuncFormatter(lambda x, pos: f"{mdates.num2date(x).hour}"))
    ax.set_yticks(range(len(plot_order))); ax.set_yticklabels(plot_order, fontsize=12, fontweight='bold') # [cite: 15, 16]
    ax.set_ylim(-0.8, len(plot_order) - 0.2)
    plt.title(f"Production Plan - Week of {start_date}", fontsize=16, pad=40, fontproperties=jp_font)
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png')
    plt.close(); buf.seek(0) # [cite: 16, 17]
    return buf, fig

# --- Streamlit UI ---
st.set_page_config(layout="wide", page_title="Production Planner")
st.title("🏭 Production Plan Visualizer & Master Report")

uploaded_file = st.file_uploader("Excelファイルをアップロード (.xlsm)", type=["xlsm"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Fill', header=None)
    available_weeks = get_available_weeks(df_raw)
    
    if available_weeks:
        col1, col2 = st.columns([1, 2])
        with col1: selected_week = st.selectbox("開始日（月曜日）を選択", available_weeks)
        with col2:
            st.write(" ")
            generate_btn = st.button("🚀 Generate Plan & Excel", use_container_width=True) # [cite: 17, 18]

        if generate_btn:
            with st.spinner('処理中...'):
                df_tasks = process_tasks(df_raw)
                img_buf, fig = generate_plot(df_tasks, selected_week) # [cite: 18]
                
                # Excel生成 (Hourly Volume)
                wb = openpyxl.Workbook(); ws_vol = wb.active; ws_vol.title = "Hourly_Volume" # [cite: 19, 20]
                start_dt = datetime.datetime.combine(selected_week, datetime.time(0, 0)) # [cite: 20]
                hour_list = [start_dt + datetime.timedelta(hours=h) for h in range(168)]
                
                thick_black = Side(style='thick', color='000000')
                
                ws_vol.cell(1, 1, "Line").font = Font(bold=True) # [cite: 20, 21]
                ws_vol.cell(1, 2, "Product").font = Font(bold=True)
                for i, h_dt in enumerate(hour_list):
                    cell = ws_vol.cell(1, 3 + i, h_dt.strftime('%m/%d %H:00'))
                    cell.font = Font(bold=True); cell.alignment = Alignment(text_rotation=90, horizontal='center') # [cite: 21, 22]

                clean_tasks = df_tasks[~df_tasks['is_maint']]
                unique_items = sorted(list(set(zip(clean_tasks['Line'], clean_tasks['Product']))))

                for r_idx, (line, product) in enumerate(unique_items):
                    row_num = r_idx + 2
                    ws_vol.cell(row_num, 1, line); ws_vol.cell(row_num, 2, product) # [cite: 22, 23]
                    for c_idx, h_dt in enumerate(hour_list):
                        h_end = h_dt + datetime.timedelta(hours=1)
                        overlap_tasks = clean_tasks[(clean_tasks['Line'] == line) & (clean_tasks['Product'] == product)]
                        ton_sum = sum(t['Ton'] * ((min(t['Finish'], h_end) - max(t['Start'], h_dt)).total_seconds() / (t['Finish'] - t['Start']).total_seconds()) for _, t in overlap_tasks.iterrows() if (min(t['Finish'], h_end) - max(t['Start'], h_dt)).total_seconds() > 0) # [cite: 23, 24]
                        if ton_sum > 0:
                            cell = ws_vol.cell(row_num, 3 + c_idx, round(ton_sum, 2)) # [cite: 24, 25]
                            cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

                for col_idx in range(3, ws_vol.max_column + 1):
                    header = str(ws_vol.cell(1, col_idx).value)
                    if header.endswith("00:00") and not (header.endswith("10:00") or header.endswith("20:00")): # [cite: 25, 26]
                        for r in range(1, ws_vol.max_row + 1):
                            ws_vol.cell(r, col_idx).border = Border(left=thick_black)

                ws_vis = wb.create_sheet("Visual_Schedule")
                ws_vis.add_image(openpyxl.drawing.image.Image(img_buf), 'B2') # [cite: 26, 27]
                out_excel = BytesIO(); wb.save(out_excel) # [cite: 27, 28]
                
                st.success("✅ 生成完了")
                st.download_button("📥 ダウンロード (Excel)", out_excel.getvalue(), f"Report_{selected_week}.xlsx")
                st.pyplot(fig) # [cite: 28]
