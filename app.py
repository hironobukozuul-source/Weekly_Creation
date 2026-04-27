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

# --- フォント設定 ---
def setup_japanese_font():
    font_path = "NotoSansJP-Regular.ttf"
    if not os.path.exists(font_path):
        url = "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf"
        r = requests.get(url)
        with open(font_path, 'wb') as f:
            f.write(r.content) [cite: 1, 2]
    
    font_prop = font_manager.FontProperties(fname=font_path) [cite: 2]
    font_manager.fontManager.addfont(font_path) [cite: 2]
    plt.rcParams['font.family'] = font_prop.get_name() [cite: 2]
    return font_prop

jp_font = setup_japanese_font() [cite: 2]

# --- 定数設定 (Shared Code Logic) ---
NAME_MAP = {'Pump': 'Pump', 'Ref1': 'Ref-1', 'Flexible': 'Flexible', 'Ref3': 'Ref-3', 'Ref4': 'Ref-4', 'Ref5': 'Ref-5', 'Awa': 'Awa'} [cite: 2]
LINE_START_COLS = {'Pump': 2, 'Ref1': 11, 'Flexible': 20, 'Ref3': 29, 'Ref4': 38, 'Ref5': 47, 'Awa': 56} [cite: 2]
MON_START = datetime.time(3, 30) [cite: 2]
DATE_COL = 63 [cite: 2]
MAINT_KEYWORDS = ['P/C', 'CLN', 'SETUP', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC', '原価改定', '段取', 'メンテナンス', '点検', '清掃', '切替', '予備', 'WAIT', 'SAMPLE', 'サンプル'] [cite: 2]

def to_time(val):
    if isinstance(val, datetime.time): return val [cite: 3]
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60) [cite: 3]
    return None [cite: 3]

@st.cache_data
def get_available_weeks(df_raw):
    dates = pd.to_datetime(df_raw.iloc[3:, DATE_COL], errors='coerce').dropna() [cite: 3]
    mondays = dates[dates.dt.weekday == 0].dt.date.unique() [cite: 3]
    return sorted(mondays) [cite: 3]

def process_tasks(df_raw):
    tasks = [] [cite: 3]
    line_config = {} [cite: 3]
    for line, start_idx in LINE_START_COLS.items(): [cite: 3]
        found_ton_col = None [cite: 4]
        for c in range(start_idx, start_idx + 10): [cite: 4]
            h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))] [cite: 4]
            if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals): [cite: 4]
                found_ton_col = c [cite: 4]
                break
        line_config[line] = {'prod': start_idx, 'start': start_idx+2, 'finish': start_idx+3, 'ton': found_ton_col or (start_idx+4)} [cite: 5]

    for i in range(3, len(df_raw)): [cite: 5]
        date_val = df_raw.iloc[i, DATE_COL] [cite: 5]
        if not isinstance(date_val, (datetime.datetime, pd.Timestamp)): continue [cite: 5]
        for line, cols in line_config.items(): [cite: 5]
            prod_raw, st_raw, fn_raw, tn_raw = df_raw.iloc[i, cols['prod']], df_raw.iloc[i, cols['start']], df_raw.iloc[i, cols['finish']], df_raw.iloc[i, cols['ton']] [cite: 5]
            if pd.isna(st_raw) or pd.isna(fn_raw) or pd.isna(prod_raw): continue [cite: 6]
            product = str(prod_raw).strip() [cite: 6]
            if product.lower() in ['nan', '連操なし', '']: continue [cite: 6]
            try: ton = float(tn_raw) if pd.notna(tn_raw) else 0.0 [cite: 6]
            except: ton = 0.0 [cite: 6]
            s_t, f_t = to_time(st_raw), to_time(fn_raw) [cite: 6]
            if s_t and f_t: [cite: 7]
                dt_s, dt_f = datetime.datetime.combine(date_val.date(), s_t), datetime.datetime.combine(date_val.date(), f_t) [cite: 7]
                if s_t < MON_START: dt_s += datetime.timedelta(days=1) [cite: 7]
                if f_t < MON_START: dt_f += datetime.timedelta(days=1) [cite: 7]
                if dt_f <= dt_s: dt_f += datetime.timedelta(days=1) [cite: 7]
                is_m = (ton <= 0) or any(kw.upper() in product.upper() for kw in MAINT_KEYWORDS) [cite: 8]
                tasks.append({'Line': line, 'Product': product, 'Start': dt_s, 'Finish': dt_f, 'Ton': ton, 'is_maint': is_m}) [cite: 8]
    return pd.DataFrame(tasks) [cite: 8]

def generate_plot(df_tasks, start_date):
    plot_start, plot_end = datetime.datetime.combine(start_date, MON_START), datetime.datetime.combine(start_date, MON_START) + datetime.timedelta(days=7) [cite: 8]
    requested_order = ['Pump', 'Ref1', 'Flexible', 'Ref3', 'Ref4', 'Ref5', 'Awa'] [cite: 8]
    plot_order = [NAME_MAP[n] for n in requested_order[::-1]] [cite: 8]
    line_to_y = {NAME_MAP[n]: i for i, n in enumerate(requested_order[::-1])} [cite: 9]

    fig, ax = plt.subplots(figsize=(28, 14), facecolor='white') [cite: 9]
    line_offset_state = {line: 30 for line in plot_order} [cite: 9]

    merged = [] [cite: 9]
    for line_key in requested_order: [cite: 9]
        line_df = df_tasks[df_tasks['Line'] == line_key].sort_values('Start') [cite: 9]
        if line_df.empty: continue [cite: 9]
        curr = line_df.iloc[0].to_dict() [cite: 9]
        curr['Segments'], curr['TotalTon'] = [(curr['Start'], curr['Finish'])], curr['Ton'] [cite: 9]
        for idx in range(1, len(line_df)): [cite: 9]
            nxt = line_df.iloc[idx].to_dict() [cite: 10]
            if nxt['Product'] == curr['Product'] and nxt['is_maint'] == curr['is_maint'] and (nxt['Start'] - curr['Finish']) <= datetime.timedelta(hours=4): [cite: 10]
                curr['Finish'], curr['TotalTon'] = nxt['Finish'], curr['TotalTon'] + nxt['Ton'] [cite: 10]
                curr['Segments'].append((nxt['Start'], nxt['Finish'])) [cite: 10]
            else:
                merged.append(curr) [cite: 10]
                curr = nxt [cite: 11]
                curr['Segments'], curr['TotalTon'] = [(curr['Start'], curr['Finish'])], curr['Ton'] [cite: 11]
        merged.append(curr) [cite: 11]

    for camp in merged: [cite: 11]
        if camp['Finish'] < plot_start or camp['Start'] > plot_end: continue [cite: 11]
        line_name = NAME_MAP[camp['Line']] [cite: 11]
        y, is_m, color = line_to_y[line_name], camp['is_maint'], ('#7F7F7F' if camp['is_maint'] else '#1F4E78') [cite: 11]
        for s_dt, f_dt in camp['Segments']: [cite: 11]
            s, e = max(mdates.date2num(s_dt), mdates.date2num(plot_start)), min(mdates.date2num(f_dt), mdates.date2num(plot_end)) [cite: 11]
            if e > s: [cite: 12]
                if is_m: ax.hlines(y, s, e, colors=color, linestyles='dotted', linewidth=2.5, zorder=3) [cite: 12]
                else: ax.hlines(y, s, e, colors=color, linewidth=5, capstyle='butt', zorder=3) [cite: 12]
        mid = mdates.date2num(max(camp['Start'], plot_start) + (min(camp['Finish'], plot_end) - max(camp['Start'], plot_start))/2) [cite: 12]
        if is_m: ax.text(mid, y + 0.1, camp['Product'], ha='center', va='bottom', fontsize=9, color='#555555', fontweight='bold', fontproperties=jp_font) [cite: 12]
        else: [cite: 13]
            y_off = line_offset_state[line_name] [cite: 13]
            line_offset_state[line_name] = -30 if y_off > 0 else 30 [cite: 13]
            ax.annotate(f"{camp['Product']}\n{camp['TotalTon']:.1f}t", xy=(mid, y), xytext=(0, y_off), textcoords='offset points', ha='center', va=('bottom' if y_off > 0 else 'top'), bbox=dict(boxstyle='square,pad=0.3', fc='white', ec=color, lw=1, alpha=0.9), arrowprops=dict(arrowstyle='->', color=color, connectionstyle='arc3'), fontsize=10, fontweight='bold', fontproperties=jp_font) [cite: 13]

    ax.set_xlim(mdates.date2num(plot_start), mdates.date2num(plot_end)) [cite: 13]
    for i in range(9): ax.axvline(mdates.date2num(plot_start + datetime.timedelta(days=i)), color='red', alpha=0.4, linewidth=2, zorder=5) [cite: 13]
    curr_h = plot_start [cite: 14]
    while curr_h <= plot_end: [cite: 14]
        ax.axvline(mdates.date2num(curr_h), color='#EEEEEE', linewidth=0.7, zorder=1) [cite: 14]
        curr_h += datetime.timedelta(hours=1) [cite: 14]
    for i in range(len(plot_order) - 1): ax.axhline(i + 0.5, color='#CCCCCC', linewidth=1.2, alpha=0.8, zorder=2) [cite: 14]
    ax.xaxis.set_major_locator(mdates.DayLocator()) [cite: 14]
    ax.xaxis.set_major_formatter(mdates.DateFormatter('\n%m/%d (%a)')) [cite: 15]
    ax.xaxis.set_minor_locator(plt.FixedLocator([mdates.date2num(plot_start + datetime.timedelta(hours=3*i)) for i in range(57)])) [cite: 15]
    ax.xaxis.set_minor_formatter(FuncFormatter(lambda x, pos: f"{mdates.num2date(x).hour}")) [cite: 15]
    ax.set_yticks(range(len(plot_order))) [cite: 15]
    ax.set_yticklabels(plot_order, fontsize=12, fontweight='bold') [cite: 16]
    ax.set_ylim(-0.8, len(plot_order) - 0.2) [cite: 16]
    plt.title(f"Production Plan - Week of {start_date}", fontsize=16, pad=40, fontproperties=jp_font) [cite: 16]
    plt.tight_layout() [cite: 16]
    buf = BytesIO() [cite: 16]
    plt.savefig(buf, format='png') [cite: 16]
    plt.close() [cite: 16]
    buf.seek(0) [cite: 17]
    return buf, fig [cite: 17]

# --- Streamlit UI ---
st.set_page_config(layout="wide", page_title="Production Planner") [cite: 17]
st.title("🏭 Production Plan Visualizer & Master Report") [cite: 17]

uploaded_file = st.file_uploader("Excelファイルをアップロード (.xlsm)", type=["xlsm"]) [cite: 17]

if uploaded_file: [cite: 17]
    df_raw = pd.read_excel(uploaded_file, sheet_name='Fill', header=None) [cite: 17]
    available_weeks = get_available_weeks(df_raw) [cite: 17]
    
    if available_weeks: [cite: 17]
        col1, col2 = st.columns([1, 2]) [cite: 17]
        with col1: selected_week = st.selectbox("開始日（月曜日）を選択", available_weeks) [cite: 17]
        with col2:
            st.write(" ")
            generate_btn = st.button("🚀 Generate Plan & Excel", use_container_width=True) [cite: 18]

        if generate_btn: [cite: 18]
            with st.spinner('処理中...'): [cite: 18]
                df_tasks = process_tasks(df_raw) [cite: 18]
                img_buf, fig = generate_plot(df_tasks, selected_week) [cite: 18]
                
                # --- Excel生成 (Monday Logic Applied: Start at 03:30) ---
                wb = openpyxl.Workbook() [cite: 19]
                ws_vol = wb.active [cite: 20]
                ws_vol.title = "Hourly_Volume" [cite: 20]
                
                # 起点を月曜の03:30に設定 [cite: 2, 8]
                start_dt = datetime.datetime.combine(selected_week, MON_START) [cite: 2, 8]
                hour_list = [start_dt + datetime.timedelta(hours=h) for h in range(168)] [cite: 20]
                
                thick_black = Side(style='thick', color='000000') [cite: 20]
                ws_vol.cell(1, 1, "Line").font = Font(bold=True) [cite: 21]
                ws_vol.cell(1, 2, "Product").font = Font(bold=True) [cite: 21]
                for i, h_dt in enumerate(hour_list): [cite: 21]
                    cell = ws_vol.cell(1, 3 + i, h_dt.strftime('%m/%d %H:%M')) [cite: 21]
                    cell.font = Font(bold=True) [cite: 21]
                    cell.alignment = Alignment(text_rotation=90, horizontal='center') [cite: 22]

                clean_tasks = df_tasks[~df_tasks['is_maint']] [cite: 22]
                unique_items = sorted(list(set(zip(clean_tasks['Line'], clean_tasks['Product'])))) [cite: 22]

                for r_idx, (line, product) in enumerate(unique_items): [cite: 22]
                    row_num = r_idx + 2 [cite: 22]
                    ws_vol.cell(row_num, 1, line) [cite: 23]
                    ws_vol.cell(row_num, 2, product) [cite: 23]
                    for c_idx, h_dt in enumerate(hour_list): [cite: 23]
                        h_end = h_dt + datetime.timedelta(hours=1) [cite: 23]
                        overlap_tasks = clean_tasks[(clean_tasks['Line'] == line) & (clean_tasks['Product'] == product)] [cite: 23]
                        ton_sum = sum(t['Ton'] * ((min(t['Finish'], h_end) - max(t['Start'], h_dt)).total_seconds() / (t['Finish'] - t['Start']).total_seconds()) 
                                      for _, t in overlap_tasks.iterrows() if (min(t['Finish'], h_end) - max(t['Start'], h_dt)).total_seconds() > 0) [cite: 24]
                        if ton_sum > 0: [cite: 24]
                            cell = ws_vol.cell(row_num, 3 + c_idx, round(ton_sum, 2)) [cite: 25]
                            cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid") [cite: 25]

                for col_idx in range(3, ws_vol.max_column + 1): [cite: 25]
                    header = str(ws_vol.cell(1, col_idx).value) [cite: 25]
                    # 日付の区切り線 (03:30) に太線を引く [cite: 26]
                    if "03:30" in header: [cite: 26]
                        for r in range(1, ws_vol.max_row + 1): [cite: 26]
                            ws_vol.cell(r, col_idx).border = Border(left=thick_black) [cite: 26]

                ws_vis = wb.create_sheet("Visual_Schedule") [cite: 26]
                img_copy = BytesIO(img_buf.getvalue()) # エラー回避のためのコピー
                ws_vis.add_image(openpyxl.drawing.image.Image(img_copy), 'B2') [cite: 27]
                
                out_excel = BytesIO() [cite: 27]
                wb.save(out_excel) [cite: 28]
                
                st.success("✅ 生成完了") [cite: 28]
                st.download_button("📥 ダウンロード (Excel)", out_excel.getvalue(), f"Report_{selected_week}.xlsx") [cite: 28]
                st.pyplot(fig) [cite: 28]
