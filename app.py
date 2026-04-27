import streamlit as st
import pandas as pd
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
import io

# --- 1. 定数・設定 ---
NAME_MAP = {
    'Pump': 'Pump', 'Ref1': 'Ref-1', 'Flexible': 'Flexible', 
    'Ref3': 'Ref-3', 'Ref4': 'Ref-4', 'Ref5': 'Ref-5', 'Awa': 'Awa'
}
LINE_START_COLS = {
    'Pump': 2, 'Ref1': 11, 'Flexible': 20, 'Ref3': 29, 
    'Ref4': 38, 'Ref5': 47, 'Awa': 56
}
MON_START = datetime.time(3, 30)
STD_START = datetime.time(7, 0)
DATE_COL = 63

# --- 2. 補助関数 ---
def to_time(val):
    if isinstance(val, datetime.time): return val
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60)
    return None

def process_tasks(df_raw):
    line_config_dynamic = {}
    for line, start_idx in LINE_START_COLS.items():
        found_ton_col = None
        for c in range(start_idx, start_idx + 10):
            h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))]
            if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals):
                found_ton_col = c
                break
        if found_ton_col is None: found_ton_col = start_idx + 4
        line_config_dynamic[line] = {'prod': start_idx, 'start': start_idx + 2, 'finish': start_idx + 3, 'ton': found_ton_col}

    tasks = []
    last_dt_per_line = {line: None for line in line_config_dynamic}
    last_date_per_line = {line: None for line in line_config_dynamic}

    for i in range(3, len(df_raw)):
        date_val = df_raw.iloc[i, DATE_COL]
        if pd.isna(date_val) or not isinstance(date_val, (datetime.datetime, pd.Timestamp)):
            continue
        
        current_date = date_val.date()
        shift_start_time = MON_START if current_date.weekday() == 0 else STD_START

        for line, cols in line_config_dynamic.items():
            product_raw = df_raw.iloc[i, cols['prod']]
            start_t_raw = df_raw.iloc[i, cols['start']]
            finish_t_raw = df_raw.iloc[i, cols['finish']]
            ton_raw = df_raw.iloc[i, cols['ton']]
            
            if pd.isna(start_t_raw) or pd.isna(finish_t_raw) or pd.isna(product_raw): continue
            product = str(product_raw).strip()
            if product.lower() in ['nan', '連操なし', '']: continue
            
            s_time = to_time(start_t_raw)
            f_time = to_time(finish_t_raw)
            try: ton = float(ton_raw) if pd.notna(ton_raw) else 0.0
            except: ton = 0.0

            if s_time and f_time:
                if last_date_per_line[line] != current_date:
                    base_dt = datetime.datetime.combine(current_date, shift_start_time)
                    last_date_per_line[line] = current_date
                else:
                    base_dt = last_dt_per_line[line]

                s_dt = datetime.datetime.combine(base_dt.date(), s_time)
                if s_dt < base_dt: s_dt += datetime.timedelta(days=1)
                f_dt = datetime.datetime.combine(s_dt.date(), f_time)
                if f_dt <= s_dt: f_dt += datetime.timedelta(days=1)
                
                last_dt_per_line[line] = f_dt
                is_maint = (ton == 0) or any(x in product for x in ['P/C', 'CLN', 'Setup', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC', '原価改定'])
                tasks.append({'Line': line, 'Product': product, 'Start': s_dt, 'Finish': f_dt, 'Ton': ton, 'is_maint': is_maint})
    return pd.DataFrame(tasks)

# --- 3. プロット関数 ---
def generate_plot(df_tasks, start_date):
    plot_start = datetime.datetime.combine(start_date, datetime.time(0, 0))
    plot_end = plot_start + datetime.timedelta(days=7)
    requested_order = ['Pump', 'Ref1', 'Flexible', 'Ref3', 'Ref4', 'Ref5', 'Awa']
    plot_order = [NAME_MAP[n] for n in requested_order[::-1]]
    line_to_y = {NAME_MAP[n]: i for i, n in enumerate(requested_order[::-1])}

    fig, ax = plt.subplots(figsize=(56, 28), facecolor='white')
    
    # Hour Strips
    mid_points = [i + 0.5 for i in range(len(plot_order) - 1)]
    h_unit = 0.038
    for y_mid in mid_points:
        ax.axhspan(y_mid - h_unit, y_mid + h_unit, color='#F8F8F8', alpha=0.9, zorder=1.5)
        curr_t = plot_start
        while curr_t < plot_end:
            if curr_t.hour % 3 == 0:
                ax.text(mdates.date2num(curr_t), y_mid, f"{curr_t.hour}", ha='center', va='center', fontsize=20, color='#666666', fontweight='bold', zorder=2)
            curr_t += datetime.timedelta(hours=3)

    line_offset_state = {line: 60 for line in plot_order}
    for line_key in requested_order:
        line_df = df_tasks[df_tasks['Line'] == line_key].sort_values('Start')
        if line_df.empty: continue
        # (Merging logic simplified for brevity)
        merged_list = []
        curr = line_df.iloc[0].to_dict()
        curr['Segments'] = [(curr['Start'], curr['Finish'])]; curr['TotalTon'] = curr['Ton']
        for idx in range(1, len(line_df)):
            nxt = line_df.iloc[idx].to_dict()
            if nxt['Product'] == curr['Product'] and nxt['is_maint'] == curr['is_maint'] and (nxt['Start'] - curr['Finish']) <= datetime.timedelta(hours=4):
                curr['Finish'] = nxt['Finish']; curr['TotalTon'] += nxt['Ton']; curr['Segments'].append((nxt['Start'], nxt['Finish']))
            else:
                merged_list.append(curr); curr = nxt; curr['Segments'] = [(curr['Start'], curr['Finish'])]; curr['TotalTon'] = curr['Ton']
        merged_list.append(curr)

        y = line_to_y[NAME_MAP[line_key]]
        for camp in merged_list:
            if camp['Finish'] < plot_start or camp['Start'] > plot_end: continue
            is_m = camp['is_maint']; color = '#1F4E78' if not is_m else '#7F7F7F'
            for s_dt, f_dt in camp['Segments']:
                s, e = max(mdates.date2num(s_dt), mdates.date2num(plot_start)), min(mdates.date2num(f_dt), mdates.date2num(plot_end))
                if e > s:
                    if is_m: ax.hlines(y, s, e, colors=color, linestyles='dotted', linewidth=4, zorder=3)
                    else:
                        ax.hlines(y, s, e, colors=color, linewidth=12, capstyle='butt', zorder=3)
                        ax.plot([s, s], [y-0.04, y+0.04], [e, e], [y-0.04, y+0.04], color=color, linewidth=3, zorder=4)
            mid_t = mdates.date2num(max(camp['Start'], plot_start) + (min(camp['Finish'], plot_end) - max(camp['Start'], plot_start))/2)
            if is_m: ax.text(mid_t, y + 0.12, camp['Product'], ha='center', va='bottom', fontsize=16, color='#444444', fontweight='bold')
            else:
                y_off = line_offset_state[NAME_MAP[line_key]]; va = 'bottom' if y_off > 0 else 'top'; line_offset_state[NAME_MAP[line_key]] *= -1
                ax.annotate(f"{camp['Product']}\n{camp['TotalTon']:.1f}t", xy=(mid_t, y), xytext=(0, y_off), textcoords='offset points', ha='center', va=va, bbox=dict(boxstyle='square,pad=0.3', fc='white', ec=color, lw=2, alpha=0.9), arrowprops=dict(arrowstyle='->', color=color, connectionstyle='arc3', lw=2), fontsize=18, fontweight='bold')

    ax.set_xlim(mdates.date2num(plot_start), mdates.date2num(plot_end))
    ax.xaxis.set_major_locator(mdates.DayLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter('\n%m/%d (%a)'))
    all_3h = [mdates.date2num(plot_start + datetime.timedelta(hours=3*i)) for i in range(57)]
    ax.xaxis.set_minor_locator(plt.FixedLocator(all_3h)); ax.xaxis.set_minor_formatter(FuncFormatter(lambda x, pos: f"{mdates.num2date(x).hour}"))
    ax.set_yticks(range(len(plot_order))); ax.set_yticklabels(plot_order, fontsize=24, fontweight='bold')
    for i in range(8): ax.axvline(mdates.date2num(plot_start + datetime.timedelta(days=i)), color='red', alpha=0.4, linewidth=4)
    plt.title(f"Production Plan - {start_date}", fontsize=32, pad=60)
    plt.tight_layout()
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png')
    plt.close()
    return img_buf

# --- 4. Streamlit UI ---
st.set_page_config(layout="wide")
st.title("Production Plan Visualizer")

uploaded_file = st.file_uploader("Excelファイルをアップロード (.xlsm)", type=["xlsm"])
start_date = st.date_input("表示開始日を選択", datetime.date(2026, 6, 1))

if uploaded_file:
    with st.spinner('データを処理中...'):
        df_raw = pd.read_excel(uploaded_file, sheet_name='Fill', header=None)
        df_tasks = process_tasks(df_raw)
        img_buf = generate_plot(df_tasks, start_date)
        
        st.image(img_buf, use_column_width=True)
        st.download_button(label="画像をダウンロード", data=img_buf.getvalue(), file_name=f"plan_{start_date}.png", mime="image/png")
