import streamlit as st
import pandas as pd
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
import io
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill

# --- 定数設定 (Original) ---
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

# --- データ処理ロジック (Original) ---
def to_time(val):
    if isinstance(val, datetime.time): return val
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60)
    return None

@st.cache_data
def get_available_weeks(df_raw):
    """ファイル内から月曜日の日付を抽出"""
    dates = pd.to_datetime(df_raw.iloc[3:, DATE_COL], errors='coerce').dropna()
    mondays = dates[dates.dt.weekday == 0].dt.date.unique()
    return sorted(mondays)

def process_tasks(df_raw):
    tasks = []
    line_config_dynamic = {}
    for line, start_idx in LINE_START_COLS.items():
        found_ton_col = None
        for c in range(start_idx, start_idx + 10):
            h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))]
            if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals):
                found_ton_col = c
                break
        line_config_dynamic[line] = {
            'prod': start_idx, 'start': start_idx + 2, 'finish': start_idx + 3, 'ton': found_ton_col or (start_idx + 4)
        }

    for i in range(3, len(df_raw)):
        date_val = df_raw.iloc[i, DATE_COL]
        if not isinstance(date_val, (datetime.datetime, pd.Timestamp)):
            continue
        for line, cols in line_config_dynamic.items():
            product_raw = df_raw.iloc[i, cols['prod']]
            start_t_raw = df_raw.iloc[i, cols['start']]
            finish_t_raw = df_raw.iloc[i, cols['finish']]
            ton_raw = df_raw.iloc[i, cols['ton']]
            if pd.isna(start_t_raw) or pd.isna(finish_t_raw) or pd.isna(product_raw):
                continue
            product = str(product_raw).strip()
            if product.lower() in ['nan', '連操なし', '']: continue
            s_time, f_time = to_time(start_t_raw), to_time(finish_t_raw)
            try: ton = float(ton_raw) if pd.notna(ton_raw) else 0.0
            except: ton = 0.0
            if s_time and f_time:
                dt_s = datetime.datetime.combine(date_val.date(), s_time)
                dt_f = datetime.datetime.combine(date_val.date(), f_time)
                if s_time < MON_START: dt_s += datetime.timedelta(days=1)
                if f_time < MON_START: dt_f += datetime.timedelta(days=1)
                if dt_f <= dt_s: dt_f += datetime.timedelta(days=1)
                is_maint = (ton == 0) or any(x in product for x in ['P/C', 'CLN', 'Setup', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC', '原価改定'])
                tasks.append({'Line': line, 'Product': product, 'Start': dt_s, 'Finish': dt_f, 'Ton': ton, 'is_maint': is_maint})
    return pd.DataFrame(tasks)

# --- 視覚化ロジック (Original - 変更なし) ---
def generate_plot(df_tasks, start_date):
    plot_start = datetime.datetime.combine(start_date, MON_START)
    plot_end = plot_start + datetime.timedelta(days=7)
    requested_order = ['Pump', 'Ref1', 'Flexible', 'Ref3', 'Ref4', 'Ref5', 'Awa']
    plot_order = [NAME_MAP[n] for n in requested_order[::-1]]
    line_to_y = {NAME_MAP[n]: i for i, n in enumerate(requested_order[::-1])}

    fig, ax = plt.subplots(figsize=(28, 14), facecolor='white')
    line_offset_state = {line: 30 for line in plot_order}

    merged = []
    for line_key in requested_order:
        line_df = df_tasks[df_tasks['Line'] == line_key].sort_values('Start')
        if line_df.empty: continue
        curr = line_df.iloc[0].to_dict()
        curr['Segments'] = [(curr['Start'], curr['Finish'])]
        curr['TotalTon'] = curr['Ton']
        for idx in range(1, len(line_df)):
            nxt = line_df.iloc[idx].to_dict()
            if nxt['Product'] == curr['Product'] and nxt['is_maint'] == curr['is_maint'] and (nxt['Start'] - curr['Finish']) <= datetime.timedelta(hours=4):
                curr['Finish'] = nxt['Finish']
                curr['TotalTon'] += nxt['Ton']
                curr['Segments'].append((nxt['Start'], nxt['Finish']))
            else:
                merged.append(curr)
                curr = nxt
                curr['Segments'] = [(curr['Start'], curr['Finish'])]
                curr['TotalTon'] = curr['Ton']
        merged.append(curr)

    for camp in merged:
        if camp['Finish'] < plot_start or camp['Start'] > plot_end: continue
        line_name = NAME_MAP[camp['Line']]
        y = line_to_y[line_name]
        is_m = camp['is_maint']
        color = '#7F7F7F' if is_m else '#1F4E78'
        for s_dt, f_dt in camp['Segments']:
            s = max(mdates.date2num(s_dt), mdates.date2num(plot_start))
            e = min(mdates.date2num(f_dt), mdates.date2num(plot_end))
            if e > s:
                if is_m: ax.hlines(y, s, e, colors=color, linestyles='dotted', linewidth=2.5, zorder=3)
                else: ax.hlines(y, s, e, colors=color, linewidth=5, capstyle='butt', zorder=3)
        mid_time_num = mdates.date2num(max(camp['Start'], plot_start) + (min(camp['Finish'], plot_end) - max(camp['Start'], plot_start))/2)
        if is_m:
            ax.text(mid_time_num, y + 0.1, camp['Product'], ha='center', va='bottom', fontsize=8, color='#555555', fontweight='bold')
        else:
            y_offset = line_offset_state[line_name]
            va_pos = 'bottom' if y_offset > 0 else 'top'
            line_offset_state[line_name] = -30 if y_offset > 0 else 30
            ax.annotate(f"{camp['Product']}\n{camp['TotalTon']:.1f}t", xy=(mid_time_num, y), xytext=(0, y_offset), textcoords='offset points',
                        ha='center', va=va_pos, bbox=dict(boxstyle='square,pad=0.3', fc='white', ec=color, lw=1, alpha=0.9),
                        arrowprops=dict(arrowstyle='->', color=color, connectionstyle='arc3'), fontsize=9, fontweight='bold')

    ax.set_xlim(mdates.date2num(plot_start), mdates.date2num(plot_end))
    for i in range(9): ax.axvline(mdates.date2num(plot_start + datetime.timedelta(days=i)), color='red', alpha=0.4, linewidth=2, zorder=5)
    curr_h = plot_start
    while curr_h <= plot_end:
        ax.axvline(mdates.date2num(curr_h), color='#EEEEEE', linewidth=0.7, zorder=1)
        curr_h += datetime.timedelta(hours=1)
    for i in range(len(plot_order) - 1): ax.axhline(i + 0.5, color='#CCCCCC', linewidth=1.2, linestyle='-', alpha=0.8, zorder=2)
    ax.xaxis.set_major_locator(mdates.DayLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('\n%m/%d (%a)'))
    ax.xaxis.set_minor_locator(plt.FixedLocator([mdates.date2num(plot_start + datetime.timedelta(hours=3*i)) for i in range(57)]))
    ax.xaxis.set_minor_formatter(FuncFormatter(lambda x, pos: f"{mdates.num2date(x).hour}"))
    ax.set_yticks(range(len(plot_order)))
    ax.set_yticklabels(plot_order, fontsize=12, fontweight='bold')
    ax.set_ylim(-0.8, len(plot_order) - 0.2)
    plt.title(f"Production Plan - Week of {start_date}", fontsize=16, pad=40)
    plt.tight_layout()
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png')
    plt.close()
    img_buf.seek(0)
    return img_buf

# --- Streamlit UI ---
st.set_page_config(layout="wide", page_title="Production Planner")
st.title("🏭 Production Plan Visualizer & Master Report")

uploaded_file = st.file_uploader("Excelファイルをアップロード (.xlsm)", type=["xlsm"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, sheet_name='Fill', header=None)
    available_weeks = get_available_weeks(df_raw)
    
    if available_weeks:
        col1, col2 = st.columns([1, 2])
        with col1:
            selected_week = st.selectbox("表示したい週の開始日（月曜日）を選択してください", available_weeks)
        with col2:
            st.write(" ")
            generate_btn = st.button("🚀 Generate Master Report", use_container_width=True)

        if generate_btn:
            with st.spinner('処理中...'):
                df_tasks = process_tasks(df_raw)
                img_buf = generate_plot(df_tasks, selected_week)
                
                # --- Excel生成ロジック ---
                wb = openpyxl.Workbook()
                ws_vol = wb.active
                ws_vol.title = "Hourly_Production_Volume"
                
                # ヘッダー作成 (00:00から168時間)
                start_dt = datetime.datetime.combine(selected_week, datetime.time(0, 0))
                hour_list = [start_dt + datetime.timedelta(hours=h) for h in range(168)]
                
                ws_vol.cell(1, 1, "Line").font = Font(bold=True)
                ws_vol.cell(1, 2, "Product").font = Font(bold=True)
                for i, h_dt in enumerate(hour_list):
                    cell = ws_vol.cell(1, 3 + i, h_dt.strftime('%m/%d %H:00'))
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(text_rotation=90, horizontal='center')

                # データ計算 (メンテナンス除外)
                clean_tasks = df_tasks[~df_tasks['is_maint']]
                unique_items = sorted(list(set(zip(clean_tasks['Line'], clean_tasks['Product']))))

                for r_idx, (line, product) in enumerate(unique_items):
                    row_num = r_idx + 2
                    ws_vol.cell(row_num, 1, line)
                    ws_vol.cell(row_num, 2, product)
                    item_tasks = clean_tasks[(clean_tasks['Line'] == line) & (clean_tasks['Product'] == product)]
                    
                    for c_idx, h_dt in enumerate(hour_list):
                        h_end = h_dt + datetime.timedelta(hours=1)
                        ton_sum = 0
                        for _, t in item_tasks.iterrows():
                            overlap = (min(t['Finish'], h_end) - max(t['Start'], h_dt)).total_seconds()
                            if overlap > 0:
                                total_dur = (t['Finish'] - t['Start']).total_seconds()
                                ton_sum += t['Ton'] * (overlap / total_dur)
                        if ton_sum > 0:
                            cell = ws_vol.cell(row_num, 3 + c_idx, round(ton_sum, 2))
                            cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

                # 日付境界線 (0:00)
                thick_side = Side(style='thick', color='000000')
                for c in range(3, ws_vol.max_column + 1):
                    header_val = str(ws_vol.cell(1, c).value)
                    if "00:00" in header_val and not any(x in header_val for x in ["10:00", "20:00"]):
                        for r in range(1, ws_vol.max_row + 1):
                            ws_vol.cell(r, c).border = Border(left=thick_side)

                # ビジュアルスケジュールシート
                ws_vis = wb.create_sheet("Visual_Schedule")
                img_buf.seek(0) # 読み取り位置をリセット
                img = openpyxl.drawing.image.Image(img_buf)
                ws_vis.add_image(img, 'B2')
                
                # Excelをメモリに書き出し
                out_excel = io.BytesIO()
                wb.save(out_excel)
                
                # UI表示
                st.success("✅ 生成完了")
                st.download_button(
                    label="📥 Master Report (Excel) をダウンロード",
                    data=out_excel.getvalue(),
                    file_name=f"Production_Master_{selected_week}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # ガントチャートを表示
                img_buf.seek(0) # Streamlit表示前にもう一度リセット
                st.image(img_buf)
