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
            f.write(r.content)
    font_prop = font_manager.FontProperties(fname=font_path)
    font_manager.fontManager.addfont(font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
    return font_prop

jp_font = setup_japanese_font()

# --- 定数設定 ---
NAME_MAP = {'Pump': 'Pump', 'Ref1': 'Ref-1', 'Flexible': 'Flexible', 'Ref3': 'Ref-3', 'Ref4': 'Ref-4', 'Ref5': 'Ref-5', 'Awa': 'Awa'}
LINE_START_COLS = {'Pump': 2, 'Ref1': 11, 'Flexible': 20, 'Ref3': 29, 'Ref4': 38, 'Ref5': 47, 'Awa': 56}
MON_START = datetime.time(3, 30) 
DATE_COL = 63
MAINT_KEYWORDS = ['P/C', 'CLN', 'SETUP', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC', '原価改定', '段取', 'メンテナンス', '点検', '清掃', '切替', '予備', 'WAIT', 'SAMPLE']

def to_time(val):
    if isinstance(val, datetime.time): return val
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60)
    return None

@st.cache_data
def get_available_weeks(df_raw):
    dates = pd.to_datetime(df_raw.iloc[3:, DATE_COL], errors='coerce').dropna()
    mondays = dates[dates.dt.weekday == 0].dt.date.unique()
    return sorted(mondays)

def process_tasks(df_raw):
    tasks = []
    line_config = {}
    for line, start_idx in LINE_START_COLS.items():
        found_ton_col = None
        for c in range(start_idx, start_idx + 10):
            h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))]
            if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals):
                found_ton_col = c
                break
        line_config[line] = {'prod': start_idx, 'start': start_idx+2, 'finish': start_idx+3, 'ton': found_ton_col or (start_idx+4)}

    # 【ステートフル・ロジック】各ラインの現在の日付と直前の終了時刻を追跡
    # キー: ライン名, 値: {'current_date': date, 'last_finish_time': time}
    line_states = {}

    for i in range(3, len(df_raw)):
        date_val = df_raw.iloc[i, DATE_COL]
        if not isinstance(date_val, (datetime.datetime, pd.Timestamp)): continue
        
        base_date = date_val.date()

        for line, cols in line_config.items():
            prod_raw, st_raw, fn_raw, tn_raw = df_raw.iloc[i, cols['prod']], df_raw.iloc[i, cols['start']], df_raw.iloc[i, cols['finish']], df_raw.iloc[i, cols['ton']]
            if pd.isna(st_raw) or pd.isna(fn_raw) or pd.isna(prod_raw): continue
            
            product = str(prod_raw).strip()
            if product.lower() in ['nan', '連操なし', '']: continue
            
            s_t, f_t = to_time(st_raw), to_time(fn_raw)
            if s_t and f_t:
                # ラインごとの状態を初期化（行が新しくなっても日付列が変わらなければ維持）
                if line not in line_states:
                    line_states[line] = {'current_date': base_date, 'last_finish_time': datetime.time(0, 0)}
                
                # 時刻が逆転した、または日付列が明示的に進んだ場合に日付を繰り越し
                # ※同じ日の行データ内でも、前のセグメントより開始時間が早ければ翌日とみなす
                if s_t < line_states[line]['last_finish_time']:
                    line_states[line]['current_date'] += datetime.timedelta(days=1)
                
                # Excelの日付列が手動で進められている場合は、それを優先して同期
                if base_date > line_states[line]['current_date']:
                    line_states[line]['current_date'] = base_date

                dt_s = datetime.datetime.combine(line_states[line]['current_date'], s_t)
                dt_f = datetime.datetime.combine(line_states[line]['current_date'], f_t)
                
                # 終了時刻が開始時刻より前の場合は日付を跨いでいる
                if f_t <= s_t:
                    dt_f += datetime.timedelta(days=1)
                
                try: ton = float(tn_raw) if pd.notna(tn_raw) else 0.0
                except: ton = 0.0
                
                is_m = (ton <= 0) or any(kw.upper() in product.upper() for kw in MAINT_KEYWORDS)
                
                tasks.append({'Line': line, 'Product': product, 'Start': dt_s, 'Finish': dt_f, 'Ton': ton, 'is_maint': is_m})
                
                # 状態を更新（直前の終了時刻を記録。翌日跨ぎの場合はその時刻をセット）
                line_states[line]['last_finish_time'] = f_t

    return pd.DataFrame(tasks)

def generate_plot(df_tasks, start_date):
    # グラフ表示は 00:00 から
    plot_start = datetime.datetime.combine(start_date, datetime.time(0, 0))
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
        curr['Segments'], curr['TotalTon'] = [(curr['Start'], curr['Finish'])], curr['Ton']
        for idx in range(1, len(line_df)):
            nxt = line_df.iloc[idx].to_dict()
            if nxt['Product'] == curr['Product'] and nxt['is_maint'] == curr['is_maint'] and (nxt['Start'] - curr['Finish']) <= datetime.timedelta(hours=4):
                curr['Finish'], curr['TotalTon'] = nxt['Finish'], curr['TotalTon'] + nxt['Ton']
                curr['Segments'].append((nxt['Start'], nxt['Finish']))
            else:
                merged.append(curr); curr = nxt; curr['Segments'], curr['TotalTon'] = [(curr['Start'], curr['Finish'])], curr['Ton']
        merged.append(curr)

    for camp in merged:
        if camp['Finish'] < plot_start or camp['Start'] > plot_end: continue
        line_name = NAME_MAP[camp['Line']]
        y, is_m, color = line_to_y[line_name], camp['is_maint'], ('#7F7F7F' if camp['is_maint'] else '#1F4E78')
        for s_dt, f_dt in camp['Segments']:
            s, e = max(mdates.date2num(s_dt), mdates.date2num(plot_start)), min(mdates.date2num(f_dt), mdates.date2num(plot_end))
            if e > s:
                if is_m: ax.hlines(y, s, e, colors=color, linestyles='dotted', linewidth=2.5, zorder=3)
                else: ax.hlines(y, s, e, colors=color, linewidth=5, capstyle='butt', zorder=3)
        mid = mdates.date2num(max(camp['Start'], plot_start) + (min(camp['Finish'], plot_end) - max(camp['Start'], plot_start))/2)
        if is_m: ax.text(mid, y + 0.1, camp['Product'], ha='center', va='bottom', fontsize=9, color='#555555', fontweight='bold', fontproperties=jp_font)
        else:
            y_off = line_offset_state[line_name]
            line_offset_state[line_name] = -30 if y_off > 0 else 30
            ax.annotate(f"{camp['Product']}\n{camp['TotalTon']:.1f}t", xy=(mid, y), xytext=(0, y_off), textcoords='offset points', ha='center', va=('bottom' if y_off > 0 else 'top'), bbox=dict(boxstyle='square,pad=0.3', fc='white', ec=color, lw=1, alpha=0.9), arrowprops=dict(arrowstyle='->', color=color, connectionstyle='arc3'), fontsize=10, fontweight='bold', fontproperties=jp_font)

    ax.set_xlim(mdates.date2num(plot_start), mdates.date2num(plot_end))
    for i in range(9): ax.axvline(mdates.date2num(plot_start + datetime.timedelta(days=i)), color='red', alpha=0.4, linewidth=2, zorder=5)
    curr_h = plot_start
    while curr_h <= plot_end:
        ax.axvline(mdates.date2num(curr_h), color='#EEEEEE', linewidth=0.7, zorder=1)
        curr_h += datetime.timedelta(hours=1)
    for i in range(len(plot_order) - 1): ax.axhline(i + 0.5, color='#CCCCCC', linewidth=1.2, alpha=0.8, zorder=2)
    ax.xaxis.set_major_locator(mdates.DayLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter('\n%m/%d (%a)'))
    ax.xaxis.set_minor_locator(plt.FixedLocator([mdates.date2num(plot_start + datetime.timedelta(hours=3*i)) for i in range(57)]))
    ax.xaxis.set_minor_formatter(FuncFormatter(lambda x, pos: f"{mdates.num2date(x).hour}"))
    ax.set_yticks(range(len(plot_order))); ax.set_yticklabels(plot_order, fontsize=12, fontweight='bold')
    ax.set_ylim(-0.8, len(plot_order) - 0.2)
    plt.title(f"Production Plan - Week of {start_date}", fontsize=16, pad=40, fontproperties=jp_font)
    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format='png')
    plt.close(); buf.seek(0)
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
            generate_btn = st.button("🚀 Generate Plan & Excel", use_container_width=True)

        if generate_btn:
            with st.spinner('処理中...'):
                df_tasks = process_tasks(df_raw)
                img_buf, fig = generate_plot(df_tasks, selected_week)
                
                # --- Excel生成 (Hourly Volume: 3:30から168時間) ---
                wb = openpyxl.Workbook(); ws_vol = wb.active; ws_vol.title = "Hourly_Volume"
                start_dt = datetime.datetime.combine(selected_week, MON_START)
                hour_list = [start_dt + datetime.timedelta(hours=h) for h in range(168)]
                
                thick_black = Side(style='thick', color='000000')
                ws_vol.cell(1, 1, "Line").font = Font(bold=True)
                ws_vol.cell(1, 2, "Product").font = Font(bold=True)
                for i, h_dt in enumerate(hour_list):
                    cell = ws_vol.cell(1, 3 + i, h_dt.strftime('%m/%d %H:%M'))
                    cell.font = Font(bold=True); cell.alignment = Alignment(text_rotation=90, horizontal='center')

                clean_tasks = df_tasks[~df_tasks['is_maint']]
                unique_items = sorted(list(set(zip(clean_tasks['Line'], clean_tasks['Product']))))

                for r_idx, (line, product) in enumerate(unique_items):
                    row_num = r_idx + 2
                    ws_vol.cell(row_num, 1, line); ws_vol.cell(row_num, 2, product)
                    for c_idx, h_dt in enumerate(hour_list):
                        h_end = h_dt + datetime.timedelta(hours=1)
                        overlap_tasks = clean_tasks[(clean_tasks['Line'] == line) & (clean_tasks['Product'] == product)]
                        ton_sum = sum(t['Ton'] * ((min(t['Finish'], h_end) - max(t['Start'], h_dt)).total_seconds() / (t['Finish'] - t['Start']).total_seconds()) for _, t in overlap_tasks.iterrows() if (min(t['Finish'], h_end) - max(t['Start'], h_dt)).total_seconds() > 0)
                        if ton_sum > 0:
                            cell = ws_vol.cell(row_num, 3 + c_idx, round(ton_sum, 2))
                            cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

                for col_idx in range(3, ws_vol.max_column + 1):
                    header = str(ws_vol.cell(1, col_idx).value)
                    if "03:30" in header:
                        for r in range(1, ws_vol.max_row + 1):
                            ws_vol.cell(r, col_idx).border = Border(left=thick_black)

                ws_vis = wb.create_sheet("Visual_Schedule")
                img_copy = BytesIO(img_buf.getvalue())
                ws_vis.add_image(openpyxl.drawing.image.Image(img_copy), 'B2')
                out_excel = BytesIO(); wb.save(out_excel)
                
                st.success("✅ 生成完了")
                st.download_button("📥 ダウンロード (Excel)", out_excel.getvalue(), f"Report_{selected_week}.xlsx")
                st.pyplot(fig)
