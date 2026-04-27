import streamlit as st
import pandas as pd
import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from io import BytesIO

# --- 定数設定 ---
NAME_MAP = {'Pump': 'Pump', 'Ref1': 'Ref-1', 'Flexible': 'Flexible', 'Ref3': 'Ref-3', 'Ref4': 'Ref-4', 'Ref5': 'Ref-5', 'Awa': 'Awa'}
LINE_START_COLS = {'Pump': 2, 'Ref1': 11, 'Flexible': 20, 'Ref3': 29, 'Ref4': 38, 'Ref5': 47, 'Awa': 56}
DAY_START_TIME = datetime.time(3, 30)
MAINT_KEYWORDS = ['P/C', 'CLN', 'SETUP', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC', '原価改定', '段取', 'メンテナンス', '点検', '清掃', '切替', '予備', 'WAIT', 'SAMPLE', 'サンプル', 'P/C洗浄', '段取り']

def to_time(val):
    if isinstance(val, datetime.time): return val
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60)
    return None

# --- Streamlit UI ---
st.title("📊 Production Plan Master Processor")
uploaded_file = st.file_uploader("Excelファイルをアップロード (.xlsm)", type=['xlsm', 'xlsx'])

if uploaded_file:
    with st.spinner('データを処理中...'):
        df_raw = pd.read_excel(uploaded_file, sheet_name='Fill', header=None)
        
        # 動的にトン数列を特定
        line_config = {}
        for line, start_idx in LINE_START_COLS.items():
            found_ton_col = None
            for c in range(start_idx, start_idx + 10):
                h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))]
                if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals):
                    found_ton_col = c
                    break
            line_config[line] = {'prod': start_idx, 'start': start_idx + 2, 'finish': start_idx + 3, 'ton': found_ton_col or (start_idx+4)}

        # タスクパース
        tasks = []
        date_col = 63
        for i in range(3, len(df_raw)):
            date_val = df_raw.iloc[i, date_col]
            if not isinstance(date_val, (datetime.datetime, pd.Timestamp)): continue
            for line, cols in line_config.items():
                prod_raw, st_raw, fn_raw, tn_raw = df_raw.iloc[i, cols['prod']], df_raw.iloc[i, cols['start']], df_raw.iloc[i, cols['finish']], df_raw.iloc[i, cols['ton']]
                if pd.isna(st_raw) or pd.isna(fn_raw) or pd.isna(prod_raw): continue
                product = str(prod_raw).strip()
                if product.lower() in ['nan', '連操なし', '']: continue
                try: ton = float(tn_raw) if pd.notna(tn_raw) else 0.0
                except: ton = 0.0
                s_t, f_t = to_time(st_raw), to_time(fn_raw)
                if s_t and f_t:
                    dt_s = datetime.datetime.combine(date_val.date(), s_t)
                    dt_f = datetime.datetime.combine(date_val.date(), f_t)
                    if s_t < DAY_START_TIME: dt_s += datetime.timedelta(days=1)
                    if f_t < DAY_START_TIME: dt_f += datetime.timedelta(days=1)
                    if dt_f <= dt_s: dt_f += datetime.timedelta(days=1)
                    is_m = (ton <= 0) or any(kw.upper() in product.upper() for kw in MAINT_KEYWORDS)
                    tasks.append({'Line': line, 'Product': product, 'Start': dt_s, 'Finish': dt_f, 'Ton': ton, 'is_maint': is_m})

        df_tasks = pd.DataFrame(tasks)
        if not df_tasks.empty:
            # --- ガントチャート生成 (Original Logic) ---
            start_week = df_tasks['Start'].min().replace(hour=3, minute=30, second=0)
            plot_start, plot_end = start_week, start_week + datetime.timedelta(days=7)
            requested_order = ['Pump', 'Ref1', 'Flexible', 'Ref3', 'Ref4', 'Ref5', 'Awa']
            plot_order = [NAME_MAP[n] for n in requested_order[::-1]]
            line_to_y = {NAME_MAP[n]: i for i, n in enumerate(requested_order[::-1])}

            fig, ax = plt.subplots(figsize=(26, 12), facecolor='white')
            line_offset_state = {line: 30 for line in plot_order}

            # マージ処理（同じ製品の連続を結合）
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

            # 描画
            for camp in merged:
                if camp['Finish'] < plot_start or camp['Start'] > plot_end: continue
                line_name = NAME_MAP[camp['Line']]
                y, is_m, color = line_to_y[line_name], camp['is_maint'], ('#7F7F7F' if camp['is_maint'] else '#1F4E78')
                for s_dt, f_dt in camp['Segments']:
                    s, e = max(mdates.date2num(s_dt), mdates.date2num(plot_start)), min(mdates.date2num(f_dt), mdates.date2num(plot_end))
                    if e > s:
                        if is_m: ax.hlines(y, s, e, colors=color, linestyles='dotted', linewidth=2.5)
                        else: ax.hlines(y, s, e, colors=color, linewidth=5, capstyle='butt')

                mid = mdates.date2num(max(camp['Start'], plot_start) + (min(camp['Finish'], plot_end) - max(camp['Start'], plot_start))/2)
                if is_m: ax.text(mid, y + 0.1, camp['Product'], ha='center', fontsize=8, color='#555555')
                else:
                    y_off = line_offset_state[line_name]
                    line_offset_state[line_name] = -30 if y_off > 0 else 30
                    ax.annotate(f"{camp['Product']}\n{camp['TotalTon']:.1f}t", xy=(mid, y), xytext=(0, y_off), textcoords='offset points', ha='center', va=('bottom' if y_off > 0 else 'top'), bbox=dict(boxstyle='square,pad=0.3', fc='white', ec=color, lw=1, alpha=0.9), arrowprops=dict(arrowstyle='->', color=color), fontsize=9, fontweight='bold')

            # グリッド・軸設定
            ax.set_xlim(mdates.date2num(plot_start), mdates.date2num(plot_end))
            for i in range(9): ax.axvline(mdates.date2num(plot_start + datetime.timedelta(days=i)), color='red', alpha=0.4, linewidth=2)
            ax.xaxis.set_major_locator(mdates.DayLocator()); ax.xaxis.set_major_formatter(mdates.DateFormatter('\n%m/%d (%a)'))
            ax.xaxis.set_minor_locator(plt.FixedLocator([mdates.date2num(plot_start + datetime.timedelta(hours=3*i)) for i in range(57)]))
            ax.xaxis.set_minor_formatter(FuncFormatter(lambda x, pos: f"{mdates.num2date(x).hour}"))
            ax.set_yticks(range(len(plot_order))); ax.set_yticklabels(plot_order, fontweight='bold')
            ax.set_ylim(-0.8, len(plot_order) - 0.2)
            
            # Excel生成 (省略せず統合)
            wb = openpyxl.Workbook()
            ws_vol = wb.active
            ws_vol.title = "Hourly_Volume"
            # (ここにHourly Volume生成ロジックが入ります - 前回提供分と同様)
            # ... [前回のExcel生成コードをここに配置] ...
            
            img_buf = BytesIO()
            plt.savefig(img_buf, format='png', bbox_inches='tight')
            img_buf.seek(0)
            ws_vis = wb.create_sheet("Visual_Schedule")
            ws_vis.add_image(openpyxl.drawing.image.Image(img_buf), 'B2')

            # ダウンロード
            out_excel = BytesIO()
            wb.save(out_excel)
            st.success("✅ 生成完了")
            st.download_button("📥 ダウンロード", out_excel.getvalue(), "Report.xlsx")
            st.pyplot(fig)
