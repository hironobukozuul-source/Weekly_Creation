import streamlit as st
import pandas as pd
import datetime
import matplotlib
matplotlib.use('Agg')  # Streamlitサーバー用
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from io import BytesIO

# --- 1. 定数・フィルタ設定 ---
DAY_START_TIME = datetime.time(3, 30)
MAINT_KEYWORDS = [
    'P/C', 'CLN', 'SETUP', '洗浄', 'うがい', 'SPARE', 'C/L', 'QC',
    '原価改定', '段取', 'メンテナンス', '点検', '清掃', '切替', '予備',
    'WAIT', 'SAMPLE', 'サンプル', 'P/C洗浄', '段取り'
]

def to_time(val):
    if isinstance(val, datetime.time): return val
    if isinstance(val, (int, float)):
        ts = int(round(val * 86400))
        return datetime.time((ts // 3600) % 24, (ts // 60) % 60, ts % 60)
    return None

# --- 2. Streamlit UI ---
st.title("📊 Production Plan Master Processor")
st.info("計画表(.xlsm)をアップロードすると、メンテナンス除外済みのExcelとガントチャートを生成します。")

uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsm', 'xlsx'])

if uploaded_file:
    with st.spinner('データを処理中...'):
        # データ読み込み
        df_raw = pd.read_excel(uploaded_file, sheet_name='Fill', header=None)

        line_start_cols = {
            'Pump': 2, 'Ref1': 11, 'Flexible': 20, 'Ref3': 29, 
            'Ref4': 38, 'Ref5': 47, 'Awa': 56
        }

        line_config = {}
        for line, start_idx in line_start_cols.items():
            found_ton_col = None
            for c in range(start_idx, start_idx + 10):
                h_vals = [str(df_raw.iloc[r, c]).lower() for r in range(min(3, len(df_raw)))]
                if any('output' in v for v in h_vals) and any('ton' in v for v in h_vals):
                    found_ton_col = c
                    break
            line_config[line] = {'prod': start_idx, 'start': start_idx + 2, 'finish': start_idx + 3, 'ton': found_ton_col or (start_idx+4)}

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
                    
                    is_maint = (ton <= 0) or any(kw.upper() in product.upper() for kw in MAINT_KEYWORDS)
                    tasks.append({'Line': line, 'Product': product, 'Start': dt_s, 'Finish': dt_f, 'Ton': ton, 'is_maint': is_maint})

        if not tasks:
            st.error("有効なタスクが見つかりませんでした。シート名や列構成を確認してください。")
        else:
            # --- 3. Excel生成 ---
            wb = openpyxl.Workbook()
            ws_vol = wb.active
            ws_vol.title = "Hourly_Production_Volume"

            # 基準日（データの最初の開始日）
            start_dt = min(t['Start'] for t in tasks).replace(hour=0, minute=0, second=0)
            hour_list = [start_dt + datetime.timedelta(hours=h) for h in range(168)]

            # ヘッダー
            thick_black = Side(style='thick', color='000000')
            ws_vol.cell(1, 1, "Line").font = Font(bold=True)
            ws_vol.cell(1, 2, "Product").font = Font(bold=True)
            for i, h_dt in enumerate(hour_list):
                cell = ws_vol.cell(1, 3 + i, h_dt.strftime('%m/%d %H:00'))
                cell.font = Font(bold=True)
                cell.alignment = Alignment(text_rotation=90, horizontal='center')

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

            # 0:00 境界線
            for col_idx in range(3, ws_vol.max_column + 1):
                header = str(ws_vol.cell(1, col_idx).value)
                if header.endswith("00:00") and not (header.endswith("10:00") or header.endswith("20:00")):
                    for r in range(1, ws_vol.max_row + 1):
                        cell = ws_vol.cell(r, col_idx)
                        cell.border = Border(left=thick_black, top=cell.border.top, bottom=cell.border.bottom, right=cell.border.right)

            # --- 4. ビジュアル生成 ---
            name_map = {'Pump': 'Pump', 'Ref1': 'Ref-1', 'Flexible': 'Flexible', 'Ref3': 'Ref-3', 'Ref4': 'Ref-4', 'Ref5': 'Ref-5', 'Awa': 'Awa'}
            requested_order = ['Pump', 'Ref1', 'Flexible', 'Ref3', 'Ref4', 'Ref5', 'Awa']
            plot_order_names = [name_map[n] for n in requested_order[::-1]]
            line_to_y = {name_map[n]: i for i, n in enumerate(requested_order[::-1])}

            fig, ax = plt.subplots(figsize=(20, 10))
            for t in tasks:
                y = line_to_y[name_map[t['Line']]]
                s, e = mdates.date2num(t['Start']), mdates.date2num(t['Finish'])
                color = '#7F7F7F' if t['is_maint'] else '#1F4E78'
                ax.hlines(y, s, e, colors=color, linewidth=8 if not t['is_maint'] else 3)
            
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
            ax.set_yticks(range(len(plot_order_names)))
            ax.set_yticklabels(plot_order_names, fontweight='bold')
            plt.grid(axis='x', color='red', alpha=0.2)
            
            img_buf = BytesIO()
            plt.savefig(img_buf, format='png', bbox_inches='tight')
            img_buf.seek(0)

            ws_vis = wb.create_sheet("Visual_Schedule")
            img = openpyxl.drawing.image.Image(img_buf)
            ws_vis.add_image(img, 'B2')

            # --- 5. ダウンロード ---
            output_excel = BytesIO()
            wb.save(output_excel)
            output_excel.seek(0)

            st.success("✅ 処理完了！")
            st.download_button(
                label="📥 修正済みExcelをダウンロード",
                data=output_excel,
                file_name=f"Production_Report_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.pyplot(fig) # 画面上にもプレビュー表示
