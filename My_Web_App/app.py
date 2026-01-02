import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import datetime, timedelta
import io

# --- 核心邏輯：資料清洗 ---
def load_data(file):
    try:
        if file.name.endswith('.xlsx'):
            df = pd.read_excel(file, engine='openpyxl')
        else:
            try:
                df = pd.read_csv(file, encoding='utf-8')
            except:
                df = pd.read_csv(file, encoding='cp950')
        
        # 欄位模糊匹配邏輯
        rename_map = {
            '店別': '區域', '班別營業|店名': '店名', '班別營業|班別': '班別',
            '班別營業|日期': '日期', '班別營業|值班者': '值班者',
            '檳榔銷售|金額': '檳榔', '營業金額|實收金額': '實收', '營業金額|結帳差額': '帳差'
        }
        
        actual_rename = {}
        for target_key, new_name in rename_map.items():
            for col in df.columns:
                if target_key in col:
                    actual_rename[col] = new_name
        
        df = df[list(actual_rename.keys())].rename(columns=actual_rename)

        for col in ['檳榔', '實收', '帳差']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        # [修正] 歸類邏輯：店別包含「日紅」或「彰化」的都算彰化
        if '區域' in df.columns:
            df['區域'] = df['區域'].astype(str).str.strip()
            df.loc[df['區域'].str.contains('日紅|彰化'), '區域'] = '彰化'
        
        # [修正] 帳差一律加負號
        if '帳差' in df.columns:
            df['帳差'] = df['帳差'] * -1
        
        report_date = pd.to_datetime(df.iloc[0]['日期']) if not df.empty else datetime.now()
        return df, report_date
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        return None, None

# --- 核心邏輯：讀取昨日累計 ---
def get_cumulative(wb, current_date):
    if current_date.day == 1: return 0, 0, 0
    try:
        names = wb.sheetnames
        prev_name = (current_date - timedelta(days=1)).strftime("%m-%d")
        ws = wb[prev_name] if prev_name in names else wb[names[-1]]
        t, ch, tc = 0, 0, 0
        for row in ws.iter_rows(min_col=12, max_col=15):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    v = ws.cell(row=cell.row, column=16).value or 0
                    if "營業總金額" in cell.value: t = v
                    elif "彰化檳榔金額" in cell.value: ch = v
                    elif "台中檳榔金額" in cell.value: tc = v
        return t, ch, tc
    except: return 0, 0, 0

# --- Streamlit 網頁佈局 ---
st.set_page_config(page_title="直營店日報產生器 V12", layout="wide")
st.title("🍹 直營店日報自動化系統")
st.info("💡 會計小提醒：若是每個月 1 號，請直接上傳原始檔即可，不需上傳累計表。")

f1 = st.file_uploader("1. 上傳當日系統原始檔 (CSV/Excel)", type=['csv', 'xlsx'])
f2 = st.file_uploader("2. 上傳目前的月累計 Excel (選用)", type=['xlsx'])

if st.button("🚀 生成報表"):
    if f1:
        df, report_date = load_data(f1)
        if df is None or df.empty:
            st.error("找不到資料，請檢查檔案格式。")
            st.stop()
            
        wb = load_workbook(f2) if f2 else Workbook()
        if 'Sheet' in wb.sheetnames: del wb['Sheet']
        
        p_t, p_ch, p_tc = get_cumulative(wb, report_date)
        sn = report_date.strftime("%m-%d")
        if sn in wb.sheetnames: del wb[sn]
        ws = wb.create_sheet(sn)

        # --- 樣式設定 ---
        thin = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
        align_c = Alignment('center', 'center', wrap_text=True)
        align_r = Alignment('right', 'center', wrap_text=True)
        align_l_top = Alignment('left', 'top', wrap_text=True)
        font_h = Font('微軟正黑體', 12, bold=True)
        font_n = Font('微軟正黑體', 10)
        font_b = Font('微軟正黑體', 10, bold=True)
        font_red = Font('微軟正黑體', 10, color="FF0000", bold=True)
        font_blue = Font('微軟正黑體', 10, color="0000FF", bold=True)
        font_green = Font('微軟正黑體', 10, color="008000", bold=True)
        font_panel = Font('微軟正黑體', 12, bold=True)
        fill_blue = PatternFill('solid', fgColor="D9E1F2")

        col_ws = {'A':12,'B':6,'C':8,'D':9,'E':9,'F':6,'G':9,'H':6,'I':6,'J':6,'K':2,'L':12,'M':6,'N':8,'O':9,'P':9,'Q':6,'R':9,'S':6,'T':6,'U':6}
        for k, v in col_ws.items(): ws.column_dimensions[k].width = v

        tw_year = report_date.year - 1911
        date_str = f" {tw_year}年{report_date.month}月{report_date.day}日"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
        ws['A1']=f"{date_str} 直營店營收報表 (彰化區)"; ws['A1'].font=font_h; ws['A1'].alignment=align_c
        ws.merge_cells(start_row=1, start_column=12, end_row=1, end_column=21)
        ws['L1']=f"{date_str} 直營店營收報表 (台中區)"; ws['L1'].font=font_h; ws['L1'].alignment=align_c

        headers = ['店名', '班別', '值班者', '檳榔金額', '實收金額', '帳差', '合計', '收款', '實差', '現金合計']
        for i, h in enumerate(headers):
            for sc in [1, 12]:
                c = ws.cell(2, sc+i, h); c.border=thin; c.alignment=align_c; c.fill=fill_blue

        def render(df_s, r, cs):
            if df_s.empty: return r
            rows = len(df_s)
            for i in range(rows):
                curr = r + i; d = df_s.iloc[i]
                ws.cell(curr, cs+1, d['班別']).alignment=align_c
                ws.cell(curr, cs+2, d['值班者']).alignment=align_c
                ws.cell(curr, cs+3, d['檳榔']).number_format='#,##0'
                ws.cell(curr, cs+4, d['實收']).number_format='#,##0'
                dv = d['帳差']
                cd = ws.cell(curr, cs+5, dv); cd.number_format='#,##0'; cd.alignment=align_c
                cd.font = font_red if dv<0 else (font_blue if dv>0 else font_n)
                for x in range(10): ws.cell(curr, cs+x).border=thin
            ws.merge_cells(r, cs, r+rows-1, cs); ws.cell(r, cs, df_s.iloc[0]['店名']).font=font_b; ws.cell(r, cs, df_s.iloc[0]['店名']).alignment=align_c
            ws.merge_cells(r, cs+6, r+rows-1, cs+6); ws.cell(r, cs+6, df_s['實收'].sum()).font=font_b; ws.cell(r, cs+6, df_s['實收'].sum()).alignment=align_c; ws.cell(r, cs+6, df_s['實收'].sum()).number_format='#,##0'
            ws.merge_cells(r, cs+9, r+rows-1, cs+9)
            return r + rows

        rL = rR = 3
        ch_d = df[df['區域']=='彰化']
        for s in list(dict.fromkeys(ch_d['店名'])): rL = render(ch_d[ch_d['店名']==s], rL, 1)
        tc_d = df[df['區域']=='台中']
        for s in list(dict.fromkeys(tc_d['店名'])): rR = render(tc_d[tc_d['店名']==s], rR, 12)

        # 彰化底部
        ws.cell(rL, 4, ch_d['檳榔'].sum()).font=font_green; ws.cell(rL, 7, ch_d['實收'].sum()).font=font_green
        for c in [4, 7]: ws.cell(rL, c).number_format='#,##0'; ws.cell(rL, c).alignment=align_c
        for c in range(1, 11): ws.cell(rL, c).border=thin
        rL += 1
        for lbl in ["班別入帳：", "轉入轉出：", "掉入調出："]:
            ws.merge_cells(rL, 1, rL+1, 10); ws.cell(rL, 1, lbl).alignment=align_l_top; ws.cell(rL, 1, lbl).font=font_n
            for ro in range(2): 
                for ci in range(1, 11): ws.cell(rL+ro, ci).border=thin
            rL += 2

        # 台中底部
        ws.cell(rR, 15, tc_d['檳榔'].sum()).font=font_green; ws.cell(rR, 18, tc_d['實收'].sum()).font=font_green
        for c in [15, 18]: ws.cell(rR, c).number_format='#,##0'; ws.cell(rR, c).alignment=align_c
        for c in range(12, 22): ws.cell(rR, c).border=thin
        rR += 1
        gr, gb, gd = df['實收'].sum(), ch_d['檳榔'].sum()+tc_d['檳榔'].sum(), df['帳差'].sum()
        ws.cell(rR, 15, gb).font=font_b; ws.cell(rR, 16, gr).font=font_b; ws.cell(rR, 18, gr).font=font_b
        cd = ws.cell(r_R, 17, gd); cd.font=font_red if gd<0 else (font_blue if gd>0 else font_b)
        for c in [15, 16, 17, 18]: ws.cell(r_R, c).number_format='#,##0'; ws.cell(r_R, c).alignment=align_c
        for c in range(12, 22): ws.cell(rR, c).border=thin
        rR += 1

        # Panel B
        ms = report_date.replace(day=1)
        dr = f"{ms.month}/{ms.day}-{report_date.month}/{report_date.day}"
        pd_data = [(f"{dr} 營業總金額：", p_t + gr), (f"{dr} 彰化檳榔金額：", p_ch + ch_d['檳榔'].sum()), (f"{dr} 台中檳榔金額：", p_tc + tc_d['檳榔'].sum())]
        curr = rR + 1
        for lbl, val in pd_data:
            ws.merge_cells(curr, 12, curr+1, 15); ws.cell(curr, 12, lbl).alignment=align_r; ws.cell(curr, 12, lbl).font=font_panel
            ws.merge_cells(curr, 16, curr+1, 19); ws.cell(curr, 16, val).number_format='#,##0'; ws.cell(curr, 16, val).font=font_panel; ws.cell(curr, 16, val).alignment=align_c
            for rr in range(curr, curr+2):
                for cc in range(12, 20): ws.cell(rr, cc).border=thin
            curr += 2
        for lbl in ["彰化區未收款：", "台中區未收款：", "", "現金正負差：", "實收總金額："]:
            if lbl:
                ws.merge_cells(curr, 12, curr+1, 15); ws.cell(curr, 12, lbl).alignment=align_r; ws.cell(curr, 12, lbl).font=font_panel
                ws.merge_cells(curr, 16, curr+1, 19)
                for rr in range(curr, curr+2):
                    for cc in range(12, 20): ws.cell(rr, cc).border=thin
                curr += 2
            else: curr += 1

        ws.page_setup.paperSize = 9; ws.page_setup.fitToWidth = 1
        out = io.BytesIO(); wb.save(out)
        st.success(f"✅ {sn} 報表處理完成！")
        st.download_button("💾 下載 Excel 報表", out.getvalue(), f"日報表_{report_date.strftime('%Y%m%d')}.xlsx")
