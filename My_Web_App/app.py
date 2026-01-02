import streamlit as st
import pandas as pd
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import datetime, timedelta
import io

# --- 核心邏輯函數 ---
def load_and_clean_data(uploaded_file):
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        else:
            df = pd.read_csv(uploaded_file)
        
        rename_map = {
            '店別': '區域', '班別營業|店名': '店名', '班別營業|班別': '班別',
            '班別營業|日期': '日期', '班別營業|值班者': '值班者',
            '檳榔銷售|金額': '檳榔', '營業金額|實收金額': '實收', '營業金額|結帳差額': '帳差'
        }
        df = df[[c for c in rename_map.keys() if c in df.columns]].rename(columns=rename_map)
        for col in ['檳榔', '實收', '帳差']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        if '區域' in df.columns:
            df.loc[df['區域'] == '日紅', '區域'] = '彰化'
        df['帳差'] = df['帳差'] * -1
        
        report_date = pd.to_datetime(df.iloc[0]['日期']) if not df.empty else datetime.now()
        return df, report_date
    except Exception as e:
        st.error(f"讀取原始檔失敗: {e}")
        return None, None

def get_cumulative_from_wb(wb, current_date):
    if current_date.day == 1: return 0, 0, 0
    try:
        sheet_names = wb.sheetnames
        target_name = (current_date - timedelta(days=1)).strftime("%m-%d")
        ws = wb[target_name] if target_name in sheet_names else wb[sheet_names[-1]]
        
        p_t, p_ch, p_tc = 0, 0, 0
        for row in ws.iter_rows(min_col=12, max_col=15):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    val = ws.cell(row=cell.row, column=16).value or 0
                    if "營業總金額" in cell.value: p_t = val
                    elif "彰化檳榔金額" in cell.value: p_ch = val
                    elif "台中檳榔金額" in cell.value: p_tc = val
        return p_t, p_ch, p_tc
    except: return 0, 0, 0

# --- Streamlit 介面 ---
st.title("🍹 直營店營收報表自動化系統")
st.write("請依序上傳檔案，系統將自動生成今日報表並計算累計。")

file_1 = st.file_uploader("1. 上傳【當日系統原始檔】(CSV 或 Excel)", type=['csv', 'xlsx'])
file_2 = st.file_uploader("2. 上傳【目前的月累計表】(選用，若 1 號新表請忽略)", type=['xlsx'])

if st.button("🚀 開始生成報表"):
    if file_1:
        df, report_date = load_and_clean_data(file_1)
        
        # 建立或讀取 Workbook
        if file_2:
            wb = load_workbook(file_2)
        else:
            wb = Workbook()
            if 'Sheet' in wb.sheetnames: del wb['Sheet']
        
        prev_t, prev_ch, prev_tc = get_cumulative_from_wb(wb, report_date)
        
        sheet_name = report_date.strftime("%m-%d")
        if sheet_name in wb.sheetnames: del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)
        
        # --- 樣式與繪製 (同 V11 邏輯) ---
        # (此處省略繪製表格的 100 行代碼，實際上傳時需包含完整 render_store 等邏輯)
        # ... [將 V11 的寫入與樣式代碼放入此處] ...
        
        # 最後將結果轉為 Bytes 流供下載
        output = io.BytesIO()
        wb.save(output)
        st.success(f"✅ {sheet_name} 報表處理完成！")
        st.download_button(
            label="💾 點我下載生成的 Excel 檔案",
            data=output.getvalue(),
            file_name=f"直營店營收報表_{report_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("請至少上傳當日原始檔。")