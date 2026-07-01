import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 3rem; }
    .main-title { font-size: 32px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #6B7280; margin-bottom: 25px; }
    /* 지표 카드 아래 여백을 40px로 확보 */
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; margin-bottom: 40px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 YAKJIN TNA Ai Operational dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">TNA Analysis summary (Sheet-specific)</div>', unsafe_allow_html=True)

def clean_string(val):
    try:
        s_val = str(val).strip().upper()
        if s_val in ['NAN', 'NONE', '<NA>', 'NAT', 'NULL', '']: return ""
        return s_val.replace(" ", "").replace("'", "").replace("#", "").replace("/", "").replace("(", "").replace(")", "").replace("-", "").replace("\n", "").replace("\r", "")
    except: return ""

def analyze_tna(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets_data = {}
    
    for sheet_name in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df_raw.empty: continue
            
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_values = [clean_string(v) for v in row.values]
            if any('STYLE' in v for v in row_values if v):
                header_idx = idx
                break
        if header_idx is None: continue
            
        df = df_raw.iloc[header_idx + 2:].copy()
        
        # 컬럼 매핑 (간소화)
        style_col, div_col, print_col, fwash_col, ex_factory_col, qty_col = None, None, None, None, None, None
        
        # 헤더 컬럼 이름들을 단순화하여 매핑
        header_row = df_raw.iloc[header_idx].astype(str)
        df.columns = [f"Col_{i}" for i in range(len(header_row))]
        
        # 매핑 로직 (정확한 키워드 탐색)
        for i, val in enumerate(header_row):
            c = clean_string(val)
            if 'STYLE' in c: style_col = f"Col_{i}"
            elif 'DIV' in c: div_col = f"Col_{i}"
            elif 'PRINT' in c: print_col = f"Col_{i}"
            elif 'FWASH' in c: fwash_col = f"Col_{i}"
            elif any(k in c for k in ['EXFAC', '1STSD']): ex_factory_col = f"Col_{i}"
            elif any(k in c for k in ['QTY', '작업수량']): qty_col = f"Col_{i}"

        if style_col is None: continue
            
        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', '']: continue
            
            # 수량 처리 (예외처리 포함)
            try:
                raw_qty = str(row.get(qty_col, 0))
                qty_val = int(float(raw_qty.replace(',', '')))
            except:
                qty_val = 0

            sheet_rows.append({
                "Style": style_raw,
                "Division": str(row.get(div_col, 'N/A')),
                "Graphic": '🟢 O' if 'O' in str(row.get(print_col, '')) else '🔴 X',
                "Wash": '🟢 O' if 'O' in str(row.get(fwash_col, '')) else '🔴 X',
                "1st Ex-Factory": str(row.get(ex_factory_col, '-')),
                "Qty": qty_val
            })
                
        if sheet_rows: all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
    return all_sheets_data

uploaded_file = st.file_uploader("TNA 엑셀 파일을 드래그하거나 선택하세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    results = analyze_tna(uploaded_file.read())
    if results:
        tabs = st.tabs(list(results.keys()))
        for num, sheet_name in enumerate(results.keys()):
            with tabs[num]:
                df_sheet = results[sheet_name]
                
                # 수량 콤마 적용
                df_disp = df_sheet.copy()
                df_disp['Qty'] = df_disp['Qty'].apply(lambda x: f"{x:,}")
                
                cols = st.columns(3)
                cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df_sheet):,}</h2></div>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{df_sheet["Qty"].sum():,}</h2></div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="metric-box"><h4>Graphic O</h4><h2>{len(df_sheet[df_sheet["Graphic"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                
                st.dataframe(df_disp, use_container_width=True, hide_index=True)
