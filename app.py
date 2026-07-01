import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 2rem; }
    .main-title { font-size: 32px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #6B7280; margin-bottom: 25px; }
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 YAKJIN TNA Ai Operational dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">TNA Analysis summary</div>', unsafe_allow_html=True)

def clean_string(val):
    try:
        s_val = str(val).strip().upper()
        if s_val in ['NAN', 'NONE', '<NA>', 'NAT', 'NULL', '']: return ""
        return s_val.replace(" ", "").replace("'", "").replace("#", "").replace("/", "").replace("(", "").replace(")", "").replace("-", "").replace("\n", "").replace("\r", "")
    except: return ""

def analyze_tna(file_bytes):
    # .csv 파일인지 .xlsx 파일인지 구분하여 읽기
    try:
        if file_bytes.name.endswith('.csv'):
            df_raw = pd.read_csv(file_bytes, header=None)
        else:
            df_raw = pd.read_excel(file_bytes, header=None)
    except:
        return {}

    all_sheets_data = {}
    
    # 헤더 인덱스 찾기
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_values = [clean_string(v) for v in row.values]
        if any('STYLE' in v for v in row_values if v):
            header_idx = idx
            break
    
    if header_idx is None: return {}
        
    row0 = df_raw.iloc[header_idx].astype(str).replace('nan', '').str.strip()
    row1 = df_raw.iloc[header_idx + 1].astype(str).replace('nan', '').str.strip() if (header_idx + 1) < len(df_raw) else row0
    
    combined_columns = []
    current_parent = ""
    for p, s in zip(row0, row1):
        if p != "": current_parent = p
        if current_parent and s and p != s: combined_columns.append(f"{current_parent} {s}")
        elif s: combined_columns.append(s)
        elif current_parent: combined_columns.append(current_parent)
        else: combined_columns.append("Unnamed")
            
    seen = {}
    unique_columns = []
    for col in combined_columns:
        if col not in seen:
            seen[col] = 0
            unique_columns.append(col)
        else:
            seen[col] += 1
            unique_columns.append(f"{col}_{seen[col]}")
            
    df = df_raw.iloc[header_idx + 2:].copy()
    df.columns = unique_columns

    # 컬럼 매핑 (매우 유연하게 검색)
    col_map = {col: clean_string(col) for col in df.columns}
    
    style_col = next((c for c, v in col_map.items() if 'STYLE' in v and '배정' not in v), None)
    division_col = next((c for c, v in col_map.items() if 'DIVISION' in v or 'DIV' in v), None)
    print_col = next((c for c, v in col_map.items() if 'PRINT' in v), None)
    emb_col = next((c for c, v in col_map.items() if 'EMB' in v or 'SEQUIN' in v), None)
    fwash_col = next((c for c, v in col_map.items() if 'FWASH' in v), None)
    gwash_col = next((c for c, v in col_map.items() if 'GWASH' in v), None)
    gdye_col = next((c for c, v in col_map.items() if 'GDYE' in v), None)
    line_start_col = next((c for c, v in col_map.items() if 'START' in v), None)
    fabric_in_fac_col = next((c for c, v in col_map.items() if 'INFAC' in v), None)
    pps_appd_col = next((c for c, v in col_map.items() if 'PPGTSAPPD' in v or 'PPAPPD' in v), None)
    qty_col = next((c for c, v in col_map.items() if 'GMTQTY' in v or 'TOTALORDERQTY' in v or '작업수량' in v), None)

    if style_col is None: return {}

    sheet_rows = []
    for _, row in df.iterrows():
        style_raw = str(row.get(style_col, '')).strip()
        if not style_raw or style_raw.lower() in ['nan', 'none', '']: continue
            
        line_start_raw = row.get(line_start_col)
        try: line_start = pd.to_datetime(line_start_raw)
        except: continue
            
        styles_list = [s.strip() for s in style_raw.replace('/', ',').split(',') if s.strip()]
        
        has_graphic = '🟢 O' if (print_col and str(row.get(print_col, '')).strip().upper() not in ['X', '']) or \
                                (emb_col and str(row.get(emb_col, '')).strip().upper() not in ['X', '']) else '🔴 X'
        has_wash = '🟢 O' if (fwash_col and str(row.get(fwash_col, '')).strip().upper() not in ['X', '']) or \
                             (gwash_col and str(row.get(gwash_col, '')).strip().upper() not in ['X', '']) else '🔴 X'
        
        fabric_status = "🔴 Late" if str(row.get(fabric_in_fac_col, '')).strip() == "" else "🟢 Ready"
        risk = "🔴 High" if fabric_status == "🔴 Late" else "🟢 Low"
        div_val = str(row.get(division_col, 'N/A')) if division_col else 'N/A'
        
        qty_val = 0
        if qty_col:
            try: qty_val = int(float(str(row.get(qty_col)).replace(',', '').replace('pcs', '').split()[0]))
            except: qty_val = 0
            
        for s_name in styles_list:
            sheet_rows.append({
                "Style": s_name, "Division": div_val, "Graphic": has_graphic, "Wash": has_wash,
                "Line Start": line_start.strftime('%m/%d'), "Fabric Status": fabric_status,
                "Risk": risk, "Qty": qty_val // len(styles_list)
            })
            
    return {"Sheet1": pd.DataFrame(sheet_rows)}

uploaded_file = st.file_uploader("파일 업로드", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    results = analyze_tna(uploaded_file)
    if results:
        df = results["Sheet1"]
        total_qty = df['Qty'].sum()
        
        cols = st.columns(5)
        cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
        cols[1].markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{len(df[df["Risk"]=="🔴 High"])}</h2></div>', unsafe_allow_html=True)
        cols[2].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{total_qty:,}</h2></div>', unsafe_allow_html=True)
        cols[3].markdown(f'<div class="metric-box"><h4>Graphic</h4><h2>{len(df[df["Graphic"]=="🟢 O"])}</h2></div>', unsafe_allow_html=True)
        cols[4].markdown(f'<div class="metric-box"><h4>Wash</h4><h2>{len(df[df["Wash"]=="🟢 O"])}</h2></div>', unsafe_allow_html=True)
        
        st.write("---")
        df['Qty'] = df['Qty'].apply(lambda x: f"{x:,}")
        st.dataframe(df, use_container_width=True)
