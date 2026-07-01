import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    /* 상단 빈 공간 제거 */
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

        style_col, division_col = None, None
        print_col, emb_col, fwash_col, gwash_col, gdye_col = None, None, None, None, None
        line_start_col, fabric_in_fac_col, pps_appd_col, ex_factory_col, qty_col = None, None, None, None, None

        for col in df.columns:
            c_clean = clean_string(col)
            if 'STYLE' in c_clean and '배정' not in c_clean: style_col = col
            elif 'DIVISION' in c_clean or 'DIV' in c_clean: division_col = col
            elif 'PRINT' in c_clean: print_col = col
            elif 'EMB' in c_clean or 'SEQUIN' in c_clean: emb_col = col
            elif 'FWASH' in c_clean or 'F/WASH' in c_clean: fwash_col = col
            elif 'GWASH' in c_clean or 'G/WASH' in c_clean: gwash_col = col
            elif 'GDYE' in c_clean or 'G/DYE' in c_clean: gdye_col = col
            elif 'START' in c_clean: line_start_col = col
            elif 'INFAC' in c_clean: fabric_in_fac_col = col
            elif 'PPGTSAPPD' in c_clean or 'PPAPPD' in c_clean: pps_appd_col = col
            elif any(k in c_clean for k in ['1STSD', 'EXFAC', 'EXFACTORY', '1STEX', 'SD', 'S/D', 'FACTORYOUT', 'EXFACTORYDATE']): ex_factory_col = col
            elif any(k in c_clean for k in ['GMTQTY', 'TOTALORDERQTY', '작업수량']) and qty_col is None: qty_col = col

        if style_col is None: continue

        if qty_col:
            df[qty_col] = df[qty_col].astype(str).str.strip()
            df.loc[df[qty_col].isin(['nan', 'NaN', 'None', '<NA>', '', ' ']), qty_col] = None
            df[qty_col] = df[qty_col].ffill()
            
        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', ''] or style_raw.upper().startswith('TOTAL'): continue
                
            line_start_raw = row.get(line_start_col)
            if str(line_start_raw).strip().lower() in ['', 'nan', 'none', 'nat', '<na>']: continue
            try: line_start = pd.to_datetime(line_start_raw)
            except: continue
                
            styles_list = [s.strip() for s in style_raw.replace('/', ',').split(',') if s.strip()]
            
            has_graphic = '🔴 X'
            if (print_col and str(row.get(print_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']) or \
               (emb_col and str(row.get(emb_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']): has_graphic = '🟢 O'
                
            has_wash = '🔴 X'
            if (fwash_col and str(row.get(fwash_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']) or \
               (gwash_col and str(row.get(gwash_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']) or \
               (gdye_col and str(row.get(gdye_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']): has_wash = '🟢 O'
            
            fabric_in_fac = str(row.get(fabric_in_fac_col, 'nan')).strip().lower()
            fabric_status = "🔴 Late" if fabric_in_fac in ['nan', 'none', 'nat', '<na>', ''] else "🟢 Ready"
            pps_status = str(row.get(pps_appd_col, '')) if pps_appd_col else "➖"
            risk = "🟢 Low" if fabric_status == "🟢 Ready" else "🔴 High"
            div_val = str(row.get(division_col, 'N/A')).strip() if division_col else 'N/
