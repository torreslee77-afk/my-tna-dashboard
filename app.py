import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size: 32px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #6B7280; margin-bottom: 25px; }
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 YAKJIN TNA Ai Operational dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">TNA Analysis summary</div>', unsafe_allow_html=True)

def clean_string(val):
    if pd.isna(val): return ""
    return str(val).strip().upper().replace(" ", "").replace("'", "").replace("#", "").replace("/", "").replace("(", "").replace(")", "").replace("-", "")

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
                
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = combined_columns

        style_col, buyer_col = None, None
        print_col, emb_col, fwash_col, gwash_col, gdye_col = None, None, None, None, None
        line_start_col, fabric_in_fac_col, pps_appd_col, ex_factory_col, qty_col = None, None, None, None, None

        for col in df.columns:
            c_clean = clean_string(col)
            if 'STYLE' in c_clean and '배정' not in c_clean: style_col = col
            elif any(k in c_clean for k in ['BUYER', 'DIVISION', '담당']): buyer_col = col
            elif 'PRINT' in c_clean: print_col = col
            elif 'EMB' in c_clean or 'SEQUIN' in c_clean: emb_col = col
            elif 'FWASH' in c_clean or 'F/WASH' in c_clean: fwash_col = col
            elif 'GWASH' in c_clean or 'G/WASH' in c_clean: gwash_col = col
            elif 'GDYE' in c_clean or 'G/DYE' in c_clean: gdye_col = col
            elif 'START' in c_clean: line_start_col = col
            elif 'INFAC' in c_clean: fabric_in_fac_col = col
            elif 'PPGTSAPPD' in c_clean or 'PPAPPD' in c_clean: pps_appd_col = col
            elif any(k in c_clean for k in ['1STSD', 'EXFAC', 'EXFACTORY', '1STEX', 'SD', 'S/D', 'FACTORYOUT', 'EXFACTORYDATE']): ex_factory_col = col
            elif 'GMTQTY' in c_clean: qty_col = col

        if style_col is None: continue

        if qty_col:
            df[qty_col] = df[qty_col].ffill()
            
        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', ''] or style_raw.upper().startswith('TOTAL'): continue
                
            line_start_raw = row.get(line_start_col)
            if pd.isna(line_start_raw) or str(line_start_raw).strip() in ['', 'nan', 'None']: continue
            try: line_start = pd.to_datetime(line_start_raw)
            except: continue
                
            styles_list = [s.strip() for s in style_raw.replace('/', ',').split(',') if s.strip()]
            
            # [수정] 에러의 원인이던 fabric_in_fac 체크를 더 안전하게 변경
            fabric_status = "🔴 Late"
            if fabric_in_fac_col and pd.notna(row.get(fabric_in_fac_col)):
                fabric_status = "🟢 Ready"
            
            # [수정] 수량 계산 로직 (기존 로직 보존)
            qty_val = 0
            if qty_col:
                raw_val = row.get(qty_col)
                try: qty_val = int(float(str(raw_val).replace(',', '').replace('pcs', '').strip()))
                except: qty_val = 0
            
            # 나머지 로직은 기존과 동일
            has_graphic = '🔴 X'
            if print_col and pd.notna(row.get(print_col)) and str(row.get(print_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_graphic = '🟢 O'
            elif emb_col and pd.notna(row.get(emb_col)) and str(row.get(emb_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_graphic = '🟢 O'
                
            has_wash = '🔴 X'
            if fwash_col and pd.notna(row.get(fwash_col)) and str(row.get(fwash_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_wash = '🟢 O'
            elif gwash_col and pd.notna(row.get(gwash_col)) and str(row.get(gwash_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_wash = '🟢 O'
            elif gdye_col and pd.notna(row.get(gdye_col)) and str(row.get(gdye_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_wash = '🟢 O'

            # ... (나머지 로직은 기존과 완전 동일) ...
            for i, single_style in enumerate(styles_list):
                allocated_qty = qty_val // len(styles_list)
                if i == 0: allocated_qty += qty_val % len(styles_list)
                sheet_rows.append({
                    "Style": single_style, "Qty": allocated_qty, "Graphic": has_graphic, "Wash": has_wash,
                    "Line Start": line_start.strftime('%m/%d'), "Fabric Status": fabric_status
                })
        if sheet_rows:
            all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
    return all_sheets_data

# (하단 UI는 기존과 동일하게 유지)
