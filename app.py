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
                
        final_columns = []
        counts = {}
        for col in combined_columns:
            if col not in counts:
                counts[col] = 1
                final_columns.append(col)
            else:
                counts[col] += 1
                final_columns.append(f"{col}_{counts[col]}")
        
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = final_columns

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
            elif 'TOTALORDERQTY' in c_clean and qty_col is None: qty_col = col

        if style_col is None: continue

        if qty_col:
            df[qty_col] = df[qty_col].replace('nan', None).ffill()
            
        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', ''] or style_raw.upper().startswith('TOTAL'): continue
            
            line_start_raw = row.get(line_start_col)
            if pd.isna(line_start_raw) or str(line_start_raw).strip() in ['', 'nan', 'None']: continue
            try: line_start = pd.to_datetime(line_start_raw)
            except: continue
            
            # [수정된 부분] 수량 계산 로직 (숫자 변환 강화)
            qty_val = 0
            if qty_col:
                raw_val = row.get(qty_col)
                try: 
                    qty_val = int(float(str(raw_val).replace(',', '').strip()))
                except: 
                    qty_val = 0
            
            styles_list = [s.strip() for s in style_raw.replace('/', ',').split(',') if s.strip()]
            
            # (기존 데이터 처리 로직은 동일하게 유지)
            has_graphic = '🟢 O' if (print_col and pd.notna(row.get(print_col)) and str(row.get(print_col)).strip() not in ['', 'nan', 'X', 'x']) else '🔴 X'
            has_wash = '🟢 O' if (fwash_col and pd.notna(row.get(fwash_col)) and str(row.get(fwash_col)).strip() not in ['', 'nan', 'X', 'x']) else '🔴 X'
            
            for i, single_style in enumerate(styles_list):
                allocated_qty = qty_val // len(styles_list) if len(styles_list) > 0 else 0
                if i == 0 and len(styles_list) > 0: allocated_qty += qty_val % len(styles_list)

                sheet_rows.append({
                    "Style": single_style,
                    "Buyer": str(row.get(buyer_col, 'YAKJIN')),
                    "Line Start": line_start.strftime('%m/%d'),
                    "Qty": allocated_qty,
                    "Graphic": has_graphic,
                    "Wash": has_wash
                })
        if sheet_rows:
            all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
    return all_sheets_data

uploaded_file = st.file_uploader("TNA 엑셀 파일을 업로드하세요.", type=["xlsx", "xls"])
if uploaded_file is not None:
    results = analyze_tna(uploaded_file.read())
    if results:
        total_qty = sum(df['Qty'].sum() for df in results.values())
        st.metric("총 오더 수량 (QTY)", f"{total_qty:,} pcs")
        tabs = st.tabs(list(results.keys()))
        for num, sheet_name in enumerate(results.keys()):
            with tabs[num]:
                st.dataframe(results[sheet_name], use_container_width=True)
