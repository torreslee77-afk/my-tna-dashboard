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
                
        # 중복 컬럼 처리
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

        # 주요 컬럼 탐색
        style_col, division_col = None, None
        print_col, emb_col, fwash_col, gwash_col, gdye_col = None, None, None, None, None
        line_start_col, fabric_in_fac_col, pps_appd_col, ex_factory_col, qty_col = None, None, None, None, None

        for col in df.columns:
            c_clean = clean_string(col)
            if 'STYLE' in c_clean and '배정' not in c_clean: style_col = col
            # [수정] 좀 더 유연하게 탐색
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
            
            # 그래픽/워시 로직
            has_graphic = '🔴 X'
            if (print_col and str(row.get(print_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']) or \
               (emb_col and str(row.get(emb_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']): has_graphic = '🟢 O'
                
            has_wash = '🔴 X'
            if (fwash_col and str(row.get(fwash_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']) or \
               (gwash_col and str(row.get(gwash_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']) or \
               (gdye_col and str(row.get(gdye_col, '')).strip().lower() not in ['', 'nan', 'none', 'x', '🔴 x']): has_wash = '🟢 O'
            
            # 상태값 처리
            fabric_in_fac = str(row.get(fabric_in_fac_col, 'nan')).strip().lower()
            fabric_status = "🔴 Late" if fabric_in_fac in ['nan', 'none', 'nat', '<na>', ''] else "🟢 Ready"
            pps_status = str(row.get(pps_appd_col, '')) if pps_appd_col else "➖"
            risk = "🟢 Low" if fabric_status == "🟢 Ready" else "🔴 High"
            
            # DIVISION 값 추출 (에러 방지)
            div_val = str(row.get(division_col, 'N/A')).strip() if division_col else 'N/A'
            
            qty_val = 0
            if qty_col:
                raw_val = row.get(qty_col)
                try: qty_val = int(float(str(raw_val).replace(',', '').replace('pcs', '').split()[0]))
                except: qty_val = 0
                
            for i, single_style in enumerate(styles_list):
                allocated_qty = qty_val // len(styles_list)
                if i == 0: allocated_qty += qty_val % len(styles_list)

                sheet_rows.append({
                    "Style": single_style, "Division": div_val, "Graphic": has_graphic, "Wash": has_wash,
                    "Line Start": line_start.strftime('%m/%d'), "Fabric Status": fabric_status,
                    "PPS Status": pps_status, "Qty": allocated_qty, "Risk": risk
                })
        if sheet_rows:
            all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
    return all_sheets_data

uploaded_file = st.file_uploader("TNA 엑셀 파일을 여기에 드래그하거나 선택하세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("분석 중..."):
        results = analyze_tna(uploaded_file.read())
        if not results: st.error("데이터 분석 실패")
        else:
            total_styles = sum(len(df) for df in results.values())
            high_risks = sum(len(df[df['Risk'] == "🔴 High"]) for df in results.values())
            total_qty = sum(df['Qty'].sum() for df in results.values())
            graphic_cnt = sum(len(df[df['Graphic'] == "🟢 O"]) for df in results.values())
            wash_cnt = sum(len(df[df['Wash'] == "🟢 O"]) for df in results.values())
            
            cols = st.columns(5)
            cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{total_styles}</h2></div>', unsafe_allow_html=True)
            cols[1].markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{high_risks}</h2></div>', unsafe_allow_html=True)
            cols[2].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{total_qty:,}</h2></div>', unsafe_allow_html=True)
            cols[3].markdown(f'<div class="metric-box"><h4>Graphic styles</h4><h2>{graphic_cnt}</h2></div>', unsafe_allow_html=True)
            cols[4].markdown(f'<div class="metric-box"><h4>Wash styles</h4><h2>{wash_cnt}</h2></div>', unsafe_allow_html=True)
            
            st.write("---")
            tabs = st.tabs(list(results.keys()))
            for num, sheet_name in enumerate(results.keys()):
                with tabs[num]:
                    st.dataframe(results[sheet_name], use_container_width=True, hide_index=True)
