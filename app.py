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

def is_valid_work(val):
    clean_val = str(val).strip().upper()
    invalid_keywords = ['', 'NAN', 'NONE', 'X', 'N/A', 'C/O', 'NAT', '<NA>', 'NULL', '0']
    return clean_val not in invalid_keywords

def analyze_tna(file_bytes):
    # 파일 형식에 따른 읽기 (업로드 오류 방지)
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
    except:
        return {}
        
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
            
        # 헤더 생성 및 중복 처리
        row0 = df_raw.iloc[header_idx].astype(str).replace('nan', '').str.strip()
        row1 = df_raw.iloc[header_idx + 1].astype(str).replace('nan', '').str.strip() if (header_idx + 1) < len(df_raw) else row0
        combined_columns = [f"{p} {s}".strip() if p and s and p!=s else (p if p else s) for p, s in zip(row0, row1)]
        
        seen = {}
        unique_columns = []
        for col in combined_columns:
            if col not in seen: seen[col] = 0; unique_columns.append(col)
            else: seen[col] += 1; unique_columns.append(f"{col}_{seen[col]}")
                
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = unique_columns

        # 컬럼 매핑
        c_map = {clean_string(c): c for c in df.columns}
        style_col = next((c_map[c] for c in c_map if 'STYLE' in c and '배정' not in c), None)
        div_col = next((c_map[c] for c in c_map if 'DIVISION' in c or 'DIV' in c), None)
        print_col = next((c_map[c] for c in c_map if 'PRINT' in c), None)
        emb_col = next((c_map[c] for c in c_map if 'EMB' in c or 'SEQUIN' in c), None)
        fwash_col = next((c_map[c] for c in c_map if 'FWASH' in c), None)
        gwash_col = next((c_map[c] for c in c_map if 'GWASH' in c), None)
        gdye_col = next((c_map[c] for c in c_map if 'GDYE' in c), None)
        line_start_col = next((c_map[c] for c in c_map if 'START' in c), None)
        infac_col = next((c_map[c] for c in c_map if 'INFAC' in c), None)
        qty_col = next((c_map[c] for c in c_map if 'GMTQTY' in c or 'TOTALORDERQTY' in c), None)

        if style_col is None: continue

        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', '']: continue
                
            line_start_raw = row.get(line_start_col)
            try: line_start = pd.to_datetime(line_start_raw)
            except: continue
                
            has_graphic = '🟢 O' if (print_col and is_valid_work(row.get(print_col))) or (emb_col and is_valid_work(row.get(emb_col))) else '🔴 X'
            has_wash = '🟢 O' if (fwash_col and is_valid_work(row.get(fwash_col))) or (gwash_col and is_valid_work(row.get(gwash_col))) or (gdye_col and is_valid_work(row.get(gdye_col))) else '🔴 X'
            
            risk = '🔴 High' if str(row.get(infac_col, '')).strip() == "" else '🟢 Low'
            div_val = str(row.get(div_col, 'N/A'))
            
            qty_val = 0
            try: qty_val = int(float(str(row.get(qty_col, 0)).replace(',', '').replace('pcs', '').split()[0]))
            except: qty_val = 0
                
            sheet_rows.append({
                "Style": style_raw, "Division": div_val, "Graphic": has_graphic, "Wash": has_wash,
                "Line Start": line_start.strftime('%m/%d'), "Risk": risk, "Qty": qty_val
            })
            
        if sheet_rows: all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
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
                    df_disp = results[sheet_name].copy()
                    df_disp['Qty'] = df_disp['Qty'].apply(lambda x: f"{x:,}")
                    st.dataframe(df_disp, use_container_width=True, hide_index=True)
