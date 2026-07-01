import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# 1. 페이지 기본 설정 및 디자인 (기존 유지)
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

# Helper: 키워드 매칭 함수
def clean_string(val):
    if pd.isna(val): return ""
    return str(val).strip().upper().replace(" ", "").replace("'", "").replace("#", "").replace("/", "").replace("(", "").replace(")", "").replace("-", "")

# 2. 엑셀 분석 엔진 함수 (핵심 구조 유지)
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

        # --- 매칭 타겟 컬럼 인덱스 찾기 ---
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
            # [변경 부분] 정확히 GMT QTY를 매칭하도록 설정
            elif c_clean == 'GMTQTY': qty_col = col

        if style_col is None: continue

        # 수량 컬럼 ffill 적용
        if qty_col:
            df[qty_col] = df[qty_col].replace('nan', None).ffill()
            
        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', ''] or style_raw.upper().startswith('TOTAL'): continue
                
            line_start_raw = row.get(line_start_col) if line_start_col else None
            if pd.isna(line_start_raw) or str(line_start_raw).strip() in ['', 'nan', 'None']: continue
            try: line_start = pd.to_datetime(line_start_raw)
            except: continue
                
            styles_list = [s.strip() for s in style_raw.replace('/', ',').split(',') if s.strip()]
            
            # --- 기존 로직 그대로 유지 ---
            has_graphic = '🔴 X'
            if print_col and pd.notna(row.get(print_col)) and str(row.get(print_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_graphic = '🟢 O'
            elif emb_col and pd.notna(row.get(emb_col)) and str(row.get(emb_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_graphic = '🟢 O'
                
            has_wash = '🔴 X'
            if fwash_col and pd.notna(row.get(fwash_col)) and str(row.get(fwash_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_wash = '🟢 O'
            elif gwash_col and pd.notna(row.get(gwash_col)) and str(row.get(gwash_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_wash = '🟢 O'
            elif gdye_col and pd.notna(row.get(gdye_col)) and str(row.get(gdye_col)).strip() not in ['', 'nan', 'X', 'x', '🔴 X']: has_wash = '🟢 O'
            
            days_buffer = 14 if ('🟢 O' in has_graphic or '🟢 O' in has_wash) else 7
            fabric_due = line_start - timedelta(days=14)
            fpp_due = line_start - timedelta(days=14)
            pps_due = line_start - timedelta(days=days_buffer)
            
            fabric_in_fac = row.get(fabric_in_fac_col)
            fabric_status = "🔴 Late" if pd.isna(fabric_in_fac) else "🟢 Ready"
            
            pp_appd_raw = str(row.get(pps_appd_col, '')).strip() if pps_appd_col else ""
            if pp_appd_raw.upper() in ['N/A']: pps_status = "⚪ N/A"
            elif pp_appd_raw.upper() in ['C/O']: pps_status = "⚪ C/O"
            elif pp_appd_raw == 'nan' or not pp_appd_raw: pps_status = "➖"
            else:
                try: pps_status = pd.to_datetime(pp_appd_raw).strftime('%m/%d')
                except: pps_status = pp_appd_raw
                
            risk = "🟢 Low"
            if fabric_status == "🔴 Late": risk = "🔴 High"
                
            buyer_val = str(row.get(buyer_col, 'YAKJIN')).strip() if buyer_col else 'YAKJIN'
            
            ex_fac_val = '-'
            if ex_factory_col and pd.notna(row.get(ex_factory_col)):
                try: ex_fac_val = pd.to_datetime(row.get(ex_factory_col)).strftime('%m/%d')
                except: ex_fac_val = str(row.get(ex_factory_col))
            
            # 수량 계산 (안전하게 처리)
            qty_val = 0
            if qty_col:
                raw_val = row.get(qty_col)
                try: qty_val = int(float(str(raw_val).replace(',', '')))
                except: qty_val = 0
                
            for i, single_style in enumerate(styles_list):
                allocated_qty = qty_val // len(styles_list)
                if i == 0: allocated_qty += qty_val % len(styles_list)

                sheet_rows.append({
                    "Style": single_style,
                    "Buyer": buyer_val,
                    "Graphic": has_graphic,
                    "Wash": has_wash,
                    "Line Start": line_start.strftime('%m/%d'),
                    "Fabric Due": fabric_due.strftime('%m/%d'),
                    "Fabric Status": fabric_status,
                    "FPP Due": fpp_due.strftime('%m/%d'),
                    "PPS Status": pps_status,
                    "1st Ex-Factory": ex_fac_val,
                    "Qty": allocated_qty,
                    "Risk": risk
                })
        if sheet_rows:
            all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
    return all_sheets_data

# 3. UI (기존 유지)
uploaded_file = st.file_uploader("TNA 엑셀 파일을 여기에 드래그하거나 선택하세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("분석 중..."):
        results = analyze_tna(uploaded_file.read())
        if not results: st.error("데이터 분석 실패")
        else:
            total_styles = sum(len(df) for df in results.values())
            high_risks = sum(len(df[df['Risk'] == "🔴 High"]) for df in results.values())
            total_qty = sum(df['Qty'].sum() for df in results.values())
            
            col1, col2, col3 = st.columns(3)
            col1.markdown(f'<div class="metric-box"><h4>총 스타일 수</h4><h2>{total_styles} 개</h2></div>', unsafe_allow_html=True)
            col2.markdown(f'<div class="metric-box"><h4>🔴 High Risk 스타일</h4><h2 style="color:red;">{high_risks} 개</h2></div>', unsafe_allow_html=True)
            col3.markdown(f'<div class="metric-box"><h4>총 오더 수량 (QTY)</h4><h2>{total_qty:,} pcs</h2></div>', unsafe_allow_html=True)
            
            st.write("---")
            tabs = st.tabs(list(results.keys()))
            for num, sheet_name in enumerate(results.keys()):
                with tabs[num]:
                    st.dataframe(results[sheet_name], use_container_width=True, hide_index=True)
