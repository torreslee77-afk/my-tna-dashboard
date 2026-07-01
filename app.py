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

# 💡 [수정 1] pd.isna()로 인한 Series/배열 에러를 원천 차단하기 위해 문자열 강제 변환 후 검사
def clean_string(val):
    try:
        s_val = str(val).strip().upper()
        if s_val in ['NAN', 'NONE', '<NA>', 'NAT', 'NULL', '']: 
            return ""
        return s_val.replace(" ", "").replace("'", "").replace("#", "").replace("/", "").replace("(", "").replace(")", "").replace("-", "").replace("\n", "").replace("\r", "")
    except:
        return ""

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
                
        # 💡 [수정 2] 컬럼명 중복으로 인한 에러 방지 로직 (컬럼명이 같으면 _1, _2 등을 부여)
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
            elif any(k in c_clean for k in ['GMTQTY', 'TOTALORDERQTY', '작업수량']) and qty_col is None: qty_col = col

        if style_col is None: continue

        # 💡 [수정 3] 수량 빈칸(ffill) 채우기를 가장 에러가 없는 인덱싱 방식으로 변경
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
            print_val = str(row.get(print_col, 'nan')).strip().lower()
            emb_val = str(row.get(emb_col, 'nan')).strip().lower()
            if print_col and print_val not in ['', 'nan', 'none', 'x', '🔴 x']: has_graphic = '🟢 O'
            elif emb_col and emb_val not in ['', 'nan', 'none', 'x', '🔴 x']: has_graphic = '🟢 O'
                
            has_wash = '🔴 X'
            fwash_val = str(row.get(fwash_col, 'nan')).strip().lower()
            gwash_val = str(row.get(gwash_col, 'nan')).strip().lower()
            gdye_val = str(row.get(gdye_col, 'nan')).strip().lower()
            if fwash_col and fwash_val not in ['', 'nan', 'none', 'x', '🔴 x']: has_wash = '🟢 O'
            elif gwash_col and gwash_val not in ['', 'nan', 'none', 'x', '🔴 x']: has_wash = '🟢 O'
            elif gdye_col and gdye_val not in ['', 'nan', 'none', 'x', '🔴 x']: has_wash = '🟢 O'
            
            days_buffer = 14 if ('🟢 O' in has_graphic or '🟢 O' in has_wash) else 7
            fabric_due = line_start - timedelta(days=14)
            fpp_due = line_start - timedelta(days=14)
            pps_due = line_start - timedelta(days=days_buffer)
            
            fabric_in_fac = str(row.get(fabric_in_fac_col, 'nan')).strip().lower()
            fabric_status = "🔴 Late" if fabric_in_fac in ['nan', 'none', 'nat', '<na>', ''] else "🟢 Ready"
            
            pp_appd_raw = str(row.get(pps_appd_col, '')).strip() if pps_appd_col else ""
            if pp_appd_raw.upper() in ['N/A']: pps_status = "⚪ N/A"
            elif pp_appd_raw.upper() in ['C/O']: pps_status = "⚪ C/O"
            elif pp_appd_raw.lower() in ['nan', 'none', '']: pps_status = "➖"
            else:
                try: pps_status = pd.to_datetime(pp_appd_raw).strftime('%m/%d')
                except: pps_status = pp_appd_raw
                
            risk = "🟢 Low"
            if fabric_status == "🔴 Late": risk = "🔴 High"
                
            buyer_val = str(row.get(buyer_col, 'YAKJIN')).strip() if buyer_col else 'YAKJIN'
            if buyer_val.lower() in ['nan', 'none', '']: buyer_val = 'YAKJIN'
            
            ex_fac_val = '-'
            if ex_factory_col:
                ex_raw = row.get(ex_factory_col)
                if str(ex_raw).strip().lower() not in ['nan', 'none', 'nat', '<na>', '']:
                    try: ex_fac_val = pd.to_datetime(ex_raw).strftime('%m/%d')
                    except: ex_fac_val = str(ex_raw)
            
            qty_val = 0
            if qty_col:
                raw_val = row.get(qty_col)
                clean_qty_str = str(raw_val).replace(',', '').replace('pcs', '').replace('PCS', '').strip()
                if clean_qty_str.lower() not in ['nan', 'none', '<na>', '']:
                    try: qty_val = int(float(clean_qty_str))
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
