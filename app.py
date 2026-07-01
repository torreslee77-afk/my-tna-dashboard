import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 업로드하는 시점의 오늘 날짜 (2026-07-01)
TODAY = pd.to_datetime('2026-07-01')

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 3rem; }
    .main-title { font-size: 32px; font-weight: bold; color: #1E3A8A; margin-bottom: 5px; }
    .sub-title { font-size: 16px; color: #6B7280; margin-bottom: 25px; }
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; margin-bottom: 40px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 YAKJIN TNA Ai Operational dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">TNA Analysis summary (Sheet-specific)</div>', unsafe_allow_html=True)

def parse_date(date_val):
    """날짜 파싱 시 연도를 2026년으로 고정"""
    if pd.isnull(date_val): return None
    try:
        s = str(date_val).strip()
        dt = pd.to_datetime(s, format='%m/%d', errors='ignore')
        if dt.year == 1900: 
            dt = dt.replace(year=2026)
        return dt
    except:
        return pd.to_datetime(date_val, errors='coerce')

def get_weeks_to_line_start(line_start):
    """남은 기간을 숫자로 표시 (예: 2.5 weeks)"""
    if pd.isnull(line_start): return "-"
    
    ls = parse_date(line_start)
    delta = (ls - TODAY).days
    
    if delta < 0: return "Passed"
    elif delta == 0: return "Today"
    else:
        weeks = round(delta / 7, 1)
        return f"{weeks} weeks"

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

        # 주요 컬럼 매핑
        style_col, div_col, print_col, fwash_col, line_start_col, line_end_col = None, None, None, None, None, None
        fabric_in_fac_col, ex_factory_col, qty_col = None, None, None

        for col in df.columns:
            c_clean = clean_string(col)
            if 'STYLE' in c_clean and '배정' not in c_clean: style_col = col
            elif any(k in c_clean for k in ['DIVISION', 'DIV']): div_col = col
            elif 'PRINT' in c_clean: print_col = col
            elif 'FWASH' in c_clean or 'F/WASH' in c_clean: fwash_col = col
            elif 'START' in c_clean: line_start_col = col
            elif 'END' in c_clean and 'START' not in c_clean: line_end_col = col
            elif 'INFAC' in c_clean: fabric_in_fac_col = col
            elif any(k in c_clean for k in ['1STSD', 'EXFAC', 'EXFACTORY', '1STEX', 'SD', 'S/D', 'FACTORYOUT', 'EXFACTORYDATE']): ex_factory_col = col
            elif any(k in c_clean for k in ['GMTQTY', 'TOTALORDERQTY', '작업수량']) and qty_col is None: qty_col = col

        if style_col is None: continue
            
        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', ''] or style_raw.upper().startswith('TOTAL'): continue
            
            styles_list = [s.strip() for s in style_raw.replace('/', ',').split(',') if s.strip()]
            
            try:
                ls_val = row.get(line_start_col)
                le_val = row.get(line_end_col)
                ex_fac_raw = row.get(ex_factory_col)
                ex_fac_str = pd.to_datetime(ex_fac_raw, errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(ex_fac_raw, errors='coerce')) else '-'
                
                qty_val = int(float(str(row.get(qty_col, 0)).replace(',', ''))) if pd.notnull(row.get(qty_col)) else 0
                allocated_qty = qty_val // len(styles_list)

                for single_style in styles_list:
                    sheet_rows.append({
                        "Style": single_style,
                        "Division": str(row.get(div_col, 'N/A')),
                        "Graphic": '🟢 O' if 'O' in str(row.get(print_col, '')) else '🔴 X',
                        "Wash": '🟢 O' if 'O' in str(row.get(fwash_col, '')) else '🔴 X',
                        "Line Start": str(ls_val),
                        "Line End": str(le_val),
                        "Weeks to Line Start": get_weeks_to_line_start(ls_val),
                        "1st Ex-Factory": ex_fac_str,
                        "Qty": allocated_qty,
                        "Risk": '🔴 High' if pd.isnull(row.get(fabric_in_fac_col)) else '🟢 Low'
                    })
            except: continue
                
        if sheet_rows: all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
    return all_sheets_data

uploaded_file = st.file_uploader("TNA 엑셀 파일을 여기에 드래그하거나 선택하세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    results = analyze_tna(uploaded_file.read())
    if results:
        tabs = st.tabs(list(results.keys()))
        for num, sheet_name in enumerate(results.keys()):
            with tabs[num]:
                df_sheet = results[sheet_name]
                df_disp = df_sheet.copy()
                df_disp['Qty'] = df_disp['Qty'].apply(lambda x: f"{x:,}")
                
                cols = st.columns(5)
                cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df_sheet):,}</h2></div>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{len(df_sheet[df_sheet["Risk"] == "🔴 High"]):,}</h2></div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{df_sheet["Qty"].sum():,}</h2></div>', unsafe_allow_html=True)
                cols[3].markdown(f'<div class="metric-box"><h4>Graphic</h4><h2>{len(df_sheet[df_sheet["Graphic"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                cols[4].markdown(f'<div class="metric-box"><h4>Wash</h4><h2>{len(df_sheet[df_sheet["Wash"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                
                st.dataframe(df_disp, use_container_width=True, hide_index=True)
