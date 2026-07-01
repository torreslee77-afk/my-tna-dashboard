import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. 페이지 설정
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

# 오늘 날짜 (2026-07-01 기준)
TODAY = pd.to_datetime('2026-07-01')

# 2. CSS 스타일 (여백 및 디자인)
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

# 3. 날짜 연산 함수
def get_weeks_to_line_start(date_val):
    if pd.isnull(date_val) or date_val == '-': return "-"
    try:
        # MM/DD 형식의 날짜 문자열을 2026년으로 변환
        dt = pd.to_datetime(f"2026/{date_val}")
        delta = (dt - TODAY).days
        if delta < 0: return "Under Production"
        if delta == 0: return "Today"
        return f"{round(delta / 7, 1)} weeks"
    except:
        return "-"

def clean_string(val):
    try:
        s_val = str(val).strip().upper()
        if s_val in ['NAN', 'NONE', '<NA>', 'NAT', 'NULL', '']: return ""
        return s_val.replace(" ", "").replace("'", "").replace("#", "").replace("/", "").replace("(", "").replace(")", "").replace("-", "").replace("\n", "").replace("\r", "")
    except: return ""

# 4. 데이터 분석 함수
def analyze_tna(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets_data = {}
    for sheet_name in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df_raw.empty: continue
        header_idx = None
        for idx, row in df_raw.iterrows():
            if any('STYLE' in clean_string(v) for v in row.values if v):
                header_idx = idx; break
        if header_idx is None: continue
        
        row0 = df_raw.iloc[header_idx].astype(str).replace('nan', '').str.strip()
        row1 = df_raw.iloc[header_idx + 1].astype(str).replace('nan', '').str.strip() if (header_idx + 1) < len(df_raw) else row0
        combined = [f"{p} {s}".strip() for p, s in zip(row0, row1)]
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = combined
        
        # 컬럼 찾기
        style_col, div_col, print_col, fwash_col, line_start_col, qty_col, fab_col = None, None, None, None, None, None, None
        for col in df.columns:
            c = clean_string(col)
            if 'STYLE' in c: style_col = col
            elif 'DIV' in c: div_col = col
            elif 'PRINT' in c: print_col = col
            elif 'FWASH' in c: fwash_col = col
            elif 'START' in c: line_start_col = col
            elif any(k in c for k in ['QTY', '수량']): qty_col = col
            elif 'INFAC' in c: fab_col = col
            
        sheet_rows = []
        for _, row in df.iterrows():
            style = str(row.get(style_col, '')).strip()
            if not style or style.lower() in ['nan', 'none']: continue
            try: qty = int(float(str(row.get(qty_col, 0)).replace(',', '')))
            except: qty = 0
            
            sheet_rows.append({
                "Style": style,
                "Division": str(row.get(div_col, 'N/A')),
                "Graphic": '🟢 O' if 'O' in str(row.get(print_col, '')) else '🔴 X',
                "Wash": '🟢 O' if 'O' in str(row.get(fwash_col, '')) else '🔴 X',
                "Line Start": str(row.get(line_start_col, '-')),
                "Weeks to Line Start": get_weeks_to_line_start(row.get(line_start_col)),
                "Qty": qty,
                "Risk": '🔴 High' if pd.isnull(row.get(fab_col)) else '🟢 Low'
            })
        if sheet_rows: all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
    return all_sheets_data

# 5. 메인 UI
uploaded_file = st.file_uploader("TNA 엑셀 파일을 드래그하세요.", type=["xlsx", "xls"])
if uploaded_file is not None:
    results = analyze_tna(uploaded_file.read())
    if results:
        tabs = st.tabs(list(results.keys()))
        for num, sheet_name in enumerate(results.keys()):
            with tabs[num]:
                df_sheet = results[sheet_name]
                df_disp = df_sheet.copy()
                df_disp['Qty'] = df_disp['Qty'].apply(lambda x: f"{x:,}")
                
                # 지표 카드
                cols = st.columns(5)
                cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df_sheet):,}</h2></div>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{len(df_sheet[df_sheet["Risk"] == "🔴 High"]):,}</h2></div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{df_sheet["Qty"].sum():,}</h2></div>', unsafe_allow_html=True)
                cols[3].markdown(f'<div class="metric-box"><h4>Graphic</h4><h2>{len(df_sheet[df_sheet["Graphic"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                cols[4].markdown(f'<div class="metric-box"><h4>Wash</h4><h2>{len(df_sheet[df_sheet["Wash"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                
                st.dataframe(df_disp, use_container_width=True, hide_index=True)
