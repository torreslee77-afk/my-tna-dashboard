import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# 1. 페이지 설정
st.set_page_config(page_title="YAKJIN Operational Dashboard", page_icon="📊", layout="wide")

# CSS 스타일 설정
st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; }
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; margin-bottom: 40px; }
    .main-title { font-size: 3em; font-weight: bold; color: #1E3A8A; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

# 사이드바 메뉴 설정
menu = st.sidebar.selectbox("메뉴 선택", ["TNA Dashboard", "AD Sample Summary"])

# --- [기존 TNA 관련 함수들] ---
def get_weeks_display(ls_val):
    if pd.isnull(ls_val) or ls_val == '-': return None
    try:
        today = datetime(2026, 7, 1)
        target_date = datetime.strptime(f"2026/{ls_val}", "%Y/%m/%d")
        delta = (target_date - today).days
        if delta < 0: return "In Production"
        return f"{round(delta / 7, 1)}"
    except: return None

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
                header_idx = idx; break
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
            if col not in seen: seen[col] = 0; unique_columns.append(col)
            else: seen[col] += 1; unique_columns.append(f"{col}_{seen[col]}")
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = unique_columns
        style_col, div_col, print_col, fwash_col, line_start_col, line_end_col = None, None, None, None, None, None
        ex_factory_col, exf_qty_col, qty_col, risk_col = None, None, None, None
        for col in df.columns:
            c_clean = clean_string(col)
            if 'STYLE' in c_clean and '배정' not in c_clean: style_col = col
            elif any(k in c_clean for k in ['DIVISION', 'DIV']): div_col = col
            elif 'PRINT' in c_clean: print_col = col
            elif 'FWASH' in c_clean or 'F/WASH' in c_clean: fwash_col = col
            elif 'START' in c_clean: line_start_col = col
            elif 'END' in c_clean and 'START' not in c_clean: line_end_col = col
            elif '1STEXFQTY' in c_clean: exf_qty_col = col
            elif '납기별수량' in c_clean and 'EXF' in c_clean: ex_factory_col = col
            elif any(k in c_clean for k in ['TOTALORDERQTY', '작업수량']) and qty_col is None: qty_col = col
            elif 'KEY' in c_clean and 'RISK' in c_clean: risk_col = col
        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', '']: continue
            ls_date = row.get(line_start_col)
            ls_str = pd.to_datetime(ls_date, errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(ls_date, errors='coerce')) else '-'
            exf_val = row.get(ex_factory_col)
            exf_date = pd.to_datetime(exf_val, errors='coerce')
            exf_str = exf_date.strftime('%m/%d') if pd.notnull(exf_date) else '-'
            exf_qty_val = row.get(exf_qty_col)
            exf_qty_display = f"{int(float(str(exf_qty_val).replace(',', ''))):,}" if pd.notnull(exf_qty_val) and str(exf_qty_val).replace('.','').replace(',','').isdigit() else '-'
            qty_val = int(float(str(row.get(qty_col, 0)).replace(',', ''))) if pd.notnull(row.get(qty_col)) else 0
            risk_raw = row.get(risk_col)
            risk_val = str(risk_raw).strip().upper()
            if pd.isnull(risk_raw) or risk_val in ['NAN', 'NONE', '']: risk_val = 'N/A'
            sheet_rows.append({
                "Division": str(row.get(div_col, 'N/A')), "Style": style_raw, "Qty": f"{qty_val:,}",
                "Graphic": '🟢 O' if 'O' in str(row.get(print_col, '')) else '🔴 X',
                "Wash": '🟢 O' if 'O' in str(row.get(fwash_col, '')) else '🔴 X',
                "To LS (Wks)": get_weeks_display(ls_str), "Line Start": ls_str,
                "Line End": pd.to_datetime(row.get(line_end_col), errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(row.get(line_end_col), errors='coerce')) else '-',
                "1st Ex-Factory": exf_str, "1st Ex-Qty": exf_qty_display, "Risk": risk_val
            })
        if sheet_rows: all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
    return all_sheets_data

# --- [기능 2: AD Sample 분석 로직 수정본] ---
def run_ad_summary():
    st.title("📦 AD Sample Summary")
    uploaded_file = st.file_uploader("AD Sample raw data 파일을 업로드하세요.", type=["xlsx"], key="ad")
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            
            # 1. 상단 집계 영역
            st.subheader("Summary")
            col1, col2 = st.columns(2)
            
            # 부서별 총 수량 집계
            with col1:
                st.write("**Department별 총 수량**")
                dept_summary = df.groupby('Department')['Requested Qty'].sum().reset_index()
                st.dataframe(dept_summary, use_container_width=True)
            
            # 날짜별 수량 집계 (가장 최근 5일)
            with col2:
                st.write("**최근 5일 샘플 발송 예정 수량**")
                df['Estimated Send Date'] = pd.to_datetime(df['Estimated Send Date'])
                daily_summary = df.groupby('Estimated Send Date')['Requested Qty'].sum().reset_index()
                daily_summary = daily_summary.sort_values('Estimated Send Date', ascending=False).head(5)
                st.dataframe(daily_summary, use_container_width=True)
            
            # 2. 상세 내역 영역 (요청하신 순서대로)
            st.subheader("Detailed Breakdown")
            
            # 사이즈별 합계를 위해 그룹화
            group_cols = ['Department', 'Class', 'Style #', 'Color', 'Estimated Send Date', 'Estimated Arrival Date', 'Size']
            detailed_df = df.groupby(group_cols)['Requested Qty'].sum().reset_index()
            
            # 컬럼 순서 재정렬
            column_order = ['Department', 'Class', 'Style #', 'Color', 'Estimated Send Date', 'Estimated Arrival Date', 'Size', 'Requested Qty']
            detailed_df = detailed_df[column_order]
            
            st.dataframe(detailed_df, use_container_width=True)
            
        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
