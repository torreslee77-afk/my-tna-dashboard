import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. 페이지 설정
st.set_page_config(page_title="YAKJIN Operational Dashboard", page_icon="📊", layout="wide")

# CSS 스타일 설정 (제목 잘림 방지용 최소화)
st.markdown("""
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    .metric-box { padding: 10px; background-color: #F3F4F6; border-radius: 8px; text-align: center; margin-bottom: 10px; }
    div[data-testid="stSidebar"] { width: 250px; }
    </style>
""", unsafe_allow_html=True)

menu = st.sidebar.selectbox("메뉴 선택", ["TNA Dashboard", "AD Sample Summary"])

# --- [공통 함수: 컬럼명 찾기] ---
def find_column(df, possible_names):
    for col in df.columns:
        clean_col = str(col).strip().upper().replace(" ", "")
        for name in possible_names:
            if name.upper().replace(" ", "") == clean_col:
                return col
    return None

# --- [TNA 함수] ---
def get_weeks_display(ls_val):
    if pd.isnull(ls_val) or ls_val == '-': return None
    try:
        today = datetime(2026, 7, 3)
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

# --- [AD Sample 함수] ---
def run_ad_summary():
    st.header("📦 AD Sample Summary")
    uploaded_file = st.file_uploader("AD Sample raw data 파일을 업로드하세요.", type=["xlsx"], key="ad_uploader")
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
            qty_col = find_column(df, ['Requested Qty', 'Qty', 'Quantity', '요청수량'])
            send_col = find_column(df, ['Estimated Send Date', 'Send Date', '발송일'])
            arr_col = find_column(df, ['Estimated Arrival Date', 'Arrival Date', '도착일'])
            dept_col = find_column(df, ['Department', 'Dept', '부서'])
            class_col = find_column(df, ['Class', '구분'])
            style_col = find_column(df, ['Style #', 'Style', '스타일'])
            color_col = find_column(df, ['Color', '컬러'])
            if not all([qty_col, send_col, arr_col, dept_col]):
                st.error(f"필수 컬럼을 찾을 수 없습니다.")
                return
            df['Qty'] = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
            df['SendDate'] = pd.to_datetime(df[send_col], errors='coerce').dt.strftime('%Y-%m-%d')
            df['ArrDate'] = pd.to_datetime(df[arr_col], errors='coerce').dt.strftime('%Y-%m-%d')
            col1, col2 = st.columns(2)
            with col1:
                st.write("**부서별 총 수량**")
                st.dataframe(df.groupby(dept_col)['Qty'].sum().reset_index(), use_container_width=True)
            with col2:
                st.write("**최근 5일 샘플 발송 예정 수량**")
                daily_sum = df.groupby('SendDate')['Qty'].sum().sort_index(ascending=False).head(5).reset_index()
                st.dataframe(daily_sum, use_container_width=True)
            st.write("**상세 내역**")
            st.dataframe(df[[dept_col, class_col, style_col, color_col, 'SendDate', 'ArrDate', 'Qty']], use_container_width=True)
        except Exception as e:
            st.error(f"데이터 처리 오류: {e}")

# --- 메인 실행 ---
if menu == "TNA Dashboard":
    st.header("📊 YAKJIN TNA AI Operational Dashboard")
    uploaded_file = st.file_uploader("TNA 엑셀 파일을 업로드하세요.", type=["xlsx", "xls"], key="tna_uploader")
    if uploaded_file:
        results = analyze_tna(uploaded_file.read())
        if results:
            all_df = pd.concat(results.values(), ignore_index=True)
            tab_names = ["All Summary"] + list(results.keys())
            tabs = st.tabs(tab_names)
            for num, tab_name in enumerate(tab_names):
                with tabs[num]:
                    df_sheet = all_df if tab_name == "All Summary" else results[tab_name]
                    df_calc = df_sheet.copy()
                    df_calc['Qty_Num'] = df_calc['Qty'].astype(str).str.replace(',', '').astype(int)
                    cols = st.columns(5)
                    cols[0].markdown(f'<div class="metric-box"><h4>Styles</h4><h2>{len(df_sheet):,}</h2></div>', unsafe_allow_html=True)
                    cols[1].markdown(f'<div class="metric-box"><h4>Qty</h4><h2>{df_calc["Qty_Num"].sum():,}</h2></div>', unsafe_allow_html=True)
                    cols[2].markdown(f'<div class="metric-box"><h4>Risk</h4><h2 style="color:red;">{len(df_sheet[df_sheet["Risk"].isin(["KEY", "HIGH RISK"])]):,}</h2></div>', unsafe_allow_html=True)
                    cols[3].markdown(f'<div class="metric-box"><h4>Wash</h4><h2>{len(df_sheet[df_sheet["Wash"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                    cols[4].markdown(f'<div class="metric-box"><h4>Graphic</h4><h2>{len(df_sheet[df_sheet["Graphic"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                    
                    def highlight_risk(df):
                        style_df = pd.DataFrame('', index=df.index, columns=df.columns)
                        mask = df['Risk'].isin(['KEY', 'HIGH RISK'])
                        style_df.loc[mask, 'Risk'] = 'background-color: #ffcccc'
                        return style_df
                    
                    st.dataframe(df_sheet.style.apply(highlight_risk, axis=None), use_container_width=True, hide_index=True)

elif menu == "AD Sample Summary":
    run_ad_summary()
