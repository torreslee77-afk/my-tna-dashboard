import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. 페이지 설정
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    /* 1. 맨 위에 여백 추가 */
    .block-container { padding-top: 3rem; }
    
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; margin-bottom: 40px; }
    
    /* 2. 글자 크기 3배 키우기 */
    .main-title { font-size: 3em; font-weight: bold; color: #1E3A8A; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 YAKJIN TNA Ai Operational dashboard</div>', unsafe_allow_html=True)

# 날짜 계산 함수 (문자열 반환)
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

        # 컬럼 매핑
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

        sheet_rows = []
        for _, row in df.iterrows():
            style_raw = str(row.get(style_col, '')).strip()
            if not style_raw or style_raw.lower() in ['nan', 'none', '']: continue
            
            ls_date = row.get(line_start_col)
            ls_str = pd.to_datetime(ls_date, errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(ls_date, errors='coerce')) else '-'
            
            sheet_rows.append({
                "Style": style_raw,
                "Division": str(row.get(div_col, 'N/A')),
                "Graphic": '🟢 O' if 'O' in str(row.get(print_col, '')) else '🔴 X',
                "Wash": '🟢 O' if 'O' in str(row.get(fwash_col, '')) else '🔴 X',
                "To LS (Wks)": get_weeks_display(ls_str),
                "Line Start": ls_str,
                "Line End": pd.to_datetime(row.get(line_end_col), errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(row.get(line_end_col), errors='coerce')) else '-',
                "1st Ex-Factory": pd.to_datetime(row.get(ex_factory_col), errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(row.get(ex_factory_col), errors='coerce')) else '-',
                "Qty": int(float(str(row.get(qty_col, 0)).replace(',', ''))) if pd.notnull(row.get(qty_col)) else 0,
                "Risk": '🔴 High' if pd.isnull(row.get(fabric_in_fac_col)) else '🟢 Low'
            })
        if sheet_rows: all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
    return all_sheets_data

uploaded_file = st.file_uploader("TNA 엑셀 파일을 업로드하세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    results = analyze_tna(uploaded_file.read())
    if results:
        tabs = st.tabs(list(results.keys()))
        for num, sheet_name in enumerate(results.keys()):
            with tabs[num]:
                df_sheet = results[sheet_name]
                
                # 요약 박스
                cols = st.columns(5)
                cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df_sheet):,}</h2></div>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{len(df_sheet[df_sheet["Risk"] == "🔴 High"]):,}</h2></div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{df_sheet["Qty"].sum():,}</h2></div>', unsafe_allow_html=True)
                cols[3].markdown(f'<div class="metric-box"><h4>Graphic</h4><h2>{len(df_sheet[df_sheet["Graphic"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                cols[4].markdown(f'<div class="metric-box"><h4>Wash</h4><h2>{len(df_sheet[df_sheet["Wash"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                
                # 색상 규칙 (apply 대신 map 활용)
                def color_rows(df):
                    styles = pd.DataFrame('', index=df.index, columns=df.columns)
                    for i, row in df.iterrows():
                        val = row['To LS (Wks)']
                        if val == "In Production": color = '#ffcccc'
                        else:
                            try:
                                v = float(val)
                                if v <= 2: color = '#ffcccc'
                                elif v <= 4: color = '#ffe6cc'
                                else: color = '#d4edda'
                            except: color = '#ffffff'
                        styles.loc[i, 'To LS (Wks)'] = f'background-color: {color}'
                    return styles

                st.dataframe(
                    df_sheet.style.apply(color_rows, axis=None), 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={
                        "Qty": st.column_config.NumberColumn("Qty", format="{:,}"),
                    }
                )
