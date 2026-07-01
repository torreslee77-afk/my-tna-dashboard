import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. 페이지 설정
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 3rem; }
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; margin-bottom: 40px; }
    .main-title { font-size: 3em; font-weight: bold; color: #1E3A8A; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 YAKJIN TNA Ai Operational dashboard</div>', unsafe_allow_html=True)

def get_weeks_display(ls_val):
    if pd.isnull(ls_val) or ls_val == '-': return None
    try:
        today = datetime(2026, 7, 1)
        target_date = datetime.strptime(f"2026/{ls_val}", "%Y/%m/%d")
        delta = (target_date - today).days
        if delta < 0: return "In Production"
        return f"{round(delta / 7, 1)}"
    except: return None

def analyze_tna(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets_data = {}
    
    for sheet_name in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df_raw.empty: continue
        
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = "".join([str(v) for v in row.values])
            if 'STYLE' in row_str.upper():
                header_idx = idx; break
        if header_idx is None: continue
            
        row0 = df_raw.iloc[header_idx].fillna("").astype(str)
        row1 = df_raw.iloc[header_idx + 1].fillna("").astype(str)
        combined_columns = [f"{p}_{s}".upper().replace(" ", "") for p, s in zip(row0, row1)]
        
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = combined_columns

        style_col, div_col, print_col, fwash_col, line_start_col, line_end_col = None, None, None, None, None, None
        fabric_in_fac_col, ex_factory_col, exf_qty_col, qty_col = None, None, None, None

        for col in df.columns:
            if 'STYLE' in col and '배정' not in col: style_col = col
            elif 'DIVISION' in col or 'DIV' in col: div_col = col
            elif 'PRINT' in col: print_col = col
            elif 'FWASH' in col: fwash_col = col
            elif 'START' in col: line_start_col = col
            elif 'END' in col and 'START' not in col: line_end_col = col
            elif 'INFAC' in col: fabric_in_fac_col = col
            elif '납기별수량' in col and 'EXF' in col: ex_factory_col = col
            elif '납기별수량' in col and 'QTY' in col: exf_qty_col = col
            elif 'TOTALORDERQTY' in col or '작업수량' in col: qty_col = col

        sheet_rows = []
        for _, row in df.iterrows():
            style_val = str(row.get(style_col, '')).strip()
            if not style_val or style_val.lower() in ['nan', 'none', '']: continue
            
            exf_date = pd.to_datetime(row.get(ex_factory_col), errors='coerce')
            exf_str = exf_date.strftime('%m/%d') if pd.notnull(exf_date) else '-'
            
            q_val = str(row.get(exf_qty_col, ''))
            clean_q = ''.join(filter(str.isdigit, q_val))
            exf_qty_display = f"{int(clean_q):,}" if clean_q else '-'
            
            ls_date = pd.to_datetime(row.get(line_start_col), errors='coerce')
            ls_str = ls_date.strftime('%m/%d') if pd.notnull(ls_date) else '-'
            
            sheet_rows.append({
                "Style": style_val,
                "Division": str(row.get(div_col, 'N/A')),
                "Graphic": '🟢 O' if 'O' in str(row.get(print_col, '')) else '🔴 X',
                "Wash": '🟢 O' if 'O' in str(row.get(fwash_col, '')) else '🔴 X',
                "To LS (Wks)": get_weeks_display(ls_str),
                "Line Start": ls_str,
                "Line End": pd.to_datetime(row.get(line_end_col), errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(row.get(line_end_col), errors='coerce')) else '-',
                "1st Ex-Factory": exf_str,
                "1st Ex-Qty": exf_qty_display,
                "Qty": int(pd.to_numeric(str(row.get(qty_col, 0)).replace(',', ''), errors='coerce') or 0),
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
                
                cols = st.columns(5)
                cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df_sheet):,}</h2></div>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{len(df_sheet[df_sheet["Risk"] == "🔴 High"]):,}</h2></div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{df_sheet["Qty"].sum():,}</h2></div>', unsafe_allow_html=True)
                cols[3].markdown(f'<div class="metric-box"><h4>Graphic</h4><h2>{len(df_sheet[df_sheet["Graphic"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                cols[4].markdown(f'<div class="metric-box"><h4>Wash</h4><h2>{len(df_sheet[df_sheet["Wash"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                
                def color_rows(df):
                    styles = pd.DataFrame('', index=df.index, columns=df.columns)
                    for i, row in df.iterrows():
                        val = row['To LS (Wks)']
                        if val == "In Production": color = '#d3d3d3'
                        else:
                            try:
                                v = float(val)
                                if v <= 2: color = '#ffcccc'
                                elif v <= 4: color = '#ffe6cc'
                                else: color = '#d4edda'
                            except: color = '#ffffff'
                        styles.loc[i, 'To LS (Wks)'] = f'background-color: {color}'
                    return styles

                st.dataframe(df_sheet.drop(columns=['Qty']).style.apply(color_rows, axis=None), use_container_width=True, hide_index=True)
