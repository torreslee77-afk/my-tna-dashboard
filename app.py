import streamlit as st
import pandas as pd
from datetime import datetime
import io

# 1. Page Configuration
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

st.markdown("""
    <style>
    .metric-box { padding: 15px; background-color: #F3F4F6; border-radius: 8px; text-align: center; margin-bottom: 40px; }
    .main-title { font-size: 32px; font-weight: bold; color: #1E3A8A; margin-bottom: 20px; }
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
        
        # Identify header
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
        
        sheet_rows = []
        for _, row in df.iterrows():
            style = str(row.get(next((c for c in df.columns if 'STYLE' in clean_string(c)), None), '')).strip()
            if not style or style.lower() in ['nan', 'none']: continue
            
            # Extract values carefully to avoid Pandas truth ambiguity
            ls_val = row.get(next((c for c in df.columns if 'START' in clean_string(c)), None))
            ls_str = pd.to_datetime(ls_val, errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(ls_val, errors='coerce')) else '-'
            
            qty_raw = row.get(next((c for c in df.columns if any(k in clean_string(c) for k in ['QTY', '수량'])), None), 0)
            
            # Safe Risk evaluation
            infac_val = row.get(next((c for c in df.columns if 'INFAC' in clean_string(c)), None))
            is_high_risk = pd.isnull(infac_val)
            
            sheet_rows.append({
                "Style": style,
                "Division": str(row.get(next((c for c in df.columns if 'DIV' in clean_string(c)), 'N/A'))),
                "Graphic": '🟢 O' if 'O' in str(row.get(next((c for c in df.columns if 'PRINT' in clean_string(c)), ''))).upper() else '🔴 X',
                "Wash": '🟢 O' if 'O' in str(row.get(next((c for c in df.columns if 'FWASH' in clean_string(c)), ''))).upper() else '🔴 X',
                "To LS (Wks)": get_weeks_display(ls_str),
                "Line Start": ls_str,
                "Line End": pd.to_datetime(row.get(next((c for c in df.columns if 'END' in clean_string(c) and 'START' not in clean_string(c)), None)), errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(row.get(next((c for c in df.columns if 'END' in clean_string(c) and 'START' not in clean_string(c)), None)), errors='coerce')) else '-',
                "1st Ex-Factory": pd.to_datetime(row.get(next((c for c in df.columns if any(k in clean_string(c) for k in ['1ST', 'EXFAC'])), None)), errors='coerce').strftime('%m/%d') if pd.notnull(pd.to_datetime(row.get(next((c for c in df.columns if any(k in clean_string(c) for k in ['1ST', 'EXFAC'])), None)), errors='coerce')) else '-',
                "Qty": int(pd.to_numeric(str(qty_raw).replace(',', ''), errors='coerce') or 0),
                "Risk": '🔴 High' if is_high_risk else '🟢 Low'
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
                df = results[sheet_name]
                # UI Components
                cols = st.columns(5)
                cols[0].markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df):,}</h2></div>', unsafe_allow_html=True)
                cols[1].markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{len(df[df["Risk"] == "🔴 High"]):,}</h2></div>', unsafe_allow_html=True)
                cols[2].markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{df["Qty"].sum():,}</h2></div>', unsafe_allow_html=True)
                cols[3].markdown(f'<div class="metric-box"><h4>Graphic</h4><h2>{len(df[df["Graphic"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                cols[4].markdown(f'<div class="metric-box"><h4>Wash</h4><h2>{len(df[df["Wash"] == "🟢 O"]):,}</h2></div>', unsafe_allow_html=True)
                
                # Apply conditional formatting
                def color_rows(row):
                    val = row['To LS (Wks)']
                    if val == "In Production": color = '#ffcccc'
                    else:
                        try:
                            v = float(val)
                            color = '#ffcccc' if v <= 2 else '#ffe6cc' if v <= 4 else '#d4edda'
                        except: color = '#ffffff'
                    return [f'background-color: {color}' if col == 'To LS (Wks)' else '' for col in row.index]

                st.dataframe(
                    df.style.apply(color_rows, axis=1), 
                    use_container_width=True, hide_index=True,
                    column_config={"Qty": st.column_config.NumberColumn("Qty", format="%d")}
                )
