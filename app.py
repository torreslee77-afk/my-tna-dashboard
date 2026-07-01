import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="YAKJIN TNA Dashboard", layout="wide")

# 스타일링
st.markdown("""
    <style>
    .block-container { padding-top: 1rem; }
    .metric-box { padding: 10px; background-color: #F8F9FA; border-radius: 8px; border: 1px solid #dee2e6; text-align: center; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 YAKJIN TNA Operational Dashboard")

def is_valid_work(val):
    s = str(val).strip().upper()
    return s not in ['', 'NAN', 'NONE', 'X', 'N/A', 'C/O', 'NAT', '<NA>', 'NULL', '0']

def analyze_tna(file_bytes):
    # CSV/Excel 읽기
    try:
        df_raw = pd.read_csv(file_bytes) if file_bytes.name.endswith('.csv') else pd.read_excel(file_bytes)
    except: return pd.DataFrame()

    # 데이터 정리: 'STYLE#' 컬럼을 기준으로 데이터 시작점 찾기
    df = df_raw.copy()
    
    # 주요 컬럼 찾기 (유연한 매칭)
    def get_col(names):
        for name in names:
            if name in df.columns: return name
        return None

    style_col = get_col(['STYLE#', 'STYLE'])
    div_col = get_col(['DIVISION', 'DIV'])
    qty_col = get_col(['TOTAL ORDER Q\'TY', 'TOTALORDERQTY', 'GMT QTY'])
    
    # 작업 컬럼들
    print_c = get_col(['PRINT'])
    emb_c = get_col(['EMB/ SEQUIN'])
    fwash_c = get_col(['F/W ASH'])
    gwash_c = get_col(['G/ WASH'])
    
    # 위험도 판단을 위한 입고 예정일 (IN FAC) 컬럼들
    # 데이터가 흩어져 있으므로, 관련 컬럼이 비어있는지 확인
    infac_cols = [c for c in df.columns if 'IN FAC' in str(c)]

    processed = []
    for _, row in df.iterrows():
        if pd.isna(row.get(style_col)): continue
        
        # 1. 그래픽/워시 작업 여부
        has_graphic = '🟢 O' if (print_c and is_valid_work(row.get(print_c))) or (emb_c and is_valid_work(row.get(emb_c))) else '🔴 X'
        has_wash = '🟢 O' if (fwash_c and is_valid_work(row.get(fwash_c))) or (gwash_c and is_valid_work(row.get(gwash_c))) else '🔴 X'
        
        # 2. Risk 판단: IN FAC 컬럼들 중 하나라도 비어있으면 Late로 간주
        is_late = any(pd.isna(row.get(c)) or str(row.get(c)).strip() == '' for c in infac_cols)
        risk = '🔴 High' if is_late else '🟢 Low'
        
        processed.append({
            "Style": row.get(style_col),
            "Division": row.get(div_col, 'N/A'),
            "Graphic": has_graphic,
            "Wash": has_wash,
            "Qty": int(row.get(qty_col, 0)) if pd.notnull(row.get(qty_col)) else 0,
            "Risk": risk
        })
    return pd.DataFrame(processed)

uploaded_file = st.file_uploader("파일을 업로드하세요", type=["csv", "xlsx"])
if uploaded_file:
    df = analyze_tna(uploaded_file)
    if not df.empty:
        # 요약 지표
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(f'<div class="metric-box"><h4>TTL Styles</h4><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-box"><h4>High Risk</h4><h2 style="color:red;">{len(df[df["Risk"]=="🔴 High"])}</h2></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-box"><h4>TTL Qty</h4><h2>{df["Qty"].sum():,}</h2></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="metric-box"><h4>Graphic styles</h4><h2>{len(df[df["Graphic"]=="🟢 O"])}</h2></div>', unsafe_allow_html=True)
        c5.markdown(f'<div class="metric-box"><h4>Wash styles</h4><h2>{len(df[df["Wash"]=="🟢 O"])}</h2></div>', unsafe_allow_html=True)
        
        st.write("---")
        df_display = df.copy()
        df_display['Qty'] = df_display['Qty'].apply(lambda x: f"{x:,}")
        st.dataframe(df_display, use_container_width=True)
