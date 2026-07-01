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
st.markdown('<div class="sub-title">엑셀 TNA 파일을 업로드하면 각 팀별 신호등 현황과 리스크를 자동으로 분석합니다.</div>', unsafe_allow_html=True)

# Helper: 복잡한 엑셀 컬럼명에서 특정 키워드를 스마트하게 매칭하는 함수
def find_column(df_columns, primary_keywords, secondary_keyword=None):
    for col in df_columns:
        col_str = str(col).strip().upper().replace(" ", "").replace("'", "").replace("#", "")
        match_primary = any(p_kw.upper() in col_str for p_kw in primary_keywords)
        if match_primary:
            if secondary_keyword:
                if secondary_keyword.upper() in col_str:
                    return col
            else:
                return col
    return None

# 2. 엑셀 분석 엔진 함수
def analyze_tna(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets_data = {}
    
    for sheet_name in xls.sheet_names:
        # .xls 파일 구조적 특성을 고려하여 1, 2행을 로우 데이터로 먼저 읽음
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df_raw.empty or len(df_raw) < 2:
            continue
            
        # 1행(대분류)과 2행(소분류)을 깨끗하게 병합하여 단일 헤더 만들기
        row0 = df_raw.iloc[0].astype(str).replace('nan', '').str.strip()
        row1 = df_raw.iloc[1].astype(str).replace('nan', '').str.strip()
        
        combined_columns = []
        current_parent = ""
        for p, s in zip(row0, row1):
            if p != "":
                current_parent = p
            
            if current_parent and s:
                combined_columns.append(f"{current_parent} {s}")
            elif s:
                combined_columns.append(s)
            elif current_parent:
                combined_columns.append(current_parent)
            else:
                combined_columns.append("Unnamed")
                
        # 헤더 아래 진짜 데이터만 분리
        df = df_raw.iloc[2:].copy()
        df.columns = combined_columns

        # --- 약진 명세서 기준 컬럼 매칭 ---
        style_col = find_column(df.columns, ['STYLE', 'STYLE#', '배정STYLE'])
        buyer_col = find_column(df.columns, ['BUYER', 'DIVISION', '담당'])
        factory_col = find_column(df.columns, ['FACTORY'])
        
        print_col = find_column(df.columns, ['PRINT'])
        emb_col = find_column(df.columns, ['EMB', 'SEQUIN'])
        fwash_col = find_column(df.columns, ['FWASH', 'F/WASH'])
        gwash_col = find_column(df.columns, ['GWASH', 'G/WASH'])
        gdye_col = find_column(df.columns, ['GDYE', 'G/DYE'])
        
        line_start_col = find_column(df.columns, ['START'], 'LINE') or find_column(df.columns, ['START'])
        fabric_in_fac_col = find_column(df.columns, ['INFAC'], 'FABRIC') or find_column(df.columns, ['INFAC'])
        pps_appd_col = find_column(df.columns, ['APPD'], 'PP') or find_column(df.columns, ['PP'])
        ex_factory_col = find_column(df.columns, ['1STS/D', 'EXF', 'SD', 'EXF'])
        qty_col = find_column(df.columns, ['TOTALORDERQTY', 'QTY'])

        if style_col is None:
            continue
            
        sheet_rows = []
        for _, row in df.iterrows():
            style = str(row.get(style_col, '')).strip()
            if not style or style == 'nan' or style.upper().startswith('TOTAL'):
                continue
                
            # Graphic / Wash 자동 판정
            has_graphic = 'X'
            if print_col and pd.notna(row.get(print_col)) and str(row.get(print_col)).strip() not in ['', 'nan', 'X', 'x']:
                has_graphic = 'O'
            elif emb_col and pd.notna(row.get(emb_col)) and str(row.get(emb_col)).strip() not in ['', 'nan', 'X', 'x']:
                has_graphic = 'O'
                
            has_wash = 'X'
            if fwash_col and pd.notna(row.get(fwash_col)) and str(row.get(fwash_col)).strip() not in ['', 'nan', 'X', 'x']:
                has_wash = 'O'
            elif gwash_col and pd.notna(row.get(gwash_col)) and str(row.get(gwash_col)).strip() not in ['', 'nan', 'X', 'x']:
                has_wash = 'O'
            elif gdye_col and pd.notna(row.get(gdye_col)) and str(row.get(gdye_col)).strip() not in ['', 'nan', 'X', 'x']:
                has_wash = 'O'
            
            line_start_raw = row.get(line_start_col) if line_start_col else None
            if pd.isna(line_start_raw):
                continue
            try:
                line_start = pd.to_datetime(line_start_raw)
            except:
                continue
            
            # 워시/그래픽 여부에 따른 버퍼 룰 적용
            days_buffer = 14 if (has_graphic == 'O' or has_wash == 'O') else 7
            fabric_due = line_start - timedelta(days=14)
            fpp_due = line_start - timedelta(days=14)
            pps_due = line_start - timedelta(days=days_buffer)
            
            fabric_in_fac = row.get(fabric_in_fac_col) if fabric_in_fac_col else None
            fabric_status = "🔴 Late" if pd.isna(fabric_in_fac) else "🟢 Ready"
            
            pp_appd_raw = str(row.get(pps_appd_col, '')).strip() if pps_appd_col else ""
            if pp_appd_raw.upper() in ['N/A']:
                pps_status = "⚪ N/A"
            elif pp_appd_raw.upper() in ['C/O']:
                pps_status = "⚪ C/O"
            elif pp_appd_raw == 'nan' or not pp_appd_raw:
                pps_status = "➖"
            else:
                try: pps_status = pd.to_datetime(pp_appd_raw).strftime('%m/%d')
                except: pps_status = pp_appd_raw
                
            risk = "🟢 Low"
            if fabric_status == "🔴 Late":
                risk = "🔴 High"
                
            buyer_val = str(row.get(buyer_col, 'YAKJIN')).strip() if buyer_col else 'YAKJIN'
            factory_val = str(row.get(factory_col, 'YV')).strip() if factory_col else 'YV'
            ex_fac_val = '-'
            if ex_factory_col and pd.notna(row.get(ex_factory_col)):
                try: ex_fac_val = pd.to_datetime(row.get(ex_factory_col)).strftime('%m/%d')
                except: ex_fac_val = str(row.get(ex_factory_col))
                
            try: qty_val = int(row.get(qty_col, 0)) if qty_col and pd.notna(row.get(qty_col)) else 0
            except: qty_val = 0
                
            sheet_rows.append({
                "Style": style,
                "Buyer": buyer_val if buyer_val != 'nan' else 'YAKJIN',
                "Factory": factory_val if factory_val != 'nan' else 'YV',
                "Graphic": has_graphic,
                "Wash": has_wash,
                "Line Start": line_start.strftime('%m/%d'),
                "Fabric Due": fabric_due.strftime('%m/%d'),
                "Fabric Status": fabric_status,
                "FPP Due": fpp_due.strftime('%m/%d'),
                "PPS Status": pps_status,
                "1st Ex-Factory": ex_fac_val,
                "Qty": qty_val,
                "Risk": risk
            })
        if sheet_rows:
            all_sheets_data[sheet_name] = pd.DataFrame(sheet_rows)
            
    return all_sheets_data

# 3. UI - 파일 업로드 섹션
uploaded_file = st.file_uploader("TNA 엑셀 파일을 여기에 드래그하거나 선택하세요.", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("AI 엔진이 엑셀 데이터를 정밀 분석 중입니다..."):
        try:
            file_bytes = uploaded_file.read()
            results = analyze_tna(file_bytes)
            
            if not results:
                st.error("분석할 수 있는 데이터나 'STYLE' 또는 'STYLE#' 컬럼을 찾지 못했습니다. 파일 구조를 확인해 주세요.")
            else:
                total_styles = sum(len(df) for df in results.values())
                high_risks = sum(len(df[df['Risk'] == "🔴 High"]) for df in results.values())
                total_qty = sum(df['Qty'].sum() for df in results.values())
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="metric-box"><h4>총 스타일 수</h4><h2>{total_styles} 개</h2></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="metric-box"><h4>🔴 High Risk 스타일</h4><h2 style="color:red;">{high_risks} 개</h2></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="metric-box"><h4>총 오더 수량 (QTY)</h4><h2>{total_qty:,} pcs</h2></div>', unsafe_allow_html=True)
                
                st.write("---")
                
                tabs = st.tabs(list(results.keys()))
                for num, sheet_name in enumerate(results.keys()):
                    with tabs[num]:
                        st.subheader(f"📂 {sheet_name} 부서 현황")
                        st.dataframe(results[sheet_name], use_container_width=True, hide_index=True)
                        
        except Exception as e:
            st.error(f"파일을 읽는 중 에러가 발생했습니다: {e}")
