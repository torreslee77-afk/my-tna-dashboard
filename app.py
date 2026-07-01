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
    """
    primary_keywords 중 하나를 포함하고, 
    만약 secondary_keyword가 지정되면 그것까지 포함하는 컬럼 인덱스/이름을 반환합니다.
    """
    for col in df_columns:
        col_str = str(col).strip().upper().replace(" ", "").replace("'", "").replace("#", "")
        
        # 기본 키워드 매칭 체크
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
        # 데이터 유실 방지를 위해 일단 헤더 없이 읽은 후 유연하게 매칭 처리
        df = pd.read_excel(xls, sheet_name=sheet_name, header=[0, 1] if xls.book.sheet_by_name(sheet_name).nrows > 1 else 0)
        
        if df.empty:
            continue
            
        # 다중 헤더(Multi-index)인 경우 컬럼명을 텍스트 하나로 병합하여 인식하기 쉽게 처리
        if isinstance(df.columns, pd.MultiIndex):
            flat_cols = []
            for col in df.columns:
                p_part = str(col[0]).strip() if not pd.isna(col[0]) and 'Unnamed' not in str(col[0]) else ""
                s_part = str(col[1]).strip() if not pd.isna(col[1]) and 'Unnamed' not in str(col[1]) else ""
                flat_cols.append(f"{p_part} {s_part}".strip())
            df.columns = flat_cols

        # --- 유연한 컬럼 매칭 알고리즘 적용 ---
        style_col = find_column(df.columns, ['STYLE', 'STYLE#', '배정Style'])
        buyer_col = find_column(df.columns, ['BUYER', 'DIVISION', '담당'])
        factory_col = find_column(df.columns, ['FACTORY'])
        
        # Graphic / Wash 공정 여부 판단 컬럼 (약진 TNA 명세서 반영)
        print_col = find_column(df.columns, ['PRINT'])
        emb_col = find_column(df.columns, ['EMB', 'SEQUIN'])
        fwash_col = find_column(df.columns, ['FWASH', 'F/WASH'])
        gwash_col = find_column(df.columns, ['GWASH', 'G/WASH'])
        gdye_col = find_column(df.columns, ['GDYE', 'G/DYE'])
        
        # 중요 날짜 및 수량 컬럼 매칭
        line_start_col = find_column(df.columns, ['START'], 'LINE') or find_column(df.columns, ['START'])
        fabric_in_fac_col = find_column(df.columns, ['INFAC'], 'FABRIC') or find_column(df.columns, ['INFAC'])
        pps_appd_col = find_column(df.columns, ['APPD'], 'PP') or find_column(df.columns, ['PP'])
        ex_factory_col = find_column(df.columns, ['1stS/D', 'EXF', 'SD'])
        qty_col = find_column(df.columns, ['TOTALORDERQTY', 'QTY'])

        # 필수 기준이 되는 Style 컬럼이 없으면 해당 시트는 패스
        if style_col is None:
            continue
            
        sheet_rows = []
        for _, row in df.iterrows():
            style = str(row.get(style_col, '')).strip()
            if not style or style == 'nan' or style.startswith('TOTAL'):
                continue
                
            # Graphic / Wash 자동 판정 (O, Yes, 1 또는 빈칸이 아니거나 방법이 적힌 경우)
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
            
            # Line Start 기준일 파싱
            line_start_raw = row.get(line_start_col) if line_start_col else None
            if pd.isna(line_start_raw):
                continue
            try:
                line_start = pd.to_datetime(line_start_raw)
            except:
                continue
            
            # Rule 적용 (Graphic/Wash가 있으면 D-14, 없으면 D-7)
            days_buffer = 14 if (has_graphic == 'O' or has_wash == 'O') else 7
            fabric_due = line_start - timedelta(days=14)
            fpp_due = line_start - timedelta(days=14)
            pps_due = line_start - timedelta(days=days_buffer)
            
            # 원단 입고 상태 판정 (Fabric IN FAC 기반)
            fabric_in_fac = row.get(fabric_in_fac_col) if fabric_in_fac_col else None
            fabric_status = "🔴 Late" if pd.isna(fabric_in_fac) else "🟢 Ready"
            
            # PPS 승인 상태 판정 (N/A, C/O 대응)
            pp_appd_raw = str(row.get(pps_appd_col, '')).strip() if pps_appd_col else ""
            if pp_appd_raw.upper() in ['N/A']:
                pps_status = "⚪ N/A"
            elif pp_appd_raw.upper() in ['C/O']:
                pps_status = "⚪ C/O"
            elif pp_appd_raw == 'nan' or not pp_appd_raw:
                pps_status = "➖"
            else:
                # 날짜 형식인 경우 이쁘게 포맷팅 시도
                try:
                    pps_status = pd.to_datetime(pp_appd_raw).strftime('%m/%d')
                except:
                    pps_status = pp_appd_raw
                
            # 리스크 판정
            risk = "🟢 Low"
            if fabric_status == "🔴 Late":
                risk = "🔴 High"
                
            # 최종 정돈된 데이터 매핑
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
                # 상단 요약 대시보드 계산
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
                
                # 시트별 탭 생성
                tabs = st.tabs(list(results.keys()))
                for num, sheet_name in enumerate(results.keys()):
                    with tabs[num]:
                        st.subheader(f"📂 {sheet_name} 부서 현황")
                        st.dataframe(results[sheet_name], use_container_width=True, hide_index=True)
                        
        except Exception as e:
            st.error(f"파일을 읽는 중 에러가 발생했습니다: {e}")
