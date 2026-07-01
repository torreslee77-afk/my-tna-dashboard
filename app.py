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

# 2. 엑셀 분석 엔진 함수
def analyze_tna(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets_data = {}
    
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if df.empty:
            continue
            
        # 대소문자 공백 상관없이 Style / Style# / STYLE# 컬럼 유연하게 찾기
        style_col = None
        for col in df.columns:
            clean_col = str(col).strip().upper()  # 대문자로 통일해서 비교
            if clean_col in ['STYLE', 'STYLE#']:
                style_col = col
                break
                
        if style_col is None:
            continue
            
        sheet_rows = []
        for _, row in df.iterrows():
            style = str(row.get(style_col, '')).strip()
            if not style or style == 'nan':
                continue
                
            # Graphic / Wash 자동 판정
            has_graphic = 'O' if any(x in str(row.get('PRINT', '')) or x in str(row.get('EMB/SEQUIN', '')) for x in ['O', 'Yes', '1']) else 'X'
            has_wash = 'O' if any(x in str(row.get('F/WASH', '')) or x in str(row.get('G/WASH', '')) or x in str(row.get('G/DYE', '')) for x in ['O', 'Yes', '1']) else 'X'
            
            # Line Start 기준일 체크
            line_start_raw = row.get('Line Start')
            if pd.isna(line_start_raw):
                continue
            line_start = pd.to_datetime(line_start_raw)
            
            # Rule 적용 (Graphic/Wash가 있으면 D-14, 없으면 D-7)
            days_buffer = 14 if (has_graphic == 'O' or has_wash == 'O') else 7
            fabric_due = line_start - timedelta(days=14)
            fpp_due = line_start - timedelta(days=14)
            pps_due = line_start - timedelta(days=days_buffer)
            
            # 신호등 상태 판정
            fabric_in_fac = row.get('Fabric IN FAC')
            fabric_status = "🔴 Late" if pd.isna(fabric_in_fac) else "🟢 Ready"
            
            pp_appd_raw = str(row.get('PP(GTS APPD)', '')).strip()
            if pp_appd_raw in ['N/A', 'n/a']:
                pps_status = "⚪ N/A"
            elif pp_appd_raw in ['C/O', 'c/o']:
                pps_status = "⚪ C/O"
            elif pp_appd_raw == 'nan' or not pp_appd_raw:
                pps_status = "➖"
            else:
                pps_status = pp_appd_raw
                
            # 리스크 판정
            risk = "🟢 Low"
            if fabric_status == "🔴 Late":
                risk = "🔴 High"
                
            sheet_rows.append({
                "Style": style,
                "Buyer": str(row.get('Buyer', 'GAP')),
                "Factory": str(row.get('Factory', 'YV')),
                "Graphic": has_graphic,
                "Wash": has_wash,
                "Line Start": line_start.strftime('%m/%d'),
                "Fabric Due": fabric_due.strftime('%m/%d'),
                "Fabric Status": fabric_status,
                "FPP Due": fpp_due.strftime('%m/%d'),
                "PPS Status": pps_status,
                "1st Ex-Factory": pd.to_datetime(row.get('1st S/D')).strftime('%m/%d') if not pd.isna(row.get('1st S/D')) else '-',
                "Qty": int(row.get('QTY', 0)) if not pd.isna(row.get('QTY')) else 0,
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
                st.error("분석할 수 있는 데이터나 'STYLE' 또는 'STYLE#' 컬럼을 찾지 못했습니다.")
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
