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

# Helper: 대소문자, 공백, 특수문자 제거 후 키워드 매칭
def clean_string(val):
    if pd.isna(val):
        return ""
    return str(val).strip().upper().replace(" ", "").replace("'", "").replace("#", "").replace("/", "")

# 2. 엑셀 분석 엔진 함수
def analyze_tna(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets_data = {}
    
    for sheet_name in xls.sheet_names:
        # 헤더 없이 모든 데이터를 통째로 읽어옴
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df_raw.empty:
            continue
            
        # --- [핵심] 병합된 셀을 고려하여 STYLE# 제목 행 찾기 ---
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_values = [clean_string(v) for v in row.values]
            has_style = any(k in row_values for k in ['STYLE', 'STYLE', '배정STYLE'])
            if not has_style:
                has_style = any(('STYLE' in v or '배정STYLE' in v) for v in row_values if v)
                
            if has_style:
                header_idx = idx
                break
                
        if header_idx is None:
            continue  # STYLE#을 못 찾은 시트는 패스
            
        # 제목 줄과 그 다음 줄을 결합하여 복잡한 2중 헤더 처리
        row0 = df_raw.iloc[header_idx].astype(str).replace('nan', '').str.strip()
        row1 = df_raw.iloc[header_idx + 1].astype(str).replace('nan', '').str.strip() if (header_idx + 1) < len(df_raw) else row0
        
        combined_columns = []
        current_parent = ""
        for p, s in zip(row0, row1):
            if p != "":
                current_parent = p
            if current_parent and s and p != s:
                combined_columns.append(f"{current_parent} {s}")
            elif s:
                combined_columns.append(s)
            elif current_parent:
                combined_columns.append(current_parent)
            else:
                combined_columns.append("Unnamed")
                
        # --- [오류 해결 핵심] 동일 이름 컬럼 고유화 (Series 에러 방지) ---
        final_columns = []
        counts = {}
        for col in combined_columns:
            if col not in counts:
                counts[col] = 1
                final_columns.append(col)
            else:
                counts[col] += 1
                final_columns.append(f"{col}_{counts[col]}")
                
        # 진짜 데이터 부분만 잘라내고 고유화된 컬럼명 부여
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = final_columns

        # --- 매칭 타겟 컬럼 인덱스 찾기 ---
        style_col, buyer_col, factory_col = None, None, None
        print_col, emb_col, fwash_col, gwash_col, gdye_col = None, None, None, None, None
        line_start_col, fabric_in_fac_col, pps_appd_col, ex_factory_col, qty_col = None, None, None, None, None

        for col in df.columns:
            c_clean = clean_string(col)
            if any(k in c_clean for k in ['STYLE', 'STYLE', '배정STYLE']): style_col = col
            elif any(k in c_clean for k in ['BUYER', 'DIVISION', '담당']): buyer_col = col
            elif 'FACTORY' in c_clean: factory_col = col
            elif 'PRINT' in c_clean: print_col = col
            elif 'EMB' in c_clean or 'SEQUIN' in c_clean: emb_col = col
            elif 'FWASH' in c_clean or 'F/WASH' in c_clean: fwash_col = col
            elif 'GWASH' in c_clean or 'G/WASH' in c_clean: gwash_col = col
            elif 'GDYE' in c_clean or 'G/DYE' in c_clean: gdye_col = col
            elif 'START' in c_clean: line_start_col = col
            elif 'INFAC' in c_clean: fabric_in_fac_col = col
            elif 'PPGTSAPPD' in c_clean or 'PPAPPD' in c_clean: pps_appd_col = col
            elif c_clean in ['1STS/D', 'EXF', 'SD', '1STSD']: ex_factory_col = col
            elif c_clean in ['TOTALORDERQTY', 'QTY']: qty_col = col

        if style_col is None:
            continue
            
        sheet_rows = []
        for _, row in df.iterrows():
            style = str(row.get(style_col, '')).strip()
            if not style or style == 'nan' or style.upper().startswith('TOTAL'):
                continue
                
            # Graphic / Wash 자동 판정
            has_graphic = 'X'
            if print_col and pd.notna(row.get(print_col)) and str(row.get(print_col)).strip() not in ['', 'nan', 'X', 'x']: has_graphic = 'O'
            elif emb_col and pd.notna(row.get(emb_col)) and str(row.get(emb_col)).strip() not in ['', 'nan', 'X', 'x']: has_graphic = 'O'
                
            has_wash = 'X'
            if fwash_col and pd.notna(row.get(fwash_col)) and str(row.get(fwash_col)).strip() not in ['', 'nan', 'X', 'x']: has_wash = 'O'
            elif gwash_col and pd.notna(row.get(gwash_col)) and str(row.get(gwash_col)).strip() not in ['', 'nan', 'X', 'x']: has_wash = 'O'
            elif gdye_col and pd.notna(row.get(gdye_col)) and str(row.get(gdye_col)).strip() not in ['', 'nan', 'X', 'x']: has_wash = 'O'
            
            # 날짜 파싱 안전장치
            line_start_raw = row.get(line_start_col) if line_start_col else None
            if pd.isna(line_start_raw): continue
            try: line_start = pd.to_datetime(line_start_raw)
            except: continue
            
            days_buffer = 14 if (has_graphic == 'O' or has_wash == 'O') else 7
            fabric_due = line_start - timedelta(days=14)
            fpp_due = line_start - timedelta(days=14)
            pps_due = line_start - timedelta(days=days_buffer)
            
            fabric_in_fac = row.get(fabric_in_fac_col) if fabric_in_fac_col else None
            fabric_status = "🔴 Late" if pd.isna(fabric_in_fac) else "🟢 Ready"
            
            pp_appd_raw = str(row.get(pps_appd_col, '')).strip() if pps_appd_col else ""
            if pp_appd_raw.upper() in ['N/A']: pps_status = "⚪ N/A"
            elif pp_appd_raw.upper() in ['C/O']: pps_status = "⚪ C/O"
            elif pp_appd_raw == 'nan' or not pp_appd_raw: pps_status = "➖"
            else:
                try: pps_status = pd.to_datetime(pp_appd_raw).strftime('%m/%d')
                except: pps_status = pp_appd_raw
                
            risk = "🟢 Low"
            if fabric_status == "🔴 Late": risk = "🔴 High"
                
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
