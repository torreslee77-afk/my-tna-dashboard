import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# 1. 페이지 기본 설정 및 디자인
st.set_page_config(page_title="YAKJIN TNA Ai Operational dashboard", page_icon="📊", layout="wide")

# (중략 - 디자인 코드는 동일)

# Helper: 키워드 매칭 함수
def clean_string(val):
    if pd.isna(val): return ""
    return str(val).strip().upper().replace(" ", "").replace("'", "").replace("#", "").replace("/", "").replace("(", "").replace(")", "").replace("-", "")

# 2. 엑셀 분석 엔진 함수
def analyze_tna(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    all_sheets_data = {}
    
    for sheet_name in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df_raw.empty: continue
            
        # 헤더 행 찾기
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_values = [clean_string(v) for v in row.values]
            if any('STYLE' in v for v in row_values if v):
                header_idx = idx
                break
        if header_idx is None: continue
            
        # 컬럼 병합 처리
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
                
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = combined_columns

        # --- [핵심] 컬럼 매칭 지정 ---
        style_col, qty_col = None, None
        
        for col in df.columns:
            c_clean = clean_string(col)
            if 'STYLE' in c_clean and '배정' not in c_clean: style_col = col
            # 수량 컬럼을 'GMTQTY'로 엄격하게 매칭
            if c_clean == 'GMTQTY': qty_col = col
            
            # (나머지 컬럼 매칭은 생략됨, 기존 코드 유지)
            # ...

        if style_col is None: continue

        # [중요] 병합된 수량(GMT QTY) 셀 전방 채우기(ffill)
        if qty_col:
            df[qty_col] = df[qty_col].replace('', None).ffill()
            
        # ... (이하 데이터 리스트 처리 로직은 기존과 동일)
