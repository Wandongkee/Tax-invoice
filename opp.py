import streamlit as st
import pandas as pd
import os
import io

# 자동 경로 인식 
current_dir = os.path.dirname(os.path.abspath(__file__))

# -------------------------------------------------------------------
# 1. 공통 유틸리티 함수
# -------------------------------------------------------------------
def safe_date(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    try:
        dt = pd.to_datetime(str(val), errors='coerce')
        if pd.isna(dt):
            return ""
        return dt.strftime('%Y-%m-%d')
    except:
        return ""

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# -------------------------------------------------------------------
# 2. 핵심 대조 로직
# -------------------------------------------------------------------
def process_tax_invoices(hometax_file, erp_file, is_sales=True):
    if is_sales:
        ht_biz_col = '공급받는자사업자등록번호'
        ht_name_col = '상호.1' 
        prefix = "매출_"
    else:
        ht_biz_col = '공급자사업자등록번호'
        ht_name_col = '상호'
        prefix = "매입_"

    df_ht = pd.read_excel(hometax_file, skiprows=5)
    df_erp = pd.read_excel(erp_file, skiprows=1)
    
    if ht_name_col not in df_ht.columns:
        ht_name_col = '상호'

    clean_biz = lambda x: str(x).replace("-", "").strip() if pd.notna(x) else ""
    clean_amt = lambda x: float(str(x).replace(",", "")) if pd.notna(x) and str(x).strip() != "" else 0

    df_ht['비교_사업자번호'] = df_ht[ht_biz_col].apply(clean_biz)
    df_ht['비교_공급가액'] = df_ht['공급가액'].apply(clean_amt)
    df_ht['비교_세액'] = df_ht['세액'].apply(clean_amt)
    df_ht['비교_작성일자'] = df_ht['작성일자'].apply(safe_date)
    df_ht['전산대조결과'] = ""
    df_ht['Matched'] = False
    
    df_erp['비교_사업자번호'] = df_erp['사업자등록번호'].apply(clean_biz)
    df_erp['비교_공급가액'] = df_erp['공급가액'].apply(clean_amt)
    df_erp['비교_세액'] = df_erp['세액'].apply(clean_amt)
    df_erp['비교_발생일자'] = df_erp['발생일자'].apply(safe_date)
    df_erp['Matched'] = False
    
    valid_erp = df_erp[df_erp['비교_사업자번호'] != ''].copy()
    wrong_invoices = []

    # [Step 1] 완벽 일치
    for ht_idx, ht_row in df_ht.iterrows():
        if not ht_row['비교_사업자번호']:
            df_ht.at[ht_idx, '전산대조결과'] = "비교제외"
            df_ht.at[ht_idx, 'Matched'] = True
            continue
            
        match_mask = (
            (valid_erp['비교_사업자번호'] == ht_row['비교_사업자번호']) &
            (valid_erp['비교_발생일자'] == ht_row['비교_작성일자']) &
            (valid_erp['비교_공급가액'] == ht_row['비교_공급가액']) &
            (valid_erp['비교_세액'] == ht_row['비교_세액']) &
            (~valid_erp['Matched'])
        )
        matching_indices = valid_erp[match_mask].index
        
        if len(matching_indices) > 0:
            erp_idx = matching_indices[0]
            valid_erp.at[erp_idx, 'Matched'] = True
            df_ht.at[ht_idx, 'Matched'] = True
            df_ht.at[ht_idx, '전산대조결과'] = "정상(일치)"

    # [Step 2] 작성일자 오류
    for ht_idx, ht_row in df_ht[~df_ht['Matched']].iterrows():
        match_mask = (
            (valid_erp['비교_사업자번호'] == ht_row['비교_사업자번호']) &
            (valid_erp['비교_공급가액'] == ht_row['비교_공급가액']) &
            (valid_erp['비교_세액'] == ht_row['비교_세액']) &
            (~valid_erp['Matched'])
        )
        matching_indices = valid_erp[match_mask].index
        
        if len(matching_indices) > 0:
            erp_idx = matching_indices[0]
            valid_erp.at[erp_idx, 'Matched'] = True
            df_ht.at[ht_idx, 'Matched'] = True
            df_ht.at[ht_idx, '전산대조결과'] = "🚨 틀린세금계산서(작성일자 오류)"
            
            wrong_invoices.append({
                '오류유형': '작성일자 오류',
                '사업자번호': ht_row[ht_biz_col],
                '상호': ht_row[ht_name_col],
                '홈텍스_작성일자': ht_row['작성일자'],
                '전산_발생일자': valid_erp.at[erp_idx, '발생일자'],
                '홈텍스_공급가액': ht_row['공급가액'],
                '전산_공급가액': valid_erp.at[erp_idx, '공급가액'],
                '홈텍스_세액': ht_row['세액'],
                '전산_세액': valid_erp.at[erp_idx, '세액'],
                '참고(전산_전표번호)': valid_erp.at[erp_idx, '전표번호'] if '전표번호' in valid_erp.columns else '',
                '참고(전산_적요)': valid_erp.at[erp_idx, '적요'] if '적요' in valid_erp.columns else ''
            })

    # [Step 3] 금액/세액 오류
    for ht_idx, ht_row in df_ht[~df_ht['Matched']].iterrows():
        match_mask = (
            (valid_erp['비교_사업자번호'] == ht_row['비교_사업자번호']) &
            (valid_erp['비교_발생일자'] == ht_row['비교_작성일자']) &
            (~valid_erp['Matched'])
        )
        matching_indices = valid_erp[match_mask].index
        
        if len(matching_indices) > 0:
            erp_idx = matching_indices[0]
            valid_erp.at[erp_idx, 'Matched'] = True
            df_ht.at[ht_idx, 'Matched'] = True
            df_ht.at[ht_idx, '전산대조결과'] = "🚨 틀린세금계산서(금액/세액 오류)"
            
            wrong_invoices.append({
                '오류유형': '금액/세액 오류',
                '사업자번호': ht_row[ht_biz_col],
                '상호': ht_row[ht_name_col],
                '홈텍스_작성일자': ht_row['작성일자'],
                '전산_발생일자': valid_erp.at[erp_idx, '발생일자'],
                '홈텍스_공급가액': ht_row['공급가액'],
                '전산_공급가액': valid_erp.at[erp_idx, '공급가액'],
                '홈텍스_세액': ht_row['세액'],
                '전산_세액': valid_erp.at[erp_idx, '세액'],
                '참고(전산_전표번호)': valid_erp.at[erp_idx, '전표번호'] if '전표번호' in valid_erp.columns else '',
                '참고(전산_적요)': valid_erp.at[erp_idx, '적요'] if '적요' in valid_erp.columns else ''
            })

    # [Step 4] 전산 누락
    for ht_idx, ht_row in df_ht[~df_ht['Matched']].iterrows():
        if ht_row['전산대조결과'] == "":
            df_ht.at[ht_idx, '전산대조결과'] = "❌ 전산에 빠짐(누락)"

    # [Step 5] 종이세금계산서 의심
    df_paper = valid_erp[~valid_erp['Matched']].copy()
    
    cols_to_drop = ['비교_사업자번호', '비교_공급가액', '비교_세액', '비교_작성일자', 'Matched']
    df_ht.drop(columns=cols_to_drop, inplace=True, errors='ignore')
    df_paper.drop(columns=cols_to_drop, inplace=True, errors='ignore')
    
    cols = df_ht.columns.tolist()
    if '전산대조결과' in cols:
        cols.remove('전산대조결과')
        cols = ['전산대조결과'] + cols
        df_ht = df_ht[cols]

    results = {
        'ht_result': to_excel_bytes(df_ht),
        'paper_result': to_excel_bytes(df_paper),
        'wrong_invoices': to_excel_bytes(pd.DataFrame(wrong_invoices)) if wrong_invoices else None,
        'prefix': prefix,
        'wrong_count': len(wrong_invoices)
    }
    
    return results

# -------------------------------------------------------------------
# 3. Streamlit 웹앱 UI 구성 (세션 상태 적용)
# -------------------------------------------------------------------
st.set_page_config(page_title="세금계산서 대조 시스템", layout="wide")
st.title("📑 세금계산서 전산/홈택스 대조 앱")

# 세션 상태(저장소) 초기화
if 'sales_results' not in st.session_state:
    st.session_state['sales_results'] = None
if 'purc_results' not in st.session_state:
    st.session_state['purc_results'] = None

tab1, tab2 = st.tabs(["🔵 매출 세금계산서 대조", "🔴 매입 세금계산서 대조"])

# --- 매출 탭 ---
with tab1:
    st.subheader("매출 데이터 업로드")
    col1, col2 = st.columns(2)
    with col1:
        ht_file_sales = st.file_uploader("홈택스 엑셀 업로드 (매출)", type=['xls', 'xlsx'], key='ht_sales')
    with col2:
        erp_file_sales = st.file_uploader("전산 엑셀 업로드 (매출)", type=['xls', 'xlsx'], key='erp_sales')

    # 대조 버튼을 누르면 세션 상태에 결과 저장
    if ht_file_sales and erp_file_sales:
        if st.button("매출 데이터 대조 시작", key='btn_sales'):
            with st.spinner("분석 중입니다..."):
                st.session_state['sales_results'] = process_tax_invoices(ht_file_sales, erp_file_sales, is_sales=True)
            st.success("✨ 매출 데이터 분석이 완료되었습니다!")

    # 세션 상태에 결과가 존재하면 다운로드 버튼 활성화 유지
    if st.session_state['sales_results'] is not None:
        res = st.session_state['sales_results']
        st.download_button("📥 1. 홈택스 원본 대조결과 다운로드", data=res['ht_result'], 
                           file_name=f"1_{res['prefix']}홈택스_대조완료.xlsx", mime="application/vnd.ms-excel", key='dl_sales_1')
        st.download_button("📥 2. 종이세금계산서 의심목록 다운로드", data=res['paper_result'], 
                           file_name=f"2_{res['prefix']}종이세금계산서_의심.xlsx", mime="application/vnd.ms-excel", key='dl_sales_2')
        
        if res['wrong_invoices']:
            st.warning(f"🚨 오입력 의심 건수: {res['wrong_count']}건 발견됨")
            st.download_button("📥 3. 틀린세금계산서 상세내역 다운로드", data=res['wrong_invoices'], 
                               file_name=f"3_{res['prefix']}틀린세금계산서_상세내역.xlsx", mime="application/vnd.ms-excel", key='dl_sales_3')
        else:
            st.info("👉 틀리게 입력된 세금계산서가 없습니다!")

# --- 매입 탭 ---
with tab2:
    st.subheader("매입 데이터 업로드")
    col1, col2 = st.columns(2)
    with col1:
        ht_file_purc = st.file_uploader("홈택스 엑셀 업로드 (매입)", type=['xls', 'xlsx'], key='ht_purc')
    with col2:
        erp_file_purc = st.file_uploader("전산 엑셀 업로드 (매입)", type=['xls', 'xlsx'], key='erp_purc')

    if ht_file_purc and erp_file_purc:
        if st.button("매입 데이터 대조 시작", key='btn_purc'):
            with st.spinner("분석 중입니다..."):
                st.session_state['purc_results'] = process_tax_invoices(ht_file_purc, erp_file_purc, is_sales=False)
            st.success("✨ 매입 데이터 분석이 완료되었습니다!")

    if st.session_state['purc_results'] is not None:
        res = st.session_state['purc_results']
        st.download_button("📥 1. 홈택스 원본 대조결과 다운로드", data=res['ht_result'], 
                           file_name=f"1_{res['prefix']}홈택스_대조완료.xlsx", mime="application/vnd.ms-excel", key='dl_purc_1')
        st.download_button("📥 2. 종이세금계산서 의심목록 다운로드", data=res['paper_result'], 
                           file_name=f"2_{res['prefix']}종이세금계산서_의심.xlsx", mime="application/vnd.ms-excel", key='dl_purc_2')
        
        if res['wrong_invoices']:
            st.warning(f"🚨 오입력 의심 건수: {res['wrong_count']}건 발견됨")
            st.download_button("📥 3. 틀린세금계산서 상세내역 다운로드", data=res['wrong_invoices'], 
                               file_name=f"3_{res['prefix']}틀린세금계산서_상세내역.xlsx", mime="application/vnd.ms-excel", key='dl_purc_3')
        else:
            st.info("👉 틀리게 입력된 세금계산서가 없습니다!")
