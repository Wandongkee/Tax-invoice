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


def to_excel_workbook(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
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
                '홈택스_작성일자': ht_row['작성일자'],
                '전산_발생일자': valid_erp.at[erp_idx, '발생일자'],
                '홈택스_공급가액': ht_row['공급가액'],
                '전산_공급가액': valid_erp.at[erp_idx, '공급가액'],
                '홈택스_세액': ht_row['세액'],
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
                '홈택스_작성일자': ht_row['작성일자'],
                '전산_발생일자': valid_erp.at[erp_idx, '발생일자'],
                '홈택스_공급가액': ht_row['공급가액'],
                '전산_공급가액': valid_erp.at[erp_idx, '공급가액'],
                '홈택스_세액': ht_row['세액'],
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

    cols_to_drop = ['비교_사업자번호', '비교_공급가액', '비교_세액', '비교_작성일자', '비교_발생일자', 'Matched']
    df_ht.drop(columns=cols_to_drop, inplace=True, errors='ignore')
    df_paper.drop(columns=cols_to_drop, inplace=True, errors='ignore')

    cols = df_ht.columns.tolist()
    if '전산대조결과' in cols:
        cols.remove('전산대조결과')
        cols = ['전산대조결과'] + cols
        df_ht = df_ht[cols]

    df_wrong = pd.DataFrame(wrong_invoices)
    combined_result = to_excel_workbook({
        '1_홈택스_대조결과': df_ht,
        '2_종이세금계산서_의심': df_paper,
        '3_틀린세금계산서_상세': df_wrong
    })

    results = {
        'combined_result': combined_result,
        'prefix': prefix,
        'wrong_count': len(wrong_invoices)
    }

    return results


def render_invoice_section(title, hometax_label, erp_label, button_label, session_key, is_sales, uploader_suffix):
    st.subheader(title)
    col1, col2 = st.columns(2)
    with col1:
        hometax_file = st.file_uploader(hometax_label, type=['xls', 'xlsx'], key=f'ht_{uploader_suffix}')
    with col2:
        erp_file = st.file_uploader(erp_label, type=['xls', 'xlsx'], key=f'erp_{uploader_suffix}')

    if hometax_file and erp_file:
        if st.button(button_label, key=f'btn_{uploader_suffix}'):
            with st.spinner("분석 중입니다..."):
                st.session_state[session_key] = process_tax_invoices(hometax_file, erp_file, is_sales=is_sales)
            st.success(f"✨ {title.replace(' 업로드', '')} 분석이 완료되었습니다!")

    if st.session_state[session_key] is not None:
        res = st.session_state[session_key]
        if res['wrong_count'] > 0:
            st.warning(f"🚨 오입력 의심 건수: {res['wrong_count']}건 발견됨")
        else:
            st.info("👉 틀리게 입력된 세금계산서가 없습니다!")

        st.download_button(
            "📥 통합 대조결과 다운로드",
            data=res['combined_result'],
            file_name=f"{res['prefix']}세금계산서_통합대조결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f'dl_{uploader_suffix}_combined'
        )


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

render_invoice_section(
    title="매출 데이터 업로드",
    hometax_label="홈택스 엑셀 업로드 (매출)",
    erp_label="전산 엑셀 업로드 (매출)",
    button_label="매출 데이터 대조 시작",
    session_key='sales_results',
    is_sales=True,
    uploader_suffix='sales'
)

st.divider()

render_invoice_section(
    title="매입 데이터 업로드",
    hometax_label="홈택스 엑셀 업로드 (매입)",
    erp_label="전산 엑셀 업로드 (매입)",
    button_label="매입 데이터 대조 시작",
    session_key='purc_results',
    is_sales=False,
    uploader_suffix='purc'
)
