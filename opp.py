import streamlit as st
import pandas as pd
import os
import io
import hashlib

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


def clean_biz_no(x):
    return str(x).replace("-", "").strip() if pd.notna(x) else ""


def to_excel_workbook(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def check_business_number_mismatch(hometax_file, erp_file):
    """홈택스 파일 상단의 사업자등록번호와 전산 파일의 사업장등록번호가
    일치하는지 확인해서, 사업장이 잘못 짝지어진 파일 쌍을 사전에 잡아낸다.
    확인이 불가능한 경우(양식이 다르거나 컬럼이 없는 경우)는 None을 반환해
    기존 처리 흐름을 막지 않는다."""
    try:
        hometax_file.seek(0)
        df_head = pd.read_excel(hometax_file, header=None, nrows=1)
        hometax_file.seek(0)
        if df_head.shape[1] < 2:
            return None
        ht_biz_no_raw = df_head.iat[0, 1]
        ht_biz_no = clean_biz_no(ht_biz_no_raw)
    except Exception:
        return None

    try:
        erp_file.seek(0)
        df_erp_head = pd.read_excel(erp_file, skiprows=1, usecols=['사업장등록번호'])
        erp_file.seek(0)
        erp_biz_nos_raw = df_erp_head['사업장등록번호'].dropna().unique().tolist()
        erp_biz_nos_clean = {clean_biz_no(v) for v in erp_biz_nos_raw}
        erp_biz_nos_clean.discard("")
    except Exception:
        return None

    if not ht_biz_no or not erp_biz_nos_clean:
        return None

    if ht_biz_no not in erp_biz_nos_clean:
        return {
            'hometax_biz_no': ht_biz_no_raw,
            'erp_biz_nos': erp_biz_nos_raw
        }
    return None


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
            (valid_erp['비교_발생일자'] != '') &
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


    # 정확히 일치하는 전표를 우선 보존한 뒤, 남은 +/- 자료만 상쇄한다.
    # 같은 거래처, 같은 유효 날짜, 공급가액/세액 모두 반대인 경우에 한정한다.
    # 이는 대조상 순액 처리이며 법적인 취소 관계를 확정하는 것은 아니다.
    def mark_offset_pairs(frame, date_col, source_name):
        pending = {}
        audit_rows = []
        pair_count = 0
        for idx, row in frame.loc[~frame['Matched']].iterrows():
            biz = row['비교_사업자번호']
            date = row[date_col]
            amount = row['비교_공급가액']
            tax = row['비교_세액']
            if not biz or not date or not pd.notna(amount) or not pd.notna(tax):
                continue
            if (amount == 0 and tax == 0) or amount * tax < 0:
                continue
            key = (biz, date, amount, tax)
            opposite = (biz, date, -amount, -tax)
            candidates = pending.get(opposite)
            if candidates:
                other_idx = candidates.pop(0)
                pair_count += 1
                pair_id = f"{source_name}-{pair_count}"
                for pair_idx in (other_idx, idx):
                    frame.at[pair_idx, 'Matched'] = True
                    if source_name == '홈택스':
                        frame.at[pair_idx, '전산대조결과'] = "상쇄(동일 거래처·작성일자, 공급가액·세액 합계 0)"
                    audit = frame.loc[pair_idx].to_dict()
                    audit.update({
                        '자료구분': source_name,
                        '상쇄그룹': pair_id,
                        '원본엑셀행': int(pair_idx) + (7 if source_name == '홈택스' else 3),
                        '처리내용': '대조상 상쇄: 원본 보존, 오류·누락·중복 의심에서 제외'
                    })
                    audit_rows.append(audit)
            else:
                pending.setdefault(key, []).append(idx)
        return audit_rows, pair_count

    ht_offsets, ht_offset_count = mark_offset_pairs(df_ht, '비교_작성일자', '홈택스')
    erp_offsets, erp_offset_count = mark_offset_pairs(valid_erp, '비교_발생일자', '전산')


    # 날짜 오류 후보와 금액 오류 후보를 함께 구성한다.
    # 양쪽 모두 후보가 하나일 때만 연결하여 원본 행 순서에 따른 오연결을 막는다.
    candidates = {}
    reverse_candidates = {}
    for ht_idx, ht_row in df_ht.loc[~df_ht['Matched']].iterrows():
        same_biz = valid_erp['비교_사업자번호'].eq(ht_row['비교_사업자번호'])
        same_amount = (
            valid_erp['비교_공급가액'].eq(ht_row['비교_공급가액']) &
            valid_erp['비교_세액'].eq(ht_row['비교_세액'])
        )
        same_date = (
            valid_erp['비교_발생일자'].eq(ht_row['비교_작성일자']) &
            valid_erp['비교_발생일자'].ne('') &
            bool(ht_row['비교_작성일자'])
        )
        indices = valid_erp.index[same_biz & (same_amount | same_date) & ~valid_erp['Matched']].tolist()
        candidates[ht_idx] = indices
        for erp_idx in indices:
            reverse_candidates.setdefault(erp_idx, []).append(ht_idx)

    review_erp = set()
    review_rows = []
    for ht_idx, indices in candidates.items():
        if not indices:
            df_ht.at[ht_idx, '전산대조결과'] = "❌ 전산에 빠짐(누락)"
            continue
        ht_row = df_ht.loc[ht_idx]
        if len(indices) != 1 or len(reverse_candidates[indices[0]]) != 1:
            df_ht.at[ht_idx, '전산대조결과'] = "🔎 확인 필요(비교 후보 복수)"
            review_erp.update(indices)
            for erp_idx in indices:
                erp_row = valid_erp.loc[erp_idx]
                review_rows.append({
                    '사유': '후보가 여러 건이므로 자동 연결하지 않음',
                    '홈택스_원본엑셀행': int(ht_idx) + 7,
                    '전산_원본엑셀행': int(erp_idx) + 3,
                    '사업자번호': ht_row[ht_biz_col],
                    '상호': ht_row[ht_name_col],
                    '홈택스_작성일자': ht_row['작성일자'],
                    '전산_발생일자': erp_row['발생일자'],
                    '홈택스_공급가액': ht_row['공급가액'],
                    '전산_공급가액': erp_row['공급가액'],
                    '홈택스_세액': ht_row['세액'],
                    '전산_세액': erp_row['세액'],
                    '전산_전표번호': erp_row.get('전표번호', ''),
                    '전산_작성부서': erp_row.get('작성부서', ''),
                    '전산_작성사원': erp_row.get('작성사원', '')
                })
            continue

        erp_idx = indices[0]
        erp_row = valid_erp.loc[erp_idx]
        error_type = ('작성일자 오류' if
                      not ht_row['비교_작성일자'] or not erp_row['비교_발생일자'] or
                      ht_row['비교_작성일자'] != erp_row['비교_발생일자']
                      else '금액/세액 오류')
        valid_erp.at[erp_idx, 'Matched'] = True
        df_ht.at[ht_idx, 'Matched'] = True
        df_ht.at[ht_idx, '전산대조결과'] = f"🚨 틀린세금계산서({error_type})"
        wrong_invoices.append({
            '오류유형': error_type,
            '사업자번호': ht_row[ht_biz_col],
            '상호': ht_row[ht_name_col],
            '작성부서': erp_row.get('작성부서', ''),
            '작성사원': erp_row.get('작성사원', ''),
            '홈택스_작성일자': ht_row['작성일자'],
            '전산_발생일자': erp_row['발생일자'],
            '홈택스_공급가액': ht_row['공급가액'],
            '전산_공급가액': erp_row['공급가액'],
            '홈택스_세액': ht_row['세액'],
            '전산_세액': erp_row['세액'],
            '참고(전산_전표번호)': erp_row.get('전표번호', ''),
            '참고(전산_적요)': erp_row.get('적요', '')
        })

    # 남은 전산 자료: 복수 후보 > 실제 동일 항목 중복 > 일반 미대조 순서.
    df_paper = valid_erp.loc[~valid_erp['Matched']].copy()
    ht_biz_numbers = set(df_ht.loc[df_ht['비교_사업자번호'] != '', '비교_사업자번호'])

    def invoice_key(row, date_col):
        return (row['비교_사업자번호'], row[date_col],
                row['비교_공급가액'], row['비교_세액'])

    exact_keys = {
        invoice_key(row, '비교_작성일자')
        for _, row in df_ht.loc[df_ht['전산대조결과'].eq('정상(일치)')].iterrows()
        if row['비교_작성일자']
    }
    remaining_key_counts = {}
    for idx, row in df_paper.iterrows():
        if idx not in review_erp and row['비교_발생일자']:
            key = invoice_key(row, '비교_발생일자')
            remaining_key_counts[key] = remaining_key_counts.get(key, 0) + 1

    def classify_unmatched(row):
        if row.name in review_erp:
            return "🔎 확인 필요(비교 후보 복수)"
        key = invoice_key(row, '비교_발생일자')
        if row['비교_발생일자'] and (key in exact_keys or remaining_key_counts.get(key, 0) > 1):
            return "🔁 중복의심(사업자·날짜·공급가액·세액 동일)"
        if row['비교_사업자번호'] in ht_biz_numbers:
            return "🔎 확인 필요(홈택스 대응 자료 없음)"
        return "📄 종이세금계산서 의심"

    df_paper.insert(0, '분류결과', [
        classify_unmatched(row) for _, row in df_paper.iterrows()
    ])

    statuses = df_ht['전산대조결과']
    erp_statuses = df_paper['분류결과'].astype('string')
    counts = {
        '정상 일치': int(statuses.eq('정상(일치)').sum()),
        '전산 누락': int(statuses.str.contains('누락', regex=False).sum()),
        '날짜·금액 오류': len(wrong_invoices),
        '상쇄 처리(쌍)': ht_offset_count + erp_offset_count,
        '홈택스 확인 필요': int(statuses.str.startswith('🔎').sum()),
        '전산 확인 필요': int(erp_statuses.str.startswith('🔎').sum()),
        '전산 중복 의심': int(erp_statuses.str.startswith('🔁').sum()),
        '종이계산서 의심': int(erp_statuses.str.startswith('📄').sum()),
        '홈택스 비교 제외': int(statuses.eq('비교제외').sum()),
        '전산 비교 제외': int(df_erp['비교_사업자번호'].eq('').sum()),
        '홈택스 상쇄(행)': ht_offset_count * 2,
        '전산 상쇄(행)': erp_offset_count * 2,
        '홈택스 원본(행)': len(df_ht),
        '전산 원본(행)': len(df_erp)
    }
    # 비교 제외 자료도 다운로드 결과에 보존한다.
    df_excluded = df_erp.loc[df_erp['비교_사업자번호'].eq('')].copy()
    df_excluded.insert(0, '분류결과', '비교제외(사업자번호 없음)')
    df_paper = pd.concat([df_paper, df_excluded], ignore_index=True)

    cols_to_drop = ['비교_사업자번호', '비교_공급가액', '비교_세액', '비교_작성일자', '비교_발생일자', 'Matched']
    df_ht.drop(columns=cols_to_drop, inplace=True, errors='ignore')
    df_paper.drop(columns=cols_to_drop, inplace=True, errors='ignore')

    cols = df_ht.columns.tolist()
    if '전산대조결과' in cols:
        cols.remove('전산대조결과')
        cols = ['전산대조결과'] + cols
        df_ht = df_ht[cols]

    df_wrong = pd.DataFrame(wrong_invoices)
    df_offsets = pd.DataFrame(ht_offsets + erp_offsets)
    df_offsets.drop(columns=cols_to_drop, inplace=True, errors='ignore')
    combined_result = to_excel_workbook({
        '1_홈택스_대조결과': df_ht,
        '2_종이세금계산서_의심': df_paper,
        '3_틀린세금계산서_상세': df_wrong,
        '4_상쇄처리_내역': df_offsets,
        '5_확인필요_후보': pd.DataFrame(review_rows),
        '6_대조결과_요약': pd.DataFrame(list(counts.items()), columns=['구분', '건수'])
    })

    results = {
        'combined_result': combined_result,
        'prefix': prefix,
        'wrong_count': len(wrong_invoices),
        'offset_pair_count': ht_offset_count + erp_offset_count,
        'counts': counts
    }

    return results


def sync_uploaded_files(state, session_key, hometax_file, erp_file):
    def fingerprint(upload):
        if upload is None:
            return None
        return (upload.name, hashlib.sha256(upload.getvalue()).hexdigest())
    identity = (fingerprint(hometax_file), fingerprint(erp_file))
    identity_key = f"{session_key}_files"
    if state.get(identity_key) != identity:
        state[session_key] = None
        state[identity_key] = identity


def render_invoice_section(title, hometax_label, erp_label, button_label, session_key, is_sales, uploader_suffix):
    st.subheader(title)
    col1, col2 = st.columns(2)
    with col1:
        hometax_file = st.file_uploader(hometax_label, type=['xls', 'xlsx'], key=f'ht_{uploader_suffix}')
    with col2:
        erp_file = st.file_uploader(erp_label, type=['xls', 'xlsx'], key=f'erp_{uploader_suffix}')

    sync_uploaded_files(st.session_state, session_key, hometax_file, erp_file)

    if hometax_file and erp_file:
        mismatch = check_business_number_mismatch(hometax_file, erp_file)
        if mismatch:
            erp_list = ', '.join(str(v) for v in mismatch['erp_biz_nos'])
            st.error(
                f"⚠️ 사업자번호 불일치 의심: 업로드하신 홈택스 파일은 "
                f"**{mismatch['hometax_biz_no']}** 사업자 목록인데, "
                f"전산 파일의 사업장등록번호({erp_list})와 일치하지 않습니다. "
                f"서로 다른 사업장(서울/여주) 파일이 잘못 짝지어진 건 아닌지 확인해주세요."
            )

        if st.button(button_label, key=f'btn_{uploader_suffix}'):
            st.session_state[session_key] = None
            with st.spinner("분석 중입니다..."):
                st.session_state[session_key] = process_tax_invoices(hometax_file, erp_file, is_sales=is_sales)
            st.success(f"✨ {title.replace(' 업로드', '')} 분석이 완료되었습니다!")

    if st.session_state[session_key] is not None:
        res = st.session_state[session_key]
        counts = res['counts']
        labels = list(counts)[:8]
        for offset in (0, 4):
            columns = st.columns(4)
            for column, label in zip(columns, labels[offset:offset + 4]):
                column.metric(label, counts[label])
        st.caption(
            "정상·오류는 연결된 건수, 상쇄는 2행을 1쌍으로 계산합니다. "
            "확인 필요는 자료별 행 수이므로 홈택스·전산 건수를 합쳐 사건 수로 보지 마세요. "
            "같은 사업자번호·날짜·금액·세액이라도 별도 거래일 수 있어 중복은 의심으로 표시합니다."
        )
        attention_labels = ['전산 누락', '날짜·금액 오류', '홈택스 확인 필요',
                            '전산 확인 필요', '전산 중복 의심', '종이계산서 의심',
                            '홈택스 비교 제외', '전산 비교 제외']
        if any(counts[label] for label in attention_labels):
            st.warning("확인이 필요한 자료가 있습니다. 결과 엑셀의 분류와 후보 내역을 확인해주세요.")
        elif counts['홈택스 원본(행)'] == 0 and counts['전산 원본(행)'] == 0:
            st.info("비교할 자료가 없습니다.")
        else:
            st.success("대조가 완료되었습니다. 정상 일치 또는 상쇄 처리되었으며 추가 확인 항목은 없습니다.")
        if counts['홈택스 비교 제외'] or counts['전산 비교 제외']:
            st.info(f"사업자번호가 없어 비교 제외: 홈택스 {counts['홈택스 비교 제외']}행 / 전산 {counts['전산 비교 제외']}행")

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

