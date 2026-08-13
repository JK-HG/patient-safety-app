import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="환자안전지표 모니터링 분석기", page_icon="🏥", layout="wide"
)

st.title("🏥 환자안전지표 현장모니터링 결과 분석기")
st.write(
    "다운로드한 환자확인율 결과표 엑셀 파일(.xlsx)을 업로드하면 보고서 양식에 맞는 통계를 자동으로 산출합니다."
)

uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=["xlsx", "xls"])


def map_job(job):
    """직종 -> 보고서용 직군 매핑"""
    if pd.isna(job):
        return "기타"
    job = str(job).strip()
    if job == "간호사":
        return "간호"
    elif job == "의사":
        return "의사"
    elif job in ["의료기사", "방사선사", "임상병리사"]:
        return "진료지원"
    elif job in ["사원", "행정원"]:
        return "행정"
    return "기타"


def map_context(item):
    """상황별 카테고리 매핑"""
    if pd.isna(item):
        return "6. 그 외 항목"

    item = str(item).strip()

    # 1. 입원/외래진료/수납
    if item in ["진료 전", "기 타 - 수납 시", "기 타 - 입원 시"]:
        return "1. 입원/외래진료/수납"
    # 2. 투약
    elif item == "의약품 투여 전":
        return "2. 투약"
    # 3. 검사/검체 채취
    elif item.startswith("검사 시행 전"):
        return "3. 검사/검체 채취"
    # 4. 처치/수술/수혈
    elif (
        item.startswith("처치 및 시술 전")
        or item.startswith("혈액제제 투여 전")
        or "수술전" in item
    ):
        return "4. 처치/수술/수혈"
    # 5. 물리치료
    elif item == "기 타 - 물리치료":
        return "5. 물리치료"
    # 6. 그 외 항목
    else:
        return "6. 그 외 항목"


def normalize_dept_name(dept):
    """부서명 단일화 정규화 (일반내과->내과, 물리치료팀/물리치료실->물리치료)"""
    if pd.isna(dept):
        return "기타"
    dept = str(dept).strip()
    if dept in ["일반내과", "내과"]:
        return "내과"
    elif dept in ["물리치료팀", "물리치료실"]:
        return "물리치료"
    return dept


def map_dept(dept):
    """부서별 카테고리 매핑"""
    if pd.isna(dept):
        return "5. 기타"

    dept = str(dept).strip()

    # 1. 진료과
    if dept in [
        "화상외과",
        "성형외과",
        "재활의학과",
        "내과",
        "일반내과",
        "소아청소년과",
    ]:
        return "1. 진료과"
    # 2. 간호부
    elif dept in [
        "외래",
        "신관3병동",
        "본관5병동",
        "본관6병동",
        "본관7병동",
        "본관8병동",
        "응급실",
        "화상중환자실",
    ]:
        return "2. 간호부"
    # 3. 진료지원
    elif dept in ["물리치료실", "물리치료팀", "물리치료", "영상의학과"]:
        return "3. 진료지원"
    # 4. 행정
    elif dept in ["수납계", "원무계"]:
        return "4. 행정"
    else:
        return "5. 기타"


def process_excel(df_raw):
    # 상단 텍스트 행 제외하고 'NO' 열 위치를 기준으로 헤더 자동 인식
    if "NO" not in df_raw.columns:
        for idx, row in df_raw.iterrows():
            if "NO" in row.values:
                df_raw.columns = df_raw.iloc[idx]
                df_raw = df_raw.iloc[idx + 1 :].reset_index(drop=True)
                break

    # 유효한 데이터 행만 필터링
    df = df_raw.dropna(subset=["NO"]).copy()

    # 1. 1차 환자성명 확인: '이름확인'이 '시행'
    df["1차확인_성공"] = df["이름확인"].astype(str).str.strip() == "시행"

    # 2. 2차 확인: '등록번호 확인' 또는 '생년월일 확인' 중 하나라도 '시행'
    df["2차확인_성공"] = (
        df["등록번호 확인"].astype(str).str.strip() == "시행"
    ) | (df["생년월일 확인"].astype(str).str.strip() == "시행")

    # 3. 정확한 환자확인: 1차와 2차를 모두 시행한 경우
    df["정확한확인_성공"] = df["1차확인_성공"] & df["2차확인_성공"]

    # 파생 변수 생성
    df["직군"] = df["직종"].apply(map_job)
    df["상황"] = df["환자확인사항"].apply(map_context)

    # 부서명 정규화 적용
    raw_dept = df["상세부서명"].fillna(df["부서"])
    df["상세부서"] = raw_dept.apply(normalize_dept_name)
    df["부서"] = df["상세부서"].apply(map_dept)

    return df


def calc_stats(df, group_col):
    stats = (
        df.groupby(group_col)
        .agg(
            전체건수=("정확한확인_성공", "count"),
            차1확인_성공=("1차확인_성공", "sum"),
            차2확인_성공=("2차확인_성공", "sum"),
            정확한확인_성공=("정확한확인_성공", "sum"),
        )
        .reset_index()
    )

    # 지정 번호순 정렬
    if group_col in ["상황", "부서"]:
        stats = stats.sort_values(by=group_col).reset_index(drop=True)

    stats["1차확인 비율"] = (
        (stats["차1확인_성공"] / stats["전체건수"] * 100).round(1).astype(str) + "%"
    )
    stats["2차확인 비율"] = (
        (stats["차2확인_성공"] / stats["전체건수"] * 100).round(1).astype(str) + "%"
    )
    stats["정확한확인 비율"] = (
        (stats["정확한확인_성공"] / stats["전체건수"] * 100)
        .round(1)
        .astype(str)
        + "%"
    )

    result_df = pd.DataFrame()
    result_df[group_col] = (
        stats[group_col].astype(str) + " (" + stats["전체건수"].astype(str) + "건)"
    )
    result_df["1차 환자성명 확인"] = (
        stats["차1확인_성공"].astype(str) + "건 (" + stats["1차확인 비율"] + ")"
    )
    result_df["2차 등록번호 확인"] = (
        stats["차2확인_성공"].astype(str) + "건 (" + stats["2차확인 비율"] + ")"
    )
    result_df["정확한 환자확인"] = (
        stats["정확한확인_성공"].astype(str)
        + "건 ("
        + stats["정확한확인 비율"]
        + ")"
    )

    return result_df, stats


if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = process_excel(df_raw)

        # 총괄 수치
        total_count = len(df)
        p1_count = df["1차확인_성공"].sum()
        p2_count = df["2차확인_성공"].sum()
        final_count = df["정확한확인_성공"].sum()
        final_rate = (
            round((final_count / total_count) * 100, 1) if total_count > 0 else 0
        )

        st.subheader("📊 1. 모니터링 총괄 현황")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("총 점검 건수", f"{total_count} 건")
        c2.metric("1차 이름확인 건수", f"{p1_count} 건")
        c3.metric("2차 등록/생년월일 확인", f"{p2_count} 건")
        c4.metric(
            "최종 정확한 환자확인",
            f"{final_count} 건",
            delta=f"이행률 {final_rate}%",
        )

        st.divider()

        # 1. 직군별
        st.subheader("👥 1) 직군별 정확한 환자 확인")
        res_job, _ = calc_stats(df, "직군")
        st.dataframe(res_job, use_container_width=True)

        # 2. 상황별
        st.subheader("📋 2) 상황별 정확한 환자 확인")
        res_context, _ = calc_stats(df, "상황")
        st.dataframe(res_context, use_container_width=True)

        # 3. 부서별
        st.subheader("🏥 3) 부서별 카테고리 정확한 환자 확인")
        res_dept, _ = calc_stats(df, "부서")
        st.dataframe(res_dept, use_container_width=True)

        # 미시행 목록
        st.divider()
        st.subheader("⚠️ 정확한 환자확인 미시행 내역 (1차 또는 2차 미시행)")
        fail_df = df[~df["정확한확인_성공"]][
            [
                "NO",
                "부서",
                "상세부서",
                "직종",
                "이름",
                "환자확인사항",
                "이름확인",
                "등록번호 확인",
                "생년월일 확인",
                "미시행 사유 또는 기타사항",
            ]
        ]
        if len(fail_df) > 0:
            st.dataframe(fail_df, use_container_width=True)
        else:
            st.success("🎉 모든 건에서 정확한 환자확인이 수행되었습니다!")

    except Exception as e:
        st.error(f"엑셀 분석 중 오류가 발생했습니다: {e}")
