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


# 1. 직종 -> 직군 매핑
def map_job(job):
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


# 2. 상황 대분류 & 소분류 매핑
def map_context_detailed(item):
    if pd.isna(item):
        return "6. 그 외 항목", "기타"

    item = str(item).strip()

    if item == "진료 전":
        return "1. 입원/외래진료/수납", "진료 전"
    elif item == "기 타 - 수납 시":
        return "1. 입원/외래진료/수납", "수납 시"
    elif item == "기 타 - 입원 시":
        return "1. 입원/외래진료/수납", "입원 시"
    elif item == "의약품 투여 전":
        return "2. 투약", "의약품 투여 전"
    elif item.startswith("검사 시행 전"):
        sub_name = (
            item.replace("검사 시행 전 - ", "")
            if " - " in item
            else "검사 시행 전"
        )
        return "3. 검사/검체 채취", sub_name
    elif item.startswith("처치 및 시술 전"):
        return "4. 처치/수술/수혈", "처치 및 시술 전"
    elif item.startswith("혈액제제 투여 전"):
        return "4. 처치/수술/수혈", "혈액제제 투여 전"
    elif "수술전" in item:
        return "4. 처치/수술/수혈", "수술 전"
    elif item == "기 타 - 물리치료":
        return "5. 물리치료", "물리치료"
    else:
        return "6. 그 외 항목", item


# 3. 부서 대분류 & 소분류 매핑
def map_dept_detailed(dept):
    if pd.isna(dept):
        return "5. 기타", "기타"

    dept = str(dept).strip()

    # 명칭 정규화
    if dept == "일반내과":
        dept = "내과"
    elif dept == "물리치료팀":
        dept = "물리치료실"

    # 1. 진료과
    if dept in [
        "화상외과",
        "성형외과",
        "재활의학과",
        "내과",
        "소아청소년과",
    ]:
        return "1. 진료과", dept
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
        return "2. 간호부", dept
    # 3. 진료지원
    elif dept in ["물리치료실", "영상의학과"]:
        return "3. 진료지원", dept
    # 4. 행정
    elif dept in ["수납계", "원무계"]:
        return "4. 행정", dept
    else:
        return "5. 기타", dept


def process_excel(df_raw):
    # 헤더 자동 찾기
    if "NO" not in df_raw.columns:
        for idx, row in df_raw.iterrows():
            if "NO" in row.values:
                df_raw.columns = df_raw.iloc[idx]
                df_raw = df_raw.iloc[idx + 1 :].reset_index(drop=True)
                break

    df = df_raw.dropna(subset=["NO"]).copy()

    # 1차, 2차, 최종 성공 여부 판단
    df["1차확인_성공"] = df["이름확인"].astype(str).str.strip() == "시행"
    df["2차확인_성공"] = (
        df["등록번호 확인"].astype(str).str.strip() == "시행"
    ) | (df["생년월일 확인"].astype(str).str.strip() == "시행")
    df["정확한확인_성공"] = df["1차확인_성공"] & df["2차확인_성공"]

    # 파생 변수
    df["직군"] = df["직종"].apply(map_job)

    # 상황 분류
    context_res = df["환자확인사항"].apply(map_context_detailed)
    df["상황_대분류"] = [r[0] for r in context_res]
    df["상황_소분류"] = [r[1] for r in context_res]

    # 부서 분류
    raw_dept = df["상세부서명"].fillna(df["부서"])
    dept_res = raw_dept.apply(map_dept_detailed)
    df["부서_대분류"] = [r[0] for r in dept_res]
    df["부서_소분류"] = [r[1] for r in dept_res]

    return df


def calc_stats_flat(df, group_col):
    """단일 컬럼 집계 (직군용)"""
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

    return result_df


def calc_stats_hierarchy(df, main_col, sub_col, main_name):
    """대분류 & 소분류 계층형 집계 (부서 및 상황용)"""
    stats = (
        df.groupby([main_col, sub_col])
        .agg(
            전체건수=("정확한확인_성공", "count"),
            차1확인_성공=("1차확인_성공", "sum"),
            차2확인_성공=("2차확인_성공", "sum"),
            정확한확인_성공=("정확한확인_성공", "sum"),
        )
        .reset_index()
    )

    stats = stats.sort_values(by=[main_col, sub_col]).reset_index(drop=True)

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
    result_df[main_name + " 구분"] = stats[main_col]
    result_df["세부 항목(소분류)"] = (
        stats[sub_col].astype(str) + " (" + stats["전체건수"].astype(str) + "건)"
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

    return result_df


if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = process_excel(df_raw)

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
        res_job = calc_stats_flat(df, "직군")
        st.dataframe(res_job, use_container_width=True)

        # 2. 상황별
        st.subheader("📋 2) 상황별 세부 항목 정확한 환자 확인")
        res_context = calc_stats_hierarchy(
            df, "상황_대분류", "상황_소분류", "상황"
        )
        st.dataframe(res_context, use_container_width=True)

        # 3. 부서별
        st.subheader("🏥 3) 부서 대분류 및 세부 부서별 정확한 환자 확인")
        res_dept = calc_stats_hierarchy(df, "부서_대분류", "부서_소분류", "부서")
        st.dataframe(res_dept, use_container_width=True)

        # 미시행 목록
        st.divider()
        st.subheader("⚠️ 정확한 환자확인 미시행 내역 (1차 또는 2차 미시행)")
        fail_df = df[~df["정확한확인_성공"]][
            [
                "NO",
                "부서_대분류",
                "부서_소분류",
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
