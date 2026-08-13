import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="환자안전지표 모니터링 분석기", page_icon="🏥", layout="wide"
)

st.title("🏥 환자안전지표 현장모니터링 결과 분석기")
st.write(
    "다운로드한 환자확인율 결과표 엑셀 파일(.xlsx)을 업로드하면 보고서 양식에 맞는 통계 및 시각화 차트를 자동으로 산출합니다."
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


# 2. 상황별 카테고리 매핑 (1~6번 단일 분류)
def map_context(item):
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
    # 헤더 자동 인식
    if "NO" not in df_raw.columns:
        for idx, row in df_raw.iterrows():
            if "NO" in row.values:
                df_raw.columns = df_raw.iloc[idx]
                df_raw = df_raw.iloc[idx + 1 :].reset_index(drop=True)
                break

    df = df_raw.dropna(subset=["NO"]).copy()

    # 1차, 2차, 최종 정확한 환자확인 판정
    df["1차확인_성공"] = df["이름확인"].astype(str).str.strip() == "시행"
    df["2차확인_성공"] = (
        df["등록번호 확인"].astype(str).str.strip() == "시행"
    ) | (df["생년월일 확인"].astype(str).str.strip() == "시행")
    df["정확한확인_성공"] = df["1차확인_성공"] & df["2차확인_성공"]

    # 파생 변수
    df["직군"] = df["직종"].apply(map_job)
    df["상황"] = df["환자확인사항"].apply(map_context)

    # 부서 분류
    raw_dept = df["상세부서명"].fillna(df["부서"])
    dept_res = raw_dept.apply(map_dept_detailed)
    df["부서_대분류"] = [r[0] for r in dept_res]
    df["부서_소분류"] = [r[1] for r in dept_res]

    return df


def calc_stats_raw(df, group_col):
    """그래프 및 통계용 기초 데이터프레임 집계"""
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

    stats["1차확인_비율"] = (
        (stats["차1확인_성공"] / stats["전체건수"] * 100).round(1)
    )
    stats["2차확인_비율"] = (
        (stats["차2확인_성공"] / stats["전체건수"] * 100).round(1)
    )
    stats["정확한확인_비율"] = (
        (stats["정확한확인_성공"] / stats["전체건수"] * 100).round(1)
    )

    return stats


def format_stats_df(stats, group_col):
    """표 출력용 문자열 포맷팅"""
    result_df = pd.DataFrame()
    result_df[group_col] = (
        stats[group_col].astype(str) + " (" + stats["전체건수"].astype(str) + "건)"
    )
    result_df["1차 환자성명 확인"] = (
        stats["차1확인_성공"].astype(str)
        + "건 ("
        + stats["1차확인_비율"].astype(str)
        + "%)"
    )
    result_df["2차 등록번호 확인"] = (
        stats["차2확인_성공"].astype(str)
        + "건 ("
        + stats["2차확인_비율"].astype(str)
        + "%)"
    )
    result_df["정확한 환자확인"] = (
        stats["정확한확인_성공"].astype(str)
        + "건 ("
        + stats["정확한확인_비율"].astype(str)
        + "%)"
    )
    return result_df


if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = process_excel(df_raw)

        total_count = len(df)
        p1_count = int(df["1차확인_성공"].sum())
        p2_count = int(df["2차확인_성공"].sum())
        final_count = int(df["정확한확인_성공"].sum())
        final_rate = (
            round((final_count / total_count) * 100, 1) if total_count > 0 else 0
        )

        # -------------------------------------------------------------
        # 1. 모니터링 총괄 현황
        # -------------------------------------------------------------
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

        # 총괄 현황 누적 막대그래프
        p1_pct = round(p1_count / total_count * 100, 1) if total_count > 0 else 0
        p2_pct = round(p2_count / total_count * 100, 1) if total_count > 0 else 0
        final_pct = round(final_count / total_count * 100, 1) if total_count > 0 else 0

        fig_total = go.Figure()
        fig_total.add_trace(
            go.Bar(
                name="시행",
                x=["1차 환자성명 확인", "2차 등록번호 확인", "정확한 환자확인"],
                y=[p1_count, p2_count, final_count],
                text=[
                    f"{p1_count}명 ({p1_pct}%)",
                    f"{p2_count}명 ({p2_pct}%)",
                    f"{final_count}명 ({final_pct}%)",
                ],
                textposition="inside",
                marker_color="#2b5c8f",
            )
        )
        fig_total.add_trace(
            go.Bar(
                name="미시행",
                x=["1차 환자성명 확인", "2차 등록번호 확인", "정확한 환자확인"],
                y=[
                    total_count - p1_count,
                    total_count - p2_count,
                    total_count - final_count,
                ],
                text=[
                    f"{total_count - p1_count}명",
                    f"{total_count - p2_count}명",
                    f"{total_count - final_count}명",
                ],
                textposition="inside",
                marker_color="#d9534f",
            )
        )
        fig_total.update_layout(
            barmode="stack",
            title="총괄 항목별 시행/미시행 인원 비율 (누적 막대)",
            yaxis_title="인원 수 (명)",
            height=380,
            margin=dict(l=20, r=20, t=40, b=20),
        )
        st.plotly_chart(fig_total, use_container_width=True)

        st.divider()

        # -------------------------------------------------------------
        # 2. 직군별 정확한 환자 확인
        # -------------------------------------------------------------
        st.subheader("👥 2) 직군별 정확한 환자 확인")
        job_order = ["의사", "간호", "진료지원", "행정"]
        raw_job = calc_stats_raw(df, "직군")

        # 4개 항목 순서 보장
        raw_job["직군"] = pd.Categorical(
            raw_job["직군"], categories=job_order, ordered=True
        )
        raw_job = raw_job.sort_values("직군").reset_index(drop=True)

        res_job_df = format_stats_df(raw_job, "직군")
        st.dataframe(res_job_df, use_container_width=True)

        # 직군별 누적 막대그래프 (1차 %, 2차 %)
        fig_job = go.Figure()
        fig_job.add_trace(
            go.Bar(
                name="1차 확인 비율 (%)",
                x=raw_job["직군"],
                y=raw_job["1차확인_비율"],
                text=raw_job["1차확인_비율"].astype(str) + "%",
                textposition="inside",
                marker_color="#3366cc",
            )
        )
        fig_job.add_trace(
            go.Bar(
                name="2차 확인 비율 (%)",
                x=raw_job["직군"],
                y=raw_job["2차확인_비율"],
                text=raw_job["2차확인_비율"].astype(str) + "%",
                textposition="inside",
                marker_color="#109618",
            )
        )
        fig_job.update_layout(
            barmode="stack",
            title="직군별 1차 및 2차 환자확인율 (누적 막대)",
            yaxis_title="비율 (%)",
            height=380,
            margin=dict(l=20, r=20, t=40, b=20),
        )
        st.plotly_chart(fig_job, use_container_width=True)

        st.divider()

        # -------------------------------------------------------------
        # 3. 상황별 정확한 환자 확인
        # -------------------------------------------------------------
        st.subheader("📋 3) 상황별 정확한 환자 확인")
        raw_context = calc_stats_raw(df, "상황")
        raw_context = raw_context.sort_values("상황").reset_index(drop=True)

        res_context_df = format_stats_df(raw_context, "상황")
        st.dataframe(res_context_df, use_container_width=True)

        # 상황별 누적 막대그래프
        fig_context = go.Figure()
        fig_context.add_trace(
            go.Bar(
                name="1차 확인 비율 (%)",
                x=raw_context["상황"],
                y=raw_context["1차확인_비율"],
                text=raw_context["1차확인_비율"].astype(str) + "%",
                textposition="inside",
                marker_color="#3366cc",
            )
        )
        fig_context.add_trace(
            go.Bar(
                name="2차 확인 비율 (%)",
                x=raw_context["상황"],
                y=raw_context["2차확인_비율"],
                text=raw_context["2차확인_비율"].astype(str) + "%",
                textposition="inside",
                marker_color="#109618",
            )
        )
        fig_context.update_layout(
            barmode="stack",
            title="상황별 1차 및 2차 환자확인율 (누적 막대)",
            yaxis_title="비율 (%)",
            height=400,
            margin=dict(l=20, r=20, t=40, b=20),
        )
        st.plotly_chart(fig_context, use_container_width=True)

        st.divider()

        # -------------------------------------------------------------
        # 4. 부서별(대분류/소분류) 정확한 환자 확인
        # -------------------------------------------------------------
        st.subheader("🏥 4) 부서별(대분류/소분류) 정확한 환자 확인")

        # 세부 부서 집계
        raw_dept_sub = (
            df.groupby(["부서_대분류", "부서_소분류"])
            .agg(
                전체건수=("정확한확인_성공", "count"),
                차1확인_성공=("1차확인_성공", "sum"),
                차2확인_성공=("2차확인_성공", "sum"),
                정확한확인_성공=("정확한확인_성공", "sum"),
            )
            .reset_index()
        )
        raw_dept_sub["1차확인_비율"] = (
            (raw_dept_sub["차1확인_성공"] / raw_dept_sub["전체건수"] * 100).round(1)
        )
        raw_dept_sub["2차확인_비율"] = (
            (raw_dept_sub["차2확인_성공"] / raw_dept_sub["전체건수"] * 100).round(1)
        )
        raw_dept_sub["정확한확인_비율"] = (
            (raw_dept_sub["정확한확인_성공"] / raw_dept_sub["전체건수"] * 100).round(1)
        )

        # 표 출력용 데이터프레임
        res_dept_df = pd.DataFrame()
        res_dept_df["부서 대분류"] = raw_dept_sub["부서_대분류"]
        res_dept_df["세부 부서(소분류)"] = (
            raw_dept_sub["부서_소분류"].astype(str)
            + " ("
            + raw_dept_sub["전체건수"].astype(str)
            + "건)"
        )
        res_dept_df["1차 환자성명 확인"] = (
            raw_dept_sub["차1확인_성공"].astype(str)
            + "건 ("
            + raw_dept_sub["1차확인_비율"].astype(str)
            + "%)"
        )
        res_dept_df["2차 등록번호 확인"] = (
            raw_dept_sub["차2확인_성공"].astype(str)
            + "건 ("
            + raw_dept_sub["2차확인_비율"].astype(str)
            + "%)"
        )
        res_dept_df["정확한 환자확인"] = (
            raw_dept_sub["정확한확인_성공"].astype(str)
            + "건 ("
            + raw_dept_sub["정확한확인_비율"].astype(str)
            + "%)"
        )

        st.dataframe(res_dept_df, use_container_width=True)

        st.markdown("#### 📈 대분류별 세부 부서 이행률 막대그래프")
        dept_categories = sorted(raw_dept_sub["부서_대분류"].unique())

        # 대분류별 탭으로 나누어 막대그래프 출력
        tabs = st.tabs(dept_categories)
        for tab, cat in zip(tabs, dept_categories):
            with tab:
                sub_df = raw_dept_sub[raw_dept_sub["부서_대분류"] == cat]
                fig_dept_sub = px.bar(
                    sub_df,
                    x="부서_소분류",
                    y="정확한확인_비율",
                    text=sub_df["정확한확인_비율"].astype(str) + "%",
                    title=f"[{cat}] 세부 부서별 정확한 환자확인율 (%)",
                    labels={
                        "부서_소분류": "세부 부서",
                        "정확한확인_비율": "이행률 (%)",
                    },
                    color_discrete_sequence=["#2b5c8f"],
                )
                fig_dept_sub.update_traces(textposition="outside")
                fig_dept_sub.update_layout(
                    yaxis=dict(range=[0, 115]),
                    height=360,
                    margin=dict(l=20, r=20, t=40, b=20),
                )
                st.plotly_chart(fig_dept_sub, use_container_width=True)

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
