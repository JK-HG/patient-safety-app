import io
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
     
     
# 2. 상황별 카테고리 매핑
def map_context(item):
    if pd.isna(item):
        return "그 외 항목"
     
    item = str(item).strip()
     
    if item in ["진료 전", "기 타 - 수납 시", "기 타 - 입원 시"]:
        return "입원/외래진료/수납"
    elif item == "의약품 투여 전":
        return "투약"
    elif item.startswith("검사 시행 전"):
        return "검사 및 검체채취"
    elif (
        item.startswith("처치 및 시술 전")
        or item.startswith("혈액제제 투여 전")
        or "수술전" in item
    ):
        return "처치/수술 및 수혈"
    elif item == "기 타 - 물리치료":
        return "물리치료"
    else:
        return "그 외 항목"
     
     
# 3. 부서 대분류 & 소분류 매핑
def map_dept_detailed(dept, job):
    dept_str = (
        str(dept).strip()
        if pd.notna(dept) and str(dept).strip() not in ["nan", "None", ""]
        else ""
    )
    job_str = (
        str(job).strip()
        if pd.notna(job) and str(job).strip() not in ["nan", "None", ""]
        else ""
    )
     
    if dept_str in ["일반내과", "내과"]:
        dept_clean = "내과"
    elif dept_str in ["물리치료팀", "물리치료실"]:
        dept_clean = "물리치료실"
    elif dept_str in ["수납계", "원무계", "수납/접수"]:
        dept_clean = "수납/접수"
    else:
        dept_clean = dept_str
     
    outpatient_depts = [
        "화상외과",
        "성형외과",
        "재활의학과",
        "내과",
        "소아청소년과",
        "외래",
    ]
     
    if job_str == "의사":
        sub_dept = dept_clean if dept_clean else "진료과"
        return "진료과", sub_dept
     
    if job_str == "간호사":
        if dept_clean in outpatient_depts:
            return "간호부", "외래"
        else:
            return "간호부", dept_clean if dept_clean else "외래"
     
    if dept_clean in ["물리치료실", "영상의학과", "수납/접수"]:
        return "진료지원 및 행정", dept_clean
     
    if dept_clean in outpatient_depts:
        return "간호부", "외래"
     
    return "기타", dept_clean if dept_clean else "기타"
     
     
# 4. 미시행 유형 분류 함수
def classify_fail_reason(row):
    p1 = row["1차확인_성공"]
    p2 = row["2차확인_성공"]
    if not p1 and p2:
        return "1차 미시행"
    elif p1 and not p2:
        return "2차 미시행"
    else:
        return "1,2차 미시행"
     
     
def process_excel(df_raw):
    if "NO" not in df_raw.columns:
        for idx, row in df_raw.iterrows():
            if "NO" in row.values:
                df_raw.columns = df_raw.iloc[idx]
                df_raw = df_raw.iloc[idx + 1 :].reset_index(drop=True)
                break
     
    df = df_raw.dropna(subset=["NO"]).copy()
     
    df["1차확인_성공"] = df["이름확인"].astype(str).str.strip() == "시행"
    df["2차확인_성공"] = (
        df["등록번호 확인"].astype(str).str.strip() == "시행"
    ) | (df["생년월일 확인"].astype(str).str.strip() == "시행")
    df["정확한확인_성공"] = df["1차확인_성공"] & df["2차확인_성공"]
     
    df["직군"] = df["직종"].apply(map_job)
    df["상황"] = df["환자확인사항"].apply(map_context)
     
    raw_dept = df["상세부서명"].fillna(df["부서"])
     
    dept_res = [
        map_dept_detailed(d, j) for d, j in zip(raw_dept, df["직종"])
    ]
    df["부서_대분류"] = [r[0] for r in dept_res]
    df["부서_소분류"] = [r[1] for r in dept_res]
     
    return df
     
     
def calc_stats_raw(df, group_col):
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
        stats["차1확인_성공"] / stats["전체건수"] * 100
    ).round(1)
    stats["2차확인_비율"] = (
        stats["차2확인_성공"] / stats["전체건수"] * 100
    ).round(1)
    stats["정확한확인_비율"] = (
        stats["정확한확인_성공"] / stats["전체건수"] * 100
    ).round(1)
     
    return stats
     
     
def format_stats_df(stats, group_col):
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
     
        col_q_curr, col_q_cmp = st.columns([1, 2])
        with col_q_curr:
            current_quarter = st.selectbox(
                "📌 **현재 업로드한 엑셀 데이터의 분기**",
                ["1분기", "2분기", "3분기", "4분기"],
                index=0,
            )
     
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(f"총 점검 건수 ({current_quarter})", f"{total_count} 건")
        c2.metric("1차 이름확인 건수", f"{p1_count} 건")
        c3.metric("2차 등록/생년월일 확인", f"{p2_count} 건")
        c4.metric(
            "최종 정확한 환자확인",
            f"{final_count} 건",
            delta=f"이행률 {final_rate}%",
        )
     
        all_quarters_list = ["1분기", "2분기", "3분기", "4분기"]
        other_quarters_options = [
            q for q in all_quarters_list if q != current_quarter
        ]
     
        with st.expander(f"⚙ **다른 분기 수치 추가 및 비교 설정**", expanded=False):
            st.info(
                "비교하고자 하는 타 분기를 선택하고 수치를 입력하세요. 입력한 분기들이 그래프에 함께 표시됩니다."
            )
            selected_other_quarters = st.multiselect(
                "비교할 타 분기를 선택하세요:",
                other_quarters_options,
                default=[],
            )
     
            quarter_data = {}
            quarter_data[current_quarter] = {
                "total": total_count,
                "p1": p1_count,
                "p2": p2_count,
                "final": final_count,
            }
     
            if selected_other_quarters:
                q_cols = st.columns(len(selected_other_quarters))
                for idx, q_name in enumerate(selected_other_quarters):
                    with q_cols[idx]:
                        st.markdown(f"##### **[{q_name}] 수치 입력**")
                        q_total = st.number_input(
                            f"총 점검 건수 ({q_name})",
                            value=300,
                            min_value=1,
                            step=1,
                            key=f"tot_{q_name}",
                        )
                        q_p1 = st.number_input(
                            f"1차 성명 확인 ({q_name})",
                            value=290,
                            min_value=0,
                            max_value=q_total,
                            step=1,
                            key=f"p1_{q_name}",
                        )
                        q_p2 = st.number_input(
                            f"2차 등록번호 확인 ({q_name})",
                            value=285,
                            min_value=0,
                            max_value=q_total,
                            step=1,
                            key=f"p2_{q_name}",
                        )
                        q_final = st.number_input(
                            f"정확한 환자확인 ({q_name})",
                            value=280,
                            min_value=0,
                            max_value=q_total,
                            step=1,
                            key=f"fin_{q_name}",
                        )
     
                        quarter_data[q_name] = {
                            "total": q_total,
                            "p1": q_p1,
                            "p2": q_p2,
                            "final": q_final,
                        }
     
        ordered_quarter_data = {
            q: quarter_data[q] for q in all_quarters_list if q in quarter_data
        }
     
        categories = ["1차 환자성명 확인", "2차 등록번호 확인", "정확한 환자확인"]
        fig_total = go.Figure()
     
        q_colors = {
            "1분기": ("#1E3A8A", "#EF4444"),
            "2분기": ("#15803D", "#F59E0B"),
            "3분기": ("#C2410C", "#991B1B"),
            "4분기": ("#6B21A8", "#DB2777"),
        }
     
        for q_name, q_val in ordered_quarter_data.items():
            tot = q_val["total"]
            p1 = q_val["p1"]
            p2 = q_val["p2"]
            fin = q_val["final"]
     
            p1_pct = round(p1 / tot * 100, 1) if tot > 0 else 0
            p2_pct = round(p2 / tot * 100, 1) if tot > 0 else 0
            fin_pct = round(fin / tot * 100, 1) if tot > 0 else 0
     
            pass_color, fail_color = q_colors.get(
                q_name, ("#1E3A8A", "#EF4444")
            )
     
            fig_total.add_trace(
                go.Bar(
                    name=f"{q_name} (시행)",
                    x=categories,
                    y=[p1, p2, fin],
                    text=[
                        f"<b>{p1}명</b><br>({p1_pct}%)",
                        f"<b>{p2}명</b><br>({p2_pct}%)",
                        f"<b>{fin}명</b><br>({fin_pct}%)",
                    ],
                    textposition="inside",
                    insidetextfont=dict(size=18, color="white"),
                    marker_color=pass_color,
                    offsetgroup=q_name,
                )
            )
     
            fail_p1, fail_p2, fail_fin = tot - p1, tot - p2, tot - fin
     
            fig_total.add_trace(
                go.Bar(
                    name=f"{q_name} (미시행)",
                    x=categories,
                    y=[fail_p1, fail_p2, fail_fin],
                    text=[
                        f"<b>{fail_p1}명</b>" if fail_p1 > 0 else "",
                        f"<b>{fail_p2}명</b>" if fail_p2 > 0 else "",
                        f"<b>{fail_fin}명</b>" if fail_fin > 0 else "",
                    ],
                    textposition="outside",
                    outsidetextfont=dict(size=18, color=fail_color),
                    marker_color=fail_color,
                    offsetgroup=q_name,
                    base=[p1, p2, fin],
                )
            )
     
        max_total = max([q["total"] for q in ordered_quarter_data.values()])
        num_q = len(ordered_quarter_data)
     
        if num_q == 1:
            dynamic_bargap = 0.62
            col_ratio = [0.5, 4.0, 0.5]
        elif num_q == 2:
            dynamic_bargap = 0.45
            col_ratio = [0.3, 4.4, 0.3]
        else:
            dynamic_bargap = 0.25
            col_ratio = [0.1, 4.8, 0.1]
     
        fig_total.update_layout(
            barmode="group",
            title=dict(
                text="<b>정확한 환자 확인율</b>",
                x=0.5,
                y=0.96,
                xanchor="center",
                yanchor="top",
                font=dict(size=28, color="black"),
            ),
            xaxis=dict(
                tickfont=dict(color="black", size=20, family="sans-serif"),
                title=None,
            ),
            yaxis=dict(
                title=None,
                tickfont=dict(size=18, color="black"),
                range=[0, max_total * 1.20],
            ),
            bargap=dynamic_bargap,
            bargroupgap=0.08,
            height=720,
            margin=dict(l=80, r=100, t=100, b=160),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.32,
                xanchor="center",
                x=0.5,
                font=dict(size=14),
                itemwidth=30,
                entrywidth=100,
                entrywidthmode="pixels",
            ),
        )
     
        col_l1, col_m1, col_r1 = st.columns(col_ratio)
        with col_m1:
            with st.container(border=True):
                st.plotly_chart(fig_total, use_container_width=True)
     
        st.divider()
     
        # -------------------------------------------------------------
        # 2. 직군별 정확한 환자 확인
        # -------------------------------------------------------------
        st.subheader("👥 2) 직군별 정확한 환자 확인")
        raw_job = calc_stats_raw(df, "직군")
     
        job_order = ["의사", "간호", "진료지원", "행정"]
        raw_job_filtered = raw_job[raw_job["직군"].isin(job_order)].copy()
        raw_job_filtered["직군"] = pd.Categorical(
            raw_job_filtered["직군"], categories=job_order, ordered=True
        )
        raw_job_filtered = raw_job_filtered.sort_values("직군").reset_index(
            drop=True
        )
     
        res_job_df = format_stats_df(raw_job_filtered, "직군")
        st.dataframe(res_job_df, use_container_width=True)
     
        fig_job = go.Figure()
        fig_job.add_trace(
            go.Bar(
                name="1차 확인 비율 (%)",
                x=raw_job_filtered["직군"],
                y=raw_job_filtered["1차확인_비율"],
                text=[
                    f"1차: {val}%" for val in raw_job_filtered["1차확인_비율"]
                ],
                textposition="inside",
                insidetextfont=dict(size=20, color="white"),
                marker_color="#3366cc",
            )
        )
        fig_job.add_trace(
            go.Bar(
                name="2차 확인 비율 (%)",
                x=raw_job_filtered["직군"],
                y=raw_job_filtered["2차확인_비율"],
                text=[
                    f"2차: {val}%" for val in raw_job_filtered["2차확인_비율"]
                ],
                textposition="inside",
                insidetextfont=dict(size=20, color="white"),
                marker_color="#109618",
            )
        )
     
        for idx, row in raw_job_filtered.iterrows():
            total_height = row["1차확인_비율"] + row["2차확인_비율"]
            fig_job.add_annotation(
                x=row["직군"],
                y=total_height + 8,
                text=f"<b>정확한 확인율: {row['정확한확인_비율']}%</b>",
                showarrow=False,
                font=dict(size=20, color="#2b5c8f"),
            )
     
        fig_job.update_layout(
            barmode="stack",
            title=dict(
                text="<b>직군별 정확한 환자 확인율</b>",
                x=0.5,
                y=0.96,
                xanchor="center",
                yanchor="top",
                font=dict(size=28, color="black"),
            ),
            xaxis=dict(tickfont=dict(color="black", size=20)),
            yaxis=dict(
                title=None,
                tickfont=dict(size=18, color="black"),
                range=[0, 235],
            ),
            bargap=0.45,
            height=660,
            margin=dict(l=80, r=100, t=100, b=130),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.20,
                xanchor="center",
                x=0.5,
                font=dict(size=18),
                entrywidth=200,
                entrywidthmode="pixels",
            ),
        )
     
        col_l2, col_m2, col_r2 = st.columns([0.5, 4, 0.5])
        with col_m2:
            with st.container(border=True):
                st.plotly_chart(fig_job, use_container_width=True)
     
        st.divider()
     
        # -------------------------------------------------------------
        # 3. 상황별 정확한 환자 확인
        # -------------------------------------------------------------
        st.subheader("📋 3) 상황별 정확한 환자 확인")
        raw_context = calc_stats_raw(df, "상황")
     
        context_order = [
            "입원/외래진료/수납",
            "투약",
            "검사 및 검체채취",
            "처치/수술 및 수혈",
            "물리치료",
        ]
        raw_context_filtered = raw_context[
            raw_context["상황"].isin(context_order)
        ].copy()
        raw_context_filtered["상황"] = pd.Categorical(
            raw_context_filtered["상황"], categories=context_order, ordered=True
        )
        raw_context_filtered = raw_context_filtered.sort_values(
            "상황"
        ).reset_index(drop=True)
     
        res_context_df = format_stats_df(raw_context_filtered, "상황")
        st.dataframe(res_context_df, use_container_width=True)
     
        fig_context = go.Figure()
        fig_context.add_trace(
            go.Bar(
                name="1차 확인 비율 (%)",
                x=raw_context_filtered["상황"],
                y=raw_context_filtered["1차확인_비율"],
                text=[
                    f"1차: {val}%"
                    for val in raw_context_filtered["1차확인_비율"]
                ],
                textposition="inside",
                insidetextfont=dict(size=20, color="white"),
                marker_color="#3366cc",
            )
        )
        fig_context.add_trace(
            go.Bar(
                name="2차 확인 비율 (%)",
                x=raw_context_filtered["상황"],
                y=raw_context_filtered["2차확인_비율"],
                text=[
                    f"2차: {val}%"
                    for val in raw_context_filtered["2차확인_비율"]
                ],
                textposition="inside",
                insidetextfont=dict(size=20, color="white"),
                marker_color="#109618",
            )
        )
     
        for idx, row in raw_context_filtered.iterrows():
            total_height = row["1차확인_비율"] + row["2차확인_비율"]
            fig_context.add_annotation(
                x=row["상황"],
                y=total_height + 8,
                text=f"<b>정확한 확인율: {row['정확한확인_비율']}%</b>",
                showarrow=False,
                font=dict(size=18, color="#2b5c8f"),
            )
     
        fig_context.update_layout(
            barmode="stack",
            title=dict(
                text="<b>상황별 정확한 환자 확인율</b>",
                x=0.5,
                y=0.96,
                xanchor="center",
                yanchor="top",
                font=dict(size=28, color="black"),
            ),
            xaxis=dict(tickfont=dict(color="black", size=20)),
            yaxis=dict(
                title=None,
                tickfont=dict(size=18, color="black"),
                range=[0, 235],
            ),
            bargap=0.4,
            height=660,
            margin=dict(l=80, r=100, t=100, b=130),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.20,
                xanchor="center",
                x=0.5,
                font=dict(size=18),
                entrywidth=200,
                entrywidthmode="pixels",
            ),
        )
     
        col_l3, col_m3, col_r3 = st.columns([0.2, 4.6, 0.2])
        with col_m3:
            with st.container(border=True):
                st.plotly_chart(fig_context, use_container_width=True)
     
        st.divider()
     
        # -------------------------------------------------------------
        # 4. 부서별(대분류/소분류) 정확한 환자 확인
        # -------------------------------------------------------------
        st.subheader("🏥 4) 부서별(대분류/소분류) 정확한 환자 확인")
     
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
            raw_dept_sub["차1확인_성공"] / raw_dept_sub["전체건수"] * 100
        ).round(1)
        raw_dept_sub["2차확인_비율"] = (
            raw_dept_sub["차2확인_성공"] / raw_dept_sub["전체건수"] * 100
        ).round(1)
        raw_dept_sub["정확한확인_비율"] = (
            raw_dept_sub["정확한확인_성공"] / raw_dept_sub["전체건수"] * 100
        ).round(1)
     
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
     
        valid_categories = ["진료과", "간호부", "진료지원 및 행정"]
        dept_categories = [
            c for c in valid_categories if c in raw_dept_sub["부서_대분류"].unique()
        ]
     
        dept_title_map = {
            "진료과": "진료과별 정확한 환자확인",
            "간호부": "간호부 정확한 환자확인",
            "진료지원 및 행정": "진료지원 및 행정 정확한 환자확인",
        }
     
        tabs = st.tabs(dept_categories)
        for tab, cat in zip(tabs, dept_categories):
            with tab:
                sub_df = raw_dept_sub[
                    raw_dept_sub["부서_대분류"] == cat
                ].copy()
     
                sub_df = sub_df[
                    ~sub_df["부서_소분류"].str.lower().isin(["기타", "none", "nan"])
                ].copy()
     
                if cat == "진료과":
                    custom_order = [
                        "화상외과",
                        "성형외과",
                        "재활의학과",
                        "내과",
                        "소아청소년과",
                    ]
                    sub_df["부서_소분류"] = pd.Categorical(
                        sub_df["부서_소분류"],
                        categories=custom_order,
                        ordered=True,
                    )
                    sub_df = sub_df.sort_values("부서_소분류").reset_index(
                        drop=True
                    )
                elif cat == "간호부":
                    custom_order = [
                        "외래",
                        "신관3병동",
                        "본관5병동",
                        "본관6병동",
                        "본관7병동",
                        "본관8병동",
                        "응급실",
                        "화상중환자실",
                    ]
                    sub_df["부서_소분류"] = pd.Categorical(
                        sub_df["부서_소분류"],
                        categories=custom_order,
                        ordered=True,
                    )
                    sub_df = sub_df.sort_values("부서_소분류").reset_index(
                        drop=True
                    )
                elif cat == "진료지원 및 행정":
                    custom_order = ["물리치료실", "영상의학과", "수납/접수"]
                    sub_df["부서_소분류"] = pd.Categorical(
                        sub_df["부서_소분류"],
                        categories=custom_order,
                        ordered=True,
                    )
                    sub_df = sub_df.sort_values("부서_소분류").reset_index(
                        drop=True
                    )
     
                chart_title = dept_title_map.get(cat, f"{cat} 정확한 환자확인")
     
                fig_dept_sub = px.bar(
                    sub_df,
                    x="부서_소분류",
                    y="정확한확인_비율",
                    text=sub_df["정확한확인_비율"].astype(str) + "%",
                    title=f"<b>{chart_title}</b>",
                    labels={
                        "부서_소분류": "세부 부서",
                        "정확한확인_비율": "이행률 (%)",
                    },
                    color_discrete_sequence=["#002060"],
                )
                fig_dept_sub.update_traces(
                    textposition="outside",
                    textfont=dict(size=20, color="black"),
                )
                fig_dept_sub.update_layout(
                    title=dict(
                        x=0.5,
                        y=0.96,
                        xanchor="center",
                        yanchor="top",
                        font=dict(size=28, color="black"),
                    ),
                    xaxis=dict(
                        title=None, tickfont=dict(color="black", size=20)
                    ),
                    yaxis=dict(
                        title=None,
                        tickfont=dict(size=18, color="black"),
                        range=[0, 115],
                        ticksuffix="%",
                    ),
                    bargap=0.5,
                    height=580,
                    margin=dict(l=80, r=60, t=100, b=90),
                )
     
                col_l4, col_m4, col_r4 = st.columns([0.5, 4, 0.5])
                with col_m4:
                    with st.container(border=True):
                        st.plotly_chart(
                            fig_dept_sub, use_container_width=True
                        )
     
        # -------------------------------------------------------------
        # 5. 미시행 분석 및 상세 목록
        # -------------------------------------------------------------
        st.divider()
        st.subheader("⚠ 정확한 환자확인 미시행 내역 및 사유 분석")
     
        fail_df = df[~df["정확한확인_성공"]].copy()
     
        if len(fail_df) > 0:
            fail_df["미시행_유형"] = fail_df.apply(
                classify_fail_reason, axis=1
            )
     
            fail_summary = (
                fail_df["미시행_유형"].value_counts().reset_index()
            )
            fail_summary.columns = ["미시행_유형", "건수"]
     
            color_map = {
                "1차 미시행": "#ff9999",
                "2차 미시행": "#ffcc99",
                "1,2차 미시행": "#e06666",
            }
     
            fig_pie = px.pie(
                fail_summary,
                values="건수",
                names="미시행_유형",
                title="<b>정확한 환자확인 미시행 사유</b>",
                color="미시행_유형",
                color_discrete_map=color_map,
                hole=0.3,
            )
     
            # 파이 차트 중앙 텍스트 추가
            total_fail = fail_summary["건수"].sum()
            fig_pie.add_annotation(
                text=f"총 미시행<br><b>{total_fail}건</b>",
                x=0.5, y=0.5, font=dict(size=20, color="black"), showarrow=False
            )
     
            fig_pie.update_traces(
                textposition="inside",
                textinfo="label+percent+value",
                texttemplate="<b>%{label}</b><br>%{value}건 (%{percent})",
                insidetextfont=dict(size=20, color="black"),
                insidetextorientation="horizontal",
            )
     
            fig_pie.update_layout(
                title=dict(
                    x=0.5,
                    y=0.96,
                    xanchor="center",
                    yanchor="top",
                    font=dict(size=28, color="black"),
                ),
                height=580,
                margin=dict(l=80, r=60, t=100, b=130),
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=-0.20,
                    xanchor="center",
                    x=0.5,
                    font=dict(size=18),
                ),
            )
     
            col_pie_l, col_pie_m, col_pie_r = st.columns([0.5, 4, 0.5])
            with col_pie_m:
                with st.container(border=True):
                    st.plotly_chart(fig_pie, use_container_width=True)
     
            st.markdown("#### 📋 미시행 상세 목록")
            show_fail_df = fail_df[
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
                    "미시행_유형",
                    "미시행 사유 또는 기타사항",
                ]
            ]
            st.dataframe(show_fail_df, use_container_width=True)
     
        else:
            st.success("🎉 모든 건에서 정확한 환자확인이 수행되었습니다!")
     
        # -------------------------------------------------------------
        # 결과 엑셀 파일 다운로드 기능
        # -------------------------------------------------------------
        st.divider()
        st.subheader("📥 결과 보고서 다운로드")
     
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            raw_job.to_excel(writer, sheet_name="직군별_분석", index=False)
            raw_context.to_excel(writer, sheet_name="상황별_분석", index=False)
            raw_dept_sub.to_excel(writer, sheet_name="부서별_분석", index=False)
            if len(fail_df) > 0:
                fail_df.to_excel(
                    writer, sheet_name="미시행_상세내역", index=False
                )
     
        st.download_button(
            label="📊 분석 결과 엑셀로 다운로드하기",
            data=buffer.getvalue(),
            file_name=f"환자확인_분석결과_{current_quarter}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.info(
            "버튼을 누르면 분석된 통계 자료와 미시행 내역이 담긴 엑셀 파일을 내려받을 수 있습니다."
        )
     
    except Exception as e:
        st.error(f"엑셀 분석 중 오류가 발생했습니다: {e}")
