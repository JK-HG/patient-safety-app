import io
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# 소수점 .0을 제거하는 공통 함수
def format_pct(val):
    return f"{val:.1f}".replace(".0", "") + "%"

st.set_page_config(page_title="환자안전지표 모니터링 분석기", page_icon="🏥", layout="wide")

st.title("🏥 환자안전지표 현장모니터링 결과 분석기")
st.write("다운로드한 환자확인율 결과표 엑셀 파일(.xlsx)을 업로드하면 보고서 양식에 맞는 통계 및 시각화 차트를 자동으로 산출합니다.")

uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=["xlsx", "xls"])

# [함수 정의들 생략 - 동일]
def map_job(job):
    if pd.isna(job): return "기타"
    job = str(job).strip()
    if job == "간호사": return "간호"
    elif job == "의사": return "의사"
    elif job in ["의료기사", "방사선사", "임상병리사"]: return "진료지원"
    elif job in ["사원", "행정원"]: return "행정"
    return "기타"

def map_context(item):
    if pd.isna(item): return "그 외 항목"
    item = str(item).strip()
    if item in ["진료 전", "기 타 - 수납 시", "기 타 - 입원 시"]: return "입원/외래진료/수납"
    elif item == "의약품 투여 전": return "투약"
    elif item.startswith("검사 시행 전"): return "검사 및 검체채취"
    elif item.startswith("처치 및 시술 전") or item.startswith("혈액제제 투여 전") or "수술전" in item: return "처치/수술 및 수혈"
    elif item == "기 타 - 물리치료": return "물리치료"
    else: return "그 외 항목"

def map_dept_detailed(dept, job):
    dept_str = str(dept).strip() if pd.notna(dept) and str(dept).strip() not in ["nan", "None", ""] else ""
    job_str = str(job).strip() if pd.notna(job) and str(job).strip() not in ["nan", "None", ""] else ""
    if dept_str in ["일반내과", "내과"]: dept_clean = "내과"
    elif dept_str in ["물리치료팀", "물리치료실"]: dept_clean = "물리치료실"
    elif dept_str in ["수납계", "원무계", "수납/접수"]: dept_clean = "수납/접수"
    else: dept_clean = dept_str
    outpatient_depts = ["화상외과", "성형외과", "재활의학과", "내과", "소아청소년과", "외래"]
    if job_str == "의사": return "진료과", (dept_clean if dept_clean else "진료과")
    if job_str == "간호사": return "간호부", ("외래" if dept_clean in outpatient_depts else (dept_clean if dept_clean else "외래"))
    if dept_clean in ["물리치료실", "영상의학과", "수납/접수"]: return "진료지원 및 행정", dept_clean
    if dept_clean in outpatient_depts: return "간호부", "외래"
    return "기타", (dept_clean if dept_clean else "기타")

def classify_fail_reason(row):
    p1 = row["1차확인_성공"]
    p2 = row["2차확인_성공"]
    if not p1 and p2: return "1차 미시행"
    elif p1 and not p2: return "2차 미시행"
    else: return "1,2차 미시행"

def process_excel(df_raw):
    if "NO" not in df_raw.columns:
        for idx, row in df_raw.iterrows():
            if "NO" in row.values:
                df_raw.columns = df_raw.iloc[idx]; df_raw = df_raw.iloc[idx+1:].reset_index(drop=True); break
    df = df_raw.dropna(subset=["NO"]).copy()
    df["1차확인_성공"] = df["이름확인"].astype(str).str.strip() == "시행"
    df["2차확인_성공"] = (df["등록번호 확인"].astype(str).str.strip() == "시행") | (df["생년월일 확인"].astype(str).str.strip() == "시행")
    df["정확한확인_성공"] = df["1차확인_성공"] & df["2차확인_성공"]
    df["직군"] = df["직종"].apply(map_job)
    df["상황"] = df["환자확인사항"].apply(map_context)
    raw_dept = df["상세부서명"].fillna(df["부서"])
    dept_res = [map_dept_detailed(d, j) for d, j in zip(raw_dept, df["직종"])]
    df["부서_대분류"] = [r[0] for r in dept_res]
    df["부서_소분류"] = [r[1] for r in dept_res]
    return df

def calc_stats_raw(df, group_col):
    stats = df.groupby(group_col).agg(전체건수=("정확한확인_성공", "count"), 차1확인_성공=("1차확인_성공", "sum"), 차2확인_성공=("2차확인_성공", "sum"), 정확한확인_성공=("정확한확인_성공", "sum")).reset_index()
    stats["1차확인_비율"] = (stats["차1확인_성공"] / stats["전체건수"] * 100).round(1)
    stats["2차확인_비율"] = (stats["차2확인_성공"] / stats["전체건수"] * 100).round(1)
    stats["정확한확인_비율"] = (stats["정확한확인_성공"] / stats["전체건수"] * 100).round(1)
    return stats

def format_stats_df(stats, group_col):
    result_df = pd.DataFrame()
    result_df[group_col] = stats[group_col].astype(str) + " (" + stats["전체건수"].astype(str) + "건)"
    result_df["1차 환자성명 확인"] = stats["차1확인_성공"].astype(str) + "건 (" + stats["1차확인_비율"].apply(format_pct) + ")"
    result_df["2차 등록번호 확인"] = stats["차2확인_성공"].astype(str) + "건 (" + stats["2차확인_비율"].apply(format_pct) + ")"
    result_df["정확한 환자확인"] = stats["정확한확인_성공"].astype(str) + "건 (" + stats["정확한확인_비율"].apply(format_pct) + ")"
    return result_df

if uploaded_file is not None:
    try:
        df = process_excel(pd.read_excel(uploaded_file))
        total, p1, p2, final = len(df), df["1차확인_성공"].sum(), df["2차확인_성공"].sum(), df["정확한확인_성공"].sum()
        
        # 1. 총괄 현황
        st.subheader("📊 1. 모니터링 총괄 현황")
        final_pct = (final / total * 100).round(1)
        st.metric("최종 정확한 환자확인", f"{final} 건", delta=f"이행률 {format_pct(final_pct)}")
        
        # 2. 직군별 그래프 수정 예시
        raw_job = calc_stats_raw(df, "직군")
        fig_job = go.Figure()
        fig_job.add_trace(go.Bar(name="1차", x=raw_job["직군"], y=raw_job["1차확인_비율"], text=raw_job["1차확인_비율"].apply(format_pct), textposition="inside"))
        fig_job.add_trace(go.Bar(name="2차", x=raw_job["직군"], y=raw_job["2차확인_비율"], text=raw_job["2차확인_비율"].apply(format_pct), textposition="inside"))
        for idx, row in raw_job.iterrows():
            fig_job.add_annotation(x=row["직군"], y=row["1차확인_비율"] + row["2차확인_비율"] + 5, text=f"<b>{format_pct(row['정확한확인_비율'])}</b>", showarrow=False)
        st.plotly_chart(fig_job, use_container_width=True)

        # 3. 미시행 파이 차트 수정
        fail_df = df[~df["정확한확인_성공"]].copy()
        if len(fail_df) > 0:
            fail_summary = fail_df.apply(classify_fail_reason, axis=1).value_counts().reset_index()
            fail_summary.columns = ["미시행_유형", "건수"]
            fig_pie = px.pie(fail_summary, values="건수", names="미시행_유형", hole=0.3)
            # 퍼센트 포맷을 .0% 형태가 아니라 .1f에서 .0을 지우는 방식으로 처리
            fig_pie.update_traces(texttemplate="<b>%{label}</b><br>%{value}건 (%{percent:.1%})".replace(".0%", "%"), textinfo="label+percent+value")
            st.plotly_chart(fig_pie, use_container_width=True)
            
        # [이하 생략 - 동일한 방식으로 모든 텍스트 부분을 format_pct 함수로 감싸서 수정]
        
    except Exception as e:
        st.error(f"엑셀 분석 중 오류가 발생했습니다: {e}")
