import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import io

st.set_page_config(page_title="환자안전지표 모니터링 분석기", page_icon="🏥", layout="wide")

st.title("🏥 환자안전지표 현장모니터링 결과 분석기")
st.write("다운로드한 환자확인율 결과표 엑셀 파일(.xlsx)을 업로드하면 보고서 양식에 맞는 통계 및 시각화 차트를 자동으로 산출합니다.")

uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=["xlsx", "xls"])

# --- 함수 정의 ---
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
    return "기타", (dept_clean if dept_clean else "기타")

def classify_fail_reason(row):
    p1, p2 = row["1차확인_성공"], row["2차확인_성공"]
    if not p1 and p2: return "1차 미시행"
    elif p1 and not p2: return "2차 미시행"
    else: return "1,2차 미시행"

def process_excel(df_raw):
    for idx, row in df_raw.iterrows():
        if "NO" in row.values:
            df_raw.columns = df_raw.iloc[idx]
            df_raw = df_raw.iloc[idx + 1 :].reset_index(drop=True)
            break
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

# --- 메인 로직 ---
if uploaded_file is not None:
    try:
        df = process_excel(pd.read_excel(uploaded_file))
        total_count = len(df)
        p1_count, p2_count = int(df["1차확인_성공"].sum()), int(df["2차확인_성공"].sum())
        final_count = int(df["정확한확인_성공"].sum())
        
        st.subheader("📊 1. 모니터링 총괄 현황")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("총 점검 건수", f"{total_count} 건")
        c2.metric("1차 이름확인", f"{p1_count} 건")
        c3.metric("2차 등록번호 확인", f"{p2_count} 건")
        c4.metric("최종 정확한 환자확인", f"{final_count} 건", delta=f"{round(final_count/total_count*100,1)}%")

        # (중간 그래프 로직은 생략했지만, 기존에 잘 되던 코드 그대로 두시면 됩니다.)
        
        # --- [결과 다운로드 기능] ---
        st.divider()
        st.subheader("📥 결과 보고서 다운로드")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # 주요 통계 계산
            raw_job = df.groupby("직군").agg(전체=("정확한확인_성공", "count"), 성공=("정확한확인_성공", "sum")).reset_index()
            raw_job.to_excel(writer, sheet_name="직군별_분석", index=False)
            df[~df["정확한확인_성공"]].to_excel(writer, sheet_name="미시행_상세내역", index=False)
        
        st.download_button(
            label="📊 분석 결과 엑셀로 다운로드하기",
            data=buffer.getvalue(),
            file_name="분석결과_보고서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("버튼을 누르면 분석된 데이터와 미시행 내역이 담긴 엑셀 파일을 다운로드합니다.")

    except Exception as e:
        st.error(f"오류 발생: {e}")
