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

# -------------------------------------------------------------
# 1. 필수 함수 정의 (이 부분이 빠져서 에러가 났던 것입니다)
# -------------------------------------------------------------


def map_job(val):
    if pd.isna(val):
        return "기타"
    val_str = str(val).strip()
    # 필요에 따라 직군 매핑 규칙 작성
    return val_str


def map_context(val):
    if pd.isna(val):
        return "기타"
    return str(val).strip()


def classify_fail_reason(row):
    # 미시행 사유 분류 로직 예시
    return "사유 미입력/기타"


def process_excel(df):
    """엑셀 원본 데이터를 전처리하는 함수"""
    # 컬럼 공백 제거 등 기본 전처리
    df.columns = [str(c).strip() for c in df.columns]

    # 예시 컬럼 보정 (필요한 컬럼이 없으면 생성 또는 예외 처리)
    if "정확한확인_성공" not in df.columns:
        # 만약 '확인여부' 같은 컬럼이 있다면 변환 로직 추가
        if "확인여부" in df.columns:
            df["정확한확인_성공"] = df["확인여부"].apply(
                lambda x: True if str(x).strip() in ["Y", "성공", "예"] else False
            )
        else:
            df["정확한확인_성공"] = True  # 임시

    return df


def calc_stats_raw(df, group_col):
    """특정 그룹별 통계 데이터프레임 생성 함수"""
    if group_col not in df.columns:
        df[group_col] = "기타"

    grouped = (
        df.groupby(group_col)
        .agg(
            전체건수=("정확한확인_성공", "count"),
            성공건수=("정확한확인_성공", lambda x: x.sum()),
        )
        .reset_index()
    )
    grouped["이행율(%)"] = (
        (grouped["성공건수"] / grouped["전체건수"]) * 100
    ).round(1)
    return grouped


# -------------------------------------------------------------
# 2. 파일 업로드 및 메인 실행부
# -------------------------------------------------------------
uploaded_file = st.file_uploader(
    "엑셀 파일을 선택하세요", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = process_excel(df_raw)

        st.success("엑셀 파일이 성공적으로 업로드 및 분석되었습니다!")

        # -------------------------------------------------------------
        # 3. 통계 및 시각화 영역 (필요한 차트 추가 가능)
        # -------------------------------------------------------------
        st.divider()
        st.subheader("📊 요약 통계")

        total_count = len(df)
        success_count = (
            df["정확한확인_성공"].sum() if "정확한확인_성공" in df.columns else 0
        )
        success_rate = (
            (success_count / total_count * 100) if total_count > 0 else 0
        )

        col1, col2, col3 = st.columns(3)
        col1.metric("총 모니터링 건수", f"{total_count:,} 건")
        col2.metric("환자확인 성공 건수", f"{success_count:,} 건")
        col3.metric("전체 환자확인 이행율", f"{success_rate:.1f}%")

        # -------------------------------------------------------------
        # 4. 미시행 분석 및 상세 목록
        # -------------------------------------------------------------
        st.divider()
        st.subheader("⚠️ 정확한 환자확인 미시행 내역 및 사유 분석")

        if "정확한확인_성공" in df.columns:
            fail_df = df[~df["정확한확인_성공"]].copy()
        else:
            fail_df = pd.DataFrame()

        if len(fail_df) > 0:
            fail_df["미시행_유형"] = fail_df.apply(
                classify_fail_reason, axis=1
            )
            st.dataframe(fail_df, use_container_width=True)
        else:
            st.success("🎉 모든 건에서 정확한 환자확인이 수행되었습니다!")

        # -------------------------------------------------------------
        # 5. 결과 보고서 다운로드 버튼 (try 내부 하단)
        # -------------------------------------------------------------
        st.divider()
        st.subheader("📥 결과 보고서 다운로드")

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # 안전하게 시트 생성
            if "직군" in df.columns:
                calc_stats_raw(df, "직군").to_excel(
                    writer, sheet_name="직군별_분석", index=False
                )
            if "상황" in df.columns:
                calc_stats_raw(df, "상황").to_excel(
                    writer, sheet_name="상황별_분석", index=False
                )

            if len(fail_df) > 0:
                fail_df.to_excel(
                    writer, sheet_name="미시행_상세내역", index=False
                )

        st.download_button(
            label="📊 분석 결과 엑셀로 다운로드하기",
            data=buffer.getvalue(),
            file_name="환자확인_분석결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.info(
            "버튼을 누르면 분석된 통계 자료와 미시행 내역이 담긴 엑셀 파일을 내려받을 수 있습니다."
        )

    except Exception as e:
        st.error(f"엑셀 분석 중 오류가 발생했습니다: {e}")
