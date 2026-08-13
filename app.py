import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import io

st.set_page_config(page_title="환자안전지표 모니터링 분석기", page_icon="🏥", layout="wide")

st.title("🏥 환자안전지표 현장모니터링 결과 분석기")
st.write("다운로드한 환자확인율 결과표 엑셀 파일(.xlsx)을 업로드하면 보고서 양식에 맞는 통계 및 시각화 차트를 자동으로 산출합니다.")

uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=["xlsx", "xls"])

# (위쪽 함수 정의 부분은 기존과 동일하므로 생략 - 그대로 유지하세요)
# [참고: 위에서 정의한 map_job, map_context, map_dept_detailed, classify_fail_reason, process_excel, calc_stats_raw, format_stats_df 함수들은 그대로 두세요]

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = process_excel(df_raw)

        # (중간의 차트 및 통계 분석 코드들 전부 유지)
        # ... [모니터링 총괄, 직군별, 상황별, 부서별 차트 코드들] ...

        # -------------------------------------------------------------
        # 5. 미시행 분석 및 상세 목록 (정상 실행 영역)
        # -------------------------------------------------------------
        st.divider()
        st.subheader("⚠️ 정확한 환자확인 미시행 내역 및 사유 분석")

        fail_df = df[~df["정확한확인_성공"]].copy()
        if len(fail_df) > 0:
            fail_df["미시행_유형"] = fail_df.apply(classify_fail_reason, axis=1)
            # ... [파이 차트 및 상세 테이블 코드] ...
            st.dataframe(fail_df, use_container_width=True)
        else:
            st.success("🎉 모든 건에서 정확한 환자확인이 수행되었습니다!")

        # -------------------------------------------------------------
        # [수정됨] 다운로드 버튼 위치를 여기(try 안쪽 하단)로 이동!
        # -------------------------------------------------------------
        st.divider()
        st.subheader("📥 결과 보고서 다운로드")

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            calc_stats_raw(df, "직군").to_excel(writer, sheet_name="직군별_분석", index=False)
            calc_stats_raw(df, "상황").to_excel(writer, sheet_name="상황별_분석", index=False)
            
            # 부서별 통계용 데이터프레임 생성
            raw_dept_sub = df.groupby(["부서_대분류", "부서_소분류"]).agg(전체건수=("정확한확인_성공", "count")).reset_index()
            raw_dept_sub.to_excel(writer, sheet_name="부서별_분석", index=False)
            
            if len(fail_df) > 0:
                fail_df.to_excel(writer, sheet_name="미시행_상세내역", index=False)

        st.download_button(
            label="📊 분석 결과 엑셀로 다운로드하기",
            data=buffer.getvalue(),
            file_name="환자확인_분석결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.info("버튼을 누르면 분석된 통계 자료와 미시행 내역이 담긴 엑셀 파일을 내려받을 수 있습니다.")

    except Exception as e:
        st.error(f"엑셀 분석 중 오류가 발생했습니다: {e}")
