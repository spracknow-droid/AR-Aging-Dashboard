import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from streamlit_plotly_events import plotly_events
import io
import numpy as np
from datetime import date
from dateutil.relativedelta import relativedelta
import calendar

# 페이지 구성을 와이드 레이아웃으로 설정
st.set_page_config(layout="wide")

st.title('매출채권 연령분석 현황')

# 사이드바 - 파일 업로드 영역
st.sidebar.header('파일 업로드')
uploaded_file = st.sidebar.file_uploader("엑셀 파일(.xlsx)을 업로드하세요", type=['xlsx'])

if uploaded_file:
    # 엑셀 파일을 읽고 마지막 행 삭제
    df = pd.read_excel(uploaded_file)
    df = df.iloc[:-1]
    st.success('파일 업로드 성공!')

    # 필수 컬럼이 모두 포함돼 있는지 확인
    required_columns = ['매출일자', '채권금액(원화)', '환종', '채권금액(외화)', '거래처명', '적요']
    if all(col in df.columns for col in required_columns):
        try:
            # '매출일자' 컬럼을 날짜 형식으로 변환 (변환 불가 시 NaT로 처리)
            df['매출일자'] = pd.to_datetime(df['매출일자'], errors='coerce')

            # 유효하지 않은 날짜 데이터가 있는 행 제거
            invalid_rows_before = len(df)
            df.dropna(subset=['매출일자'], inplace=True)
            invalid_rows_after = len(df)

            # 제거된 행이 있을 경우 사용자에게 경고 메시지 표시
            if invalid_rows_before > invalid_rows_after:
                removed_count = invalid_rows_before - invalid_rows_after
                st.warning(f"경고: 유효하지 않은 '매출일자' 데이터가 포함된 행 {removed_count}개가 제거되었습니다.")

            # 최신 날짜 찾기
            latest_date = df['매출일자'].max()

            # 최신 날짜의 해당 월 마지막 일 계산
            year = latest_date.year
            month = latest_date.month
            last_day = calendar.monthrange(year, month)[1]
            standard_date = date(year, month, last_day)

            # 사이드바에 매출채권 기준일자 표시
            st.sidebar.markdown("---")
            st.sidebar.subheader("매출채권 기준일자")
            st.sidebar.info(standard_date.strftime("%Y년 %m월 %d일"))

            # --- 레이아웃: 너비 비율로 3개 컬럼 생성 (20:35:45) ---
            col1, col2, col3 = st.columns([20, 35, 45])

            # 칼럼 1: 총 잔액 표시 및 파이차트
            with col1:
                st.subheader("매출채권 잔액")
                total_receivables = df['채권금액(원화)'].sum()
                # 1) 총 매출채권 잔액을 백만원 단위로 표기 (소수점 제거)
                total_receivables_mil = total_receivables / 1_000_000
                st.metric(label="매출채권 총 잔액 합계", value=f"{total_receivables_mil:,.0f} 백만원")

                st.markdown("---")
