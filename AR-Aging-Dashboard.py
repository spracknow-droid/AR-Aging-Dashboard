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

            # --- 레이아웃: 너비 비율로 3개 컬럼 생성 (20:40:40) ---
            col1, col2, col3 = st.columns([20, 40, 40])

            # 칼럼 1: 총 잔액 표시 및 파이차트
            with col1:
                st.subheader("매출채권")
                total_receivables = df['채권금액(원화)'].sum()
                # 총 매출채권 잔액을 백만원 단위로 표기 (소수점 제거)
                total_receivables_mil = total_receivables / 1_000_000
                st.metric(label="총 잔액", value=f"{total_receivables_mil:,.0f} 백만원")

                st.markdown("---")

                # 통화별 데이터 집계
                currency_summary = df.groupby('환종')['채권금액(원화)'].sum().reset_index()

                # 파이 차트 생성 및 표시
                fig_pie = px.pie(
                    currency_summary,
                    values='채권금액(원화)',
                    names='환종',
                    title='통화별 잔액 비율',
                    color_discrete_sequence=px.colors.qualitative.Plotly
                )
                fig_pie.update_traces(
                    marker=dict(line=dict(color='#000000', width=1)),
                    textinfo='percent+label',
                    pull=[0.05, 0, 0, 0]
                )
                st.plotly_chart(fig_pie, use_container_width=True)

            # 칼럼 2: 채권 연령 현황 시각화
            with col2:
                st.subheader("채권 연령 현황")

                # 채권 연령을 계산하는 함수
                def get_age_group(invoice_date, standard_date):
                    days_diff = (standard_date - invoice_date.date()).days
                    if days_diff <= 30:
                        return '<1 mo.'
                    elif days_diff <= 90:
                        return '1-3 mo.'
                    elif days_diff <= 180:
                        return '3-6 mo.'
                    elif days_diff <= 365:
                        return '6-12 mo.'
                    else:
                        return '>1 yr.'

                # '채권 연령' 컬럼 추가
                df['채권 연령'] = df['매출일자'].apply(lambda x: get_age_group(x, standard_date))

                # 연령별 데이터 집계
                age_order = ['<1 mo.', '1-3 mo.', '3-6 mo.', '6-12 mo.', '>1 yr.']
                age_summary = df.groupby('채권 연령')['채권금액(원화)'].sum().reindex(age_order, fill_value=0).reset_index()

                # 막대 그래프 생성
                fig_bar = px.bar(
                    age_summary,
                    x='채권 연령',
                    y='채권금액(원화)',
                    title='매출채권 연령별 잔액',
                    labels={'채권 연령': '채권 연령', '채권금액(원화)': '합계 금액 (원)'},
                    color_discrete_sequence=px.colors.qualitative.Plotly
                )

                fig_bar.update_traces(texttemplate='%{y:,.0f}', textposition='outside')
                fig_bar.update_layout(
                    xaxis={'categoryorder': 'array', 'categoryarray': age_order},
                    uniformtext_minsize=7,
                    uniformtext_mode='hide',
                    yaxis={'tickformat': ','},
                    # 여기에 마진(여백) 설정 추가
                    margin=dict(l=150, r=20, t=50, b=70) # 예시 값. 필요에 따라 조정 가능

                )

                # plotly_events를 사용하여 그래프 표시 및 클릭 이벤트 처리
                selected_age_group = plotly_events(fig_bar, click_event=True, key="bar_chart_age")

                # 클릭 이벤트에 따른 상세 정보 표 생성
                if selected_age_group:
                    st.markdown("---")
                    clicked_age = selected_age_group[0]['x']
                    st.subheader(f"'{clicked_age}' 상세 내역")
                    
                    # 클릭된 연령 그룹에 해당하는 원본 데이터 필터링
                    age_group_df = df[df['채권 연령'] == clicked_age].copy()
                    
                    # '채권금액(원화)'를 기준으로 내림차순 정렬
                    age_group_df.sort_values(by='채권금액(원화)', ascending=False, inplace=True)
                    
                    # 필요한 컬럼만 선택하고 소계 행 추가
                    detail_df = age_group_df[['거래처명', '채권금액(원화)', '채권금액(외화)']].copy()
                    
                    # 소계 계산
                    total_krw = detail_df['채권금액(원화)'].sum()
                    total_foreign = detail_df['채권금액(외화)'].sum()
                    
                    # 소계 행 생성
                    subtotal_row = pd.DataFrame([['**소계**', total_krw, total_foreign]], columns=detail_df.columns)
                    
                    # 소계 행을 데이터프레임에 추가
                    detail_df = pd.concat([detail_df, subtotal_row], ignore_index=True)
                    
                    # 숫자 포맷팅
                    detail_df['채권금액(원화)'] = detail_df['채권금액(원화)'].apply(lambda x: f'{x:,.0f}')
                    detail_df['채권금액(외화)'] = detail_df['채권금액(외화)'].apply(lambda x: f'{x:,.2f}')
                    
                    # 결과 표 표시
                    st.dataframe(detail_df, use_container_width=True, hide_index=True)

            # 칼럼 3: 6개월 이상 미회수 채권 Top 5 표시
            with col3:
                st.subheader("6개월 이상 미회수 채권 (Top 5)")

                threshold_date = standard_date - relativedelta(months=6)
                overdue_df = df[df['매출일자'].dt.date <= threshold_date].copy()

                # 거래처명 약식 표기를 위한 딕셔너리
                name_mapping = {
                    'Jiangsu Shekoy Semiconductor New Material Co., Ltd': 'Shekoy',
                    'UP Electronic Materials(Taiwan) Limited': 'UPTW',
                    'Changxin Memory Technologies, Inc. (CXMT)': 'CXMT',
                    'CHJS(Chengdu High tech Jin Science)': 'CHJS',
                    'SemiLink Materials, LLC.' : 'SemiLink'
                }
                # 딕셔너리를 사용하여 거래처명 컬럼의 값을 변경
                overdue_df['거래처명'] = overdue_df['거래처명'].replace(name_mapping)

                # '채권금액(원화)'를 기준으로 집계하도록 수정
                top5_df = overdue_df.groupby('거래처명')['채권금액(원화)'].sum().reset_index()
                top5_df = top5_df.sort_values(by='채권금액(원화)', ascending=False).head(5)

                # y축 정렬을 위해 거래처명을 카테고리형으로 변환
                top5_df['거래처명'] = pd.Categorical(top5_df['거래처명'], categories=top5_df['거래처명'].unique(), ordered=True)

                if not top5_df.empty:
                    # 가로 막대그래프 생성 및 표시
                    fig2 = px.bar(
                        top5_df,
                        y='거래처명',
                        x='채권금액(원화)',
                        orientation='h',
                        title='6개월 이상 미회수 채권 TOP 5',
                        labels={'거래처명': '거래처명', '채권금액(원화)': '합계 금액 (원)'},
                        color_discrete_sequence=px.colors.qualitative.Bold
                    )
                    
                    fig2.update_yaxes(categoryorder='total ascending')
                    fig2.update_traces(texttemplate='%{x:,.0f}', textposition='outside', cliponaxis=False)
                    fig2.update_layout(
                        margin=dict(t=50, l=120, r=20, b=80),
                        uniformtext_minsize=7,
                        uniformtext_mode='hide'
                    )

                    # plotly_events를 사용하여 그래프 표시 및 클릭 이벤트 처리
                    selected_customer = plotly_events(fig2, click_event=True, key="bar_chart_top5")

                    # 클릭 이벤트에 따른 상세 정보 표 생성
                    if selected_customer:
                        st.markdown("---")
                        # 클릭된 막대의 거래처명 가져오기
                        clicked_customer = selected_customer[0]['y']
                        st.subheader(f"'{clicked_customer}' 상세 내역")

                        # 클릭된 거래처에 해당하는 원본 데이터 필터링
                        customer_detail_df = overdue_df[overdue_df['거래처명'] == clicked_customer].copy()

                        # 필요한 컬럼만 선택하여 포맷팅
                        customer_detail_df['매출일자'] = customer_detail_df['매출일자'].dt.strftime('%Y-%m-%d')
                        customer_detail_df['채권금액(원화)'] = customer_detail_df['채권금액(원화)'].apply(lambda x: f'{x:,.0f}')

                        # 결과 표 표시
                        st.dataframe(customer_detail_df[['매출일자', '채권금액(원화)', '적요']], use_container_width=True, hide_index=True)

                else:
                    st.info("6개월 이상 미회수된 채권이 없습니다.")

        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
            st.info("엑셀 파일의 '매출일자', '채권금액(원화)', '환종', '채권금액(외화)', '거래처명', '적요' 컬럼 형식을 확인해주세요.")

    else:
        st.error("필수 컬럼이 누락되었습니다. '매출일자', '채권금액(원화)', '환종', '채권금액(외화)', '거래처명', '적요' 컬럼을 모두 포함해주세요.")
