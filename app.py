import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import openpyxl

# 데이터 파일 경로
file_path = 'cafe.xlsx'

# 데이터 불러오기
try:
    cafe_data = pd.read_excel(file_path)
except FileNotFoundError:
    st.error(f"파일을 찾을 수 없습니다: {file_path}")
    st.stop()
except Exception as e:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다: {e}")
    st.stop()

# order_date를 datetime 형식으로 변환
cafe_data['order_date'] = pd.to_datetime(cafe_data['order_date'], errors='coerce')

# 데이터가 올바르게 로드되었는지 확인
if cafe_data.empty:
    st.error("데이터가 비어 있습니다. 올바른 데이터 파일을 업로드했는지 확인하십시오.")
    st.stop()

# order_date가 NaT인 경우 제외
cafe_data = cafe_data.dropna(subset=['order_date'])

# 연도와 월 컬럼 추가
cafe_data['year'] = cafe_data['order_date'].dt.year
cafe_data['month'] = cafe_data['order_date'].dt.month

# Streamlit 대시보드 설정
st.title('카페 매출 대시보드')

# 연도와 제품 선택
years = sorted(cafe_data['year'].unique())
if not years:
    st.error("사용 가능한 연도가 없습니다.")
    st.stop()

products = sorted(cafe_data['item'].unique())
if not products:
    st.error("사용 가능한 제품이 없습니다.")
    st.stop()

selected_year = st.sidebar.selectbox('연도 선택', years)
selected_product = st.sidebar.selectbox('제품 선택', products)

# 선택된 연도와 제품에 따라 데이터 필터링
filtered_data = cafe_data[(cafe_data['year'] == selected_year) & 
                          (cafe_data['item'] == selected_product)]

# 필터링된 데이터가 비어 있는 경우 처리
if filtered_data.empty:
    st.warning(f"{selected_year}년에는 {selected_product}의 데이터가 없습니다.")
else:
    # 월별 매출 데이터 계산
    monthly_sales = filtered_data.groupby('month')['price'].sum().reset_index()

    # 매출 표 출력
    st.write(f'{selected_year}년 {selected_product} 매출 데이터')
    st.write(filtered_data)

    # 막대 그래프 그리기
    st.write(f'{selected_year}년 {selected_product} 월별 매출')
    st.bar_chart(monthly_sales.set_index('month'))

# 메뉴 선택
menu = st.selectbox('메뉴 선택', products)
menu_data = cafe_data[(cafe_data['year'] == selected_year) & (cafe_data['item'] == menu)]
menu_sales = menu_data.groupby('month')['price'].sum().reset_index()

if menu_sales.empty:
    st.warning(f"{selected_year}년에는 {menu}의 데이터가 없습니다.")
else:
    st.write(f'{selected_year}년 {menu} 월별 매출')
    st.line_chart(menu_sales.set_index('month'))

# 매출 분석
st.subheader('매출 분석')
yearly_sales = cafe_data.groupby('year')['price'].sum().reset_index()
monthly_sales_overall = cafe_data.groupby(['year', 'month'])['price'].sum().reset_index()

if yearly_sales.empty:
    st.warning("연도별 매출 데이터가 없습니다.")
else:
    st.write('연도별 매출 분석')
    st.bar_chart(yearly_sales.set_index('year'))

if monthly_sales_overall.empty:
    st.warning(f"{selected_year}년의 월별 매출 데이터가 없습니다.")
else:
    st.write(f'{selected_year}년 월별 매출 분석')
    monthly_sales_year = monthly_sales_overall[monthly_sales_overall['year'] == selected_year]
    st.line_chart(monthly_sales_year.set_index('month')['price'])

# 가장 많이 팔린 음료와 가장 적게 팔린 음료 분석
st.subheader('가장 많이 팔린 음료와 가장 적게 팔린 음료 분석')
yearly_sales_by_product = cafe_data[cafe_data['year'] == selected_year].groupby('item')['price'].sum().reset_index()

if yearly_sales_by_product.empty:
    st.warning(f"{selected_year}년의 제품별 매출 데이터가 없습니다.")
else:
    most_sold_product = yearly_sales_by_product.loc[yearly_sales_by_product['price'].idxmax()]
    least_sold_product = yearly_sales_by_product.loc[yearly_sales_by_product['price'].idxmin()]

    st.write(f"**{selected_year}년 가장 많이 팔린 음료:** {most_sold_product['item']} (매출: {most_sold_product['price']:,}원)")
    st.write(f"**{selected_year}년 가장 적게 팔린 음료:** {least_sold_product['item']} (매출: {least_sold_product['price']:,}원)")
