import streamlit as st
import pandas as pd
import xlrd
import requests
from io import BytesIO
from zipfile import ZipFile
import os

# 이미지 다운로드 함수
def download_image(url, file_name):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.content
        else:
            st.write(f"이미지 다운로드 실패: {file_name} (상태 코드: {response.status_code})")
            return None
    except Exception as e:
        st.write(f"에러 발생: {e}")
        return None

# Streamlit 앱 구성
st.title("띵커벨 이미지 다운로더")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xls 형식)", type=["xls"])

if uploaded_file is not None:
    # 업로드된 파일을 읽기
    workbook = xlrd.open_workbook(file_contents=uploaded_file.read())
    
    # "보드_page_1" 시트 선택
    try:
        sheet = workbook.sheet_by_name('보드_page_1')
    except xlrd.biffh.XLRDError:
        st.error("지정한 시트 '보드_page_1'을 찾을 수 없습니다.")
        st.stop()

    # 데이터 프레임으로 변환
    data = []
    for row_idx in range(sheet.nrows):
        data.append(sheet.row_values(row_idx))
    
    df = pd.DataFrame(data)
    
    st.write("엑셀 데이터:")
    st.dataframe(df.head())

    # 다운로드할 파일을 담을 임시 저장소 (메모리 상에서 zip 파일 생성)
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zip_file:
        current_folder = None
        for idx, row in df.iterrows():
            if idx < 8:  # 8행 이전은 무시
                continue
            
            no_column = row[0]  # A열 (No.열)
            attachment_url = row[5]  # F열 (첨부파일 URL)
            
            # A열이 숫자가 아닌 경우 새로운 폴더 생성 (폴더 이름 구분)
            if isinstance(no_column, str) and no_column.strip() != '':
                current_folder = no_column.strip()

            # 이미지 다운로드 및 폴더에 저장
            if current_folder and isinstance(attachment_url, str) and attachment_url.startswith('http'):
                file_name = f"image_{idx}.jpg"
                image_data = download_image(attachment_url, file_name)
                
                if image_data:
                    # 폴더별로 이미지 저장
                    folder_path = f"{current_folder}/"  # 폴더 이름
                    zip_file.writestr(f"{folder_path}{file_name}", image_data)
        
    # zip 파일을 다운로드할 수 있게 제공
    zip_buffer.seek(0)
    st.download_button(
        label="띵커벨 이미지 다운로드 (폴더별 구분된 ZIP 파일)",
        data=zip_buffer,
        file_name="띵커벨 이미지.zip",
        mime="application/zip"
    )