import streamlit as st
import os
import pandas as pd
import xlrd
import requests
from pathlib import Path

# 바탕화면 경로 가져오기
desktop_path = Path(os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop'))

# 폴더 생성 함수
def create_folder(folder_name):
    folder_path = desktop_path / folder_name
    folder_path.mkdir(parents=True, exist_ok=True)
    st.write(f"폴더 생성: {folder_path}")
    return folder_path

# 이미지 다운로드 함수
def download_image(url, folder, file_name):
    file_path = folder / file_name
    try:
        response = requests.get(url)
        if response.status_code == 200:
            with open(file_path, 'wb') as f:
                f.write(response.content)
            st.write(f"이미지 다운로드 완료: {file_name}")
        else:
            st.write(f"이미지 다운로드 실패: {file_name} (상태 코드: {response.status_code})")
    except Exception as e:
        st.write(f"에러 발생: {e}")

# Streamlit 앱 구성
st.title("띵커벨 보드 이미지 다운로더")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xls 형식)", type=["xls"])

if uploaded_file is not None:
    # 업로드된 파일을 읽기, 특정 시트 "보드_page_1" 사용
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

    # 데이터 처리
    current_folder = None
    for idx, row in df.iterrows():
        if idx < 8:  # 8행 이전은 무시
            continue
        
        if len(row) > 5:  # F열(5번째 열)이 있는지 확인
            no_column = row[0]  # A열 (No.열)
            attachment_url = row[5]  # F열 (첨부파일 URL)
            
            if no_column == '※보드 소유자 닉네임 앞에는 * 표시가 붙습니다.':
                st.write("작업이 완료되었습니다.")
                break
            
            # A열이 숫자가 아닌 경우 폴더 생성
            if isinstance(no_column, str) and no_column.strip() != '':
                current_folder = create_folder(no_column.strip())
            
            # 이미지 다운로드
            if current_folder and isinstance(attachment_url, str) and attachment_url.startswith('http'):
                file_name = f"image_{idx-8}.jpg"
                download_image(attachment_url, current_folder, file_name)
            
            # A열이 비어 있으면 종료
            if no_column == '':
                st.write("작업이 완료되었습니다.")
                break
        else:
            st.warning(f"{idx} 행에 F열이 없습니다. 건너뜁니다.")