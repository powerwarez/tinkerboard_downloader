import streamlit as st
import pandas as pd
import xlrd
import requests
from io import BytesIO
from zipfile import ZipFile

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

# ZIP 파일 생성 함수
def create_zip_file(df, download_status, progress_bar):
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zip_file:
        current_folder = None
        file_counter = 1  # 파일 번호 초기화
        total_files = len(df) - 8  # 처리할 파일 개수 (8행 이후)
        completed_files = 0  # 완료된 파일 카운트

        for idx, row in df.iterrows():
            if idx < 8:  # 8행 이전은 무시
                continue
            
            no_column = row[0]  # A열 (No.열)
            attachment_url = row[5]  # F열 (첨부파일 URL)
            
            # A열이 숫자가 아닌 경우 새로운 폴더 생성 (폴더 이름 구분)
            if isinstance(no_column, str) and no_column.strip() != '':
                current_folder = no_column.strip()
                file_counter = 1  # 새로운 폴더가 생기면 파일 카운터 초기화

            # 이미지 다운로드 및 폴더에 저장
            if current_folder and isinstance(attachment_url, str) and attachment_url.startswith('http'):
                # 파일 번호에 맞게 파일명 지정
                file_name = f"image_{file_counter:03}.jpg"
                image_data = download_image(attachment_url, file_name)
                
                if image_data:
                    # 폴더별로 이미지 저장
                    folder_path = f"{current_folder}/"  # 폴더 이름
                    zip_file.writestr(f"{folder_path}{file_name}", image_data)
                    # 다운로드 상태 업데이트
                    download_status.write(f"다운로드 완료: {file_name}")
                    
                    # 진행 바 업데이트
                    completed_files += 1
                    progress_bar.progress(completed_files / total_files)

                    file_counter += 1  # 파일 번호 증가

    zip_buffer.seek(0)
    return zip_buffer

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
    st.image("https://huggingface.co/spaces/powerwarez/gailabicon/resolve/main/gailab07.png", width=50)
    st.write("제작: 교사 서동성")
    st.wirte("띵커벨 이미지가 준비되면 다운로드 버튼이 생깁니다. 잠시 기다려주세요.")
    # ZIP 파일이 이미 생성되었는지 확인
    if 'zip_file' not in st.session_state:
        st.write("이미지 다운로드 중입니다. 잠시만 기다려주세요...")

        # 다운로드 상태와 진행 바를 업데이트할 공간
        download_status = st.empty()  # 상태를 업데이트할 공간
        progress_bar = st.progress(0)  # 진행률 바

        # ZIP 파일 생성 및 세션 상태에 저장
        st.session_state.zip_file = create_zip_file(df, download_status, progress_bar)
        st.write("이미지 다운로드가 완료되었습니다!")

    # ZIP 파일 다운로드 버튼
    st.download_button(
        label="띵커벨 이미지 다운로드 (폴더별 구분된 ZIP 파일)",
        data=st.session_state.zip_file,
        file_name="띵커벨이미지.zip",
        mime="application/zip"
    )