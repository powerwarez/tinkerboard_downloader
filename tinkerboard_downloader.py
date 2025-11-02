import streamlit as st
import pandas as pd
import xlrd
import requests
from io import BytesIO
from zipfile import ZipFile
import os
import time

# 세션 생성 (쿠키 유지)
session = requests.Session()
session_initialized = False

# 세션 초기화 함수 (메인 페이지 방문하여 쿠키 받기)
def initialize_session():
    global session_initialized
    if session_initialized:
        return
    
    try:
        st.write("🌐 띵커벨 사이트에 접속 중...")
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
            'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1'
        }
        
        # 메인 페이지 방문
        response = session.get('https://www.tkbell.co.kr/', headers=headers, timeout=30)
        st.write(f"✅ 메인 페이지 접속 완료 (상태 코드: {response.status_code})")
        st.write(f"🍪 받은 쿠키: {dict(session.cookies)}")
        
        # 이미지 서버 도메인도 방문
        response = session.get('https://b.tkbell.co.kr/', headers=headers, timeout=30)
        st.write(f"✅ 이미지 서버 접속 완료 (상태 코드: {response.status_code})")
        st.write(f"🍪 현재 쿠키: {dict(session.cookies)}")
        
        time.sleep(1)  # 1초 대기
        session_initialized = True
    except Exception as e:
        st.write(f"⚠️ 세션 초기화 에러: {e}")

# 이미지 다운로드 함수
def download_image(url, file_name):
    try:
        # 첫 번째 다운로드 시도 전에 세션 초기화
        initialize_session()
        
        # 디버깅: URL 출력
        st.write(f"🔍 시도 중인 URL: {url}")
        st.write(f"📝 파일명: {file_name}")
        
        # User-Agent 헤더 추가하여 브라우저처럼 보이게 함
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Referer': 'https://www.tkbell.co.kr/',
            'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8',
            'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Sec-Fetch-Dest': 'image',
            'Sec-Fetch-Mode': 'no-cors',
            'Sec-Fetch-Site': 'same-site',
            'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"'
        }
        
        st.write(f"📤 요청 헤더: {headers}")
        st.write(f"🍪 사용 중인 쿠키: {dict(session.cookies)}")
        
        # Session을 사용하여 쿠키 유지
        response = session.get(url, headers=headers, timeout=30, allow_redirects=True)
        
        st.write(f"📥 응답 상태 코드: {response.status_code}")
        st.write(f"📥 응답 헤더: {dict(response.headers)}")
        
        if response.status_code == 200:
            st.write(f"✅ 다운로드 성공: {file_name}")
            time.sleep(0.5)  # 요청 간 0.5초 지연
            return response.content
        else:
            st.write(f"❌ 이미지 다운로드 실패: {file_name} (상태 코드: {response.status_code})")
            time.sleep(1)  # 실패 시 1초 지연
            return None
    except Exception as e:
        st.write(f"⚠️ 에러 발생: {e}")
        import traceback
        st.write(f"상세 에러: {traceback.format_exc()}")
        return None

# ZIP 파일 생성 함수
def create_zip_file(df, download_status, progress_bar, excel_file_name):
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
                # URL에서 파일 이름 추출
                file_name = attachment_url.split('/')[-1]  # URL의 마지막 부분을 파일명으로 사용
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
    # 엑셀 파일 이름 가져오기 (확장자 제거)
    excel_file_name = os.path.splitext(uploaded_file.name)[0]
    
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
    # Arrow 직렬화 에러 방지를 위해 문자열로 변환
    display_df = df.head().astype(str)
    st.dataframe(display_df)
    st.image("https://huggingface.co/spaces/powerwarez/gailabicon/resolve/main/gailab07.png", width=50)
    st.write("제작: 교사 서동성")
    st.write("띵커벨 이미지가 준비되면 다운로드 버튼이 생깁니다. 잠시 기다려주세요.")
    
    # ZIP 파일이 이미 생성되었는지 확인
    if 'zip_file' not in st.session_state:
        st.write("이미지 다운로드 중입니다. 잠시만 기다려주세요...")

        # 다운로드 상태와 진행 바를 업데이트할 공간
        download_status = st.empty()  # 상태를 업데이트할 공간
        progress_bar = st.progress(0)  # 진행률 바

        # ZIP 파일 생성 및 세션 상태에 저장
        st.session_state.zip_file = create_zip_file(df, download_status, progress_bar, excel_file_name)
        st.write("이미지 다운로드가 완료되었습니다!")

    # ZIP 파일 다운로드 버튼
    st.download_button(
        label="띵커벨 이미지 다운로드 (폴더별 구분된 ZIP 파일)",
        data=st.session_state.zip_file,
        file_name="띵커벨이미지.zip",
        mime="application/zip"
    )