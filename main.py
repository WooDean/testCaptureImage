import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import coordinate_to_tuple
import tempfile
from PIL import Image as PILImage

# 페이지 전체 스타일 설정
st.markdown(
"""
<style>
body {
    color: #333;
    background-color: #f0f2f6;
    font-family: Arial, sans-serif;
    margin: 0;
    padding: 0;
}
h1 {
    color: #333;
    text-align: center;
    margin-bottom: 20px;
}
.container {
    max-width: 800px;
    margin: auto;
    padding: 20px;
    background-color: #fff;
    border-radius: 8px;
    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
}
.widget-label {
    color: #555;
}
.upload-section {
    margin-bottom: 20px;
}
.upload-container {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
}
.uploaded-image {
    width: 100px;
    height: auto;
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 0 5px rgba(0, 0, 0, 0.1);
}
.download-btn {
    margin-top: 20px;
    text-align: center;
}
</style>
""",
unsafe_allow_html=True
)

# 스트림릿 앱 제목 설정
st.title('테스트결과서 마지막 시트에 캡처 삽입하기')

# 사용자 설정 입력
with st.container():
    # st.write('이미지를 업로드하면, 가로 최대 설정한 장수까지 배열하여 삽입합니다.')
    st.write('**이미지 크기와 간격을 조절하여 적절하게 사용하세요! (모두 픽셀 단위) **')
    max_images_per_row = st.number_input('가로에 배치할 이미지 수', min_value=1, value=6)
    cell_width = st.number_input('이미지 너비 간격: 너비를 크게 잡을수록 이미지 간 가로 간격이 좁아져요', value=100)
    cell_height = st.number_input('이미지 높이 간격: 높이를 크게 잡을수록 이미지 간 세로 간격이 좁아져요', value=20)
    start_cell = st.text_input('이미지 시작 셀 주소 (예: A2)', value='A2')
    image_width = st.number_input('이미지 가로 크기', min_value=100, value=250)
    image_height = st.number_input('이미지 세로 크기', min_value=100, value=500)

# 사용자가 업로드할 엑셀 파일
uploaded_excel = st.file_uploader("테스트결과서 엑셀 파일을 업로드하세요", type=['xlsx'])

# 이미지 업로드
uploaded_files = st.file_uploader("이미지 파일을 선택하세요(여러 개 선택 가능)", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])

if uploaded_files and uploaded_excel:
    original_filename = uploaded_excel.name  # 업로드한 파일의 원래 이름
    # 로드된 워크북 처리
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_excel.getvalue())
        wb = load_workbook(tmp.name)
    
    # 새로운 시트를 생성하여 이미지를 붙여넣을 준비
    ws = wb.create_sheet(title='이미지 캡처')

    image_size = (image_width, image_height)  # 사용자가 입력한 크기를 토대로 이미지의 크기 설정
    start_row, start_col = coordinate_to_tuple(start_cell)  # 시작 셀 주소로부터의 행, 열 좌표 추출

    for index, uploaded_file in enumerate(uploaded_files):
        # 이미지의 행 위치와 열 위치 계산
        row = start_row + (index // max_images_per_row) * (image_height // cell_height + 2)
        col = start_col + (index % max_images_per_row) * (image_width // cell_width + 2)

        # 이미지를 PIL로 열고 크기 조정
        pil_image = PILImage.open(uploaded_file)
        pil_image = pil_image.resize(image_size)

        # 임시 파일로 이미지 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            pil_image.save(tmp.name)
            img = OpenpyxlImage(tmp.name)
            img.anchor = ws.cell(row=row, column=col).coordinate
            ws.add_image(img)

    # 워크북 저장
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
        wb.save(tmp_excel.name)
        # 사용자가 엑셀 파일을 다운로드할 수 있도록 함
        with open(tmp_excel.name, "rb") as f:
            st.download_button('수정된 엑셀 파일 다운로드', f, file_name=original_filename)

else:
    st.write('이미지 파일과 엑셀 파일을 업로드해주세요.')













