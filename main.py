import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import coordinate_to_tuple
import tempfile
from PIL import Image as PILImage

# 스타일 설정
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

# 앱 제목
st.title('테스트결과서 도우미')

# 플랫폼 선택
platforms = st.multiselect('플랫폼을 선택하세요 (복수 선택 가능)', ['iOS', 'Android', '기타'])

# 업로드할 엑셀 파일
uploaded_excel = st.file_uploader("테스트결과서 엑셀 파일을 업로드하세요", type=['xlsx'])

# 이미지 업로드 설정
uploaded_images = {}
for platform in platforms:
    uploaded_images[platform] = st.file_uploader(f"{platform} 이미지 파일을 업로드하세요", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])

# 설정 입력
with st.container():
    st.write('**테스트결과서와 캡처 이미지를 업로드하면 선택한 플랫폼에 맞는 시트에 이미지를 붙여줍니다. 이미지 크기와 간격을 조절하여 적절하게 사용하세요. (모두 픽셀 단위)**')
    max_images_per_row = st.number_input('가로에 배치할 이미지 수', min_value=1, value=6)
    start_cell = st.text_input('이미지 시작 셀 주소', value='A2')
    image_width = st.number_input('이미지 가로 크기', min_value=100, value=250)
    image_height = st.number_input('이미지 세로 크기', min_value=100, value=500)
    cell_width = st.number_input('이미지 너비 간격', value=100)
    cell_height = st.number_input('이미지 높이 간격', value=20)

if uploaded_excel and platforms:
    original_filename = uploaded_excel.name
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_excel.getvalue())
        wb = load_workbook(tmp.name)

        for platform in platforms:
            if uploaded_images[platform]:
                ws = wb.create_sheet(title=f'{platform} 이미지 캡처')
                start_row, start_col = coordinate_to_tuple(start_cell)
                for index, uploaded_file in enumerate(uploaded_images[platform]):
                    row = start_row + (index // max_images_per_row) * (image_height // cell_height + 2)
                    col = start_col + (index % max_images_per_row) * (image_width // cell_width + 2)

                    pil_image = PILImage.open(uploaded_file)
                    pil_image = pil_image.resize((image_width, image_height))

                    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_img:
                        pil_image.save(tmp_img.name)
                        img = OpenpyxlImage(tmp_img.name)
                        img.anchor = ws.cell(row=row, column=col).coordinate
                        ws.add_image(img)

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
            wb.save(tmp_excel.name)
            with open(tmp_excel.name, "rb") as f:
                st.download_button('수정된 엑셀 파일 다운로드', f, file_name=original_filename)
else:
    st.write('이미지 파일, 엑셀 파일 및 플랫폼을 선택해주세요.')













