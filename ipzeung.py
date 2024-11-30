import fitz  # PyMuPDF
from PIL import Image
import glob
import os

# 이미지 크기 및 위치 설정 함수 (사용자가 원하는 대로 조정 가능)
def get_image_position_and_size(pdf_page, img_width=100, img_height=100):
    page_width = pdf_page.rect.width
    page_height = pdf_page.rect.height
    # 오른쪽 상단에 배치되도록 위치 설정
    x0 = page_width - img_width - 10  # 오른쪽 여백
    y0 = 10  # 상단 여백
    return fitz.Rect(x0, y0, x0 + img_width, y0 + img_height)

# PDF 파일에 이미지를 추가하는 함수
def add_image_to_pdf(pdf_path, image_path, output_path):
    # PDF 파일 열기
    pdf_document = fitz.open(pdf_path)
    
    # 이미지 불러오기
    image = Image.open(image_path)
    img_width, img_height = image.size
    
    # 첫 페이지에 이미지 추가
    page = pdf_document[0]
    rect = get_image_position_and_size(page, img_width, img_height)
    page.insert_image(rect, filename=image_path)
    
    # 수정된 PDF 저장
    pdf_document.save(output_path)
    pdf_document.close()

# 폴더 내 PDF 파일을 찾고 이미지 추가 수행
pdf_files = glob.glob("노*.어쩌고.pdf")  # '노'로 시작하고 '어쩌고.pdf'로 끝나는 파일 찾기

for pdf_file in pdf_files:
    base_name = os.path.splitext(pdf_file)[0]  # 예: 노1 또는 노1-1
    image_path = base_name + ".jpg"  # 기본 이미지 파일 경로 설정

    # 이미지 파일이 존재하지 않으면 사용자에게 경로를 입력받음
    if not os.path.exists(image_path):
        image_path = input(f"{pdf_file}에 추가할 이미지 파일 경로를 입력하세요: ")

        # 입력한 경로가 유효하지 않으면 다시 물어봄
        while not os.path.exists(image_path):
            print("유효하지 않은 경로입니다. 다시 입력해 주세요.")
            image_path = input(f"{pdf_file}에 추가할 이미지 파일 경로를 입력하세요: ")

    output_pdf = base_name + "_output.pdf"  # 출력 파일명 설정
    add_image_to_pdf(pdf_file, image_path, output_pdf)
    print(f"{pdf_file}의 첫 페이지에 {image_path}를 추가하여 {output_pdf}로 저장했습니다.")
