from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os

# ChromeOptions 생성
chrome_options = Options()

# WebDriver 설정
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# 엑셀 파일에서 URL과 파일 이름 읽기
input_excel_filename = 'urls_and_filenames.xlsx'  # 여기에 입력할 엑셀 파일 이름을 지정하세요
urls_df = pd.read_excel(input_excel_filename)

# 각 URL에 대해 웹 스크래핑 수행
for index, row in urls_df.iterrows():
    url = row['URL']  # URL이 있는 열 이름
    excel_filename = row['Filename']  # 파일 이름이 있는 열 이름

    driver.get(url)

    # 페이지 로딩 대기
    time.sleep(5)

    # 제목, 전화번호, 주소, 리본 개수를 저장할 리스트
    restaurants_info = []

    while True:
        # 현재 페이지에서 모든 레스토랑 항목 추출
        restaurants = driver.find_elements(By.CSS_SELECTOR, 'li.rl-col.restaurant-thumb-item')
        
        for restaurant in restaurants:
            # 제목 추출
            try:
                title = restaurant.find_element(By.CSS_SELECTOR, 'h3').text
            except:
                title = 'No Title'
            
            # 전화번호 추출
            try:
                phone = restaurant.find_element(By.XPATH, ".//dt[contains(text(),'전화')]/following-sibling::dd/a").text
            except:
                phone = 'No Phone Number'
            
            # 주소 추출
            try:
                address = restaurant.find_element(By.XPATH, ".//dt[contains(text(),'주소')]/following-sibling::dd").text
            except:
                address = 'No Address'
            
            # 리본 이미지 개수 추출
            try:
                ribbons = restaurant.find_elements(By.CSS_SELECTOR, 'img.img-ribbon')
                ribbon_count = len(ribbons)
            except:
                ribbon_count = 0
            
            # 제목, 전화번호, 주소, 리본 개수 저장
            restaurant_info = {
                'title': title,
                'phone': phone,
                'address': address,
                'ribbon_count': ribbon_count
            }
            
            # 중복 여부 확인 후 새로운 데이터에 추가
            if restaurant_info not in restaurants_info:
                restaurants_info.append(restaurant_info)

        # 현재 데이터 엑셀 파일로 저장
        if restaurants_info:
            # 기존 데이터 로드
            if os.path.exists(excel_filename):
                existing_data = pd.read_excel(excel_filename)
            else:
                existing_data = pd.DataFrame(columns=['title', 'phone', 'address', 'ribbon_count'])

            # 중복 데이터 제거
            new_data_df = pd.DataFrame(restaurants_info)
            merged_data = pd.concat([existing_data, new_data_df]).drop_duplicates().reset_index(drop=True)
            merged_data.to_excel(excel_filename, index=False)
            print(f"Data saved to {excel_filename}")

        # 다음 페이지 버튼을 찾고 클릭
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, 'li.next a')
            next_button.click()
            
            # 다음 페이지 로딩 대기
            time.sleep(5)
            
        except:
            # '다음' 버튼이 없는 경우 루프를 종료
            print("No more pages to navigate.")
            break

# 브라우저 종료
driver.quit()
