from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import pandas as pd
import time
import os

# ChromeOptions 생성
chrome_options = Options()

# WebDriver 설정
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# 엑셀 파일 이름 설정
excel_filename = 'restaurants_info.xlsx'

# 기존 저장된 엑셀 파일 읽기
if os.path.exists(excel_filename):
    existing_data = pd.read_excel(excel_filename)
else:
    existing_data = pd.DataFrame(columns=['title', 'phone', 'address', 'ribbon_count'])

# 접속할 웹 페이지 URL
# url = 'https://your-website-url.com'  # 여기에 실제 URL로 교체
url = 'https://www.bluer.co.kr/search?tabMode=single&searchMode=map&zone1=%EC%84%9C%EC%9A%B8%20%EA%B0%95%EB%B6%81&zone2=01.%20%ED%99%8D%EB%8C%80%EC%95%9E%2F%EC%84%9C%EA%B5%90%EB%8F%99&zone2Lat=37.553335937996756&zone2Lng=126.92488217617958&sort=&listType=&year='

driver.get(url)

# 페이지 로딩 대기
time.sleep(3)

# 새로운 데이터를 저장할 리스트
new_data = []

def extract_restaurant_info():
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
        
        # 중복 여부 확인 후 새로운 데이터에 추가
        restaurant_info = {
            'title': title,
            'phone': phone,
            'address': address,
            'ribbon_count': ribbon_count
        }
        
        if not ((existing_data['title'] == title) & 
                (existing_data['phone'] == phone) & 
                (existing_data['address'] == address)).any():
            new_data.append(restaurant_info)

def save_to_excel():
    global existing_data
    if new_data:
        # 새 데이터를 DataFrame으로 변환
        new_data_df = pd.DataFrame(new_data)
        
        # 기존 데이터와 병합하여 저장
        updated_data = pd.concat([existing_data, new_data_df], ignore_index=True)
        updated_data.to_excel(excel_filename, index=False)
        
        # 기존 데이터 업데이트
        existing_data = updated_data
        print(f"{len(new_data)} new records added to the Excel file.")
    else:
        print("No new data to save.")

def wait_for_page_to_load():
    # 페이지가 완전히 로드될 때까지 기다림
    time.sleep(3)  # 기본적인 대기 시간 (필요 시 추가 대기 방식 적용 가능)
    while True:
        try:
            driver.find_element(By.CSS_SELECTOR, 'li.rl-col.restaurant-thumb-item')
            break
        except:
            time.sleep(1)  # 페이지가 로드될 때까지 계속 대기

def navigate_and_extract():
    while True:
        # 현재 페이지의 모든 식당 정보 추출
        extract_restaurant_info()
        
        # 새로운 데이터 엑셀 저장
        save_to_excel()
        
        # 페이지 번호 찾기
        page_numbers = driver.find_elements(By.CSS_SELECTOR, 'ul.pagination li')
        page_texts = [p.text for p in page_numbers if p.text.isdigit()]
        
        if not page_texts:
            break
        
        # 페이지 번호 클릭
        for i in range(len(page_texts)):
            try:
                print(f"Clicking page: {page_texts[i]}")
                page_numbers[i].find_element(By.TAG_NAME, 'a').click()  # 페이지 번호 클릭
                wait_for_page_to_load()  # 페이지 로딩 대기
                extract_restaurant_info()  # 식당 정보 추출
                save_to_excel()  # 엑셀 저장
            except Exception as e:
                print(f"Error occurred while clicking page {page_texts[i]}: {e}")
        
        # '다음' 버튼 클릭
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, 'li.next a')
            if 'hidden' in next_button.get_attribute('class'):
                # '다음' 버튼이 숨겨져 있으면 종료
                break
            
            # '다음' 버튼 클릭
            next_button.click()
            
            # 페이지 로딩 대기
            wait_for_page_to_load()
            print("Moved to next set of pages.")
            
        except Exception as e:
            print("No more pages to navigate or error occurred.")
            break

# 탐색 및 정보 추출 함수 호출
navigate_and_extract()

# 브라우저 종료
driver.quit()
