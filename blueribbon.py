# from selenium import webdriver
from selenium.webdriver.common.by import By
# from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
# ChromeOptions 생성
chrome_options = Options()

# WebDriver 설정
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
# ChromeDriver 경로 설정 (다운로드한 ChromeDriver 경로로 설정)
# driver_path = '/path/to/chromedriver'

# 브라우저 드라이버 실행
# driver = webdriver.Chrome(executable_path=driver_path)
# driver = webdriver.Chrome(ChromeDriverManager().install())
# 접속할 웹 페이지 URL
url = 'https://www.bluer.co.kr/search?tabMode=single&searchMode=map&zone1=%EC%84%9C%EC%9A%B8%20%EA%B0%95%EB%B6%81&zone2=01.%20%ED%99%8D%EB%8C%80%EC%95%9E%2F%EC%84%9C%EA%B5%90%EB%8F%99&zone2Lat=37.553335937996756&zone2Lng=126.92488217617958&sort=&listType=&year='

# 웹 페이지 열기
driver.get(url)

time.sleep(5)
# 페이지 내 모든 <li> 태그를 찾음
restaurant_items = driver.find_elements(By.CLASS_NAME, 'restaurant-thumb-item')

# 제목과 주소 추출
for item in restaurant_items:
    # 제목 추출 (h3 태그 내 텍스트)
    title = item.find_element(By.TAG_NAME, 'h3').text
    
    # 주소 추출 (dt 태그에서 "주소"라는 텍스트를 찾고, 다음 dd 태그의 텍스트 추출)
    dt_elements = item.find_elements(By.TAG_NAME, 'dt')
    address = 'No address found'
    for dt in dt_elements:
        if '주소' in dt.text:
            address = dt.find_element(By.XPATH, 'following-sibling::dd').text
            break
    
    # 결과 출력
    print(f'제목: {title}')
    print(f'주소: {address}')
    print('-' * 50)

# 브라우저 종료
driver.quit()
