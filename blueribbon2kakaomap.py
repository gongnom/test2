import pandas as pd
import requests
import json
import os
import time

# 카카오맵 API 키 설정
API_KEY = "286444fc54e840821eadab83bbbcfa1e"

# 즐겨찾기 그룹 생성
def create_favorite_group(group_name):
    url = "https://dapi.kakao.com/v2/local/favorite/group/create.json"
    headers = {"Authorization": f"KakaoAK {API_KEY}"}
    data = {"name": group_name}
    response = requests.post(url, headers=headers, data=data)
    
    if response.status_code != 200:
        print(f"Error creating group: {response.status_code}")
        print(f"Response content: {response.text}")
        return None

    try:
        return response.json()['id']
    except KeyError:
        print(f"Unexpected response format: {response.json()}")
        return None

# 즐겨찾기 추가
def add_favorite(group_id, name, address):
    url = "https://dapi.kakao.com/v2/local/search/address.json"
    headers = {"Authorization": f"KakaoAK {API_KEY}"}
    params = {"query": address}
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code != 200 or not response.json()['documents']:
        print(f"Error finding address: {address}")
        print(f"Response: {response.text}")
        return None

    result = response.json()['documents'][0]

    url = "https://dapi.kakao.com/v2/local/favorite/add.json"
    headers = {"Authorization": f"KakaoAK {API_KEY}"}
    data = {
        "group_id": group_id,
        "name": name,
        "x": result['x'],
        "y": result['y'],
        "address": address
    }
    response = requests.post(url, headers=headers, data=data)
    
    if response.status_code != 200:
        print(f"Error adding favorite: {response.status_code}")
        print(f"Response content: {response.text}")
        return None

    return response.json()

# 엑셀 파일 처리
def process_excel_file(file_path):
    df = pd.read_excel(file_path)
    group_name = os.path.splitext(os.path.basename(file_path))[0]
    group_id = create_favorite_group(group_name)
    
    if group_id is None:
        print(f"Failed to create group for {file_path}")
        return

    for _, row in df.iterrows():
        name = row['title']  # '장소명'에서 'title'로 변경
        address = row['address']  # '주소'에서 'address'로 변경
        try:
            result = add_favorite(group_id, name, address)
            if result:
                print(f"Added {name}: {result}")
            else:
                print(f"Failed to add {name}")
        except Exception as e:
            print(f"Error adding {name}: {str(e)}")
        time.sleep(1)  # API 요청 간 1초 대기

# 메인 실행 코드
def main():
    current_dir = os.getcwd()
    excel_files = [f for f in os.listdir(current_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
    
    for file in excel_files:
        print(f"Processing file: {file}")
        process_excel_file(os.path.join(current_dir, file))
        print(f"Finished processing {file}\n")

if __name__ == "__main__":
    main()