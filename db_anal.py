import pandas as pd
import re
import os

# 입력 파일 경로와 동일한 폴더에 결과 저장
input_file_path = 'ex1.xlsx'  # 사용자가 지정한 파일 경로
output_file_path = os.path.join(os.path.dirname(input_file_path), 'column_analysis_with_adjusted_percentage.xlsx')

# 엑셀 파일 로드
data = pd.read_excel(input_file_path)

# '무응답'을 NaN으로 대체하고, 모든 값이 NaN인 행 제거
filtered_data = data.replace("무응답", None).dropna(how="all")

# 결과를 저장할 딕셔너리 초기화
results = {}

# 함수: 대괄호 외부의 ','로만 분리
def split_outside_brackets(text):
    return re.split(r',(?![^\[]*\])', text)

# 각 열별로 개수와 비율 계산
for column in filtered_data.columns:
    # 해당 열의 유효한 값만 선택
    non_null_data = filtered_data[column].dropna()
    
    if column == '상담분야':
        # 대괄호 외부의 ','로만 분리
        split_entries = non_null_data.apply(split_outside_brackets)
        # 리스트 평탄화
        all_entries = [item.strip() for sublist in split_entries for item in sublist]
        # 개수 계산
        counts = pd.Series(all_entries).value_counts()
    else:
        counts = non_null_data.value_counts()
    
    # 비율 계산
    percentages = (counts / counts.sum() * 100)
    
    # 비율 합이 정확히 100%가 되도록 조정
    percentages = percentages.round(1)
    percentage_difference = 100 - percentages.sum()
    if percentage_difference != 0:
        percentages.iloc[-1] += percentage_difference  # 마지막 항목에 차이 반영
    
    # 결과 저장
    results[column] = pd.DataFrame({
        "Count": counts,
        "Percentage (%)": percentages
    })

# 엑셀 파일로 결과 저장
with pd.ExcelWriter(output_file_path) as writer:
    for column, df in results.items():
        # 합계 계산
        totals = pd.DataFrame({
            "Count": [df["Count"].sum()],
            "Percentage (%)": [df["Percentage (%)"].sum()]
        }, index=["Total"])
        # 합계 행 추가
        df_with_totals = pd.concat([df, totals])
        # 엑셀 시트에 저장 (시트 이름은 30자 이내로 제한)
        df_with_totals.to_excel(writer, sheet_name=column[:30])

print(f"결과 파일이 저장되었습니다: {output_file_path}")

