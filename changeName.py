# changeName.py

import os
import re

# 파일들이 있는 폴더 경로 지정
input_path = r"C:\khk\docuwork\input"
output_path = r"C:\khk\docuwork\output"

# 폴더 내 jpg 파일만 정렬해서 가져오기
files = sorted([
    f for f in os.listdir(input_path)
    if f.endswith('.jpg') and f.startswith('svm')
])

for idx, filename in enumerate(files, start=1):
    # 정규식으로 페이지 번호 추출 (예: 285p)
    match = re.search(r'_(\d+p)', filename)
    if not match:
        print(f"패턴 불일치, 건너뜀: {filename}")
        continue

    page = match.group(1)  # "285p"
    new_name = f"[데에빅분기]{page}_svm{idx:02d}.jpg"

    old_path = os.path.join(input_path, filename)
    new_path = os.path.join(output_path, new_name)

    os.rename(old_path, new_path)
    print(f"{filename}  →  {new_name}")

print("완료!")