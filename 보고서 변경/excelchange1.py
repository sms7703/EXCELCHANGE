# 선택 + 진행도 (20230926최종)

import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm

# 변경 내용을 엑셀 파일에 직접 적용하는 함수
def update_excel(file_path_with_date, new_file_path, replacement):
    workbook = openpyxl.load_workbook(file_path_with_date)
    sheet = workbook.active

    # 변경할 내용 적용
    for cell_coord, new_value in replacement.items():
        cell = sheet[cell_coord]
        cell.value = new_value

    # 변경 내용 저장
    workbook.save(new_file_path)
    workbook.close()

# 폴더 선택 대화 상자를 통해 디렉토리 경로 설정
root = tk.Tk()
root.withdraw()  # 기본 창 숨김

# 사용자 선택: 직접 타이핑 또는 TXT 파일 사용
user_choice = input("원하는 작업을 선택하세요:\n1 직접 타이핑\n2 TXT 파일 사용\n")

if user_choice == '1':  # 직접 타이핑 선택
    # 사용자로부터 변경할 내용 입력 받기
    replacement = {}
    while True:
        cell_coord = input("변경할 셀 좌표 입력 (예:A1, A2, A3), 끝내려면 '끝'을 입력: ")
        if cell_coord == '끝':
            break
        new_value = input(f"{cell_coord}의 새로운 내용 입력: ")
        replacement[cell_coord] = new_value

elif user_choice == '2':  # TXT 파일 사용 선택
    # 변경할 내용을 포함한 TXT 파일 선택
    txt_file_path = filedialog.askopenfilename(title="변경할 내용이 포함된 TXT 파일을 선택하세요", filetypes=[("Text files", "*.txt")])
    if not txt_file_path:
        print("TXT 파일 선택이 취소되었습니다.")
        exit(0)

    # TXT 파일 내용을 읽어서 변경할 내용으로 구성
    replacement = {}
    with open(txt_file_path, 'r', encoding='utf-8') as txt_file:
        for line in txt_file:
            parts = line.strip().split(':')
            if len(parts) == 2:
                cell_coord, new_value = parts[0], parts[1]
                replacement[cell_coord] = new_value
            else:
                print(f"잘못된 형식의 라인을 무시합니다: {line.strip()}")

# 변경할 엑셀 파일을 포함한 폴더 선택
source_directory = filedialog.askdirectory(title="변경할 엑셀 파일을 포함한 폴더를 선택하세요")
if not source_directory:
    print("경로 취소.")
    exit(0)

# 변경된 엑셀 파일을 저장할 폴더 선택
destination_directory = filedialog.askdirectory(title="변경된 엑셀 파일을 저장할 폴더를 선택하세요")
if not destination_directory:
    print("경로 취소.")
    exit(0)

# 변경 내용을 마지막으로 확인
print("다음 변경 내용을 확인합니다:")
for cell_coord, new_value in replacement.items():
    print(f"{cell_coord}: '{new_value}'로 변경")

confirm = input("변경 사항을 최종적으로 저장하시겠습니까? (Y/N): ").strip().lower()
if confirm == 'y':
    print("변경 사항을 저장합니다.")

    # 디렉토리 내의 모든 파일을 검색하고 엑셀 파일을 찾습니다.
    for root, dirs, files in os.walk(source_directory):
        for file_name in tqdm(files, desc="파일 변경 진행중"):
            if file_name.endswith('.xlsx'):
                # 파일 경로 설정
                file_path_with_date = os.path.join(root, file_name)

                # 변경 내용을 저장할 파일 경로 생성
                new_file_name = file_name  # 새 파일 이름을 원본 파일 이름으로 설정
                new_file_path = os.path.join(destination_directory, new_file_name)

                # 파일에 변경 내용을 적용하고 저장
                update_excel(file_path_with_date, new_file_path, replacement)

    print('모든 파일 변경 완료')
else:
    print("변경 사항을 취소하였습니다.")
