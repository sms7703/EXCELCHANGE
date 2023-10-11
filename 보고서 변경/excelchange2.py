# 선택 + 진행도 + 로그 (20231010)
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm
from datetime import datetime  # datetime 모듈 추가

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

while True:
    # 사용자 선택: 직접 타이핑, TXT 파일 사용, 또는 종료
    user_choice = input("원하는 작업을 선택하세요:\n1 직접 타이핑\n2 TXT 파일 사용\n3 종료\n")

    if user_choice == '3':  # 종료 선택
        print("프로그램을 종료합니다.")
        break

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
        continue

    # 변경된 엑셀 파일을 저장할 폴더 선택
    destination_directory = filedialog.askdirectory(title="변경된 엑셀 파일을 저장할 폴더를 선택하세요")
    if not destination_directory:
        print("경로 취소.")
        continue

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

        # 변경 내용 로그 파일 생성 및 변경 시간 추가
        log_file_name = 'change_log.txt'
        log_file_path = os.path.join(destination_directory, log_file_name)
        with open(log_file_path, 'a') as log_file:  # 'a' 모드로 열어서 뒤에 추가
            log_file.write(f"변경 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")  # 현재 날짜와 시간 추가
            for cell_coord, new_value in replacement.items():
                log_file.write(f"{cell_coord}: '{new_value}'로 변경\n")

        print(f'변경 내용이 로그 파일 "{log_file_name}"에 저장되었습니다.')
    else:
        print("변경 사항을 취소하였습니다.")

"""
    pip 필요 리스트
    
        pip install tqdm   (진행도 추가)
        pip install openpyxl   (엑셀파일 패키지)
        pip install tinyaes (그래픽( 파이썬 기본 내장))
        
        pip install pyinstaller (실행파일로 변환)
        pip install Pyinstaller==5.7.0   (6버전시 오류 생겨서 5.7.0) 
        -------------------

        #pyinstaller --onefile excelchange(fin)3.py  
        #pyinstaller --key="12345" excelchange(fin)3.py (불가)

"""