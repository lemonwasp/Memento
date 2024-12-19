# 주의 : chromedriver의 버전이 131.0.6778.205(윈도우 64비트)이므로 크롬 설정을 이에 맞게 업데이트 해주어야함


import openpyxl as xl
import random
import os
import PySimpleGUI as sg
from selenium import webdriver # 브라우저를 제어
from selenium.webdriver.common.by import By # CSS 셀렉터나 태그 이름 등으로 페이지 요소를 찾는 방법을 제공
from selenium.webdriver.chrome.service import Service # WebDriver 서비스(ex. ChromeDriver)를 관리
from selenium.webdriver.chrome.options import Options # 브라우저의 설정(ex. 헤드리스 모드)을 구성
from selenium.webdriver.support.ui import WebDriverWait # 대기를 걸 수 있음
from selenium.webdriver.support import expected_conditions as EC # 예상조건 WebDriverWait와 같이 사용


# Selenium WebDriver 초기화
def init_driver():
    chrome_options = Options() 
    chrome_options.add_argument("--headless") # UI 없이 Chrome 실행 (헤드리스 모드)
    chrome_options.add_argument("--disable-gpu") #GPU가속을 비활성화 (헤드리스 모드에서 발생할 수 있는 문제 방지)
    chrome_options.add_argument("--no-sandbox") # 일부 환경(Docker 등)에서 권한 문제를 방지
    service = Service(os.path.join(os.path.dirname(__file__), "chromedriver.exe")) # ChromeDriver 실행 파일의 경로 지정, Selenium과 Chrome 사이의 브릿지 역할
    return webdriver.Chrome(service=service, options=chrome_options) # 지정된 옵션과 서비스를 사용해 Selenium WebDriver를 Chrome브라우저를 초기화


# 특정 단어의 예문 가져오기
def fetch_example(driver, word, example_index=4): # index초반의 예문은 예문으로써 적절하지 않은 경우가 많아서 적당히 뒤쪽의 index를 사용
    try:
        # 케임브리지 영어사전 URL 설정정
        url = f"https://dictionary.cambridge.org/dictionary/english/{word}"
        driver.get(url)
        
        # 모든 예문 요소 가져오기
        examples = WebDriverWait(driver, 10).until( # 불러올 때까지 대기해줌
            # 직접 html/css구조를 열어서 살펴보니 이 Selector 안에 예문이 들어가 있었음
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.deg")) 
        )
        
        # 예문 텍스트 추출
        example_sentences = [example.text.strip() for example in examples]
        return example_sentences[example_index] if 0 <= example_index < len(example_sentences) else None
    except Exception as e:
        print(f"오류 발생: {e}")
        return None


# 기존 파일에서 외운 단어를 제외하고 새 파일 생성
# 함수로 처리하면 더 깔끔할거 같아서 함수로 묶음
def create_filtered_file():
    try:
        book = xl.load_workbook("./word_memorizing/toeic_word.xlsx")
        sheet = book.active

        new_book = xl.Workbook()
        new_sheet = new_book.active

        # 새 단어장의 컬럼 너비 설정
        new_sheet.column_dimensions["A"].width = 15
        new_sheet.column_dimensions["B"].width = 80

        # 새 단어장의 헤더 추가
        new_sheet.append(["영단어", "한글 뜻"])

        for row in range(2, sheet.max_row + 1):  # 첫 번째 행은 헤더이므로 제외
            word = sheet.cell(row=row, column=2).value  # 원래 단어장의 B열: 영단어
            meaning = sheet.cell(row=row, column=3).value  # 원래 단어장의 C열: 한글 뜻
            status = sheet.cell(row=row, column=4).value  # 원래 단어장의 D열: ㅇ 여부

            if status != "ㅇ":  # 외운 단어가 아닌 경우
                new_sheet.append([word, meaning])  # 새 단어장에 추가

        # 새 파일 저장
        new_book.save("./word_memorizing/new_word_list.xlsx")
        print("새 단어장이 성공적으로 생성되었습니다.")
    except Exception as e:
        print(f"새 파일 생성 중 오류 발생: {e}")


# 암기율 계산
def calculate_memorization_rate():
    try:
        book = xl.load_workbook("./word_memorizing/toeic_word.xlsx")
        sheet = book.active

        total_words = 0
        memorized_words = 0

        for row in range(2, sheet.max_row + 1):  # 첫 번째 행은 헤더이므로 제외
            total_words += 1
            if sheet.cell(row=row, column=4).value == "ㅇ":  # D열: 암기 여부
                memorized_words += 1

        # 암기율 계산
        if total_words > 0:
            rate = (memorized_words / total_words) * 100
            return f"{rate:.2f}% ({memorized_words}/{total_words})"
        else:
            return "단어가 없습니다."
    except Exception as e:
        print(f"암기율 계산 중 오류 발생: {e}")
        return "오류 발생"


# 새 파일에서 단어 추출
def fetch_words_from_excel(num_words):
    try:
        new_book = xl.load_workbook("./word_memorizing/new_word_list.xlsx")
        sheet = new_book.active

        word_list = []
        for row in range(2, sheet.max_row + 1):  # 첫 번째 행은 헤더이므로 제외
            word = sheet.cell(row=row, column=1).value  # 새 단어장의 A열: 영단어
            meaning = sheet.cell(row=row, column=2).value  # 새 단어장의 B열: 한글 뜻
            word_list.append((word, meaning))

        random.shuffle(word_list)  # 단어를 랜덤으로 섞음
        return word_list[:num_words]
    except Exception as e:
        print(f"단어 추출 중 오류 발생: {e}")
        return []


# 단어 표시 창
def word_display_window(word_list):
    current_index = 0
    driver = init_driver()

    word_layout = [
        [sg.Text("단어:"), sg.Text(word_list[current_index][0], key="word_display")],
        [sg.Button("한글 뜻, 예문 표시")],
        [sg.Text("한글 뜻:"), sg.Text("", key="korean_meaning")],
        [sg.Text("예문:"), sg.Text("", key="example_sentence")],
        [sg.Checkbox("암기완료", key="memorized")],
        [sg.Button("다음 단어"), sg.Button("종료")]
    ]

    window = sg.Window("영단어 학습", word_layout)

    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, "종료"):
            break

        if event == "한글 뜻, 예문 표시":
            window["korean_meaning"].update(word_list[current_index][1])
            example = fetch_example(driver, word_list[current_index][0])
            window["example_sentence"].update(example if example else "예문을 가져올 수 없습니다.")

        if event == "다음 단어":
            # 암기완료 체크 시 원래 파일에 업데이트
            if values["memorized"]:
                book = xl.load_workbook("./word_memorizing/toeic_word.xlsx")
                sheet = book.active
                for row in range(2, sheet.max_row + 1):
                    if sheet.cell(row=row, column=2).value == word_list[current_index][0]:
                        sheet.cell(row=row, column=4, value="ㅇ")
                        book.save("./word_memorizing/toeic_word.xlsx")
                        break

            if current_index < len(word_list) - 1:
                current_index += 1
                window["word_display"].update(word_list[current_index][0])
                window["korean_meaning"].update("")
                window["example_sentence"].update("")
                window["memorized"].update(False)
            else:
                sg.popup("모든 단어를 학습했습니다!")
                break

    driver.quit()
    window.close()


# 메인 함수
def main():
    create_filtered_file()

    # 암기율 계산
    memorization_rate = calculate_memorization_rate()

    base_layout = [
        [sg.Text("외울 단어 개수"), sg.InputText("10", size=(20, 1), key="num_words")],
        [sg.Text(f"암기율: {memorization_rate}")],
        [sg.Button("단어 표시")]
    ]

    window = sg.Window("영어 단어 학습", base_layout)

    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, "종료"):
            break

        if event == "단어 표시":
            try:
                num_words = int(values["num_words"])
                if num_words > 0:
                    window.close()
                    word_list = fetch_words_from_excel(num_words)
                    if word_list:
                        word_display_window(word_list)
                    else:
                        sg.popup("표시할 단어가 없습니다.")
                    break
                else:
                    sg.popup("양수를 입력해주세요.")
            except ValueError:
                sg.popup("유효한 숫자를 입력하세요.")

    window.close()


if __name__ == "__main__":
    main()
