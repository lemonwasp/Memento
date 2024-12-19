# import openpyxl as xl
# import random
# import os
# import PySimpleGUI as sg
# from selenium import webdriver # 브라우저를 제어
# from selenium.webdriver.common.by import By # CSS 셀렉터나 태그 이름 등으로 페이지 요소를 찾는 방법을 제공
# from selenium.webdriver.chrome.service import Service # WebDriver 서비스(ex. ChromeDriver)를 관리
# from selenium.webdriver.chrome.options import Options # 브라우저의 설정(ex. 헤드리스 모드)을 구성
# from selenium.webdriver.support.ui import WebDriverWait # 대기를 걸 수 있음
# from selenium.webdriver.support import expected_conditions as EC # 예상조건 WebDriverWait와 같이 사용


# # Selenium WebDriver 초기화
# def init_driver():
#     chrome_options = Options()
#     chrome_options.add_argument("--headless") # UI 없이 Chrome 실행 (헤드리스 모드)
#     chrome_options.add_argument("--disable-gpu") #GPU가속을 비활성화 (헤드리스 모드에서 발생할 수 있는 문제 방지)
#     chrome_options.add_argument("--no-sandbox") # 일부 환경(Docker 등)에서 권한 문제를 방지
#     service = Service(os.path.join(os.path.dirname(__file__), "chromedriver.exe")) # ChromeDriver 실행 파일의 경로 지정, Selenium과 Chrome 사이의 브릿지 역할
#     return webdriver.Chrome(service=service, options=chrome_options) # 지정된 옵션과 서비스를 사용해 Selenium WebDriver를 Chrome브라우저를 초기화


# # 특정 단어의 예문 가져오기
# def fetch_example(driver, word, example_index=4): # index초반의 예문은 예문으로써 적절하지 않은 경우가 많아서 적당히 뒤쪽의 index를 사용
#     try:
#         # 케임브리지 영어사전 URL 설정
#         url = f"https://dictionary.cambridge.org/dictionary/english/{word}"
#         driver.get(url)
        
#         # 모든 예문 요소 가져오기
#         examples = WebDriverWait(driver, 10).until( # 불러올 때까지 대기해줌
#             EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.deg")) # 직접 html/css구조를 열어서 살펴보니 이 Selector 안에 예문이 들어가 있었음
#         )
        
#         # 예문 텍스트 추출
#         example_sentences = [example.text.strip() for example in examples]
#         if 0 <= example_index < len(example_sentences) :
#             return example_sentences[example_index]
#         else :
#             return None
#     except Exception as e:
#         print(f"오류 발생: {e}")
#         return None


# # 기존 파일에서 외운 단어를 제외하고 새 파일 생성
# # 함수로 처리하면 더 깔끔할거 같아서 함수로 묶음
# def create_filtered_file():
#     try:
#         book = xl.load_workbook("./practice/toeic_word.xlsx")
#         sheet = book.active
#         new_book = xl.Workbook()
#         new_sheet = new_book.active

#         new_sheet.append(["영단어", "한글 뜻", "예문"])

#         for row in range(1, sheet.max_row + 1):
#             word = sheet.cell(row=row, column=1).value
#             meaning = sheet.cell(row=row, column=2).value
#             status = sheet.cell(row=row, column=3).value
#             if status != "ㅇ":  # 외운 단어 제외
#                 new_sheet.append([word, meaning, ""])

#         new_book.save("./practice/new_word_list.xlsx")
#     except Exception as e:
#         print(f"새 파일 생성 중 오류 발생: {e}")


# # 새 파일에서 단어 추출
# def fetch_words_from_excel(num_words):
#     try:
#         new_book = xl.load_workbook("./practice/new_word_list.xlsx")
#         sheet = new_book.active

#         word_list = []
#         for row in range(2, sheet.max_row + 1):
#             word = sheet.cell(row=row, column=1).value
#             meaning = sheet.cell(row=row, column=2).value
#             word_list.append((word, meaning))

#         random.shuffle(word_list)
#         return word_list[:num_words]
#     except Exception as e:
#         print(f"단어 추출 중 오류 발생: {e}")
#         return []


# # 단어 표시 창
# def word_display_window(word_list):
#     current_index = 0
#     driver = init_driver()

#     word_layout = [
#         [sg.Text("단어:"), sg.Text(word_list[current_index][0], key="word_display")],
#         [sg.Button("한글 뜻, 예문 표시")],
#         [sg.Text("한글 뜻:"), sg.Text("", key="korean_meaning")],
#         [sg.Text("예문:"), sg.Text("", key="example_sentence")],
#         [sg.Checkbox("암기완료", key="memorized")],
#         [sg.Button("다음 단어"), sg.Button("종료")]
#     ]

#     window = sg.Window("영단어 학습", word_layout)

#     while True:
#         event, values = window.read()

#         if event in (sg.WINDOW_CLOSED, "종료"):
#             break

#         if event == "한글 뜻, 예문 표시":
#             window["korean_meaning"].update(word_list[current_index][1])
#             example = fetch_example(driver, word_list[current_index][0])
#             window["example_sentence"].update(example if example else "예문을 가져올 수 없습니다.")

#         if event == "다음 단어":
#             # 암기완료 체크 시 원래 파일에 업데이트
#             if values["memorized"]:
#                 book = xl.load_workbook("./practice/toeic_word.xlsx")
#                 sheet = book.active
#                 for row in range(1, sheet.max_row + 1):
#                     if sheet.cell(row=row, column=1).value == word_list[current_index][0]:
#                         sheet.cell(row=row, column=3, value="ㅇ")
#                         book.save("./practice/toeic_word.xlsx")
#                         break

#             if current_index < len(word_list) - 1:
#                 current_index += 1
#                 window["word_display"].update(word_list[current_index][0])
#                 window["korean_meaning"].update("")
#                 window["example_sentence"].update("")
#                 window["memorized"].update(False)
#             else:
#                 sg.popup("모든 단어를 학습했습니다!")
#                 break

#     driver.quit()
#     window.close()


# # 메인 함수이자 시작 창
# def main():
#     create_filtered_file()

#     base_layout = [
#         [sg.Text("외울 단어 개수"), sg.InputText("10", size=(20, 1), key="num_words")],
#         [sg.Button("단어 표시")]
#     ]
#     # 창 생성
#     window = sg.Window("영어 단어 학습", base_layout)

#     # 이벤트 루프
#     while True:
#         event, values = window.read()
#         # 종료 조건
#         if event in (sg.WINDOW_CLOSED, "종료"):
#             break
#         # 단어 표시 버튼을 클릭하면
#         if event == "단어 표시":
#             try:
#                 num_words = int(values["num_words"])
#                 if num_words > 0:
#                     window.close() # 시작 창 닫기
#                     # 입력된 개수에 맞는 단어 리스트 생성
#                     word_list = fetch_words_from_excel(num_words)
#                     # 단어 표시 화면 호출
#                     if word_list:
#                         word_display_window(word_list)
#                     else:
#                         sg.popup("표시할 단어가 없습니다.")
#                     break
#                 else:
#                     sg.popup("양수를 입력해주세요.")
#             except ValueError:
#                 sg.popup("유효한 숫자를 입력하세요.")

#     window.close()


# if __name__ == "__main__":
#     main()



# import openpyxl as xl
# import random
# import os
# import PySimpleGUI as sg
# from selenium import webdriver # 브라우저를 제어
# from selenium.webdriver.common.by import By # CSS 셀렉터나 태그 이름 등으로 페이지 요소를 찾는 방법을 제공
# from selenium.webdriver.chrome.service import Service # WebDriver 서비스(ex. ChromeDriver)를 관리
# from selenium.webdriver.chrome.options import Options # 브라우저의 설정(ex. 헤드리스 모드)을 구성
# from selenium.webdriver.support.ui import WebDriverWait # 대기를 걸 수 있음
# from selenium.webdriver.support import expected_conditions as EC # 예상조건 WebDriverWait와 같이 사용



# chrome_options = Options()
# chrome_options.add_argument("--headless") # UI 없이 Chrome 실행 (헤드리스 모드)
# chrome_options.add_argument("--disable-gpu") #GPU가속을 비활성화 (헤드리스 모드에서 발생할 수 있는 문제 방지)
# chrome_options.add_argument("--no-sandbox") # 일부 환경(Docker 등)에서 권한 문제를 방지
# service = Service(os.path.join(os.path.dirname(__file__),'chromedriver.exe')) # ChromeDriver 실행 파일의 경로 지정, Selenium과 Chrome 사이의 브릿지 역할
# driver = webdriver.Chrome(service=service, options=chrome_options) # 지정된 옵션과 서비스를 사용해 Selenium WebDriver를 Chrome브라우저를 초기화



# book = xl.load_workbook("./practice/toeic_word.xlsx")
# sheet = book.active
# oCount = 0
# newOCount = 0
# wordDictionary = {}
# newWordDictionary = {}



# # fetch_example 함수: 특정 단어의 예문 가져오기
# def fetch_example(driver, word, example_index=4):
#     try:
#         # 케임브리지 영어사전 URL 설정
#         url = f"https://dictionary.cambridge.org/dictionary/english/{word}"
#         driver.get(url)

#         # 모든 예문 요소 가져오기
#         examples = WebDriverWait(driver, 10).until(
#             EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.deg"))
#         )

#         # 예문 텍스트 추출
#         example_sentences = [example.text.strip() for example in examples]

#         # 특정 예문 가져오기
#         if 0 <= example_index < len(example_sentences):
#             return example_sentences[example_index]
#         else:
#             return None
#     except Exception as e:
#         print(f"오류 발생: {e}")
#         return None
# # # fetch로 영어사전에서 예문 가져오는 함수
# # def fetch_example(word, example_index):
# #     try:
# #         # URL 설정 및 브라우저 이동(케임브리지 영어사전에서 데이터를 가져옴옴)
# #         url = f"https://dictionary.cambridge.org/dictionary/english/{word}"
# #         driver.get(url)

# #         # 모든 예문 요소 가져오기
# #         examples = WebDriverWait(driver, 10).until(
# #             EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.deg"))
# #         )

# #         # 예문 텍스트 추출
# #         example_sentences = [example.text.strip() for example in examples]

# #         # 특정 예문 가져오기
# #         if 0 <= example_index < len(example_sentences):
# #             return example_sentences[example_index]  # 특정 인덱스의 예문 반환
# #         else:
# #             return f"지정한 인덱스 {example_index}에 해당하는 예문이 없습니다."
# #     except Exception as e:
# #         print(f"오류 발생: {e}")
# #         return None
# #     finally:
# #         driver.quit()

# # # 테스트
# # word = "refute"
# # example_index = 6  # 7번째 예문 (인덱스 6)
# # selected_example = fetch_example(word, example_index)

# # if selected_example:
# #     print(selected_example)
# # else:
# #     print(f"{word}의 예문을 가져오는 데 실패했습니다.")



# # 함수로 처리하면 더 깔끔할거 같아서 함수로 묶음
# def createNewFile(finalWordList) :
#     new_book = xl.Workbook()
#     new_sheet = new_book.active
#     new_sheet.column_dimensions['A'].width = 15
#     new_sheet.column_dimensions['B'].width = 80
#     new_sheet.column_dimensions['C'].width = 100

#     for row, rowVal in enumerate(finalWordList) :
#         cE = new_sheet.cell(row+1, 1)
#         cK = new_sheet.cell(row+1, 2)
#         cE.value = rowVal[0]
#         cK.value = rowVal[1]

#     new_book.save("./practice/new_word_list.xlsx")



# def fetch_words_from_excel(num_words) :
#     # try :
#     #     # 새 단어장 불러오기
#     #     new_book = xl.load_workbook("./practice/new_word_list.xlsx")
#     #     new_sheet = new_book.active
        
#     #     new_word_list = []
#     #     for r in range(1, new_sheet.max_row +1) : # 값이 들어가 있는 셀의 최대값이 어디인지 찾아야함. 일단 임의의 값으로 1223을 줌
#     #         new_check = new_sheet.cell(row = r, column = 3).value
#     #         if new_check == "ㅇ" :
#     #             new_o_count += 1
#     #         else :
#     #             newWordDictionary[new_sheet.cell(row = r, column = 1).value] = new_sheet.cell(row = r, column = 2).value
                
#     #     newWordList = list(newWordDictionary.items())
        
#     #     random.shuffle(newWordList)
        
#     #     for idx, (word, meaning) in enumerate(newWordList) :
#     #         try :
#     #             example_sentence = fetch_example_sentence(word)
#     #             if example_sentence :
#     #                 new_sheet.cell(row=idx + 1, column=3, value=example_sentence[0]) # 여기서 예문을 적어어줌
#     #         except Exception as e :
#     #             print(f"Error fetching example for '{word}': {e}")
        
#     #     percentage = round((1222-new_o_count)/1222*100, 2) # 여기 고쳐야 함, 전체 갯수를 몰라서 1222 넣었는데 이렇게 하면 오류 날거임

#     #     # print(list(enumerate(wordList)))
#     #     print("{0}/1222 {1}%".format((1222-newOCount), percentage)) # 46번째 줄과 마찬가지로 여기도 고쳐야함

#     #     createNewFile(newWordList)
#     try:
#         # 새 단어장 불러오기
#         new_book = xl.load_workbook("./practice/new_word_list.xlsx")
#         new_sheet = new_book.active

#         new_word_list = []
#         for row in range(1, new_sheet.max_row + 1):
#             english_word = new_sheet.cell(row=row, column=1).value
#             meaning = new_sheet.cell(row=row, column=2).value
#             learned_status = new_sheet.cell(row=row, column=3).value

#             # 외운 단어(ㅇ)를 제외
#             if learned_status != "ㅇ":
#                 new_word_list.append((english_word, meaning))

#         # 랜덤으로 단어 추출
#         random.shuffle(new_word_list)
#         return new_word_list[:num_words]
        
        
        
#     except FileNotFoundError:
#         for r in range(2, 1223) :
#             check = sheet.cell(row = r, column = 4).value
#             if check == "ㅇ" : 
#                 oCount += 1
#             else :
#                 wordDictionary[sheet.cell(row = r, column = 2).value] = sheet.cell(row = r, column = 3).value
        
#         wordList = list(wordDictionary.items()) # items()함수를 써야지 키와 값이 함께 튜플에 담김
#         wordList = wordList

#         random.shuffle(wordList)
        
#         for idx, (word, meaning) in enumerate(wordList) :
#             try :
#                 example_sentence = fetch_example_sentence(word)
#                 if example_sentence :
#                     new_sheet.cell(row=idx + 1, column=3, value=example_sentence[0]) # 여기서 예문을 적어어줌
#             except Exception as e :
#                 print(f"Error fetching example for '{word}': {e}")
        
#         # percentage = round(oCount/1222*100, 2)
        
#         # print(list(enumerate(wordList)))
#         # print("{0}/1222 {1}%".format(oCount, percentage))
        
#         createNewFile(wordList)
    


# # 단어 표시 창
# # 시작 창에서 입력한 단어 갯수만큼 반복해서 화면에 표시해줘야함
# # 다음 단어 버튼을 클릭하면 해당 단어 화면은 닫히고 다음 화면이 뜸 
# # 화면 레이아웃 구성
# def word_display_window(word_list) :
#     current_index = 0
#     word_layout = [
#         [sg.Text('단어 :'), sg.Text('----', key='word_display')],
#         [sg.Button('한글 뜻, 예문 표시', pad=((200,0), 0))],
#         [sg.Text('한글 뜻 :'), sg.Text('', key='korean_meaning')],
#         [sg.Text('예문 :'), sg.Text('', key='example_sentence')],
#         [sg.Check('암기완료', False, pad=((240,0), 0))],
#         [sg.Button('다음 단어', pad=((250,0), 0))]
#     ]
#     # 창 생성
#     word_window = sg.Window('영단어', word_layout)

#     #이벤트 루프
#     while True :
#         event, values = word_window.read()
#         #종료 조건
#         if event in ('Exit', None) : break
#         # 한국어 뜻, 예문 표시
#         if event == '한글 뜻, 예문 표시' :
#             word_window['word_display'].update(f"test")
#             word_window['korean_meaning'].update(f"test")
#         if event == '다음 단어' :
#             if current_index < len(word_list) - 1 :
#                 current_index += 1
#                 word_window['word_display'].update(word_list[current_index])
#                 word_window['korean_meaning'].update('') # 한국어 뜻 초기화
#                 word_window['example_sentence'].update('') # 예문 초기화
#             else :
#                 sg.popup('목표 달성!! 축하합니다^^')
#                 break
        




# # 시작 창이자 메인함수
# # 화면 레이아웃 구성
# def main() :
#     base_layout = [
#         [sg.Text('외울 단어 개수'),sg.InputText('10', size=(20,1), key='num_words')],
#         [sg.Button('단어 표시', pad=((200, 0), 0))]
#     ]
#     # 창 생성
#     base_window = sg.Window('영어 단어 개수 고르기', base_layout)

#     # 이벤트 루프
#     while True :
#         event, values = base_window.read()
#         # 종료 조건
#         if event in ('Exit', None) : break
#         # 단어 표시 버튼을 클릭하면
#         if event == '단어 표시' :
#             try :
#                 num_words = int(values['num_words'])
#                 if num_words > 0 :
#                     base_window.close() # 시작 창 닫기
#                     # 개수에 맞는 단어 리스트 생성!!!!!!!실제 데이터로 수정 필요
#                     word_list = fetch_words_from_excel(num_words)
#                     # 단어 표시 화면 호출
#                     if word_list :
#                         word_display_window(word_list)
#                     else :
#                         sg.popup('표시 가능한 단어가 없습니다')
#                     break
#                 else :
#                     sg.popup('양수를 입력해주세요')
#             except ValueError :
#                 sg.popup('유효한 숫자를 입력하세요')
#         base_window.close()
        
        

# if __name__ == '__main__' : # 이 스크립트에서만 실행할 수 있게함
#     main()
