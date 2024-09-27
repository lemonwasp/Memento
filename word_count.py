import openpyxl as xl
import random

book = xl.load_workbook("./practice/toeic_word.xlsx")
sheet = book.active
oCount = 0
wordDictionary = {}


for r in range(2, 1223) :
    check = sheet.cell(row = r, column = 4).value
    if check == "ㅇ" : 
        oCount += 1
    else :
        wordDictionary[sheet.cell(row = r, column = 2).value] = sheet.cell(row = r, column = 3).value


wordList = list(wordDictionary.items())

random.shuffle(wordList)

percentage = round(oCount/1222*100, 2)

new_book = xl.Workbook()
new_sheet = new_book.active
new_sheet.column_dimensions['A'].width = 15
new_sheet.column_dimensions['B'].width = 80

for row, rowVal in enumerate(wordList) :
    cE = new_sheet.cell(row+1, 1)
    cK = new_sheet.cell(row+1, 2)
    cE.value = rowVal[0]
    cK.value = rowVal[1]

new_book.save("./practice/new_word_list.xlsx")

# print(list(enumerate(wordList)))
print("{0}/1222 {1}%".format(oCount, percentage))
