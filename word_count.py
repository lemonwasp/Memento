import openpyxl as xl

book = xl.load_workbook("./practice/toeic_word.xlsx")
sheet = book.active
oCount = 0


for r in range(2, 1223) :
    check = sheet.cell(row = r, column = 4).value
    if check == "ㅇ" : 
        oCount += 1
        
percentage = round(oCount/1222*100, 2)
        
print("{0}/1222 {1}%".format(oCount, percentage))