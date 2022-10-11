import openpyxl
from mtranslate import translate
import time
from datetime import datetime

file=input('введите название файла  ')+'.xlsx'
wb = openpyxl.load_workbook(file)
sheet1 = wb[input('введите название листа  ')]
count=0
d=int(input('введите номер строки, с которой начать '))
A='D'+str(d)
B='D'+input('введите номер последней строки ')
r= 14900 #размер перевода за раз
print('программа начала переводить', datetime.now())
for i in sheet1[A:B]:
    for cell in i:
        if count <= r:
           count += len(cell.value)
           translation = translate(cell.value, 'en')
           cell.value = translation
        else:
            time.sleep(2)
            print(datetime.now())
            r += 14900
wb.save(file)
print("перевод завершен")
print('количество переведенных символов', count)
print('программа завершила перевод в', datetime.now())
pause



