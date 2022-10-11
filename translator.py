import openpyxl
from mtranslate import translate
import time
from datetime import datetime

file=input('write name of the file  ')+'.xlsx'
wb = openpyxl.load_workbook(file)
sheet1 = wb[input('write name of the list  ')]
count=0
d=int(input('write number of the first line '))
A='D'+str(d)
B='D'+input('write number of the last line ')
r= 14900 #tranlate range
print('program start translating', datetime.now())
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
print("translate is finished", datetime.now())
print('number of transleted symbols', count)





