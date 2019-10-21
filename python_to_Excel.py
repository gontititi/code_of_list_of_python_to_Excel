import numpy as np
import openpyxl as px
Excel_name= 'apple'
wb = px.Workbook()
sheet = wb.active
l = np.arange(200).tolist()
for i in range(0, len(l)):
    sheet[chr(ord('A')+(i%6)) + str((i//6)+1)] = l[i]

wb.save(str(Excel_name + '.xlsx'))
print('a')