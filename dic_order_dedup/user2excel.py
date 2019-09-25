import os.path
import openpyxl

f = open('nnp.user', 'r', encoding='UTF8')

xlsx = openpyxl.Workbook()
output = xlsx.active

for i in range(116050):
    test = [f.readline()]
    #print(test)
    output.append(test)

xlsx.save("test.xlsx")