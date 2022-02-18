import xlrd3
import xlwt
from datetime import datetime
wb = xlrd3.open_workbook('task_support.xlsx')
sheet = wb.sheet_by_index(1)
s1 = sheet.row_values(0)
s2 = sheet.col_values(0)
data = []
for i in range(len(s2)):
    data.append(sheet.row_values(i))
count_even = 0
count_0_5 = 0
count_tue = 0
count_tue_F = 0
count_tue_G = 0
for i in range(2, len(data)):
    if data[i][1] % 2 == 0:
        count_even += 1
    word = ''.join(data[i][3].split())
    if ',' in word:
        index = word.find(',')
        digit = float(word[:index] + '.' + word[index + 1:])
        if digit < 0.5:
            count_0_5 += 1
    elif float(''.join(data[i][3].split())) < 0.5:
        count_0_5 += 1
    if 'Tue' in data[i][4]:
        count_tue += 1
    p = data[i][5].split()[0]
    date_sort_F = datetime.strptime(p, "%Y-%m-%d")
    if date_sort_F.weekday() == 1:
        count_tue_F += 1
    k = data[i][6].split('-')
    date_sort_G = datetime.strptime(k[2] + '-' + k[0] + '-' + k[1], "%Y-%m-%d")
    if date_sort_G.weekday() == 1 and 21 <= int(k[1]) <= 31:
        count_tue_G += 1
print(f'В столбце "B" {count_even} четных чисел.')
count_simple = 0
counter = 0
for i in range(2, len(data)):
    for j in range(1, int(data[i][2]) + 1):
        if data[i][2] % j == 0:
            counter += 1
    if counter == 2:
        count_simple += 1
    counter = 0
font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.colour_index = 2
font0.bold = True
style0 = xlwt.XFStyle()
style0.font = font0
wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
ws.write(0, 0, f'В столбце "B" {count_even} четных чисел.', style0)
ws.write(1, 0, f'В столбце "C" {count_simple} простых чисел.', style0)
ws.write(2, 0, f'В столбце "D" {count_0_5} чисел меньших 0.5.', style0)
ws.write(3, 0, f'В столбце "E" {count_tue} вторников.', style0)
ws.write(4, 0, f'В столбце "F" {count_tue_F} вторников.', style0)
ws.write(5, 0, f'В столбце "G" {count_tue_G} вторников.', style0)
wb.save('example.xls')




