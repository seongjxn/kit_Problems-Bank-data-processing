from PBDP_Excel import *

wb = create_excel_file()
save_excel_file(wb, 'test.xlsx')


wb = open_excel_file('data.xlsx')
data = read_problems_from_excel_file(wb)


for i in range(len(data)):
    print(data[i])