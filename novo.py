import xlsxwriter

workbook = xlsxwriter.Workbook('escalaHomeOffice.xlsx')
worksheet = workbook.add_worksheet()

print(worksheet.table[0][0].number)

workbook.close()