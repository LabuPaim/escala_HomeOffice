import xlsxwriter
from Dados import *
from Utils import *
from Funcionarios import *
from datetime import datetime, date, timedelta

workbook = xlsxwriter.Workbook('escalaHomeOffice.xlsx')
worksheet = workbook.add_worksheet()

DATAS = []
DATAS.append(format_date(inputDataInicial))
dataFinal = format_date(inputDataFinal)

count = 0
while DATAS[-1] != dataFinal:
    td = timedelta(1)
    DATAS.append(DATAS[-1] + td)
    count += 1

ROW = 0
COLUMN = 0
for item in titulos:
    merge_format = workbook.add_format({
        'bold': True,
        'border':   1,
        'align': 'center',
        'valign':   'vcenter',
        'fg_color': '#305496',
        'font_color': '#ffff'
    })
    worksheet.merge_range(ROW, COLUMN, ROW + 1, COLUMN, item, merge_format)
    COLUMN += 1


for data in DATAS:
    merge_format = workbook.add_format({
        'bold': True,
        'num_format': 'd/mm/yyyy',
        'border':   1,
        'align': 'center',
        'valign':   'vcenter',
        'fg_color': '#305496',
        'font_color': '#ffff'
    })
    worksheet.write_datetime(ROW, COLUMN, data, merge_format)
    merge_format = workbook.add_format({
        # 'bold': True,
        'num_format': 'dddd',
        'border':   1,
        'align': 'center',
        'valign':   'vcenter',
        'fg_color': '#305496',
        'font_color': '#ffff'
    })
    worksheet.write_datetime(ROW + 1, COLUMN, data, merge_format)
    COLUMN += 1

ROW = 2
COLUMN = 0
for funcionario in funcionarios:
    merge_format = workbook.add_format({'fg_color': '#D9D9D9',})
    worksheet.write(ROW, COLUMN, funcionario.nome, merge_format)
    worksheet.write(ROW, COLUMN + 1, funcionario.cargo, merge_format)
    for i in range(len(DATAS)):
        homeCell = ""
        merge_format = ""

        match DATAS[i].weekday():
            case 5 | 6:
                pass
            case 0:
                homeCell = "P"
            case _:
                if (funcionario.home % 2) == 0:
                    homeCell = "H"
                if (funcionario.home % 2) == 1:
                    homeCell = "P"
        
        match homeCell:
            case "":
                merge_format = workbook.add_format({'fg_color': '#D9D9D9',})
            case "H":
                merge_format = workbook.add_format({'fg_color': '#B8CCE4',})
            case "P":
                merge_format = workbook.add_format({'fg_color': '#EA7A7A',})

        worksheet.write(ROW, i+2, homeCell, merge_format)
        funcionario.home += 1
    ROW += 1

worksheet.autofit()
workbook.close()
