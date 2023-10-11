import xlsxwriter
from Dados import *
from Utils import *
from Funcionarios import *
from datetime import timedelta

# Constantes para cores
COR_CABECALHO = '#305496'
COR_FUNDO_CLARO = '#D9D9D9'
COR_FUNDO_HOME = {
    'H': '#B8CCE4',
    'P': '#EA7A7A',
}

workbook = xlsxwriter.Workbook('escalaHomeOffice.xlsx')
worksheet = workbook.add_worksheet()

DATAS = [format_date(inputDataInicial)]
dataFinal = format_date(inputDataFinal)

while DATAS[-1] != dataFinal:
    td = timedelta(1)
    DATAS.append(DATAS[-1] + td)

ROW = 0
COLUMN = 0

def define_formato(**kwargs):
    return workbook.add_format(kwargs)

# Define formatos comuns
merge_format_bold = define_formato(
    bold=True,
    border=1,
    align='center',
    valign='vcenter',
    fg_color=COR_CABECALHO,
    font_color='#ffff'
)

for item in titulos:
    worksheet.merge_range(ROW, COLUMN, ROW + 1, COLUMN, item, merge_format_bold)
    COLUMN += 1

ROW = 2
COLUMN = 0

for funcionario in funcionarios:
    merge_format_fundo_claro = define_formato(fg_color=COR_FUNDO_CLARO)
    worksheet.write(ROW, COLUMN, funcionario.nome, merge_format_fundo_claro)
    worksheet.write(ROW, COLUMN + 1, funcionario.cargo, merge_format_fundo_claro)
    for i in range(len(DATAS)):
        homeCell = ""
        merge_format = ""

        if DATAS[i].weekday() in (5, 6):
            pass
        elif DATAS[i].weekday() == 0:
            homeCell = "P"
        else:
            homeCell = "H" if funcionario.home % 2 == 0 else "P"

        merge_format = define_formato(fg_color=COR_FUNDO_HOME.get(homeCell, COR_FUNDO_CLARO))
        worksheet.write(ROW, i + 2, homeCell, merge_format)
        funcionario.home += 1

    ROW += 1

worksheet.autofit()
workbook.close()
