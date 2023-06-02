from openpyxl import load_workbook
from openpyxl import workbook
from PySimpleGUI import PySimpleGUI as sg   
from openpyxl.worksheet.filters import (
    FilterColumn,
    CustomFilter,
    CustomFilters,
    DateGroupItem,
     FilterColumn,
     Filters
    )


sg.theme('Reddit')
layout = [
    [sg.Text("Teste")]
    ,[]
    ,[sg.Button("ok")]
]



janela = sg.Window('Filtro', layout)

while True:
    eventos, valores = janela.read()

    if eventos == sg.WINDOW_CLOSED:
        break
    
    wb = load_workbook('relatorio.xlsx')
    ws = wb.active


    def Apagarcolunas():
        # apagando as colunas indesejadas  
        ws.delete_cols(1,3)
        ws.delete_cols(2,3)
        ws.delete_cols(3)
        ws.delete_cols(4,5)
        ws.delete_cols(5,20)
        ws.delete_cols(5)
        ws.delete_cols(6)
        ws.delete_cols(6)
        ws.delete_cols(7,12)
        # apagando linhas de cabeçalho
        ws.delete_rows(1,4)
    def Filtros_Alvaro():
        filters = ws.auto_filter
        filters.ref = "A2:F1000"
        col = FilterColumn(colId=2) # para coluna C (equipe responsavel)
        col.filters = Filters(filter=["EDNALDO"]) #adiciona os valores desejados
        filters.filterColumn.append(col) #adiona os filtros na planilha
        ws.auto_filter.add_sort_condition("C2:C1000")

    wb.save('relatorio.xlsx')

janela.mainloop()

wb.save('relatorio.xlsx')

