from openpyxl import load_workbook
from openpyxl import workbook
import pandas as pd
from PySimpleGUI import PySimpleGUI as sg   
from openpyxl.worksheet.filters import (
    FilterColumn,
    CustomFilter,
    CustomFilters,
    DateGroupItem,
     FilterColumn,
     Filters
    )


sg.theme('SystemDefault')
layout = [
    [sg.Text('Selecione um Relatorio')]
    ,[sg.Input(), sg.FileBrowse(file_types=(('Arquivos Excel', '*.xls'),))]
    ,[sg.Text("Qual é o Relatorio?"), sg.Checkbox('SIM', key='SIM'), sg.Checkbox('NÃO',key='NÃO', default=True)]
    ,[sg.Text("Quem está usando?")]
    ,[sg.Radio('Alvaro','RADIO1',key='Alvaro' , default=True), sg.Radio('Mateus','RADIO1', key='Mateus'), sg.Radio('Marcia','RADIO1', key='Marcia')
    ,sg.Radio('Sueli','RADIO1', key='Sueli'), sg.Radio('Livia','RADIO1', key='Livia')]
    ,[sg.Button("Cancelar"), sg.Button("Filtrar")]
]

janela = sg.Window('Filtro', layout)

while True:
    eventos, valores = janela.read()

    if eventos == sg.WINDOW_CLOSED or eventos == 'Cancelar':
        break
    #conversão de arquivos
    elif eventos == 'Filtrar':
            xls_file = valores[0]
            xlsx_file = xls_file.replace('.xls', '.xlsx')
            try:
                # Leitura do arquivo .xls
                df = pd.read_excel(xls_file)

                # Gravação do arquivo .xlsx
                df.to_excel(xlsx_file, index=False)
                print(valores)
                sg.popup(f'O arquivo {xls_file} foi convertido para {xlsx_file} com sucesso!',
                        title='Conversão Concluída')
            except Exception as e:
                sg.popup(f'Ocorreu um erro na conversão: {e}', title='Erro')

    if "NÃO" == True in valores:
          if "Alvaro" == True in valores:
                Apagarcolunas()
                Filtros_Alvaro()


    wb = load_workbook(xlsx_file)
    ws = wb.active

#FILTROS
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
    def Filtros_Mateus():
            filters = ws.auto_filter
            filters.ref = "A2:F1000"
            col = FilterColumn(colId=2) # para coluna C (equipe responsavel)
            col.filters = Filters(filter=[""]) #adiciona os valores desejados
            filters.filterColumn.append(col) #adiona os filtros na planilha
            ws.auto_filter.add_sort_condition("C2:C1000")
    def Filtros_Marcia():
            filters = ws.auto_filter
            filters.ref = "A2:F1000"
            col = FilterColumn(colId=2) # para coluna C (equipe responsavel)
            col.filters = Filters(filter=[""]) #adiciona os valores desejados
            filters.filterColumn.append(col) #adiona os filtros na planilha
            ws.auto_filter.add_sort_condition("C2:C1000")
    def Filtros_Sueli():
            filters = ws.auto_filter
            filters.ref = "A2:F1000"
            col = FilterColumn(colId=2) # para coluna C (equipe responsavel)
            col.filters = Filters(filter=[""]) #adiciona os valores desejados
            filters.filterColumn.append(col) #adiona os filtros na planilha
            ws.auto_filter.add_sort_condition("C2:C1000")
    def Filtros_Livia():
            filters = ws.auto_filter
            filters.ref = "A2:F1000"
            col = FilterColumn(colId=2) # para coluna C (equipe responsavel)
            col.filters = Filters(filter=[""]) #adiciona os valores desejados
            filters.filterColumn.append(col) #adiona os filtros na planilha
            ws.auto_filter.add_sort_condition("C2:C1000")

    wb.save(xlsx_file)

janela.mainloop()

wb.save('relatorio.xlsx')

