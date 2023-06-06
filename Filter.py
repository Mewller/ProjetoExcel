from openpyxl import load_workbook
import pandas as pd
from PySimpleGUI import PySimpleGUI as sg   
from openpyxl.worksheet.filters import (FilterColumn,FilterColumn,Filters)


sg.theme('SystemDefault')
layout = [
    [sg.Text('Selecione um Relatorio')]
    ,[sg.Input(), sg.FileBrowse(file_types=(('Arquivos Excel', '*.xls'),))]
    ,[sg.Text("Qual é o Relatorio?"), sg.Checkbox('SIM', key='SIM', default=True), sg.Checkbox('NÃO',key='NÃO')]
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
                sg.popup(f'O Arquivo {xlsx_file} foi filtrado com sucesso!',
                        title='Relatorio Filtrado!')
            except Exception as e:
                sg.popup(f'Ocorreu um erro na conversão: {e}', title='Erro')

    
    wb = load_workbook(xlsx_file)
    ws = wb.active

#FILTROS
    def Apagarcolunas_nao():
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

    def Apagarcolunas_sim():
        ws.delete_cols(1,3)
        ws.delete_cols(2,3)
        ws.delete_cols(3)
        ws.delete_cols(4,5)
        ws.delete_cols(5,6)
        ws.delete_cols(12,8)
        ws.delete_cols(13,1)
        ws.delete_cols(15,6)
        # apagando linhas de cabeçalho
        ws.delete_rows(1,4)
        
        



    def Filtros_Alvaro():
        filters = ws.auto_filter
        filters.ref = "A2:F1000"
        col = FilterColumn(colId=2) # para coluna C (equipe responsavel)
        col.filters = Filters(filter=["EDNALDO RIBEIRO, HASKELL", "HASKELL, EDNALDO RIBEIRO","HASKELL, SORAYA GALVAO","SORAYA GALVAO, HASKELL","HASKELL, FERNANDO MENESES","FERNANDO MENESES, HASKELL","ADRIELE FERREIRA, OH MY!","ADRIELE FERREIRA","HASKELL, ADRIELE FERREIRA","HASKELL, ARIANA MORAES","ARIANA MORAES, HASKELL","HASKELL, MAGDA SUELI","HASKELL, RENATA BAPTISTA","OH MY!, ROBERTA COUTINHO","ROBERTA COUTINHO , OH MY!",]) #adiciona os valores desejados
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


    

#tamanho da coluna do relatorio do não
    def espaco_nao():
        worksheet = wb
        sheet = worksheet.active
        sheet.column_dimensions['A'].width = 11
        sheet.column_dimensions['B'].width = 37
        sheet.column_dimensions['C'].width = 25
        sheet.column_dimensions['D'].width = 70
        sheet.column_dimensions['E'].width = 17
        sheet.column_dimensions['F'].width = 20
#tamanho da coluna do relatorio do sim        
    def espaco_sim():
        worksheet = wb
        sheet = worksheet.active
        sheet.column_dimensions['A'].width = 11
        sheet.column_dimensions['B'].width = 37
        sheet.column_dimensions['C'].width = 25
        sheet.column_dimensions['D'].width = 40
        sheet.column_dimensions['E'].width = 10
        sheet.column_dimensions['F'].width = 10
        sheet.column_dimensions['G'].width = 10
        sheet.column_dimensions['H'].width = 10
        sheet.column_dimensions['I'].width = 10
        sheet.column_dimensions['J'].width = 10
        sheet.column_dimensions['K'].width = 10
        sheet.column_dimensions['L'].width = 9
        sheet.column_dimensions['M'].width = 16
        sheet.column_dimensions['N'].width = 20 
        


    if valores["Alvaro"] == True:
            if valores["NÃO"] == True:
                Apagarcolunas_nao()
                Filtros_Alvaro()
                espaco_nao()
            else:
                Apagarcolunas_sim()
                Filtros_Alvaro()
                espaco_sim()
      

    wb.save(xlsx_file)



