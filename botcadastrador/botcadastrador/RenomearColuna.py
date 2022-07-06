from openpyxl import Workbook, load_workbook
import os

planilha = load_workbook('Pasta1.xlsx')

#Alterado  uma CELULA
planilha.active['A1'] = 'LEGAL'

#Alterando CELULAS especificas de uma COLUNA
for celula in planilha.active['A']:
    linha = celula.row
    if celula.value == 'D':
        planilha.active[f'B{linha}'] = 'ALTERADO'

#Alterando COLUNA inteira
for celula in planilha.active['C']:
    linha = celula.row
    planilha.active[f'C{linha}'] = 'ALTERADO'

planilha.save('Planilha.xlsx')
os.startfile('Planilha.xlsx')