# -*- coding: utf-8 -*-
"""
Created on Sat Oct 13 04:42:34 2018

@author: Saurater (Sam) Faraday
"""

import pandas
from openpyxl import load_workbook

#Os 2 arquivos citados abaixo já devem existir na pasta local
#Arquivo de Resumo Corrigido
arquivo_resumo_corrigido = 'Resumo ANUAL-2017.xlsx'

#Arquivo de Resumo  Original
arquivo_resumo_original = 'AANUAL-2017.xls'

salvador_excel = pandas.ExcelWriter(arquivo_resumo_corrigido, engine='openpyxl')
pasta_excel = load_workbook(arquivo_resumo_corrigido)
salvador_excel.book = pasta_excel

xlsx = pandas.ExcelFile(arquivo_resumo_original )

for sheet in xlsx.sheet_names:
    planilha = pandas.read_excel(arquivo_resumo_original ,sheet)
    planilha.to_excel(salvador_excel, sheet_name=sheet, header=None, index=False)
    salvador_excel.save()
print("Execução Concluída") 
#Fim
