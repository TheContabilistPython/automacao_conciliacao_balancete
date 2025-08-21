from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import time
import os
import pygetwindow as gw
import csv
import customtkinter as ctk
import subprocess
import sys

pyautogui.FAILSAFE = False

# Função para obter o código da empresa, o mês e o ano do usuário
# Receber os argumentos
company_code = sys.argv[1]
month_year = sys.argv[2]
company_name = sys.argv[3]

#####################

# Load the Excel workbook and select the active worksheet
excel_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

numeros_procurados = [
    115,
    116,
    117,
    118,
    119,
    120,
    121,
    122,
    2694,
    2695,
    2696,
    2697,
    2698,
    2699,
    3211,
    3351,
    3754,
    3820,
    98,
    99,
    100,
    101,
    102,
    103,
    104,
    105,
    106,
    107,
    108,
    109,
    110
]

# Mapeamento de cell_a para os valores correspondentes
valor_map = {
    115: '(-) DEPRECIAÇÃO ACUMULADA',
    116: '(-) Máquinas e equipamentos',
    117: '(-) Móveis e Utensílios',
    118: '(-) Ferramentas',
    119: '(-) Imóveis',
    120: '(-) Veículos',
    121: '(-) Computadores Periféricos',
    122: '(-) Edificações e Construções',
    2694: '(-) Veículos Usados',
    2695: '(-) Aeronaves',
    2696: '(-) Benfeitorias em Imóveis de Terceiros',
    2697: '(-) Aparelhos Telefônicos',
    2698: '(-) Instalações Diversas',
    2699: '(-) Rouparias',
    3211: '(-) Obras de Arte',
    3351: '(-) Máquinas e Equipamentos Industriais',
    3754: '(-) Salas Comerciais',
    3820: '(-) Imobilizado Diversos',
    96: 'IMOBILIZADO',
    97: 'BENS E DIREITOS EM USO',
    98: 'Máquinas e Equipamentos',
    99: 'Móveis e Utensílios',
    100: 'Ferramentas',
    101: 'Imóveis',
    102: 'Veículos',
    103: 'Computadores e periféricos',
    104: 'Aparelhos Telefônicos',
    105: 'Rouparias',
    106: 'Instalações Diversas',
    107: 'Benfeitorias em imóveis de terceiros',
    108: 'Aeronaves',
    109: 'Terrenos',
    110: 'Edificações e Construções'
}

for row in ws.iter_rows(min_row=2):
    cell_a = row[0].value  # Coluna A (índice 0)
    if cell_a in numeros_procurados:
        valor_coluna_h = row[7].value  # Coluna H (índice 7)
        print(f"Número {cell_a} encontrado: Valor na coluna H = {valor_coluna_h}")
        
        try:
            if cell_a in valor_map:
                row[8].value = "OK"
        except TypeError:
            continue

# Salvar as alterações de volta no arquivo Excel
wb.save(excel_path)

