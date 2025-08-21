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
import re

pyautogui.FAILSAFE = False

# Função para obter o código da empresa, o mês e o ano do usuário
# Receber os argumentos
#company_code = sys.argv[1]
#month_year = sys.argv[2]
#company_name = sys.argv[3]

company_code = '16701'
month_year = '012025'
company_name = 'EXTENSION COMERCIAL IMPORTADORA E EXPORTADORA LTDA'

month_year_formated = f'{month_year[0:2]}-{month_year[2:]}'

#####################

caminho_html_grrf_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de valores para GRRF  GFD -  Sequência 23 - Ordem 5 01-{month_year_formated} 31-{month_year_formated}.htm"
total_grrf_a_recolher = 0
try:
    with open(caminho_html_grrf_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="1. Total GFD:")

        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-1].get_text(strip=True)
            try:
                # Remove non-numeric characters except for comma and dot
                last_value_cleaned = re.sub(r'[^\d,]', '', last_value)
                # Replace comma with dot for float conversion
                last_value_float = float(last_value_cleaned.replace(',', '.'))
                total_grrf_a_recolher += last_value_float
            except ValueError:
                continue
        if total_grrf_a_recolher is not None:
            formatted_value = f"{total_grrf_a_recolher:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            print(f"total_grrf_a_recolher: {formatted_value}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_grrf_a_recolher} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_grrf_a_recolher}: {e}")

#####################

caminho_html_irrfpf_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de bases IRRF -  Sequência 23 - Ordem 3 01-{month_year_formated} 31-{month_year_formated}.htm"
total_irrfpf_a_recolher = 0
try:
    with open(caminho_html_irrfpf_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="Valor")

        total_irrfpf_a_recolher = None
        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s16')
            if len(values) >= 2:
                try:
                    last_value = values[-1].get_text(strip=True)
                    # Extract only the numbers and format the value
                    last_value_numbers = ''.join(filter(str.isdigit, last_value.replace(',', '')))
                    total_irrfpf_a_recolher = float(f"{int(last_value_numbers[:-2])}.{last_value_numbers[-2:]}")
                except ValueError:
                    continue
        if total_irrfpf_a_recolher is not None:
            formatted_value = f"{total_irrfpf_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"total_irrfpf_a_recolher: {formatted_value}")
except FileNotFoundError:
    print(f"Erro: O arquivo {total_irrfpf_a_recolher} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_irrfpf_a_recolher}: {e}")


#####################

caminho_html_prolabore_a_pagar = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Espelho e resumo da FOLHA -  Sequência 23 - Ordem 1 01-{month_year_formated} 31-{month_year_formated}.htm"
total_prolabore_a_pagar = 0
try:
    with open(caminho_html_prolabore_a_pagar, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="Líquido")

        prolabore_a_recolher = None
        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s46')
            if len(values) >= 2:
                try:
                    last_value = values[-3].get_text(strip=True)
                    # Extract only the numbers and format the value
                    last_value_numbers = ''.join(filter(str.isdigit, last_value.replace(',', '')))
                    prolabore_a_recolher = float(f"{int(last_value_numbers[:-2])}.{last_value_numbers[-2:]}")
                except ValueError:
                    continue
        if prolabore_a_recolher is not None:
            formatted_value = f"{prolabore_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"total_prolabore_a_recolher: {formatted_value}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_prolabore_a_pagar} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_prolabore_a_pagar}: {e}")

#####################

caminho_html_fgts_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de bases FGTS -  Sequência 23 - Ordem 4 01-{month_year_formated} 31-{month_year_formated}.htm"
try:
    with open(caminho_html_fgts_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="Valor")

        total_fgts_a_recolher = None
        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s14')
            if values:
                try:
                    last_value = values[-1].get_text(strip=True)
                    # Extract only the numbers and format the value
                    last_value_numbers = ''.join(filter(str.isdigit, last_value.replace(',', '')))
                    total_fgts_a_recolher = float(f"{int(last_value_numbers[:-2])}.{last_value_numbers[-2:]}")
                except ValueError:
                    continue
        if total_fgts_a_recolher is not None:
            formatted_value = f"{total_fgts_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"total_fgts_a_recolher: {formatted_value}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_fgts_a_recolher} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_fgts_a_recolher}: {e}")

#####################

caminho_html_inss_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de bases INSS -  Sequência 23 - Ordem 2 01-{month_year_formated} 31-{month_year_formated}.htm"
try:
    with open(caminho_html_inss_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has menos de 100 linhas, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="GPS - >")

        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s40')
            if values:
                try:
                    last_value = values[-1].get_text(strip=True)
                    # Extract only the numbers and format the value
                    last_value_numbers = ''.join(filter(str.isdigit, last_value.replace(',', '')))
                    total_inss_a_recolher = float(f"{int(last_value_numbers[:-2])}.{last_value_numbers[-2:]}")
                except ValueError:
                    continue
        if total_inss_a_recolher is not None:
            formatted_value = f"{total_inss_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"total_inss_a_recolher: {formatted_value}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_inss_a_recolher} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_inss_a_recolher}: {e}")

#####################

caminho_html_sal_a_pagar = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Espelho e resumo da FOLHA -  Sequência 23 - Ordem 1 01-{month_year_formated} 31-{month_year_formated}.htm"
total_sal_a_pagar = 0
try:
    with open(caminho_html_sal_a_pagar, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="Líquido")

        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s69')
            if len(values) >= 2:
                try:
                    second_value = float(values[1].get_text(strip=True).replace('.', '').replace(',', '.'))
                    last_value = float(values[-1].get_text(strip=True).replace('.', '').replace(',', '.'))
                    total_sal_a_pagar = second_value + last_value
                except ValueError:
                    continue
        if total_sal_a_pagar is not None:
            formatted_value = f"{total_sal_a_pagar:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"total_sal_a_pagar: {formatted_value}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_sal_a_pagar} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_sal_a_pagar}: {e}")

#####################

caminho_html_13_prov = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de provisão de 13º salário -  Sequência 23 - Ordem 7 01-{month_year_formated} 31-{month_year_formated}.htm"
try:
    with open(caminho_html_13_prov, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="(=) Saldo final da conta")

        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s23')
            if len(values) >= 4:
                try:
                    sal_prov = float(values[0].get_text(strip=True).replace('.', '').replace(',', '.'))
                    fgts_prov = float(values[1].get_text(strip=True).replace('.', '').replace(',', '.'))
                    inss_prov = float(values[2].get_text(strip=True).replace('.', '').replace(',', '.'))
                    pis_prov = float(values[3].get_text(strip=True).replace('.', '').replace(',', '.'))
                except ValueError:
                    continue
        if sal_prov is not None:
            formatted_sal_prov = f"{sal_prov:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"13_SAL_PROV: {formatted_sal_prov}")
        if fgts_prov is not None:
            formatted_fgts_prov = f"{fgts_prov:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"FGTS_PROV: {formatted_fgts_prov}")
        if inss_prov is not None:
            formatted_inss_prov = f"{inss_prov:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"INSS_PROV: {formatted_inss_prov}")
        if pis_prov is not None:
            formatted_pis_prov = f"{pis_prov:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"PIS_PROV: {formatted_pis_prov}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_13_prov} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_13_prov}: {e}")

##########################################################################################################################

caminho_html_ferias_prov = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de provisão de férias -  Sequência 23 - Ordem 6 01-{month_year_formated} 31-{month_year_formated}.htm"
try:
    with open(caminho_html_ferias_prov, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="(=) Saldo final da conta")

        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s30')
            if len(values) >= 4:
                try:
                    ferias_prov = float(values[0].get_text(strip=True).replace('.', '').replace(',', '.'))
                    fgts_prov_ferias = float(values[1].get_text(strip=True).replace('.', '').replace(',', '.'))
                    inss_prov_ferias = float(values[2].get_text(strip=True).replace('.', '').replace(',', '.'))
                    pis_prov_ferias = float(values[3].get_text(strip=True).replace('.', '').replace(',', '.'))
                except ValueError:
                    continue
        if ferias_prov is not None:
            formatted_ferias_prov = f"{ferias_prov:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"FERIAS_PROV: {formatted_ferias_prov}")
        if fgts_prov_ferias is not None:
            formatted_fgts_prov_ferias = f"{fgts_prov_ferias:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"FGTS_PROV_FERIAS: {formatted_fgts_prov_ferias}")
        if inss_prov_ferias is not None:
            formatted_inss_prov_ferias = f"{inss_prov_ferias:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"INSS_PROV_FERIAS: {formatted_inss_prov_ferias}")
        if pis_prov_ferias is not None:
            formatted_pis_prov_ferias = f"{pis_prov_ferias:,.2f}".replace('.', ',').replace(',', '', 1)
            print(f"PIS_PROV_FERIAS: {formatted_pis_prov_ferias}")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_ferias_prov}: {e}")

# Load the Excel workbook and select the active worksheet
excel_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

numeros_procurados = [176, 177, 178, 179, 180, 181, 160, 171, 172, 169, 184, 174]

# Mapeamento de cell_a para os valores correspondentes
valor_map = {
    176: 'sal_prov',
    177: 'ferias_prov',
    178: 'fgts_prov',
    179: 'fgts_prov_ferias',
    180: 'inss_prov',
    181: 'inss_prov_ferias',
    160: 'total_sal_a_pagar',
    171: 'total_inss_a_recolher',
    172: 'total_fgts_a_recolher',
    169: 'prolabore_a_recolher',
    184: 'total_irrfpf_a_recolher',
    174: 'total_grrf_a_recolher'
}

for row in ws.iter_rows(min_row=2):
    cell_a = row[0].value  # Coluna A (índice 0)
    if cell_a in numeros_procurados:
        valor_coluna_f = row[5].value  # Coluna F (índice 5)
        valor_coluna_g = row[6].value  # Coluna G (índice 6)
        valor_coluna_h = row[7].value  # Coluna H (índice 7)
        print(f"Número {cell_a} encontrado: Valor na coluna H = {valor_coluna_h}")
        
        # Comparar os valores e escrever "OK" ou "Verificar" na coluna I
        try:
            if cell_a in valor_map:
                valor_nome = valor_map[cell_a]
                try:
                    valor = globals()[valor_nome]
                    if cell_a == 174:
                        if abs(valor_coluna_g - valor) <= 0.10 and valor_coluna_h == 0:
                            row[8].value = "OK"
                        else:
                            row[8].value = "Verificar"
                    else:
                        valor_comparacao = valor_coluna_h
                        if abs(valor_comparacao - valor) <= 0.10:
                            row[8].value = "OK"
                        else:
                            row[8].value = "Verificar"
                except NameError:
                    print(f"Erro: {valor_nome} não definido. Continuando a execução...")
                    row[8].value = "Verificar"
                except KeyError:
                    print(f"Erro: {valor_nome} não encontrado. Continuando a execução...")
                    row[8].value = "Verificar"
        except TypeError:
            continue

# Salvar as alterações de volta no arquivo Excel
wb.save(excel_path)

time.sleep(5)

# Caminho para o outro script#
#caminho_outro_script = f"C:\\fiscal\\FISCAL_V2.py"

# Passar os dados para o próximo script como argumentos
#subprocess.run([
#    "python", caminho_outro_script,
#   company_code, month_year, company_name
#])