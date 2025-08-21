from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import tkinter as tk
from tkinter import simpledialog
import time
import os
import pygetwindow as gw

pyautogui.FAILSAFE = False

# Função para obter o código da empresa, o mês e o ano do usuário
def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Esconder a janela principal
    
    # Solicitar o nome da empresa
    company_name = simpledialog.askstring(title="Nome da Empresa", prompt="Digite o nome da empresa:")

    # Solicitar o código da empresa
    company_code = simpledialog.askstring(title="Código da Empresa", prompt="Digite o código da empresa:")

    # Solicitar o mês e o ano
    month_year = simpledialog.askstring(title="Mês e Ano", prompt="Digite o mês e o ano (MMYYYY):")

    return company_code, month_year, company_name

# Obter o código da empresa e o mês e o ano do usuário
company_code, month_year, company_name = get_user_input()

day_month_year = '01' + month_year

#####################

caminho_html_prolabore_a_pagar = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Espelho e resumo da FOLHA Sequência 23 - Ordem 1.htm"
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

        last_total_prolabore_a_recolher = None
        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s68')
            if len(values) >= 2:
                try:
                    last_value = values[-3].get_text(strip=True)
                    # Extract only the numbers and format the value
                    last_value_numbers = ''.join(filter(str.isdigit, last_value.replace(',', '')))
                    last_total_prolabore_a_recolher = float(f"{int(last_value_numbers[:-2])}.{last_value_numbers[-2:]}")
                except ValueError:
                    continue
        if last_total_prolabore_a_recolher is not None:
            print(f"total_prolabore_a_recolher: {last_total_prolabore_a_recolher:.2f}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_prolabore_a_pagar} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_prolabore_a_pagar}: {e}")


#####################

caminho_html_fgts_a_recolher = f"C:\\relatorios_folha\\E16101_BASESFGTS_102024.htm"
try:
    with open(caminho_html_fgts_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="Valor")

        last_total_fgts_a_recolher = None
        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s12')
            if values:
                try:
                    last_value = values[-1].get_text(strip=True)
                    # Extract only the numbers and format the value
                    last_value_numbers = ''.join(filter(str.isdigit, last_value.replace(',', '')))
                    last_total_fgts_a_recolher = float(f"{int(last_value_numbers[:-2])}.{last_value_numbers[-2:]}")
                except ValueError:
                    continue
        if last_total_fgts_a_recolher is not None:
            print(f"total_fgts_a_recolher: {last_total_fgts_a_recolher:.2f}")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_fgts_a_recolher} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_fgts_a_recolher}: {e}")

#####################

caminho_html_inss_a_recolher = f"C:\\relatorios_folha\\Relatório de bases ^INSS.htm"
try:
    with open(caminho_html_inss_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
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
                    print(f"total_inss_a_recolher: {total_inss_a_recolher:.2f}")
                except ValueError:
                    continue
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_inss_a_recolher} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_inss_a_recolher}: {e}")

#####################

caminho_html_sal_a_pagar = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Espelho e resumo da FOLHA Sequência 23 - Ordem 1.htm"
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
            values = row.find_all('td', class_='s68')
            if len(values) >= 2:
                try:
                    second_value = float(values[1].get_text(strip=True).replace('.', '').replace(',', '.'))
                    last_value = float(values[-1].get_text(strip=True).replace('.', '').replace(',', '.'))
                    total_sal_a_pagar = second_value + last_value
                    print(f"total_sal_a_pagar: {total_sal_a_pagar:.2f}".replace('.', ','))
                except ValueError:
                    continue
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_sal_a_pagar} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_sal_a_pagar}: {e}")

#####################

caminho_html_13_prov = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de provisão de 13º salário Sequência 23 - Ordem 7.htm"
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
                    print(f"13_SAL_PROV: {sal_prov:.2f}".replace('.', ','))
                    print(f"FGTS_PROV: {fgts_prov:.2f}".replace('.', ','))
                    print(f"INSS_PROV: {inss_prov:.2f}".replace('.', ','))
                    print(f"PIS_PROV: {pis_prov:.2f}".replace('.', ','))
                except ValueError:
                    continue
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_13_prov} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_13_prov}: {e}")

##########################################################################################################################  

caminho_html_ferias_prov = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de provisão de férias Sequência 23 - Ordem 6.htm"
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
                    print(f"FERIAS_PROV: {ferias_prov:.2f}".replace('.', ','))
                    print(f"FGTS_PROV_FERIAS: {fgts_prov_ferias:.2f}".replace('.', ','))
                    print(f"INSS_PROV_FERIAS: {inss_prov_ferias:.2f}".replace('.', ','))
                    print(f"PIS_PROV_FERIAS: {pis_prov_ferias:.2f}".replace('.', ','))
                except ValueError:
                    continue
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_ferias_prov}: {e}")

# Load the Excel workbook and select the active worksheet
excel_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

numeros_procurados = [176, 177, 178, 179, 180, 181, 160, 171, 172, 169]

for row in ws.iter_rows(min_row=2):
    cell_a = row[0].value  # Coluna A (índice 0)
    if cell_a in numeros_procurados:
        valor_coluna_h = row[7].value  # Coluna H (índice 7)
        print(f"Número {cell_a} encontrado: Valor na coluna H = {valor_coluna_h}")
        
        # Comparar os valores e escrever "OK" ou "Verificar" na coluna I
        try:
            if cell_a == 176:
                try:
                    if abs(valor_coluna_h - sal_prov) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: sal_prov não definido. Continuando a execução...")
                    row[8].value = "Verificar"
            elif cell_a == 177:
                try:
                    if abs(valor_coluna_h - ferias_prov) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: ferias_prov não definido. Continuando a execução...")
            elif cell_a == 178:
                try:
                    if abs(valor_coluna_h - fgts_prov) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: fgts_prov não definido. Continuando a execução...")
            elif cell_a == 179:
                try:
                    if abs(valor_coluna_h - fgts_prov_ferias) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: fgts_prov_ferias não definido. Continuando a execução...")
            elif cell_a == 180:
                try:
                    if abs(valor_coluna_h - inss_prov) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: inss_prov não definido. Continuando a execução...")
            elif cell_a == 181:
                try:
                    if abs(valor_coluna_h - inss_prov_ferias) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: inss_prov_ferias não definido. Continuando a execução...")
            elif cell_a == 160:
                try:
                    if abs(valor_coluna_h - total_sal_a_pagar) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: total_sal_a_pagar não definido. Continuando a execução...")
            elif cell_a == 171:
                try:
                    if abs(valor_coluna_h - total_inss_a_recolher) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: total_inss_a_recolher não definido. Continuando a execução...")
            elif cell_a == 172:
                try:
                    if abs(valor_coluna_h - last_total_fgts_a_recolher) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: last_total_fgts_a_recolher não definido. Continuando a execução...")
            elif cell_a == 169:
                try:
                    if abs(valor_coluna_h - last_total_prolabore_a_recolher) <= 0.10:
                        row[8].value = "OK"
                    else:
                        row[8].value = "Verificar"
                except NameError:
                    print("Erro: last_total_prolabore_a_recolher não definido. Continuando a execução...")
        except TypeError:
            continue

# Salvar as alterações de volta no arquivo Excel
wb.save(excel_path)
