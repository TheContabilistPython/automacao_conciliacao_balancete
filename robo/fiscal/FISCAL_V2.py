from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import time
import os
import pygetwindow as gw
import csv
import customtkinter as ctk
import re
import subprocess
import sys

pyautogui.FAILSAFE = False

#company_code = sys.argv[1]
#month_year = sys.argv[2]
#company_name = sys.argv[3]

company_code = '16701'
month_year = '012025'
company_name = 'EXTENSION COMERCIAL IMPORTADORA E EXPORTADORA LTDA'
#####################
month_year_formated = f'{month_year[0:2]}-{month_year[2:]}'

caminho_html_csll_presumido_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração IRLP -  Sequência 22 - Ordem 9 01-{month_year_formated} 31-{month_year_formated}.htm"
csll_presumido_recolher = 0
try:
    with open(caminho_html_csll_presumido_recolher, 'r', encoding='utf-8') as file:
        content = file.read()
        
    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "CSLL ACUMULADO TRIMESTRE" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            csll_presumido_recolher += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue
        
    csll_presumido_recolher = round(float(csll_presumido_recolher), 2)
    if csll_presumido_recolher == 0:
        csll_presumido_recolher = None
    else:
        try:
            if csll_presumido_recolher != 0:
                print(f"csll_presumido_recolher: {csll_presumido_recolher:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar csll_presumido_recolher do html")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_csll_presumido_recolher} não existe. Continuando a execução...")

#####################


caminho_html_irpj_presumido_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração IRLP -  Sequência 22 - Ordem 9 01-01-2025 31-01-2025.htm"
irpj_a_recolher_presumido = 0
try:
    with open(caminho_html_irpj_presumido_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "IRPJ ACUMULADO TRIMESTRE" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            irpj_a_recolher_presumido += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    irpj_a_recolher_presumido = round(float(irpj_a_recolher_presumido), 2)
    if irpj_a_recolher_presumido == 0:
        irpj_a_recolher_presumido = None
    else:
        try:
            if irpj_a_recolher_presumido != 0:
                print(f"irpj_presumido_recolher: {irpj_a_recolher_presumido:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar irpj_presumido_recolher do html")
    if irpj_a_recolher_presumido is None:
        irpj_a_recolher_presumido = 0
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_irpj_presumido_recolher} não existe. Continuando a execução...")

#####################

caminho_html_icms_cp_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC -  Sequência 22 - Ordem 1 01-01-2025 31-01-2025.htm"
total_icms_cp_recolher = 0
try:
    with open(caminho_html_icms_cp_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "(=) Imposto a recolher pela utilização do crédito presumido" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_icms_cp_recolher += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_icms_cp_recolher = round(float(total_icms_cp_recolher), 2)
    if total_icms_cp_recolher == 0:
        total_icms_cp_recolher = None
    else:
        try:
            if total_icms_cp_recolher != 0:
                print(f"ICMS_cp_recolher: {total_icms_cp_recolher:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar icms_cp_recolher do html")
    if total_icms_cp_recolher is None:
        total_icms_cp_recolher = 0
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_icms_cp_recolher} não existe. Continuando a execução...")

#####################

caminho_issqn_a_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Conferência serviços -  Sequência 22 - Ordem 10 01-01-2025 31-01-2025.htm"
issqn_a_recolher = 0
try:
    with open(caminho_issqn_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    result = soup.find('td', string="TOTAL GERAL:")

    if result:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-2].get_text(strip=True)
        print(f"iss_qn_a_recolher: {last_value}")  # Adiciona uma instrução de depuração aqui
        try:
            # Remove caracteres não numéricos, exceto vírgula e ponto
            last_value_cleaned = re.sub(r'[^\d,]', '', last_value)
            issqn_a_recolher = float(last_value_cleaned.replace('.', '').replace(',', '.'))
        except ValueError:
            print(f"Erro ao converter o valor: {last_value}")  # Adiciona uma instrução de depuração aqui
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_issqn_a_recolher} não existe. Continuando a execução...")

#####################

caminho_sn_a_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Demonstrativo simples nacional -  Sequência 22 - Ordem 8 01-01-2025 31-01-2025.htm"
sn_a_recolher = 0
try:
    with open(caminho_sn_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Valor do Simples a Pagar" in text)

    if not results:
        print("Nenhum resultado encontrado para 'Valor do Simples a Pagar'")

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        print(f"Valor encontrado: {last_value}")  # Adiciona uma instrução de depuração aqui
        try:
            # Remove caracteres não numéricos, exceto vírgula e ponto
            last_value_cleaned = re.sub(r'[^\d,]', '', last_value)
            sn_a_recolher += float(last_value_cleaned.replace('.', '').replace(',', '.'))
        except ValueError:
            print(f"Erro ao converter o valor: {last_value}")  # Adiciona uma instrução de depuração aqui
            continue

##########################

    total_cofins_recup = round(float(sn_a_recolher), 2)
    if sn_a_recolher == 0:
        sn_a_recolher = None
    else:
        try:
            if sn_a_recolher != 0:
                print(f"sn_recolher: {sn_a_recolher:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar SN a recolher do html")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_sn_a_recolher} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_sn_a_recolher}: {e}")

#####################

caminho_html_icms_fundos = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC -  Sequência 22 - Ordem 1 01-01-2025 31-01-2025.htm"
total_icms_fundos = 0
try:
    with open(caminho_html_icms_fundos, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
        
    # Procurar pelo valor do FUNDO SOCIAL A RECOLHER
    fundo_social_result = soup.find('td', string="(=) FUNDO SOCIAL A RECOLHER")
    if fundo_social_result:
        row = fundo_social_result.find_parent('tr')
        fundo_social_value_td = row.find_all('td', class_='s9')[-1]
        fundo_social_value = fundo_social_value_td.get_text(strip=True)
        fundo_social_value_numbers = ''.join(filter(str.isdigit, fundo_social_value.replace(',', '')))
        fundo_social_a_recolher = float(f"{int(fundo_social_value_numbers[:-2])}.{fundo_social_value_numbers[-2:]}")
    else:
        fundo_social_a_recolher = 0

    # Procurar pelo valor do FUMDES A RECOLHER
    fumdes_result = soup.find('td', string="(=) FUMDES A RECOLHER")
    if fumdes_result:
        row = fumdes_result.find_parent('tr')
        fumdes_value_td = row.find_all('td', class_='s9')[-1]
        fumdes_value = fumdes_value_td.get_text(strip=True)
        fumdes_value_numbers = ''.join(filter(str.isdigit, fumdes_value.replace(',', '')))
        fumdes_a_recolher = float(f"{int(fumdes_value_numbers[:-2])}.{fumdes_value_numbers[-2:]}")
    else:
        fumdes_a_recolher = 0

    # Somar os valores
    fundos_a_recolher = fundo_social_a_recolher + fumdes_a_recolher
    formatted_value = f"{fundos_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
    print(f"fundos_a_recolher: {formatted_value}")

except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_icms_fundos} não existe. Continuando a execução...")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_icms_fundos}: {e}")

#####################

caminho_html_icms_a_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC -  Sequência 22 - Ordem 1 01-01-2025 31-01-2025.htm"
total_icms_a_recolher = 0
try:
    with open(caminho_html_icms_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "(=)Imposto a recolher" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_icms_a_recolher += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_icms_a_recolher = round(float(total_icms_a_recolher), 2)
    if total_icms_a_recolher == 0:
        total_icms_a_recolher = None
    else:
        try:
            if total_icms_a_recolher != 0:
                print(f"icms_a_recolher: {total_icms_a_recolher}")
        except TypeError:
            print(f"erro ao formatar icms_a_recolher do html")
    if total_icms_a_recolher is None:
        total_icms_a_recolher = 0
    print(f"icms_a_recolher: {total_icms_a_recolher}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_icms_a_recolher} não existe. Continuando a execução...")

#####################

caminho_html_ipi_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração IPI -   Sequência 22 - Ordem 2 01-01-2025 31-01-2025.htm"
total_ipi_recolher = 0

try:
    with open(caminho_html_ipi_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Valor do saldo devedor do IPI a recolher" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_ipi_recolher += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_ipi_recolher = round(float(total_ipi_recolher), 2)
    if total_ipi_recolher == 0:
        total_ipi_recolher = None
    else:
        try:
            if total_ipi_recolher != 0:
                print(f"ipi_recolher: {total_ipi_recolher:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar ipi_recolher do html")
    if total_ipi_recolher is None:
        total_ipi_recolher = 0
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_ipi_recolher} não existe. Continuando a execução...")


#####################

caminho_html_cofins_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins -  Sequência 22 - Ordem 4 01-01-2025 31-01-2025.htm"
total_cofins_recolher = 0
try:
    with open(caminho_html_cofins_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Valor total da contribuição a recolher/pagar no período (08 + 12)" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_cofins_recolher += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_cofins_recolher = round(float(total_cofins_recolher), 2)
    if total_cofins_recolher == 0:
        total_cofins_recolher = None
    else:
        try:
            if total_cofins_recolher != 0:
                print(f"cofins_recolher: {total_cofins_recolher:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar cofins_recolher do html")
    if total_cofins_recolher is None:
        total_cofins_recolher = 0
    print(f"cofins_recolher: {total_cofins_recolher}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_cofins_recolher} não existe. Continuando a execução...")

#####################

caminho_html_pis_recolher = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins -  Sequência 22 - Ordem 4 01-01-2025 31-01-2025.htm"
total_pis_recolher = 0
try:
    with open(caminho_html_pis_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Valor total da contribuição a recolher/pagar no período (08 + 12)" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-2].get_text(strip=True)
        try:
            total_pis_recolher += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_pis_recolher = round(float(total_pis_recolher), 2)
    if total_pis_recolher == 0:
        total_pis_recolher = None
    else:
        try:
            if total_pis_recolher != 0:
                print(f"PIS_recolher: {total_pis_recolher:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar PIS_recolher do html")
    if total_pis_recolher is None:
        total_pis_recolher = 0
    print(f"PIS_recolher: {total_pis_recolher}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_pis_recolher} não existe. Continuando a execução...")

#####################

with open(f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de retenções -  Sequência 22 - Ordem 5 01-01-2025 31-01-2025.htm", 'r', encoding='utf-8') as file:
    html_content = file.read()

# Parsear o HTML
soup = BeautifulSoup(html_content, 'html.parser')

linha = soup.find_all('td', class_='s8')  # Busca todas as células com classe 's8'

# Inicializa linha_completa como None
linha_completa = None

# Filtrar com base no texto 'Total Geral'
for td in linha:
    if 'Total Geral' in td.get_text():
        linha_completa = td.find_parent('tr')  # Encontra a linha (<tr>) associada à célula

# Verifica se linha_completa foi definida antes de usá-la
if linha_completa:
    # Extrair valores dinâmicos, ignorando a primeira observação e formatando os valores
    valores = []
    for td in linha_completa.find_all('td')[1:]:
        try:
            valor = float(td.get_text(strip=True).replace('.', '').replace(',', '.'))
            valores.append(round(valor, 2))
        except ValueError:
            continue

    # Atribuir valores às variáveis
    Pis_retido = valores[1]
    Cofins_retido = valores[2]
    csll_retido = valores[3]
    irrf_retido = valores[4]
    iss_retido = valores[5]
    inss_retido = valores[6]

    # Calcular o total de CSRF retido
    csrf_retido = Pis_retido + Cofins_retido + csll_retido

    # Exibir valores retidos
    print(f"IRRF Retido: {irrf_retido:.2f}")
    print(f"ISS Retido: {iss_retido:.2f}")
    print(f"INSS Retido: {inss_retido:.2f}")
    print(f"CSRF Retido: {csrf_retido:.2f}")

#####################

caminho_html_icms_cp_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC -  Sequência 22 - Ordem 1 01-01-2025 31-01-2025.htm"
total_icms_cp_recup = 0
try:
    with open(caminho_html_icms_cp_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "(=) Saldo credor das antecipações para o mês seguinte" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-2].get_text(strip=True)
        try:
            total_icms_cp_recup += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_icms_cp_recup = round(float(total_icms_cp_recup), 2)
    if total_icms_cp_recup == 0:
        total_icms_cp_recup = None
    else:
        try:
            if total_icms_cp_recup != 0:
                print(f"ICMS_cp_recup: {total_icms_cp_recup:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar ICMS_cp_recup do html")
    if total_icms_cp_recup is None:
        total_icms_cp_recup = 0
    print(f"ICMS_cp_recup: {total_icms_cp_recup}")  # Print even if zero
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_icms_cp_recup} não existe. Continuando a execução...")

#####################

caminho_html_cofins_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins -  Sequência 22 - Ordem 4 01-01-2025 31-01-2025.htm"
total_cofins_recup = 0
try:
    with open(caminho_html_cofins_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Totais Cofins" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_cofins_recup += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_cofins_recup = round(float(total_cofins_recup), 2)
    try:
        print(f"COFINS_recup: {total_cofins_recup:.2f}".replace('.', ','))
    except TypeError:
        print(f"erro ao formatar COFINS_recup do html")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_cofins_recup} não existe. Continuando a execução...")

#####################


caminho_html_pis_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins -  Sequência 22 - Ordem 4 01-01-2025 31-01-2025.htm"
total_pis_recup = 0
try:
    with open(caminho_html_pis_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Totais Pis/Pasep" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_pis_recup += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_pis_recup = round(float(total_pis_recup), 2)
    if total_pis_recup == 0:
        total_pis_recup = None
    else:
        try:
            if total_pis_recup != 0:
                print(f"PIS_recup: {total_pis_recup:.2f}".replace('.', ','))
        except TypeError:
            print(f"erro ao formatar PIS_recup do html")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_pis_recup} não existe. Continuando a execução...")
    
    
########################


html_path_icms_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração DimeSC -  Sequência 22 - Ordem 1 01-01-2025 31-01-2025.htm"

try:
    with open(html_path_icms_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    # Cria um objeto BeautifulSoup
    soup = BeautifulSoup(content, 'html.parser')

    # Busca a linha que começa com <td colspan="2" class="s9">190</td>
    result = soup.find('td', {'colspan': '2', 'class': 's9'}, string='190')

    # Exibe a linha inteira
    if result:
        row = result.find_parent('tr')
        observations = row.find_all('td')
        if len(observations) > 1:
            ICMS_recup = observations[2].text
            ICMS_recup = ICMS_recup.replace('.', '').replace(',', '.')
            ICMS_recup = float(ICMS_recup)
            print(f"ICMS_recup: {ICMS_recup:.2f}".replace('.', ','))
except FileNotFoundError:
    print(f"Erro: O arquivo {html_path_icms_recup} não existe. Continuando a execução...")

# ...existing code...

caminho_html_csll_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de CSLL -  Sequência 22 - Ordem 4 01-01-2025 31-01-2025.htm"
total_csll_recup = 0
try:
    with open(caminho_html_csll_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Totais CSLL" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_csll_recup += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_csll_recup = round(float(total_csll_recup), 2)
    try:
        print(f"CSLL_recup: {total_csll_recup:.2f}".replace('.', ','))
    except TypeError:
        print(f"erro ao formatar CSLL_recup do html")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_csll_recup} não existe. Continuando a execução...")

# ...existing code...

caminho_html_irrf_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de IRRF -  Sequência 22 - Ordem 4 01-01-2025 31-01-2025.htm"
total_irrf_recup = 0
try:
    with open(caminho_html_irrf_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Totais IRRF" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_irrf_recup += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_irrf_recup = round(float(total_irrf_recup), 2)
    try:
        print(f"IRRF_recup: {total_irrf_recup:.2f}".replace('.', ','))
    except TypeError:
        print(f"erro ao formatar IRRF_recup do html")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_irrf_recup} não existe. Continuando a execução...")

# ...existing code...

caminho_html_pis_recup = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório de apuração de PISpasep e Cofins -  Sequência 22 - Ordem 4 01-01-2025 31-01-2025.htm"
total_pis_recup = 0
try:
    with open(caminho_html_pis_recup, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results = soup.find_all('td', string=lambda text: text and "Totais Pis/Pasep" in text)

    for result in results:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_pis_recup += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_pis_recup = round(float(total_pis_recup), 2)
    try:
        print(f"PIS_recup: {total_pis_recup:.2f}".replace('.', ','))
    except TypeError:
        print(f"erro ao formatar PIS_recup do html")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_pis_recup} não existe. Continuando a execução...")

# ...existing code...

# Caminho para o arquivo HTML
caminho_html_icms_st = f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório Consulta da Apuração da GIA ST -  Sequência 22 - Ordem 6 01-01-2025 31-01-2025.htm"
total_icms_st = 0
try:
    with open(caminho_html_icms_st, 'r', encoding='utf-8') as file:
        content = file.read()

    soup = BeautifulSoup(content, 'html.parser')
    results_devolucoes = soup.find_all('td', class_='s6', string="ICMS de Devoluções de Mercadorias")
    results_st_devido = soup.find_all('td', class_='s6', string="ICMS-ST Devido")

    for result in results_devolucoes + results_st_devido:
        row = result.find_parent('tr')
        last_value = row.find_all('td')[-1].get_text(strip=True)
        try:
            total_icms_st += float(last_value.replace('.', '').replace(',', '.'))
        except ValueError:
            continue

    total_icms_st = round(float(total_icms_st), 2)
    print(f"total_icms_st: {total_icms_st:.2f}".replace('.', ','))
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_icms_st} não existe. Continuando a execução...")

# Abrir a planilha e procurar pelos números na coluna A
excel_path = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active
  
###################################  
    
numeros_procurados = [43, 48, 41, 42, 46, 47, 2654, 617, 185, 2707, 186, 197, 196, 198, 195, 199, 201, 202, 191, 190, 192, 203]


###################################

def extract_saldo_credor(html_path_ipi_a_recup):
    try:
        with open(f"C:\\relatorios_fiscal\\Empresa {company_code} - {company_name} - Relatório apuração IPI -  Sequência 22 - Ordem 2 01-01-2025 31-01-2025.htm", 'r', encoding='utf-8') as file:
            html_content = file.read()
            
        soup = BeautifulSoup(html_content, 'html.parser')
        saldo_credor_element = soup.find('td', string="SALDO CREDOR PERIODO SEGUINTE")

        if saldo_credor_element:
            ipi_a_recup = saldo_credor_element.find_next_sibling('td').get_text(strip=True)
            ipi_a_recup = float(ipi_a_recup.replace('.', '').replace(',', '.'))
            return ipi_a_recup
        else:
            print("Variável 'SALDO CREDOR PERIODO SEGUINTE' não encontrada.")
            return None
    except FileNotFoundError:
        print(f"Erro: O arquivo {html_path_ipi_a_recup} não existe. Continuando a execução...")
        return None
    
html_path_ipi_a_recup = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}.xlsx'
ipi_a_recup = extract_saldo_credor(html_path_ipi_a_recup)
if ipi_a_recup is not None:
    print(f"IPI a recup: {ipi_a_recup:.2f}".replace('.', ','))
else:
    print("IPI a recup: None")
    ipi_a_recup = 0  # Ensure ipi_a_recup is defined


# Mapeamento de cell_a para os valores correspondentes
valor_map = {
    41: 'ICMS_recup',
    42: 'ipi_a_recup',
    46: 'total_pis_recup',
    47: 'total_cofins_recup',
    2654: 'total_icms_cp_recup',
    617: 'csrf_retido',
    185: 'irrf_retido',
    2707: 'inss_retido',
    186: 'iss_retido',
    197: 'total_pis_recolher',
    196: 'total_cofins_recolher',
    198: 'total_ipi_recolher',
    195: 'total_icms_a_recolher',
    199: 'fundos_a_recolher',
    202: 'sn_a_recolher',
    201: 'issqn_a_recolher',
    191: 'total_icms_cp_recolher',
    190: 'irpj_a_recolher_presumido',
    192: 'csll_presumido_recolher',
    48: 'total_csll_recup',
    43: 'total_irrf_recup',
    203: 'total_icms_st'
}


# Extract the month from the user's input
input_month = month_year[:2]

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
                    if cell_a == 190 and input_month in ['01', '02', '05', '08', '11']:
                        valor_comparacao = valor_coluna_g - valor_coluna_f
                        if abs(valor_comparacao - valor) <= 0.10:
                            row[8].value = "OK"
                        else:
                            row[8].value = "Verificar"
                    elif cell_a == 192 and input_month in ['01', '02', '05', '08', '11']:
                        valor_comparacao = valor_coluna_g - valor_coluna_f
                        if abs(valor_comparacao - valor) <= 0.10:
                            row[8].value = "OK"
                        else:
                            row[8].value = "Verificar"
                    elif cell_a == 203:
                        if abs(valor_coluna_g - valor) <= 0.10:
                            row[8].value = "Verificar Pagamentos"
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

# Caminho para o outro script
caminho_outro_script = f"C:\\patrimonio\\PATRIMONIO.PY"

subprocess.run([
    "python", caminho_outro_script,
    company_code, month_year, company_name
])
