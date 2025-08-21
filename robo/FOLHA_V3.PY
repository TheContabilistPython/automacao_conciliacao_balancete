from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import time
import os
import pygetwindow as gw
import csv
import customtkinter as ctk
import tkinter as tk
import subprocess
import sys
import re
import calendar

pyautogui.FAILSAFE = False

# Função para obter o código da empresa, o mês e o ano do usuário
# Receber os argumentos
#company_code = sys.argv[1]
#month_year = sys.argv[2]
#company_name = sys.argv[3]

def definir_competencia():
    def on_ok():
        nonlocal competencia, codigo_empresa
        competencia = entry_comp.get()
        codigo_empresa = entry_cod.get()
        dialog.destroy()
        root.quit()

    def on_timeout():
        nonlocal timeout
        timeout = True
        dialog.destroy()
        root.quit()

    # Lê empresas do CSV
    empresas = []
    try:
        with open(r"C:\projeto\empresas.csv", newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';')
            for row in reader:
                empresas.append({'codigo': row['codigo'], 'nome': row['nome']})
    except Exception as e:
        empresas = []

    competencia = None
    codigo_empresa = None
    timeout = False

    root = tk.Tk()
    root.withdraw()

    dialog = tk.Toplevel(root)
    dialog.title("Definir Competência e Empresa")
    dialog.geometry("500x430")
    dialog.attributes('-topmost', True)

    label1 = tk.Label(dialog, text="Digite a competência no formato mmaaaa (ex: 052025):")
    label1.pack(pady=5)
    entry_comp = tk.Entry(dialog)
    entry_comp.pack(pady=5)
    entry_comp.focus()

    label2 = tk.Label(dialog, text="Digite o código da empresa:")
    label2.pack(pady=5)
    entry_cod = tk.Entry(dialog)
    entry_cod.pack(pady=5)

    label_search = tk.Label(dialog, text="Buscar empresa pelo nome:")
    label_search.pack(pady=5)
    entry_search = tk.Entry(dialog)
    entry_search.pack(pady=5)

    label3 = tk.Label(dialog, text="Códigos e empresas disponíveis:")
    label3.pack(pady=5)
    listbox = tk.Listbox(dialog, height=10, width=60)
    listbox.pack(pady=5, fill=tk.BOTH, expand=True)

    def update_listbox(filter_text=""):
        listbox.delete(0, tk.END)
        filtered = [
            emp for emp in empresas
            if filter_text.lower() in emp['nome'].lower()
        ] if filter_text else empresas
        for emp in filtered:
            listbox.insert(tk.END, f"{emp['codigo']} - {emp['nome']}")

    def on_listbox_select(event):
        selection = listbox.curselection()
        if selection:
            valor = listbox.get(selection[0])
            codigo = valor.split(" - ")[0]
            entry_cod.delete(0, tk.END)
            entry_cod.insert(0, codigo)

    def on_search_change(event):
        filter_text = entry_search.get()
        update_listbox(filter_text)

    listbox.bind('<<ListboxSelect>>', on_listbox_select)
    entry_search.bind('<KeyRelease>', on_search_change)

    update_listbox()

    # Botão OK (Confirmar escolha)
    btn = tk.Button(dialog, text="OK (Confirmar escolha)", command=on_ok, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
    btn.pack(pady=15)

    dialog.after(20000, on_timeout)
    root.mainloop()

    if timeout:
        return None, None

    # Validação simples
    if (
        competencia and len(competencia) == 6 and competencia[:2].isdigit() and competencia[2:].isdigit()
        and codigo_empresa and codigo_empresa.isdigit()
    ):
        mes = int(competencia[:2])
        if 1 <= mes <= 12:
            return competencia, codigo_empresa
    return None, None

competencia, codigo_empresa = definir_competencia()
if competencia is None or codigo_empresa is None:
    print("A competência ou o código da empresa não foi definido ou o tempo expirou. Encerrando o programa.")
    sys.exit(0)

print(f"Competência escolhida: {competencia}")
print(f"Código da empresa escolhido: {codigo_empresa}")
time.sleep(5)
sys.stdout.reconfigure(encoding='utf-8')

# Busca o nome da empresa no CSV
company_name = None
try:
    with open(r"C:\projeto\empresas.csv", newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile, delimiter=';')
        for row in reader:
            if row['codigo'] == codigo_empresa:
                company_name = row['nome']
                break
except Exception as e:
    print(f"Erro ao ler arquivo de empresas: {e}")
    sys.exit(0)

if company_name is None:
    print(f"Empresa com código {codigo_empresa} não encontrada no arquivo CSV.")
    sys.exit(0)

company_code = codigo_empresa
month_year = competencia

# Função para obter o último dia do mês
def get_last_day_of_month(month_year):
    month = int(month_year[:2])
    year = int(month_year[2:])
    return calendar.monthrange(year, month)[1]

# Formatar as datas
month_year_formated = f'{month_year[0:2]}-{month_year[2:]}'
last_day = get_last_day_of_month(month_year)

#####################

caminho_html_grrf_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de valores para GRRF  GFD -  Sequência 23 - Ordem 5 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"
total_grrf_a_recolher = None
try:
    with open(caminho_html_grrf_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
        total_grrf_a_recolher = None
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="1. Total GFD:")

        valor_encontrado = 0
        for result in results:
            row = result.find_parent('tr')
            last_value = row.find_all('td')[-1].get_text(strip=True)
            try:
                # Remove non-numeric characters except for comma and dot
                last_value_cleaned = re.sub(r'[^\d,]', '', last_value)
                # Replace comma with dot for float conversion
                last_value_float = float(last_value_cleaned.replace(',', '.'))
                valor_encontrado += last_value_float
            except ValueError:
                continue
        
        if valor_encontrado > 0:
            total_grrf_a_recolher = valor_encontrado
            formatted_value = f"{total_grrf_a_recolher:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            print(f"total_grrf_a_recolher: {formatted_value}")
        else:
            total_grrf_a_recolher = None
            print("total_grrf_a_recolher: Valor não encontrado no HTML")
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_grrf_a_recolher} não existe. Continuando a execução...")
    total_grrf_a_recolher = None
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_grrf_a_recolher}: {e}")
    total_grrf_a_recolher = None

#####################

caminho_html_irrfpf_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de bases IRRF -  Sequência 23 - Ordem 3 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"
total_irrfpf_a_recolher = None
try:
    with open(caminho_html_irrfpf_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
        total_irrfpf_a_recolher = None
    else:
        soup = BeautifulSoup(content, 'html.parser')
        results = soup.find_all('td', string="Valor")

        for result in results:
            row = result.find_parent('tr')
            values = row.find_all('td', class_='s16')
            if len(values) >= 2:
                try:
                    last_value = values[-1].get_text(strip=True)
                    # Extract only the numbers and format the value
                    last_value_numbers = ''.join(filter(str.isdigit, last_value.replace(',', '')))
                    if len(last_value_numbers) >= 2:
                        total_irrfpf_a_recolher = float(f"{int(last_value_numbers[:-2])}.{last_value_numbers[-2:]}")
                        formatted_value = f"{total_irrfpf_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
                        print(f"total_irrfpf_a_recolher: {formatted_value}")
                        break
                except (ValueError, IndexError):
                    continue
        
        if total_irrfpf_a_recolher is None:
            print("total_irrfpf_a_recolher: Valor não encontrado no HTML")
            
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_irrfpf_a_recolher} não existe. Continuando a execução...")
    total_irrfpf_a_recolher = None
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_irrfpf_a_recolher}: {e}")
    total_irrfpf_a_recolher = None


#####################
print(f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Espelho e resumo da FOLHA -  Sequência 23 - Ordem 1 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm")
caminho_html_prolabore_a_pagar = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Espelho e resumo da FOLHA -  Sequência 23 - Ordem 1 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"
prolabore_a_recolher = None
try:
    with open(caminho_html_prolabore_a_pagar, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
        prolabore_a_recolher = None
    else:
        soup = BeautifulSoup(content, 'html.parser')
        # Busca pela linha com "Líquido" ou "LÃquido" nas classes s66
        results = soup.find_all('td', class_='s66', string="Líquido")
        if not results:
            results = soup.find_all('td', class_='s66', string="LÃquido")

        for result in results:
            row = result.find_parent('tr')
            # Busca todos os td com class s67 na linha
            values_s67 = row.find_all('td', class_='s67')
            if len(values_s67) >= 4:  # Precisa ter pelo menos 4 valores s67
                try:
                    # O terceiro valor s67 deve ser o prolabore (30.073,36)
                    prolabore_text = values_s67[2].get_text(strip=True)
                    # Remove pontos e substitui vírgula por ponto
                    clean_value = prolabore_text.replace('.', '').replace(',', '.')
                    prolabore_a_recolher = float(clean_value)
                    if prolabore_a_recolher > 0:
                        formatted_value = f"{prolabore_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
                        print(f"total_prolabore_a_recolher: {formatted_value}")
                        break
                except (ValueError, IndexError):
                    continue
        
        # Se não encontrou na linha principal, busca na linha alternativa
        if prolabore_a_recolher is None:
            print("Prolabore não encontrado na linha principal, buscando na linha alternativa...")
            # Busca pela linha com "Líquido - >" ou "LÃquido - >" nas classes s28
            results_alt = soup.find_all('td', class_='s28', string=lambda text: text and ("Líquido - >" in text or "LÃquido - >" in text))
            
            # Se não encontrou, tenta outras variações
            if not results_alt:
                results_alt = soup.find_all('td', class_='s28', string=lambda text: text and ("Líquido" in text or "LÃquido" in text) and "->" in text)
            
            for result in results_alt:
                row = result.find_parent('tr')
                # Busca todos os td com class s29 na linha
                values_s29 = row.find_all('td', class_='s29')
                if values_s29:
                    try:
                        # O primeiro valor s29 deve ser o prolabore alternativo (6.172,30)
                        prolabore_text = values_s29[0].get_text(strip=True)
                        # Remove pontos e substitui vírgula por ponto
                        clean_value = prolabore_text.replace('.', '').replace(',', '.')
                        prolabore_a_recolher = float(clean_value)
                        if prolabore_a_recolher > 0:
                            formatted_value = f"{prolabore_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
                            print(f"total_prolabore_a_recolher (alternativo): {formatted_value}")
                            break
                    except (ValueError, IndexError):
                        continue
                        
        if prolabore_a_recolher is None:
            print("total_prolabore_a_recolher: Valor não encontrado no HTML")
            
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_prolabore_a_pagar} não existe. prolabore_a_recolher definido como None.")
    prolabore_a_recolher = None
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_prolabore_a_pagar}: {e}")
    prolabore_a_recolher = None

#####################

caminho_html_fgts_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de bases FGTS -  Sequência 23 - Ordem 4 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"
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

caminho_html_inss_a_recolher = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de bases INSS -  Sequência 23 - Ordem 2 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"
total_inss_a_recolher = None
try:
    with open(caminho_html_inss_a_recolher, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has menos de 100 linhas, skipping analysis.")
        total_inss_a_recolher = None
    else:
        soup = BeautifulSoup(content, 'html.parser')
        # Busca por elementos que contenham "GPS"
        results = soup.find_all('td', string=re.compile(r'GPS.*>'))

        for result in results:
            row = result.find_parent('tr')
            # Procurar por células com class 's40' que contêm "(Líquido)" ou "(LÃ­quido)"
            values = row.find_all('td', class_='s40')
            for value_cell in values:
                cell_text = value_cell.get_text(strip=True)
                if ("(Líquido)" in cell_text or "(LÃ­quido)" in cell_text or "Líquido)" in cell_text or "LÃ­quido)" in cell_text):
                    try:
                        # Extrair apenas o valor numérico antes de "(Líquido)"
                        value_part = cell_text.replace("(Líquido)", "").replace("(LÃ­quido)", "").replace("Líquido)", "").replace("LÃ­quido)", "").strip()
                        # Remove pontos e substitui vírgula por ponto
                        clean_value = value_part.replace('.', '').replace(',', '.')
                        total_inss_a_recolher = float(clean_value)
                        formatted_value = f"{total_inss_a_recolher:,.2f}".replace('.', ',').replace(',', '', 1)
                        print(f"total_inss_a_recolher: {formatted_value}")
                        break
                    except (ValueError, IndexError):
                        continue
            if total_inss_a_recolher is not None:
                break
    
    if total_inss_a_recolher is None:
        print("total_inss_a_recolher: Valor não encontrado no HTML")
        
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_inss_a_recolher} não existe. Continuando a execução...")
    total_inss_a_recolher = None
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_inss_a_recolher}: {e}")
    total_inss_a_recolher = None

#####################

caminho_html_sal_a_pagar = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Espelho e resumo da FOLHA -  Sequência 23 - Ordem 1 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"
total_sal_a_pagar = None
try:
    with open(caminho_html_sal_a_pagar, 'r', encoding='utf-8') as file:
        content = file.read()

    # Skip analysis if the file has less than 100 lines
    if len(content.splitlines()) < 100:
        print("HTML file has less than 100 lines, skipping analysis.")
        total_sal_a_pagar = None
    else:
        soup = BeautifulSoup(content, 'html.parser')
        # Busca pela linha com "Líquido" ou "LÃquido" nas classes s44
        results = soup.find_all('td', class_='s44', string="Líquido")
        if not results:
            results = soup.find_all('td', class_='s44', string="LÃquido")

        for result in results:
            row = result.find_parent('tr')
            # Busca todos os td com class s45 na linha
            values_s45 = row.find_all('td', class_='s45')
            if len(values_s45) >= 1:  # Precisa ter pelo menos 1 valor s45
                try:
                    # O primeiro valor s45 deve ser o total_sal_a_pagar
                    sal_text = values_s45[0].get_text(strip=True)
                    # Remove pontos e substitui vírgula por ponto
                    clean_value = sal_text.replace('.', '').replace(',', '.')
                    total_sal_a_pagar = float(clean_value)
                    if total_sal_a_pagar > 0:
                        formatted_value = f"{total_sal_a_pagar:,.2f}".replace('.', ',').replace(',', '', 1)
                        print(f"total_sal_a_pagar: {formatted_value}")
                        break
                except (ValueError, IndexError):
                    continue
        
        if total_sal_a_pagar is None:
            print("total_sal_a_pagar: Valor não encontrado no HTML")
            
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_sal_a_pagar} não existe. total_sal_a_pagar definido como None.")
    total_sal_a_pagar = None
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_sal_a_pagar}: {e}")
    total_sal_a_pagar = None

#####################

caminho_html_13_prov = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de provisão de 13º salário -  Sequência 23 - Ordem 7 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"

# Initialize variables
sal_prov = None
fgts_prov = None
inss_prov = None
pis_prov = None

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
            if row:
                values = row.find_all('td', class_='s21')
                if len(values) >= 4:
                    try:
                        sal_prov = float(values[0].get_text(strip=True).replace('.', '').replace(',', '.'))
                        fgts_prov = float(values[1].get_text(strip=True).replace('.', '').replace(',', '.'))
                        inss_prov = float(values[2].get_text(strip=True).replace('.', '').replace(',', '.'))
                        pis_prov = float(values[3].get_text(strip=True).replace('.', '').replace(',', '.'))
                        break
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
    sal_prov = None
    fgts_prov = None
    inss_prov = None
    pis_prov = None
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_13_prov}: {e}")
    sal_prov = None
    fgts_prov = None
    inss_prov = None
    pis_prov = None

##########################################################################################################################

caminho_html_ferias_prov = f"C:\\relatorios_folha\\Empresa {company_code} - {company_name} - Relatório de provisão de férias -  Sequência 23 - Ordem 6 01-{month_year_formated} {last_day:02d}-{month_year_formated}.htm"

# Initialize variables
ferias_prov = None
fgts_prov_ferias = None
inss_prov_ferias = None
pis_prov_ferias = None

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
            if row:
                values = row.find_all('td', class_='s28')
                if len(values) >= 4:
                    try:
                        ferias_prov = float(values[0].get_text(strip=True).replace('.', '').replace(',', '.'))
                        fgts_prov_ferias = float(values[1].get_text(strip=True).replace('.', '').replace(',', '.'))
                        inss_prov_ferias = float(values[2].get_text(strip=True).replace('.', '').replace(',', '.'))
                        pis_prov_ferias = float(values[3].get_text(strip=True).replace('.', '').replace(',', '.'))
                        break
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
except FileNotFoundError:
    print(f"Erro: O arquivo {caminho_html_ferias_prov} não existe. Continuando a execução...")
    ferias_prov = None
    fgts_prov_ferias = None
    inss_prov_ferias = None
    pis_prov_ferias = None
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_html_ferias_prov}: {e}")
    ferias_prov = None
    fgts_prov_ferias = None
    inss_prov_ferias = None
    pis_prov_ferias = None

# Load the Excel workbook and select the active worksheet
nome_empresa_arquivo = company_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
excel_path = f'C:\\projeto\\planilhas\\CONCILIACAO_{nome_empresa_arquivo}_{month_year}.xlsx'
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
        # Converter valores para float se não forem None
        def to_float(val):
            if val is None:
                return None
            if isinstance(val, (int, float)):
                return float(val)
            try:
                return float(str(val).replace('.', '').replace(',', '.'))
            except Exception:
                return None
        valor_coluna_g_float = to_float(valor_coluna_g)
        valor_coluna_h_float = to_float(valor_coluna_h)
        # Comparar os valores e escrever "OK" ou "Verificar" na coluna I
        try:
            if cell_a in valor_map:
                valor_nome = valor_map[cell_a]
                try:
                    valor = globals()[valor_nome]
                    print(f"Comparando: {valor_nome} (valor={valor}) com coluna_h={valor_coluna_h_float} e coluna_g={valor_coluna_g_float}")
                    
                    # Define a tolerância de 10 centavos para todas as comparações
                    tolerancia = 0.10
                    
                    # Se o valor não foi encontrado no HTML (None), marcar como "Verificar"
                    if valor is None:
                        ws.cell(row=row[0].row, column=9, value="Verificar")
                        ws.cell(row=row[0].row, column=10, value=f"{valor_nome}: Valor não encontrado no HTML")
                        print(f"Linha {row[0].row}: Escreveu Verificar - Valor não encontrado no HTML")
                    elif cell_a == 174:
                        # Para o caso 174, verifica coluna G e também coluna H se não for 0 ou None
                        condicao_g = valor_coluna_g_float is not None and valor is not None and abs(valor_coluna_g_float - valor) <= tolerancia
                        condicao_h = valor_coluna_h_float == 0 or valor_coluna_h_float is None or (valor_coluna_h_float is not None and abs(valor_coluna_h_float - valor) <= tolerancia)
                        
                        ws.cell(row=row[0].row, column=10, value=f"GRRF: G={valor_coluna_g_float} e H={valor_coluna_h_float} vs {valor_nome}={valor}")
                        if condicao_g and condicao_h:
                            ws.cell(row=row[0].row, column=9, value="OK")
                            print(f"Linha {row[0].row}: Escreveu OK (diferença G: {abs(valor_coluna_g_float - valor) if valor_coluna_g_float is not None and valor is not None else 'N/A'})")
                        else:
                            ws.cell(row=row[0].row, column=9, value="Verificar")
                            print(f"Linha {row[0].row}: Escreveu Verificar (diferença G: {abs(valor_coluna_g_float - valor) if valor_coluna_g_float is not None and valor is not None else 'N/A'})")
                    else:
                        # Para todos os outros casos, aplica a tolerância de 10 centavos
                        valor_comparacao = valor_coluna_h_float
                        ws.cell(row=row[0].row, column=10, value=f"{valor_nome}: H={valor_comparacao} vs {valor_nome}={valor}")
                        if valor_comparacao is not None and valor is not None and abs(valor_comparacao - valor) <= tolerancia:
                            ws.cell(row=row[0].row, column=9, value="OK")
                            print(f"Linha {row[0].row}: Escreveu OK (diferença: {abs(valor_comparacao - valor)})")
                        else:
                            ws.cell(row=row[0].row, column=9, value="Verificar")
                            diferenca = abs(valor_comparacao - valor) if valor_comparacao is not None and valor is not None else 'N/A'
                            print(f"Linha {row[0].row}: Escreveu Verificar (diferença: {diferenca})")
                except NameError:
                    print(f"Erro: {valor_nome} não definido. Continuando a execução...")
                    ws.cell(row=row[0].row, column=9, value="Verificar")
                    ws.cell(row=row[0].row, column=10, value=f"Erro: {valor_nome} não definido")
                except KeyError:
                    print(f"Erro: {valor_nome} não encontrado. Continuando a execução...")
                    ws.cell(row=row[0].row, column=9, value="Verificar")
                    ws.cell(row=row[0].row, column=10, value=f"Erro: {valor_nome} não encontrado")
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