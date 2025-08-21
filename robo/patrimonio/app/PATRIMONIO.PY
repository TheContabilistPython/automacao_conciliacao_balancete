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

pyautogui.FAILSAFE = False

# Função para obter o código da company_code, o mês e o ano do usuário
def get_user_input():
    ctk.set_appearance_mode("dark")  # Modo de aparência escuro
    ctk.set_default_color_theme("blue")  # Tema de cor padrão

    root = ctk.CTk()
    root.withdraw()  # Esconder a janela principal
    
    # Carregar dados da company_code do arquivo CSV
    company_data = []
    with open(r'C:\\projeto\\empresas.csv', newline='',) as csvfile:
        reader = csv.DictReader(csvfile, delimiter=';')
        for row in reader:
            company_data.append(f"{row['company_code']} - {row['company_name']}")

    # Criar uma nova janela para a entrada do usuário
    input_window = ctk.CTkToplevel(root)
    input_window.title("Conciliador Automático de Folha")
    input_window.geometry("480x400")

    # Solicitar o nome da company_code com sugestões
    ctk.CTkLabel(input_window, text="Nome da company_code:").pack(padx=10, pady=10)
    company_name_var = ctk.StringVar()
    company_name_entry = ctk.CTkEntry(input_window, textvariable=company_name_var, width=300)
    company_name_entry.pack(padx=10, pady=10)

    suggestion_listbox = ctk.CTkTextbox(input_window, width=300, height=100)
    suggestion_listbox.pack(padx=10, pady=10)

    def update_suggestions():
        value = company_name_var.get().lower()
        suggestion_listbox.delete("1.0", ctk.END)
        for item in company_data:
            if value in item.lower():
                suggestion_listbox.insert(ctk.END, item + "\n")

    def on_key_release(event):
        if hasattr(on_key_release, 'after_id'):
            input_window.after_cancel(on_key_release.after_id)
        on_key_release.after_id = input_window.after(1000, update_suggestions)

    company_name_entry.bind('<KeyRelease>', on_key_release)

    def on_listbox_select(event):
        selected_company = suggestion_listbox.get("insert linestart", "insert lineend").strip()
        company_name_var.set(selected_company)
        # Destacar a company_code selecionada
        suggestion_listbox.tag_remove("highlight", "1.0", ctk.END)
        suggestion_listbox.tag_add("highlight", "insert linestart", "insert lineend")
        suggestion_listbox.tag_config("highlight", background="yellow", foreground="black")

    suggestion_listbox.bind('<ButtonRelease-1>', on_listbox_select)

    # Solicitar o mês e o ano
    ctk.CTkLabel(input_window, text="Mês e Ano (MMYYYY):").pack(padx=10, pady=10)
    month_year_var = ctk.StringVar()
    month_year_entry = ctk.CTkEntry(input_window, textvariable=month_year_var)
    month_year_entry.pack(padx=10, pady=10)

    def on_submit():
        selected_company = company_name_var.get()
        company_code, company_name = selected_company.split(' - ', 1)
        month_year = month_year_var.get()
        input_window.destroy()
        root.quit()
        global user_input
        user_input = (company_code, month_year, company_name)

    submit_button = ctk.CTkButton(input_window, text="Conciliar", command=on_submit)
    submit_button.pack(padx=10, pady=10)

    root.mainloop()
    return user_input

# Obter o código da company_code e o mês e o ano do usuário
company_code, month_year, company_name = get_user_input()

day_month_year = '01' + month_year

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

