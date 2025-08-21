from bs4 import BeautifulSoup
import openpyxl
import pyautogui
import time
import os
import pygetwindow as gw
import csv
import customtkinter as ctk
import pandas as pd
import subprocess


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

# Press Win + R
pyautogui.hotkey('win', 'r')
time.sleep(1)  # Wait for the Run dialog to open

# Type the path to the file
pyautogui.write('C:\\projeto\\UNICO.EXE.lnk')
time.sleep(1)  # Wait for the typing to complete

# Press Enter
pyautogui.press('enter')

# Wait for 10 seconds
time.sleep(15)

# Type 'contabil'
pyautogui.write('contabil')

# Press Tab
pyautogui.press('tab')

# Type '1234'
pyautogui.write('1234')

# Press Enter
pyautogui.press('enter')

# Wait for 5 seconds
time.sleep(10)

pyautogui.press('alt')
time.sleep(0.5)
pyautogui.press('r')
time.sleep(0.5)
pyautogui.press('q')
time.sleep(2)
pyautogui.write('22')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.write(company_code)
pyautogui.press('enter')
time.sleep(1)
pyautogui.write(day_month_year)
time.sleep(1)
pyautogui.press('tab')
pyautogui.write(month_year)

time.sleep(2)

pyautogui.leftClick(96, 123)
time.sleep(1)
pyautogui.leftClick(99, 184)    
time.sleep(1)
pyautogui.leftClick(695, 399)
time.sleep(2)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.write(f"C:\\relatorios_fiscal")
time.sleep(1)
pyautogui.press('enter')

def wait_for_window(window_title, timeout=60):
    """Wait for a specific window to appear."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = gw.getWindowsWithTitle(window_title)
        if windows:
            return windows[0]
        time.sleep(1)
    raise TimeoutError(f"Window with title '{window_title}' not found within {timeout} seconds.")

# Wait for the specific window to appear after downloading the report
try:
    report_window = wait_for_window("Inconsistências", timeout=120)
    print(f"Window '{report_window.title}' detected.")
    
    time.sleep(3)
    pyautogui.hotkey('alt', 'tab')
    time.sleep(1)
    pyautogui.press('esc')
    time.sleep(1)
    pyautogui.press('esc')
    time.sleep(1)
    pyautogui.press('esc')
    # Re-execute the specified sequence
    pyautogui.press('alt')
    time.sleep(0.5)
    pyautogui.press('r')
    time.sleep(0.5)
    pyautogui.press('q')
    time.sleep(2)
    pyautogui.write('23')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(company_code)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.write(day_month_year)
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.write(month_year)

    time.sleep(2)

    pyautogui.leftClick(96, 123)
    time.sleep(1)
    pyautogui.leftClick(99, 184)    
    time.sleep(1)
    pyautogui.leftClick(695, 399)
    time.sleep(2)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(f"C:\\relatorios_folha")
    time.sleep(1)
    pyautogui.press('enter')

except TimeoutError as e:
    print(e)
    exit(1)
