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

pyautogui.hotkey('ctrl', '3')
time.sleep(1)
pyautogui.press('alt')
time.sleep(1)
pyautogui.press('i')
time.sleep(1)
pyautogui.press('l')
time.sleep(3)
pyautogui.write(company_code)
time.sleep(1)
pyautogui.press('enter')
time.sleep(2)

pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('pgdn')
time.sleep(10)
pyautogui.leftClick(62, 120)
time.sleep(10)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('esc')
time.sleep(1)
pyautogui.press('esc')
time.sleep(1)
pyautogui.press('esc')
time.sleep(1)
pyautogui.press('up')
time.sleep(1)
pyautogui.press('up')
time.sleep(1)
pyautogui.press('enter')


time.sleep(4)

##############################
##        COMANDOS          ##
##############################

pyautogui.hotkey('ctrl', '1')
time.sleep(1)
pyautogui.press('alt')
time.sleep(1)
pyautogui.press('o')
time.sleep(1)
pyautogui.press('l')
time.sleep(3)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('pgdn')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('enter')
time.sleep(20)
pyautogui.leftClick(312, 124)
time.sleep(5)
pyautogui.leftClick(231, 132)
time.sleep(2)
pyautogui.write('PIS Regime Cumulativo')
time.sleep(2) 
pyautogui.press('enter')
time.sleep(2)
pyautogui.press('enter')
time.sleep(2)
pyautogui.leftClick(231, 132)
time.sleep(3)
pyautogui.write('PIS Regime N')
time.sleep(2)
pyautogui.press('enter')  
time.sleep(2)
pyautogui.press('enter')
time.sleep(1)
pyautogui.leftClick(231, 132)
time.sleep(3)
pyautogui.write('COFINS Regime N')
time.sleep(2)
pyautogui.press('enter')  
time.sleep(2)
pyautogui.press('enter')
time.sleep(3)
pyautogui.leftClick(231, 132)
pyautogui.write('COFINS Regime Cumulativo')
time.sleep(2)
pyautogui.press('enter')  
time.sleep(2)
pyautogui.press('enter')
time.sleep(3)

# alt e m 92 tab tab month_year
time.sleep(1)
pyautogui.press('alt')
time.sleep(1)
pyautogui.press('e')
time.sleep(1)
pyautogui.press('m')
time.sleep(1)
pyautogui.write('92')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('esc')
time.sleep(3)

# ctrl 0 alt e 1 tab tab month_year enter time.sleep(20)

pyautogui.hotkey('ctrl', '0')
time.sleep(3)
pyautogui.press('alt')
time.sleep(1)
pyautogui.press('e')
time.sleep(1)
pyautogui.press('1')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('enter')
time.sleep(30)
pyautogui.press('enter')
time.sleep(3)

#alt e s tab month_year enter time.sleep(20) enter

pyautogui.press('alt')
time.sleep(1)
pyautogui.press('e')
time.sleep(1)
pyautogui.press('s')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('enter')
time.sleep(35)
pyautogui.press('enter')
time.sleep(5)

#alt r g g "5" enter tab month_year tab tab tab tab enter

pyautogui.press('alt')
time.sleep(1)
pyautogui.press('r')
time.sleep(1)
pyautogui.press('g')
time.sleep(1)
pyautogui.press('g')
time.sleep(1)
pyautogui.write('5')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.write(month_year)
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('tab')
time.sleep(1)
pyautogui.press('enter')
time.sleep(3)


##############################
##     baixar balancete     ##
##############################
pyautogui.hotkey('ctrl', '2')

time.sleep(4)

# Press Ctrl + B
pyautogui.hotkey('ctrl', 'b')

time.sleep(5)

# Type the date with "01" at the beginning
pyautogui.write('01' + month_year)

pyautogui.press('enter')

# Wait for 2 seconds
time.sleep(2)

# Press Enter again
pyautogui.press('enter')

# Wait for 2 seconds
time.sleep(2)

# Click on specific coordinates
pyautogui.click(x=260, y=160)
time.sleep(1)
pyautogui.click(x=627, y=411)

# Wait for 4 seconds
time.sleep(4)

# Click on specific coordinates
pyautogui.click(x=95, y=160)
time.sleep(1)
pyautogui.click(x=97, y=282)
time.sleep(1)
pyautogui.press('enter')

# Wait for 2 seconds
time.sleep(2)

pyautogui.press('enter')

time.sleep(2)

# Write the path with the company code and month
save_path = f'C:\\projeto\\planilhas\\balancete_{company_code}_{month_year}.csv'
pyautogui.write(save_path)

pyautogui.press('enter')

# Wait for the file to be saved
time.sleep(10)  # Increased wait time to ensure the file is saved

# Load the CSV file with the correct encoding and separator
try:
    df = pd.read_csv(save_path, encoding='latin1', sep=';', on_bad_lines='skip')
except FileNotFoundError:
    print(f"Erro: O arquivo {save_path} não existe. Continuando a execução...")
except pd.errors.ParserError as e:
    print(f"Erro ao analisar o arquivo CSV: {e}")
    with open(save_path, 'r', encoding='latin1') as file:
        lines = file.readlines()
        for i, line in enumerate(lines):
            if len(line.split(';')) != 8:
                print(f"Problema na linha {i + 1}: {line.strip()}")

# Ensure the first column is treated as string
df.iloc[:, 0] = df.iloc[:, 0].astype(str)

# Check if the value '143' is in the first column
if '143' in df.iloc[:, 0].values:
    # Display a message to the user using ctk
    root = ctk.CTk()
    root.withdraw()
    
    # Create a custom message box
    message_box = ctk.CTkToplevel(root)
    message_box.title("Aviso")
    message_box.geometry("600x150")
    ctk.CTkLabel(message_box, text="Saldo em *CONTAS DE COMPENSAÇÃO* avisar FISCAL!").pack(padx=20, pady=20)
    ctk.CTkButton(message_box, text="OK", command=message_box.destroy).pack(pady=10)
    
    root.mainloop()
else:
    # Check if the value '259' is in the first column
    if '259' in df.iloc[:, 0].values:
        # Find the row with the number '259' in column A
        index_259 = df[df.iloc[:, 0] == '259'].index[0]

        # Keep only the rows up to the row with the number '259'
        df = df.iloc[:index_259 + 1]

        # Save the modified CSV file with the correct separator
        df.to_csv(save_path, index=False, encoding='latin1', sep=';')
    else:
        print("O valor '259' não foi encontrado na coluna A.")

    # Open the CSV file
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.write(save_path)
    pyautogui.press('enter')
    time.sleep(4)

    # Select all content in the CSV (Ctrl + T) and copy (Ctrl + C)
    pyautogui.hotkey('ctrl', 't')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)

    # Open Notepad
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.write('notepad')
    pyautogui.press('enter')
    time.sleep(2)

    # Paste the content into Notepad (Ctrl + V)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    # Select all content in Notepad (Ctrl + A) and copy (Ctrl + C)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)

    # Close Notepad without saving
    pyautogui.hotkey('alt', 'f4')
    time.sleep(1)

    # Open the reconciliation file
    pyautogui.hotkey('win', 'r')
    time.sleep(2)
    pyautogui.write('C:\\projeto\\planilhas\\CONCILIACAO_EMPRESA_XX_XXXX.xlsx')
    pyautogui.press('enter')
    time.sleep(5)

    # Go to cell A1
    pyautogui.hotkey('ctrl', 'home')
    time.sleep(1)

    # Paste the content (Ctrl + V)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)

    pyautogui.press('f12')

    time.sleep(2)

    save_path_1 = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{company_code}_{month_year}'
    pyautogui.write(save_path_1)

    pyautogui.press('enter')

    # Wait for the file to be saved
    time.sleep(5)

    # Delete the CSV file
    retry_count = 3
    while retry_count > 0:
        try:
            os.remove(save_path)
            print(f"Arquivo {save_path} excluído com sucesso.")
            break
        except FileNotFoundError:
            print(f"Erro: O arquivo {save_path} não existe. Continuando a execução...")
            break
        except PermissionError:
            print(f"Erro: O arquivo {save_path} está sendo usado por outro processo. Tentando fechar o arquivo...")
            # Try to close the file using pyautogui
            pyautogui.hotkey('alt', 'tab')
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(1)
            retry_count -= 1
        except Exception as e:
            print(f"Erro ao excluir o arquivo {save_path}: {e}")
            break

    # Open cmd and run taskkill command to forcefully close UNICO.EXE
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.write('cmd')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.write('taskkill /IM UNICO.EXE /F')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.write('taskkill /IM EXCEL.EXE /F')
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f4')

time.sleep(5)

# Caminho para o outro script
caminho_outro_script = f"C:\\FOLHA\\FOLHA_V3.PY"

subprocess.run([
    "python", caminho_outro_script,
    company_code, month_year, company_name
])
