import pytesseract
import pyautogui
import cv2
import numpy as np
import time
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox, scrolledtext
from openpyxl import load_workbook
import shutil
import tempfile
import os

# Função para abrir o aplicativo e fazer login
def abrir_aplicativo():
    pyautogui.hotkey('win', 'r')
    pyautogui.write(r'C:\Users\contabil18\Desktop\UNICO.LNK')
    pyautogui.press('enter')
    time.sleep(10)
    pyautogui.write('eduardoap')
    pyautogui.press('enter')
    pyautogui.write('blackeyes')
    pyautogui.press('enter')

# Função para capturar a tela e encontrar o texto "NENHUMA"
def encontrar_texto():
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    time.sleep(10)
    screenshot = pyautogui.screenshot()
    screenshot.save(r'C:\Users\contabil18\Documents\projectRD\screenshot.png')
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
    _, binary_image = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
    text = pytesseract.image_to_data(binary_image, output_type=pytesseract.Output.DICT)
    target_text = "NENHUMA"

    for i in range(len(text['text'])):
        if int(text['conf'][i]) > 0:
            x, y, w, h = text['left'][i], text['top'][i], text['width'][i], text['height'][i]
            cv2.rectangle(screenshot, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.putText(screenshot, text['text'][i], (x, y + h + 20), cv2.FONT_HERSHEY_SIMPLEX, 0.3, (0, 255, 0), 1)
    cv2.imwrite(r'C:\Users\contabil18\Documents\projectRD\detected_text.png', screenshot)

    for i in range(len(text['text'])):
        if target_text in text['text'][i]:
            x, y, w, h = text['left'][i], text['top'][i], text['width'][i], text['height'][i]
            center_x, center_y = x + w // 2, y + h // 2
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click()
            break

# Função para processar uma empresa específica
def processar_empresa(codigo_empresa, mes, ano):
    time.sleep(2)
    codigo_empresa_str = str(int(float(codigo_empresa)))
    print(f"Escrevendo o código da empresa: {codigo_empresa_str}")
    pyautogui.write(codigo_empresa_str)
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'b')
    time.sleep(2)
    today = datetime.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
    date_to_input = last_day_of_previous_month.replace(day=1).strftime('%d%m%Y')
    print(f"Escrevendo a data: {date_to_input}")
    pyautogui.write(date_to_input, interval=0.125)
    time.sleep(2)
    pyautogui.moveTo(258, 156)
    pyautogui.click(button='left')
    pyautogui.moveTo(626, 409)
    pyautogui.click(button='left')
    time.sleep(2)
    pyautogui.hotkey('alt', 'f8')
    pyautogui.press('enter')
    time.sleep(8)
    screenshot_after = pyautogui.screenshot()
    screenshot_after.save(r'C:\Users\contabil18\Documents\projectRD\screenshot_after.png')

    try:
        icon_location = pyautogui.locateOnScreen(r'C:\Users\contabil18\Documents\projectRD\icone.png')
        if icon_location:
            icon_center = pyautogui.center(icon_location)
            pyautogui.moveTo(icon_center)
            time.sleep(2)
            pyautogui.click(button='left')
            time.sleep(2)
            pyautogui.moveTo(707, 253)
            pyautogui.click(button='left')
            time.sleep(10)
            pyautogui.press('enter')
            time.sleep(2)
            caminho_balancete = f'C:\\Users\\contabil18\\Documents\\projectRD\\balancetes\\balancete_{codigo_empresa_str}'
            pyautogui.write(caminho_balancete)
            pyautogui.press('enter')
            time.sleep(2)
            processar_balancete(caminho_balancete, codigo_empresa_str, mes, ano)
        else:
            print("Ícone não encontrado na tela.")
    except pyautogui.ImageNotFoundException:
        print("Imagem não encontrada na tela.")

# Função para processar o balancete baixado
def processar_balancete(caminho_balancete, codigo_empresa, mes, ano):
    try:
        df = pd.read_csv(caminho_balancete + '.csv', encoding='latin1', sep=';')
        print("Dados da primeira coluna e seus tipos:")
        for value in df.iloc[:, 0]:
            print(f"Valor: {value}, Tipo: {type(value)}")
        if '259' in df.iloc[:, 0].astype(str).values:
            linha_259 = df[df.iloc[:, 0].astype(str) == '259'].index[0]
            df = df.iloc[:linha_259 + 1, :]
            df.to_csv(caminho_balancete + '_processado.csv', index=False, sep=';', encoding='latin1')
            print(f"Arquivo processado salvo em: {caminho_balancete}_processado.csv")
            with open(caminho_balancete + '_processado.csv', 'r', encoding='latin1') as file:
                data = file.readlines()[1:]
                text_area.delete(1.0, tk.END)
                text_area.insert(tk.END, ''.join(data))
            colar_dados_no_excel(df, codigo_empresa, mes, ano)
        else:
            print("String '259' não encontrada na primeira coluna.")
    except Exception as e:
        print(f"Erro ao processar o balancete: {e}")

# Função para colar os dados no arquivo Excel
def colar_dados_no_excel(df, codigo_empresa, mes, ano):
    try:
        caminho_excel_original = r'C:\Users\contabil18\Documents\projectRD\CONCILIACA_EMPRESA_XX_XXXX.xlsx'
        caminho_excel_novo = f'C:\\Users\\contabil18\\Documents\\projectRD\\CONCILIACA_{codigo_empresa}_{mes}_{ano}.xlsx'
        shutil.copyfile(caminho_excel_original, caminho_excel_novo)
        workbook = load_workbook(caminho_excel_novo)
        sheet = workbook.active
        with tempfile.NamedTemporaryFile(delete=False, mode='w', encoding='latin1', suffix='.txt') as temp_file:
            temp_file_path = temp_file.name
            df.to_csv(temp_file_path, index=False, sep='\t', encoding='latin1')
        os.startfile(temp_file_path)
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(2)
        pyautogui.hotkey('alt', 'f4')
        os.startfile(caminho_excel_novo)
        time.sleep(5)
        pyautogui.hotkey('ctrl', 'g')
        pyautogui.write('A1')
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'u')
        pyautogui.hotkey('ctrl', 's')
        print(f"Dados colados no arquivo Excel: {caminho_excel_novo}")
    except Exception as e:
        print(f"Erro ao colar dados no Excel: {e}")

# Função para copiar o conteúdo da área de texto para a área de transferência
def copiar_para_area_de_transferencia():
    root.clipboard_clear()
    root.clipboard_append(text_area.get(1.0, tk.END))
    root.update()
    messagebox.showinfo("Informação", "Conteúdo copiado para a área de transferência.")

# Função para iniciar o processo com o código da empresa fornecido
def iniciar_processo():
    codigo_empresa = entry_codigo_empresa.get()
    mes = entry_mes.get()
    ano = entry_ano.get()
    if codigo_empresa and mes and ano:
        abrir_aplicativo()
        encontrar_texto()
        processar_empresa(codigo_empresa, mes, ano)
    else:
        messagebox.showwarning("Aviso", "Por favor, insira o código da empresa, mês e ano.")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Processamento de Empresa")

tk.Label(root, text="Código da Empresa:").pack(pady=10)
entry_codigo_empresa = tk.Entry(root)
entry_codigo_empresa.pack(pady=5)

tk.Label(root, text="Mês:").pack(pady=10)
entry_mes = tk.Entry(root)
entry_mes.pack(pady=5)

tk.Label(root, text="Ano:").pack(pady=10)
entry_ano = tk.Entry(root)
entry_ano.pack(pady=5)

tk.Button(root, text="Iniciar Processo", command=iniciar_processo).pack(pady=20)

text_area = scrolledtext.ScrolledText(root, width=80, height=20)
text_area.pack(pady=10)

tk.Button(root, text="Copiar para Área de Transferência", command=copiar_para_area_de_transferencia).pack(pady=10)

root.mainloop()