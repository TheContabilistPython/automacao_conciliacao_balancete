import os

if os.path.exists('tkinter.py'):
    print("Renomeie o arquivo 'tkinter.py' para evitar conflitos.")
else:
    print("Nenhum conflito de nome de arquivo encontrado.")