import os
from bs4 import BeautifulSoup
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog

# Função para selecionar o arquivo HTML
def choose_file():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
    return file_path

# Função para processar o arquivo HTML
def process_html_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()

    # Parsear o HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Encontrar todas as células com colspan="6"
    cells_colspan_6 = soup.find_all('td', attrs={'colspan': '6', 'class': 's4'})
    data_colspan_6 = [cell.get_text() for cell in cells_colspan_6]

    # Encontrar todas as células com colspan="4"
    cells_colspan_3 = soup.find_all('td', attrs={'colspan': '3', 'class': 's14'})
    # Extrair apenas os números das células colspan="4"
    data_colspan_3 = []
    for cell in cells_colspan_3:
        numbers = re.findall(r'\d+', cell.get_text())
        data_colspan_3.extend(numbers)
        
    # Encontrar todas as células com colspan="7" e class="s16"
    cells_colspan_7_s16 = soup.find_all('td', attrs={'colspan': '10', 'class': 's2'})
    # Extrair datas no formato "00/00/0000"
    data_colspan_7_s16 = []
    for cell in cells_colspan_7_s16:
        dates = re.findall(r'\b\d{2}/\d{4}\b', cell.get_text())
        data_colspan_7_s16.extend(dates)
    
    return data_colspan_6, data_colspan_3, data_colspan_7_s16

# Permitir ao usuário escolher o arquivo HTML
html_file = choose_file()

# Verificar se um arquivo foi selecionado
if html_file:
    # Processar o arquivo HTML
    data_colspan_6, data_colspan_3, data_colspan_7_s16 = process_html_file(html_file)
    
    # Criar DataFrames a partir dos dados
    df_colspan_6 = pd.DataFrame(data_colspan_6, columns=['Empresa'])
    df_colspan_3 = pd.DataFrame(data_colspan_3, columns=['Quantidade de Funcionário'])
    df_colspan_7 = pd.DataFrame(data_colspan_7_s16, columns=['Data'])

    # Concatenar os DataFrames
    combined_df = pd.concat([df_colspan_6, df_colspan_3, df_colspan_7], axis=1)

    # Permitir ao usuário escolher o local de salvamento
    save_folder = filedialog.askdirectory()
    if save_folder:
        save_path = os.path.join(save_folder, 'dados.xlsx')
        combined_df.to_excel(save_path, index=False)
        print("Dados exportados para Excel com sucesso!")
    else:
        print("Nenhum local de salvamento selecionado.")
else:
    print("Nenhum arquivo selecionado.")