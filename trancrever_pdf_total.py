import os
import re
from PyPDF2 import PdfReader
import pandas as pd
from tkinter import Tk, filedialog, simpledialog

def extract_and_write_lines_from_pdf(pdf_file, excel_writer):
    """
    Extrai todas as linhas de texto de um arquivo PDF e escreve as linhas no arquivo Excel.
    
    Args:
    pdf_file (str): Caminho para o arquivo PDF.
    excel_writer (pd.ExcelWriter): Objeto ExcelWriter para escrever no arquivo Excel.
    """
    textos = []
    with open(pdf_file, 'rb') as file:
        reader = PdfReader(file)
        num_pages = len(reader.pages)
        for page_num in range(num_pages):
            page = reader.pages[page_num]
            page_text = page.extract_text()
            lines = page_text.split('\n')
            for line in lines:
                textos.append([line])
    
    # Criar um DataFrame pandas com as linhas extraídas
    df = pd.DataFrame(textos, columns=["Linhas"])
    
    # Obter o nome da planilha a partir do nome do PDF
    nome_planilha = os.path.splitext(os.path.basename(pdf_file))[0]
    
    # Escrever o DataFrame na planilha correspondente
    df.to_excel(excel_writer, sheet_name=nome_planilha, index=False)

def selecionar_pasta(mensagem):
    root = Tk()
    root.withdraw()
    pasta_selecionada = filedialog.askdirectory(title=mensagem)
    return pasta_selecionada

def selecionar_nome_arquivo(mensagem):
    root = Tk()
    root.withdraw()
    nome_arquivo = simpledialog.askstring("Nome do Arquivo", mensagem)
    return nome_arquivo

# Selecionar a pasta contendo os arquivos PDF
pasta_pdf = selecionar_pasta("Selecione a pasta com os arquivos PDF")
if not pasta_pdf:
    print("Nenhuma pasta selecionada. Encerrando o programa.")
    exit()

# Selecionar a pasta de destino para o arquivo Excel
pasta_destino = selecionar_pasta("Selecione a pasta de destino para o arquivo Excel")
if not pasta_destino:
    print("Nenhuma pasta de destino selecionada. Encerrando o programa.")
    exit()

# Selecionar o nome do arquivo Excel de saída
nome_arquivo_excel = selecionar_nome_arquivo("Digite o nome do arquivo Excel para salvar:")

if nome_arquivo_excel is None:
    print("Nenhum nome de arquivo fornecido. Encerrando o programa.")
    exit()
    
# Caminho para o arquivo Excel de saída
excel_path = os.path.join(pasta_destino, f"{nome_arquivo_excel}.xlsx")

# Criar um objeto ExcelWriter para escrever no arquivo Excel
with pd.ExcelWriter(excel_path) as writer:
    # Iterar sobre os arquivos na pasta
    for arquivo in os.listdir(pasta_pdf):
        if arquivo.endswith(".pdf"):
            pdf_path = os.path.join(pasta_pdf, arquivo)
            extract_and_write_lines_from_pdf(pdf_path, writer)

print("Transcrição concluída. Arquivo Excel salvo em:", excel_path)
