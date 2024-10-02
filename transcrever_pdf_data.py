import os
import re
import PyPDF2
from openpyxl import Workbook
import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox  # Importando o módulo messagebox

def extract_and_write_lines_with_date_from_pdf(pdf_file, target_date, excel_file):
    """
    Extrai linhas contendo uma data específica de um arquivo PDF e escreve o resultado em um arquivo Excel.
    
    Args:
    pdf_file (str): Caminho para o arquivo PDF.
    target_date (str): Data específica a ser procurada nas linhas.
    excel_file (str): Caminho para o arquivo Excel de saída.
    """
    wb = Workbook()
    ws = wb.active
    ws.append(["Linhas com a Data"])
    
    with open(pdf_file, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        num_pages = len(reader.pages)
        for page_num in range(num_pages):
            page = reader.pages[page_num]
            page_text = page.extract_text()
            lines = page_text.split('\n')
            for line in lines:
                if re.search(r'\b\d{1,2}/\d{1,2}/\d{4}\b', line):  # Procura por uma data no formato DD/MM/YYYY
                    if target_date in line:
                        ws.append([line])
    
    wb.save(excel_file)

def get_user_inputs():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    
    pdf_file_name = simpledialog.askstring("Input", "Digite o nome do arquivo PDF (sem extensão):")
    target_date = simpledialog.askstring("Input", "Digite a data específica (formato DD/MM/YYYY):")
    
    pdf_folder_path = r''
    excel_file_path = r''
    pdf_file_path = os.path.join(pdf_folder_path, f'{pdf_file_name}.pdf')
    
    extract_and_write_lines_with_date_from_pdf(pdf_file_path, target_date, excel_file_path)
    
    messagebox.showinfo("Sucesso", f"Processamento concluído com sucesso!\nOs resultados foram salvos em:\n{excel_file_path}")

if __name__ == "__main__":
    get_user_inputs()
