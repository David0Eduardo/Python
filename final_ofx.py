import os
import tkinter as tk
from tkinter import filedialog, messagebox
from ofxparse import OfxParser
from openpyxl import Workbook
import sys
import tempfile

def transcrever_ofxs_para_excel(folder_path, excel_path):
    # Dicionário para armazenar os valores de balanço com base no nome do arquivo OFX
    balanco_por_arquivo = {}
    
    if not folder_path or not excel_path:
        return

    workbook = Workbook()
    ws = workbook.active  # Corrigindo aqui

    # Criar uma sheet com todos os dados
    all_data_sheet = workbook.active  # Corrigindo aqui
    all_data_sheet.title = "All Data"

    # Iterar sobre todos os arquivos OFX na pasta selecionada
    for filename in os.listdir(folder_path):
        if filename.endswith(".ofx"):
            ofx_file_path = os.path.join(folder_path, filename)
            with open(ofx_file_path, 'rb') as file:
                ofx_data = file.read().decode('utf-8', errors='ignore')  # Decodifica o arquivo OFX como UTF-8
            # Cria um arquivo temporário para escrever os dados OFX
            with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                temp_file.write(ofx_data.encode('utf-8'))
                temp_file_path = temp_file.name
            # Usa o arquivo temporário para fazer a análise com o OfxParser
            ofx = OfxParser.parse(open(temp_file_path, 'rb'))
            # Verifica se o balanço final está disponível no arquivo OFX
            balanco_final = ofx.account.statement.balance if hasattr(ofx.account.statement, 'balance') else None
            balanco_por_arquivo[ofx_file_path] = balanco_final
            # Adicionar as informações da sheet atual à sheet 'All Data'
            adicionar_dados_para_all_data_sheet(ofx, os.path.splitext(filename)[0], all_data_sheet, balanco_final)
            # Remove o arquivo temporário
            os.unlink(temp_file_path)

    # Salvar o arquivo Excel
    workbook.save(excel_path)

    # Mensagem de conclusão e opção para rodar novamente
    mensagem_conclusao = messagebox.askquestion("Conclusão", "Transcrição concluída com sucesso! Deseja executar novamente?")
    if mensagem_conclusao == "yes":
        selecionar_pasta_ofx()
    else:
        sys.exit()

def adicionar_dados_para_all_data_sheet(ofx, filename, all_data_sheet, balanco_final):
    all_data_sheet.append(["Unidade", "Data", "Tipo", "Valor", "Memo"])
    for transaction in ofx.account.statement.transactions:
        row = [
            filename,
            transaction.date.strftime("%Y-%m-%d"),
            transaction.type,
            transaction.amount,
            transaction.memo
        ]
        all_data_sheet.append(row)
    # Adiciona o balanço final após todas as transações
    if balanco_final is not None:
        all_data_sheet.append([filename, "", "Balanço Final", balanco_final, ""])
        
def selecionar_pasta_ofx():
    # Selecionar a pasta contendo os arquivos OFX
    global folder_path
    folder_path = filedialog.askdirectory()
    if folder_path:
        button_excel.pack(pady=10)  # Exibir botão para salvar o Excel após selecionar a pasta OFX

def selecionar_pasta_salvar_excel():
    # Selecionar o local e o nome para salvar o arquivo Excel
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if excel_path:
        transcrever_ofxs_para_excel(folder_path, excel_path)

# Interface Tkinter
root = tk.Tk()
root.title("Transcrição de OFX para Excel")

button_ofx = tk.Button(root, text="Selecionar pasta com arquivos OFX", command=selecionar_pasta_ofx, padx=10)
button_ofx.pack(pady=10)

button_excel = tk.Button(root, text="Salvar como Excel", command=selecionar_pasta_salvar_excel, padx=10)

root.mainloop()
