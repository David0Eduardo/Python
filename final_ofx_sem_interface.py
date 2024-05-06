import os
from ofxparse import OfxParser
from openpyxl import Workbook
import tempfile
import sys

def transcrever_ofxs_para_excel(folder_path, excel_dir):
    if not os.path.exists(excel_dir):
        os.makedirs(excel_dir)

    balanco_por_arquivo = {}
    
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
            balanco_final = ofx.account.statement.balance if hasattr(ofx.account.statement, 'balance') else None  # Corrigindo aqui
            balanco_por_arquivo[ofx_file_path] = balanco_final
            # Adicionar as informações da sheet atual à sheet 'All Data'
            adicionar_dados_para_all_data_sheet(ofx, os.path.splitext(filename)[0], all_data_sheet, balanco_final)
            # Remove o arquivo temporário
            os.unlink(temp_file_path)

    # Salvar o arquivo Excel
    excel_path = os.path.join(excel_dir, "dados.xlsx")
    workbook.save(excel_path)

    print("Transcrição concluída com sucesso!")

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
    folder_path = input("Digite o caminho da pasta contendo os arquivos OFX: ")
    if folder_path:
        selecionar_pasta_salvar_excel()

def selecionar_pasta_salvar_excel():
    # Selecionar o local para salvar o arquivo Excel
    excel_dir = input("Digite o diretório onde deseja salvar o arquivo Excel: ")
    if excel_dir:
        transcrever_ofxs_para_excel(folder_path, excel_dir)

# Iniciar o programa
selecionar_pasta_ofx()
