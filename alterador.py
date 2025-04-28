import tkinter as tk
from tkinter import messagebox, simpledialog
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import os
from listas import var_planilhas
from permissoes import var_permissoes
from e_mails import var_emails

# Definir escopos da API do Google
API_SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

# Função para autenticar a conta Google e obter o token de acesso
def autenticar_google():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', API_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            pass
    return creds

# Função para criar os serviços da API Google Sheets e Drive
def criar_servico_google():
    creds = autenticar_google()
    service_sheets = build('sheets', 'v4', credentials=creds)
    service_drive = build('drive', 'v3', credentials=creds)
    return service_sheets, service_drive

# Função para listar planilhas no Google Drive
def listar_planilhas_drive(service_drive):
    try:
        results = service_drive.files().list(q="mimeType='application/vnd.google-apps.spreadsheet'", fields="files(id, name)").execute()
        planilhas = results.get('files', [])
        return planilhas
    except HttpError as error:
        messagebox.showerror("Erro", f"Erro ao listar planilhas: {error}")
        return []

# Função para listar pastas no Google Drive
def listar_pastas_drive(service_drive):
    try:
        results = service_drive.files().list(q="mimeType='application/vnd.google-apps.folder'", fields="files(id, name)").execute()
        pastas = results.get('files', [])
        return pastas
    except HttpError as error:
        messagebox.showerror("Erro", f"Erro ao listar pastas: {error}")
        return []

# Função para criar pasta no Google Drive
def criar_pasta_drive(service_drive, pasta_nome, pasta_pai_id=None):
    try:
        file_metadata = {
            'name': pasta_nome,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [pasta_pai_id] if pasta_pai_id else []
        }
        pasta = service_drive.files().create(body=file_metadata, fields='id').execute()
        return pasta['id']
    except HttpError as error:
        messagebox.showerror("Erro", f"Erro ao criar pasta: {error}")
        return None

# Função para copiar a planilha para outra pasta no Google Drive
def copiar_planilha_para_drive(service_drive, planilha_id, pasta_id, nome_planilha_original):
    try:
        # Renomear a planilha copiada para incluir o prefixo "Copia"
        nome_copia = "Copia - " + nome_planilha_original
        
        file_metadata = {
            'name': nome_copia,  # Nome da cópia com prefixo "Copia"
            'parents': [pasta_id]
        }
        
        copied_file = service_drive.files().copy(fileId=planilha_id, body=file_metadata).execute()
        messagebox.showinfo("Sucesso", f"Planilha '{nome_copia}' copiada com sucesso para a pasta '{copied_file['name']}'!")
    except HttpError as error:
        messagebox.showerror("Erro", f"Ocorreu um erro ao copiar a planilha: {error}")

# Função para aplicar permissões nos arquivos do Google Drive
def aplicar_permissoes_drive(service_drive, planilha_id, emails):
    try:
        for email in emails:
            # Aplicar permissão de edição (editor) para cada e-mail
            permission = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': email
            }
            service_drive.permissions().create(fileId=planilha_id, body=permission).execute()
        messagebox.showinfo("Sucesso", "Permissões de edição aplicadas com sucesso!")
    except HttpError as error:
        messagebox.showerror("Erro", f"Erro ao aplicar permissões: {error}")

# Função para remover proteção de células (colunas/abas) no Google Sheets
def remover_protecao_celulas(service_sheets, planilha_id, abas, colunas):
    try:
        for aba in abas:
            # Identificar a aba no Google Sheets
            sheet_metadata = service_sheets.spreadsheets().get(spreadsheetId=planilha_id).execute()
            sheet = next(sheet for sheet in sheet_metadata['sheets'] if sheet['properties']['title'] == aba)
            sheet_id = sheet['properties']['sheetId']
            
            # Remover proteção de células nas colunas selecionadas
            for coluna in colunas:
                range_ = f"{aba}!{coluna}1:{coluna}1000"
                requests = [{
                    'updateCells': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 0,
                            'endRowIndex': 1000,
                            'startColumnIndex': ord(coluna) - 65,
                            'endColumnIndex': ord(coluna) - 65 + 1
                        },
                        'rows': [{
                            'values': [{
                                'userEnteredValue': {'stringValue': 'Permissão aplicada'}
                            }]
                        }],
                        'fields': 'userEnteredValue'
                    }
                }]
                service_sheets.spreadsheets().batchUpdate(spreadsheetId=planilha_id, body={'requests': requests}).execute()
        messagebox.showinfo("Sucesso", "Proteção removida das células selecionadas!")
    except HttpError as error:
        messagebox.showerror("Erro", f"Erro ao remover proteção das células: {error}")

# Função para aplicar permissões no Google Sheets (edição de colunas e abas)
def aplicar_permissoes_sheets(service_sheets, planilha_id, abas, colunas):
    try:
        remover_protecao_celulas(service_sheets, planilha_id, abas, colunas)
        messagebox.showinfo("Sucesso", "Permissões de edição de células aplicadas com sucesso!")
    except HttpError as error:
        messagebox.showerror("Erro", f"Erro ao aplicar permissões no Google Sheets: {error}")

# Função para aplicar as permissões em planilhas e abas
def aplicar_permissoes(service_sheets, service_drive, lista_processadas):
    try:
        planilhas = var_planilhas
        abas = var_permissoes["Abas"]
        colunas = var_permissoes["Coluna"]
        emails = var_emails

        # Listar planilhas no Drive
        planilhas_drive = listar_planilhas_drive(service_drive)
        for planilha in planilhas:
            for planilha_drive in planilhas_drive:
                if planilha == planilha_drive['name']:
                    planilha_id = planilha_drive['id']
                    aplicar_permissoes_drive(service_drive, planilha_id, emails)
                    aplicar_permissoes_sheets(service_sheets, planilha_id, abas, colunas)
                    lista_processadas.insert(tk.END, f"{planilha} - Abas: {', '.join(abas)} | Colunas: {', '.join(colunas)} | E-mails: {', '.join(emails)}")

        messagebox.showinfo("Sucesso", "Permissões aplicadas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao aplicar permissões: {e}")

# Função para selecionar planilha e pasta de destino para copiar a planilha
def selecionar_planilha_e_pasta(service_sheets, service_drive):
    # Listar planilhas e pastas
    planilhas = listar_planilhas_drive(service_drive)
    pastas = listar_pastas_drive(service_drive)

    if not planilhas or not pastas:
        messagebox.showerror("Erro", "Não há planilhas ou pastas disponíveis.")
        return
    
    # Exibir lista de planilhas para o usuário escolher
    lista_planilhas = [planilha['name'] for planilha in planilhas]
    selected_planilha = simpledialog.askstring("Escolha a Planilha", "Escolha uma planilha para copiar:", initialvalue=lista_planilhas[0])
    
    if selected_planilha not in lista_planilhas:
        messagebox.showerror("Erro", "Planilha não encontrada!")
        return
    
    planilha_id = next(planilha['id'] for planilha in planilhas if planilha['name'] == selected_planilha)
    nome_planilha_original = selected_planilha
    
    # Pedir para o usuário digitar o caminho completo da pasta no formato Testes\\unit02\\02-2025
    caminho_pasta = simpledialog.askstring("Escolha o Caminho da Pasta", "Digite o caminho completo da pasta (exemplo: Testes\\unit02\\02-2025):")
    
    if not caminho_pasta:
        messagebox.showerror("Erro", "Caminho da pasta não fornecido!")
        return
    
    # Dividir o caminho da pasta em partes
    pastas_selecionadas = caminho_pasta.split('\\')

    # Iniciar a criação das pastas (se não existirem)
    pasta_pai_id = None
    for pasta in pastas_selecionadas:
        pasta_id = criar_pasta_drive(service_drive, pasta, pasta_pai_id)
        if pasta_id is None:
            return
        pasta_pai_id = pasta_id

    # Copiar a planilha para a pasta selecionada
    copiar_planilha_para_drive(service_drive, planilha_id, pasta_pai_id, nome_planilha_original)

# Função para criar a interface gráfica
def criar_interface():
    root = tk.Tk()
    root.title("Gestão de Planilhas e Permissões")
    root.geometry("800x600")

    # Criar lista para exibir as planilhas processadas
    lista_processadas = tk.Listbox(root, width=100, height=10)
    lista_processadas.pack(pady=10)

    # Botão para aplicar permissões
    btn_aplicar_permissoes = tk.Button(root, text="Aplicar Permissões", command=lambda: aplicar_permissoes(service_sheets, service_drive, lista_processadas))
    btn_aplicar_permissoes.pack(pady=5)

    # Botão para copiar planilha
    btn_copiar_planilha = tk.Button(root, text="Copiar Planilha", command=lambda: selecionar_planilha_e_pasta(service_sheets, service_drive))
    btn_copiar_planilha.pack(pady=5)

    root.mainloop()

# Iniciar a autenticação e a interface gráfica
if __name__ == "__main__":
    service_sheets, service_drive = criar_servico_google()
    criar_interface()
