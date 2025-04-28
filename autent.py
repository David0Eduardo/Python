import os
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

# Arquivo de credenciais
CLIENT_SECRET_FILE = 'credentials.json'  # Verifique se o arquivo credentials.json está correto
API_SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def autenticar_google():
    creds = None
    # O arquivo token.json armazena os tokens de acesso e refresh
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', API_SCOPES)
    
    # Se não houver credenciais ou as credenciais expiraram
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, API_SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Salva as credenciais para a próxima execução
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

# Usando o serviço do Google Sheets ou Google Drive
def criar_servico_google():
    creds = autenticar_google()
    service_sheets = build('sheets', 'v4', credentials=creds)
    service_drive = build('drive', 'v3', credentials=creds)
    return service_sheets, service_drive

# Testar a autenticação
service_sheets, service_drive = criar_servico_google()
