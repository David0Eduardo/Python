import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Caminho para o JSON das credenciais OAuth (download do Console Google)
CREDENTIALS_OAUTH_FILE = 'credentials_oauth.json'
TOKEN_PICKLE = 'token.pickle'

# Escopos para usar o Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

# ID da pasta de destino no Google Drive
PASTA_DESTINO_ID = ''

# Pasta local com arquivos .docx
PASTA_LOCAL = r".DAVID\Python\ENVIAR DOC PARA O DRIVE CONVERTIDO"

def autenticar_oauth():
    creds = None
    # Tenta carregar token salvo localmente
    if os.path.exists(TOKEN_PICKLE):
        with open(TOKEN_PICKLE, 'rb') as token:
            creds = pickle.load(token)

    # Se não tem credenciais válidas, pede login no navegador
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_OAUTH_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Salva token para próximas execuções
        with open(TOKEN_PICKLE, 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service

def upload_docx_com_conversao(caminho_arquivo, nome_arquivo_sem_extensao, service):
    media = MediaFileUpload(
        caminho_arquivo,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

    file_metadata = {
        'name': nome_arquivo_sem_extensao,
        'mimeType': 'application/vnd.google-apps.document',
        'parents': [PASTA_DESTINO_ID]
    }

    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name, webViewLink'
    ).execute()

    print(f"✅ {file['name']} convertido e enviado como Google Docs.")
    print(f"🔗 Link: {file['webViewLink']}\n")

def enviar_todos_docx_convertidos():
    service = autenticar_oauth()
    for nome_arquivo in os.listdir(PASTA_LOCAL):
        if nome_arquivo.lower().endswith('.docx'):
            caminho_completo = os.path.join(PASTA_LOCAL, nome_arquivo)
            nome_sem_extensao = os.path.splitext(nome_arquivo)[0]
            upload_docx_com_conversao(caminho_completo, nome_sem_extensao, service)

if __name__ == "__main__":
    enviar_todos_docx_convertidos()
