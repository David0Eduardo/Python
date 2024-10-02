import requests
import pandas as pd
import os

def obter_dados_da_api():
    url = ""

    headers = {
        "accept": "",
        "authorization": "="
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()  # Retorna os dados no formato JSON
    else:
        print('Erro na solicitação:', response.status_code)
        return []

# Função para processar e salvar dados
def processar_dados(dados, caminho_csv, caminho_excel):
    # Converter os dados para um DataFrame do pandas
    df = pd.DataFrame(dados)

    # Salvar os dados em um arquivo CSV
    df.to_csv(caminho_csv, index=False)

    # Salvar os dados em um arquivo Excel
    df.to_excel(caminho_excel, index=False)

# Função principal
def main():
    # Armazenar os dados em uma lista temporária
    dados_temp = obter_dados_da_api()

    if dados_temp:
        # Definir caminhos de salvamento
        caminho_csv = r''
        caminho_excel = r''

        # Verificar se o diretório existe, se não, criá-lo
        os.makedirs(os.path.dirname(caminho_csv), exist_ok=True)

        # Processar os dados
        processar_dados(dados_temp, caminho_csv, caminho_excel)

        # Limpar a lista temporária
        dados_temp.clear()  # Limpa a lista para liberar memória e garantir que não há dados antigos

if __name__ == "__main__":
    main()
