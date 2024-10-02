import pyodbc
import requests

def connect_to_database(server, instance, database, username, password):
    # Construindo a string de conexão para a instância específica
    connection_string = (
        f'DRIVER={{SQL Server}};'
        f'SERVER={server}\\{instance};'  # Inclui o nome da instância
        f'DATABASE={database};'
        f'UID={username};'
        f'PWD={password};'
    )
    
    try:
        # Estabelecendo a conexão com o banco de dados SQL Server
        conn = pyodbc.connect(connection_string)
        print(f"Conexão estabelecida com sucesso no banco de dados '{database}' na instância '{instance}'!")

        return conn  # Retorna a conexão para uso posterior

    except pyodbc.Error as e:
        # Captura e imprime erros específicos relacionados ao ODBC
        print("Erro ao conectar-se ao banco de dados:")
        print("Código do erro:", e.args[0])
        print("Descrição do erro:", e.args[1])
        return None

    except Exception as e:
        # Captura e imprime qualquer outro tipo de erro
        print("Um erro inesperado ocorreu:")
        print(e)
        return None

def obter_dados_da_api(url, headers):
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()  # Retorna os dados no formato JSON
    else:
        print('Erro na solicitação:', response.status_code)
        return []

def criar_tabela(conn):
    cursor = conn.cursor()
    cursor.execute('''
        IF OBJECT_ID('dbo.base', 'U') IS NOT NULL
            DROP TABLE dbo.base;
        
        CREATE TABLE dbo.Atividades (
            register DATETIME,
            
        );
    ''')
    conn.commit()

def inserir_dados(conn, dados):
    cursor = conn.cursor()
    for item in dados:
        cursor.execute('''
            INSERT INTO dbo.Atividades (
                register
            ) VALUES (?)
        ''', (
            item.get('register'),

        ))
    conn.commit()

def main():
    # Configurações do banco de dados
    server = 'localhost'  # Endereço IP do servidor SQL Server
    instance = 'teste'     # Nome da instância do SQL Server
    database = 'base'       # Nome do banco de dados específico
    username = 'python'      # Nome do usuário
    password = ''   # Senha do usuário    

    # Conectar ao banco de dados
    conn = connect_to_database(server, instance, database, username, password)
    
    if conn:
        # Configurações da API
        url = ""  # Substitua pela URL da sua API
        headers = {
            "accept": "",
            "authorization": ""
        }

        # Obter dados da API
        dados = obter_dados_da_api(url, headers)
        
        if dados:
            # Criar tabela
            criar_tabela(conn)
            
            # Inserir dados
            inserir_dados(conn, dados)

        # Fechar conexão
        conn.close()

if __name__ == "__main__":
    main()
