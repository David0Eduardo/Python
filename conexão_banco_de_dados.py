import pyodbc

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

        # Criando um cursor para interações com o banco de dados
        cursor = conn.cursor()

        # Fechando o cursor e a conexão
        cursor.close()
        conn.close()

    except pyodbc.Error as e:
        # Captura e imprime erros específicos relacionados ao ODBC
        print("Erro ao conectar-se ao banco de dados:")
        print("Código do erro:", e.args[0])
        print("Descrição do erro:", e.args[1])

    except Exception as e:
        # Captura e imprime qualquer outro tipo de erro
        print("Um erro inesperado ocorreu:")
        print(e)

# Configurações do banco de dados
server = 'localhost'  # Endereço IP do servidor SQL Server
instance = 'TESTES'     # Nome da instância do SQL Server
database = 'base01'       # Nome do banco de dados específico
username = 'python'      # Nome do usuário
password = ''   # Senha do usuário

# Chamada para conectar ao banco de dados específico na instância 'ASUS_DB'
connect_to_database(server, instance, database, username, password)
