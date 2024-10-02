import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Caminho para a pasta contendo o arquivo de dicionário
caminho_dicionario = r""

# Função para carregar frases de um arquivo .txt
def carregar_frases(caminho_arquivo_txt):
    try:
        with open(caminho_arquivo_txt, 'r', encoding='utf-8') as file:
            frases = [linha.strip() for linha in file if linha.strip()]
        return frases
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao ler o arquivo de frases: {e}")
        return []

# Função para processar o arquivo Excel
def processar_excel(caminho_arquivo, caminho_saida, frases_classificar):
    try:
        # Carregar o arquivo Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Verificar se a coluna 'Memo' existe no DataFrame
        if 'Memo' not in df.columns:
            raise ValueError("A coluna 'Memo' não foi encontrada no arquivo Excel.")
        
        # Inicializar a coluna "J" com 'outros'
        df['J'] = 'outros'
        
        # Atualizar a coluna "J" para linhas que contêm qualquer uma das frases da lista
        mask = df['Memo'].apply(lambda x: any(frase in str(x) for frase in frases_classificar))
        df.loc[mask, 'J'] = 'a classificar'
        
        # Salvar o arquivo atualizado
        df.to_excel(caminho_saida, index=False)
        
        messagebox.showinfo("Sucesso", f"Arquivo processado e salvo como '{caminho_saida}'")
    except ValueError as ve:
        messagebox.showerror("Erro", str(ve))
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Função para abrir o diálogo de seleção de arquivos e salvar o arquivo
def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if caminho_arquivo:
        # Usar o caminho do arquivo de dicionário especificado
        frases_classificar = carregar_frases(caminho_dicionario)
        if not frases_classificar:
            return  # Se não conseguiu carregar as frases, interrompe a execução

        # Abrir o diálogo para salvar o arquivo de saída
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialfile=caminho_arquivo.replace('.xlsx', '_atualizado.xlsx')
        )
        if caminho_saida:
            processar_excel(caminho_arquivo, caminho_saida, frases_classificar)

# Configuração da interface gráfica
root = tk.Tk()
root.title("Processador de Excel")

# Botão para selecionar o arquivo
btn_selecionar = tk.Button(root, text="Selecionar Arquivo Excel", command=selecionar_arquivo)
btn_selecionar.pack(padx=20, pady=20)

# Iniciar a interface gráfica
root.mainloop()
