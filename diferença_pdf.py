import PyPDF2
import difflib
import tkinter as tk
from tkinter import filedialog

def extrair_texto(pdf_path):
    texto = ""
    contador_imagens = 0
    with open(pdf_path, 'rb') as file:
        leitor_pdf = PyPDF2.PdfReader(file)
        for pagina in range(len(leitor_pdf.pages)):
            texto += leitor_pdf.pages[pagina].extract_text()
            contador_imagens += len(leitor_pdf.pages[pagina].images)
    return texto, contador_imagens

def extrair_diferencas(texto1, texto2):
    diffs = difflib.ndiff(texto1.splitlines(), texto2.splitlines())
    return '\n'.join(x for x in diffs if x.startswith('- ') or x.startswith('+ '))

def contar_imagens(pdf_path):
    contador_imagens = 0
    with open(pdf_path, 'rb') as file:
        leitor_pdf = PyPDF2.PdfReader(file)
        for pagina in range(len(leitor_pdf.pages)):
            contador_imagens += len(leitor_pdf.pages[pagina].images)
    return contador_imagens

def selecionar_arquivo():
    arquivo_path = filedialog.askopenfilename(title="Selecione o arquivo PDF")
    return arquivo_path

def comparar_arquivos():
    arquivo1_path = selecionar_arquivo()
    arquivo2_path = selecionar_arquivo()

    texto_arquivo1, imagens_arquivo1 = extrair_texto(arquivo1_path)
    texto_arquivo2, imagens_arquivo2 = extrair_texto(arquivo2_path)

    diferencas = extrair_diferencas(texto_arquivo1, texto_arquivo2)

    resultado_texto.delete(1.0, tk.END)
    resultado_texto.insert(tk.END, "Diferenças encontradas:\n")
    resultado_texto.insert(tk.END, diferencas)
    resultado_texto.insert(tk.END, f"\n\nQuantidade de imagens no arquivo 1: {imagens_arquivo1}")
    resultado_texto.insert(tk.END, f"\nQuantidade de imagens no arquivo 2: {imagens_arquivo2}")

# Criar a janela principal
janela = tk.Tk()
janela.title("Comparador de Arquivos PDF")

# Botão para comparar arquivos
botao_comparar = tk.Button(janela, text="Comparar Arquivos", command=comparar_arquivos)
botao_comparar.grid(row=0, column=0, pady=10, padx=10, sticky="ew")

# Área de texto para exibir o resultado
resultado_texto = tk.Text(janela, height=30, width=90)
resultado_texto.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

# Configurar o redimensionamento da janela para preencher toda a área disponível
janela.grid_rowconfigure(1, weight=1)
janela.grid_columnconfigure(0, weight=1)

# Executar o loop principal da janela
janela.mainloop()
