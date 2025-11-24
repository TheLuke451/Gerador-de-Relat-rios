# funcoes.py
import pandas as pd
import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def carregar_arquivo(arquivo):
    try:
        if arquivo.endswith(".csv"):
            df = pd.read_csv(arquivo)
        else:
            df = pd.read_excel(arquivo)
    except Exception as e:
        print("Erro ao encontrar o arquivo:", e)
        raise
    return df

def mapear_colunas(df):
    # Remove espaços em branco das colunas e padroniza para minúsculas
    df.columns = df.columns.str.strip().str.lower()

    # Mostra as colunas disponíveis para o usuário
    print("Colunas disponíveis no arquivo:", df.columns.tolist())

    # Solicita ao usuário os nomes das colunas relevantes
    coluna_nome = input("Digite o nome da coluna que representa o nome: ").strip().lower()
    coluna_salario = input("Digite o nome da coluna que representa o salário: ").strip().lower()
    coluna_departamento = input("Digite o nome da coluna que representa o departamento: ").strip().lower()

    # Verifica se as colunas existem no DataFrame
    if coluna_nome not in df.columns or coluna_salario not in df.columns or coluna_departamento not in df.columns:
        print(f"Coluna(s) inexistente(s), tente novamente.")
        exit()

    return coluna_nome, coluna_salario, coluna_departamento

def escolher_arquivo():
    Tk().withdraw() # Esconde a janela principal do Tkinter
    arquivo = askopenfilename(title="Selecione o arquivo CSV ou Excel", filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx;*.xls")])
    return arquivo