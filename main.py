import pandas as pd
import openpyxl
from funcoes import carregar_arquivo, mapear_colunas, escolher_arquivo

# Carrega o arquivo CSV ou Excel
print("Selecione o arquivo CSV ou Excel: ")
arquivo = escolher_arquivo()
df = carregar_arquivo(arquivo)

# Mapeia as colunas
coluna_nome, coluna_salario, coluna_departamento = mapear_colunas(df)

# Converte a coluna "Salario" para float
df[coluna_salario] = df[coluna_salario].astype(float)

# Calcula a soma dos salários
total_salarios = df[coluna_salario].sum()

# Calcula a média dos salários
media_salarios = df[coluna_salario].mean()

# Conta o número de funcionários por departamento
contagem_departamentos = df[coluna_departamento].value_counts()

# Encontra o funcionário com o maior salário
maior_salario = df.loc[df[coluna_salario].idxmax()]

# Encontra o funcionário com o menor salário
menor_salario = df.loc[df[coluna_salario].idxmin()]

# Salva os resultados em um novo arquivo CSV
resultados = {
    "Total de Salários": [total_salarios], 
    "Média de Salários": [media_salarios],
    "Funcionário com o Maior Salário": [maior_salario[coluna_nome]],
    "Maior Salário": [maior_salario[coluna_salario]],
    "Funcionário com o Menor Salário": [menor_salario[coluna_nome]],
    "Menor Salário": [menor_salario[coluna_salario]],
}

resultados_df = pd.DataFrame(resultados)
with pd.ExcelWriter("relatorio_salarios.xlsx") as writer:
    # Aba Resumo
    resultados_df.to_excel(writer, index=False, sheet_name="Resumo")
    # Aba Funcionarios
    df.to_excel(writer, sheet_name="Funcionarios" , index=False)
    # Aba Dados por Departamento
    df.groupby(coluna_departamento).agg(
        Funcionarios = (coluna_nome,"count"),
        Total_Salario = (coluna_salario,"sum"),
        Media_Salario = (coluna_salario,"mean")
    ).to_excel(writer, sheet_name="Dados por Departamento")

input("Pressione ENTER para fechar...")

