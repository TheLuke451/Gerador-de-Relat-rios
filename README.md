
# 📌 Gerador de Relatórios em Python (Python + Pandas)

Este projeto é uma ferramenta simples e prática para gerar relatórios estruturados a partir de arquivos CSV
ou excel contendo informações de funcionários, salários e seus respectivos departamentos.

Ele automatiza tarefas como: 

* Calculo de Soma e Média Salarial
* Identificação do Maior e Menor Salário
* Contagem de Funcionários por Departamento
* Criação de um relatório completo em Excel com várias abas organizadas

O objetivo é facilitar a análise de dados de RH ou financeiros de forma rápida e acessível, mesmo para quem não tem
familiaridade com planilhas.

## 🚀 Funcionalidades 

* Seleção do arquivo através de janela (Tkinter)
* Leitura automatizada de CSV ou Excel
* Padronização de nomes das colunas
* Mapeamento Flexível (usuário pode informar os nomes exatos das colunas)
* Processamento e análise de:
    * Total de Salários
    * Média Salarial
    * Maior Salário
    * Menor Salário
    * Quantidade de Funcionários por Departamento
* Geração automática de arquivo Excel com Três Abas:
    * **Resumo**
    * **Funcionários**
    * **Dados por Departamento**
* Versão Executável (.exe).

## ⚙️ Como usar

### 1. Abra o Aplicativo

Se estiver usando o .exe, basta executar.

### 2. Selecione o arquivo

Você pode escolher entre:
    
   * CSV
   * XLSX (Excel)

### 3. Informe as colunas

Digite quais colunas representam:

   * Nome
   * Salário
   * Departamento

### 4. O relatório será gerado automaticamente

Ele aparecerá na mesma pasta, com o nome:
**relatorio_salarios.xlsx**

## 🔧 Tecnologias Utilizadas 

   * Python
   * Pandas
   * Openpyxl
   * Tkinter
   * PyInstaller (Para gerar o executável)

## 📚 Objetivo do projeto

Este projeto foi feito com fins de estudo e portfólio com foco em:
    
   * Manipulação de dados com pandas
   * Boas práticas de organização de código
   * Integração com Tkinter
   * Geração de relatórios automatizados
   * Empacotamento com PyInstaller

Autor: **TheLuke451**



