#!/usr/bin/env python
# coding: utf-8

# Este código tem como função principal gerar um relatório de vendas em formato PDF a partir de dados armazenados em um arquivo Excel. 
# 
# Ele realiza a leitura do arquivo Excel, processa os dados para extrair informações relevantes sobre vendas, como período do relatório, detalhes dos produtos ou serviços vendidos, e informações financeiras como total e média de vendas. 
# 
# O código também cria um gráfico visual representando a receita por tamanho do produto, e inclui seções para observações e insights. Todo esse conteúdo é formatado e organizado em um documento PDF, que é então salvo na área de trabalho do usuário, dentro de uma pasta específica.
# 
# O código inclui tratamento de erros para lidar com possíveis problemas na leitura do arquivo ou na geração do relatório, e é flexível o suficiente para se adaptar a diferentes estruturas de dados no arquivo Excel de entrada, permitindo ajustes nos nomes das colunas e nos tipos de informações apresentadas no relatório final.

# In[1]:


import os


# In[2]:


# Define o caminho da área de trabalho
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# Define o nome da nova pasta
folder_name = "pdfrelatorio"

# Cria o caminho completo da nova pasta
new_folder_path = os.path.join(desktop_path, folder_name)

# Cria a nova pasta
try:
    os.makedirs(new_folder_path, exist_ok=True)  # 'exist_ok=True' evita erro se a pasta já existir
    print(f"Pasta '{folder_name}' criada com sucesso na área de trabalho.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")


# In[3]:


import os
print(os.getcwd())


# In[4]:


import os

def encontrar_arquivo(nome_arquivo, diretorio_inicial):
    for raiz, dirs, arquivos in os.walk(diretorio_inicial):
        if nome_arquivo in arquivos:
            return os.path.join(raiz, nome_arquivo)
    return None

# Nome do arquivo que você está procurando
nome_arquivo = "vendas.xlsx"

# Diretório inicial para a busca (por exemplo, a área de trabalho ou o diretório do usuário)
diretorio_inicial = os.path.expanduser("~")  # Isso define o diretório inicial como o diretório do usuário

# Chama a função para encontrar o arquivo
caminho_arquivo = encontrar_arquivo(nome_arquivo, diretorio_inicial)

if caminho_arquivo:
    print(f"Arquivo encontrado: {caminho_arquivo}")
else:
    print(f"Arquivo '{nome_arquivo}' não encontrado.")


# In[5]:


get_ipython().system('pip install pandas openpyxl fpdf matplotlib')


# In[6]:


import os
import pandas as pd
from fpdf import FPDF
import matplotlib.pyplot as plt

# Função para criar gráficos
def criar_grafico(df):
    plt.figure(figsize=(10, 6))
    df.groupby('tamanho')['preco'].sum().plot(kind='bar')
    plt.title('Receita por Tamanho')
    plt.xlabel('Tamanho')
    plt.ylabel('Receita')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('grafico_receita.png')
    plt.close()

# Função para gerar o relatório em PDF
def gerar_relatorio(df, caminho_pdf):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Título
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, 'Relatório de Vendas', ln=True, align='C')

    # Dados Gerais
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, 'Dados Gerais', ln=True)
    
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, f'Período do relatório: {df["data"].min().date()} a {df["data"].max().date()}', ln=True)
    pdf.cell(0, 10, 'Responsável pelas vendas: Nome do Vendedor', ln=True)  # Substitua conforme necessário
    pdf.cell(0, 10, 'Localização: Região/Cidade/Unidade', ln=True)  # Substitua conforme necessário

    # Informações sobre produtos ou serviços
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, 'Informações Sobre os Produtos ou Serviços', ln=True)

    pdf.set_font("Arial", '', 12)
    for index, row in df.iterrows():
        if all(col in row for col in ['tamanho', 'preco']):
            pdf.cell(0, 10, f'Tamanho: {row["tamanho"]}, Preço: R${row["preco"]:.2f}', ln=True)

    # Detalhes Financeiros
    total_vendas = df['preco'].sum()
    media_vendas = total_vendas / df['data'].nunique() if df['data'].nunique() > 0 else 0

    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, 'Detalhes Financeiros', ln=True)

    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, f'Total de Vendas: R${total_vendas:.2f}', ln=True)
    pdf.cell(0, 10, f'Média de Vendas: R${media_vendas:.2f}', ln=True)

    # Gráfico
    criar_grafico(df)
    pdf.image('grafico_receita.png', x=30, w=150)

    # Observações e Insights
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, 'Observações e Insights', ln=True)

    pdf.set_font("Arial", '', 12)
    pdf.multi_cell(0, 10,
                   "Análise de resultados: Pontos positivos e negativos no desempenho.\n"
                   "Recomendações: Estratégias sugeridas para melhorar as vendas.")

    # Salva o PDF
    pdf.output(caminho_pdf)

# Caminho da pasta onde o relatório será salvo
caminho_pasta = os.path.join(os.path.expanduser("~"), "Desktop", "pdfrelatorio")
os.makedirs(caminho_pasta, exist_ok=True)

# Caminho do arquivo Excel (ajustado conforme solicitado)
caminho_excel = r'C:\Users\55329\Downloads\vendas.xlsx'

# Inicializa caminho_pdf fora do bloco try-except para evitar NameError
caminho_pdf = os.path.join(caminho_pasta, "relavendas.pdf")

# Tenta ler o arquivo Excel e trata possíveis erros
try:
    df = pd.read_excel(caminho_excel)

    # Verifica as colunas disponíveis no DataFrame
    print("Colunas disponíveis no DataFrame:", df.columns.tolist())

except FileNotFoundError:
    print(f"Erro: O arquivo '{caminho_excel}' não foi encontrado. Verifique se ele existe.")
except Exception as e:
    print(f"Ocorreu um erro ao ler o arquivo: {e}")
else:
   # Gera o relatório apenas se não houver erros ao ler o arquivo
   gerar_relatorio(df, caminho_pdf)

print(f"Relatório gerado com sucesso em: {caminho_pdf}")


# 

# In[ ]:




