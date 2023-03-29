import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from unidecode import unidecode

# Função que recebe um DOI e retorna os autores do artigo correspondente
def get_authors(doi):
    url = f'https://api.crossref.org/works/{doi}'
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if 'author' in data['message']:
            authors = data['message']['author']
            authors_list = []
            for author in authors:
                if 'given' in author and 'family' in author:
                    author_name = f"{author['given']} {author['family']}"
                    author_name = unidecode(author_name).lower()
                    authors_list.append(author_name)
            return authors_list
        else:
            return None
    else:
        return None

# Cria uma janela para o usuário selecionar o arquivo
file_path = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])

# Se o usuário cancelar a seleção do arquivo, encerra o programa
if not file_path:
    exit()

# Pergunta ao usuário o nome da planilha e da coluna que contém os DOIs
sheet_name = simpledialog.askstring(title="Nome da planilha", prompt="Digite o nome da planilha que contém a coluna de DOIs:")
column_name = simpledialog.askstring(title="Nome da coluna", prompt="Digite o nome da coluna que contém os DOIs:")

# Carrega o arquivo excel selecionado pelo usuário
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Verifica se o nome da coluna existe no arquivo
if column_name not in df.columns:
    print(f"Coluna '{column_name}' não encontrada no arquivo.")
    exit()

# Faz a busca na API CrossRef para obter os autores de cada DOI na coluna selecionada
authors_column = []
for doi in df[column_name]:
    authors = get_authors(doi)
    if authors:
        authors_column.append(', '.join(authors))
    else:
        authors_column.append(None)

# Adiciona a nova coluna com os autores no dataframe
df['Autores'] = authors_column



# Cria um novo arquivo excel com a coluna de autores adicionada
new_file_path = filedialog.asksaveasfilename(title="Salvar como", defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
if new_file_path:
    df.to_excel(new_file_path, sheet_name=sheet_name, index=False)
    print("Arquivo salvo com sucesso!")
else:
    print("Operação cancelada pelo usuário.")
