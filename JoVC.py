import pandas as pd
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import filedialog, simpledialog

# Função que recebe uma lista de strings e substitui as similares
def substituir_strings_similares(strings):
    strings_substituidas = []
    for i in range(len(strings)):
        if i not in strings_substituidas:
            string1 = strings[i]
            for j in range(i+1, len(strings)):
                if j not in strings_substituidas:
                    string2 = strings[j]
                    if SequenceMatcher(None, string1, string2).ratio() > 0.8:
                        if len(string1) >= len(string2):
                            strings[j] = string1
                            strings_substituidas.append(j)
                        else:
                            strings[i] = string2
                            strings_substituidas.append(i)
    return strings

# Cria uma janela para o usuário selecionar o arquivo
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])

# Se o usuário cancelar a seleção do arquivo, encerra o programa
if not file_path:
    exit()

# Pergunta ao usuário o nome da planilha e da coluna a ser padronizada
sheet_name = simpledialog.askstring(title="Nome da planilha", prompt="Digite o nome da planilha que contém a coluna a ser padronizada:")
column_name = simpledialog.askstring(title="Nome da coluna", prompt="Digite o nome da coluna a ser padronizada:")

# Carrega o arquivo excel selecionado pelo usuário
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Verifica se o nome da coluna existe no arquivo
if column_name not in df.columns:
    print(f"Coluna '{column_name}' não encontrada no arquivo.")
    exit()

# Substitui as strings similares na coluna selecionada
df[column_name] = substituir_strings_similares(df[column_name].tolist())

# Cria um novo arquivo excel com a coluna padronizada
new_file_path = filedialog.asksaveasfilename(title="Salvar como", defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
if new_file_path:
    df.to_excel(new_file_path, sheet_name=sheet_name, index=False)
    print("Arquivo salvo com sucesso!")
else:
    print("Operação cancelada pelo usuário.")
