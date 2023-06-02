from tkinter import filedialog, messagebox
from openpyxl.utils.dataframe import dataframe_to_rows

import openpyxl
import pandas as pd

messagebox.showinfo(title='Introducao', message='Selecione primeiramente o arquivo de origem e depois o arquivo de destino ')

path_origem = filedialog.askopenfilenames()
path_origem = path_origem[0]
path_destino = filedialog.askopenfilenames()
path_destino = path_destino[0]

try:
    df1 = pd.read_excel(path_origem)
    df2 = pd.read_excel(path_destino, sheet_name='TESTE')
except Exception as e:
    messagebox.showerror(title='ERROR', message='Erro ao ler os arquivos Excel: ' + str(e))

try:
    id_a_serem_procurados = df2['CONTRATO'].values
except Exception as e:
    messagebox.showerror(title='ERROR', message='Erro ao buscar id a serem consultados ' + str(e))

dados_filtrados = []

for id in id_a_serem_procurados:
    try:
        filtro = df1['TRATO_EBT_NET'] == id
        dados_filtrados.append(df1.loc[filtro])
    except Exception as e:
        messagebox.showerror(title='ERROR', message='Erro ao buscar dados ' + str(e))

messagebox.showinfo(title='Geracao de Dados', message='Iniciando processo de gravação dos dados')
df = pd.concat(dados_filtrados)

try:
    workbook = openpyxl.load_workbook(path_destino)

    worksheet = workbook.create_sheet('RELATORIO', index=0)

    for row in dataframe_to_rows(df, index=False, header=True):
        worksheet.append(row)

    workbook.save(path_destino)

    messagebox.showinfo(title='SUCESSO!', message='Dados gerados com sucesso!!')
except Exception as e:
    messagebox.showerror(title='ERROR', message='Erro ao salvar dados na planilha ' + str(e))