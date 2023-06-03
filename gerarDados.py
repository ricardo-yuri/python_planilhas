from tkinter import filedialog, messagebox
from openpyxl.utils.dataframe import dataframe_to_rows

import tkinter as tk
import openpyxl
import pandas as pd


def selecionar_arquivo_origem():
    global path_origem;
    path_origem = filedialog.askopenfilename()
    lbl_arquivo_origem.config(text="Arquivo de origem: " + str(path_origem))


def selecionar_arquivo_destino():
    global path_destino;
    path_destino = filedialog.askopenfilename()
    lbl_arquivo_destino.config(text="Arquivo de destino: " + str(path_destino))


def cancelar():
    janela.destroy()


def gerar():
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

    janela.destroy()


janela = tk.Tk()

btn_origem = tk.Button(janela, text="Selecionar arquivo de origem", command=selecionar_arquivo_origem)
btn_origem.grid(row=0, column=0)

lbl_arquivo_origem = tk.Label(janela, text="Arquivo de origem:")
lbl_arquivo_origem.grid(row=0, column=1)

btn_destino = tk.Button(janela, text="Selecionar arquivo de destino", command=selecionar_arquivo_destino)
btn_destino.grid(row=1, column=0)

lbl_arquivo_destino = tk.Label(janela, text="Arquivo de destino:")
lbl_arquivo_destino.grid(row=1, column=1)

btn_gerar = tk.Button(janela, text="Gerar", command=gerar)
btn_gerar.grid(row=2, column=0)

btn_cancelar = tk.Button(janela, text="Cancelar", command=cancelar)
btn_cancelar.grid(row=2, column=1)

janela.mainloop()
