import pandas as pd
import openpyxl

path_origem = input("Nome do arquivo Excel de origem: ")
path_destino = input("Nome do arquivo Excel de destino: ")

try:
    df1 = pd.read_excel(path_origem)
    df2 = pd.read_excel(path_destino, sheet_name='TESTE')
except Exception as e:
    print("Erro ao ler os arquivos Excel: ", e)


try:
    id_a_serem_procurados = df2['CONTRATO'].values
    print('Quantidade de id a serem procuradas: ' + id_a_serem_procurados.size)
except Exception as e:
    print("Erro ao buscar id's a serem consultados ", e)

dados_filtrados = []

for id in id_a_serem_procurados:
    try:
        filtro = df1['TRATO_EBT_NET'] == id
        dados_filtrados.append(df1.loc[filtro])
    except Exception as e:
        print("Erro ao buscar dados ", e)

df = pd.concat(dados_filtrados)

try:
    with pd.ExcelWriter(path_destino, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        writer.book = openpyxl.load_workbook(path_destino)
        df.to_excel(writer, sheet_name='RELATORIO', index=False)
except Exception as e:
    print("Erro ao salvar dados na planilha")
