import os
import glob
import time
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
#importação das bibliotecas necessarias

def formatacao_planilha(nome_planilha):
        #função utilizada para formatar a planilha e deixar as colunas e linhas com tamanho ideal

        wb = openpyxl.load_workbook(nome_planilha) #Abre a planilha excel

        ws = wb.active #pega a aba atual da planilha
        ws.title = "Dados" #nomeia a aba

        for col in ws.columns: # aqui ele vai fazer uma verificação em cada coluna, pegando o tamanho atual dela
                max_length = 0
                column_letter = get_column_letter(col[0].column) #aqui ela vai ir coluna por coluna, a column_letter transforma os numeros do col[0] e transforma numa letra correspondete, por exemplo 1 = A

                for cell in col:
                        if cell.value is not None:
                                max_length = max(max_length, len(str(cell.value)))

                ws.column_dimensions[column_letter].width = max_length + 2

        wb.save(nome_planilha)

def remover_data(df):
        print("deseja remover as colunas com datas? Y/N")
        pergunta = input().upper()

        if pergunta == "Y":

                colunas_alvo = df.filter(like='Data de conclusão').columns
                df = df.drop(columns = colunas_alvo)
                print("remoção realizada")

        return df

arquivos_csv = glob.glob('*.csv')

if not arquivos_csv:
	print("nenhum arquivo csv encontrado")
else:
	print(f"Encontrado(s) {len(arquivos_csv)} arquivos para conversão.\n")

for arquivo_csv in arquivos_csv:
        nome_base = os.path.splitext(arquivo_csv)[0]
        arquivo_excel = f"{nome_base}.xlsx"

        if os.path.exists(arquivo_excel):
                print(f"o arquivo {arquivo_excel} ja existe, pulando conversao\n")
                continue
        else:
           print(f"convertendo {arquivo_csv} para {arquivo_excel}..\n")

        try:
                df= pd.read_csv(arquivo_csv, sep=',', encoding='utf-8')
                df = remover_data(df)
                df.to_excel(arquivo_excel, index =False)

                formatacao_planilha(arquivo_excel)


                if os.path.exists(arquivo_excel):
                        caminho_arquivo = os.path.abspath(arquivo_excel)
                        print(f"o arquivo {arquivo_excel} passou a existir na pasta {caminho_arquivo}\n")
                else:
                        print("Arquivo nao encontrado")
                print("conversao realizada com sucesso\n")

        except Exception as e:
                print(f"Erro ao converter {arquivo_csv} : {e}")
                time.sleep(5.5)
                break

print(f"\nProcesso Finalizado!\n")
sair = input("aperte enter para finalizar")