import os
import glob

from termcolor import colored
#importação das bibliotecas necessarias

print(colored('         Programa de Conversão CSV', 'light_blue'))

def conversao(arquivo_excel, df):
        with pd.ExcelWriter(arquivo_excel, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Dados", index= False)
                formatacao_planilha(writer.sheets["Dados"])


def formatacao_planilha(ws):
        for col in ws.columns:
                largura = max((len(str(c.value)) for c in col if c.value is not None), default = 0)
                ws.column_dimensions[col[0].column_letter].width = largura + 2

def remover_data(df, arquivo_excel):
        print(f"deseja remover as colunas com datas do arquivo {arquivo_excel}? S/N")
        pergunta = input().upper().strip()

        if pergunta == "Y" or pergunta == "S":

                colunas_alvo = df.filter(like='Data de conclusão').columns
                df = df.drop(columns = colunas_alvo)
                print("remoção realizada")
        return df

def main():
        arquivos_csv = glob.glob('*.csv')

        if not arquivos_csv:
                print("nenhum arquivo csv encontrado")
                print("insira um arquivo" , colored('CSV', 'white'))
        else:
                print(f"Encontrado(s) {len(arquivos_csv)} arquivos para conversão.\n")

                for arquivo_csv in arquivos_csv:
                        nome_base = os.path.splitext(arquivo_csv)[0]
                        arquivo_excel = f"{nome_base}.xlsx"

                        if os.path.exists(arquivo_excel):
                                print(f"o arquivo {arquivo_excel} ja existe, pulando conversão\n")
                                continue
                        
                        print(f"convertendo {arquivo_csv} para {arquivo_excel}..\n")

                        try:
                                df = pd.read_csv(arquivo_csv, sep=',', encoding='utf-8')
                                df = remover_data(df, arquivo_excel)
                                conversao(arquivo_excel, df)

                                if os.path.exists(arquivo_excel):
                                        caminho_arquivo = os.path.abspath(arquivo_excel)
                                        print(f"o arquivo {arquivo_excel} passou a existir na pasta {caminho_arquivo}\n")
                                else:
                                        print(colored('Arquivo não encontrado', 'red'))
                                print("conversao realizada com sucesso\n")

                        except Exception as e:
                                print(colored (f'Erro aos converter {arquivo_csv} : {e}', 'red'))
                                if os.path.exists(arquivo_excel):
                                        os.remove(arquivo_excel)

if __name__ == "__main__":
        import pandas as pd
        main()

        print(f"\nProcesso Finalizado!\n")
        input("aperte enter para finalizar")