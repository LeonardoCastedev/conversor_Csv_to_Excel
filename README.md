# Programa de Conversão CSV para Excel

Script em Python que converte todos os arquivos `.csv` de uma pasta em planilhas `.xlsx`, já com as colunas ajustadas ao tamanho do conteúdo.

## O que ele faz

- Procura todos os arquivos `.csv` na pasta onde o programa está.
- Converte cada um para `.xlsx` (mesmo nome, nova extensão), com os dados em uma aba chamada **Dados**.
- Ajusta a largura de cada coluna ao maior conteúdo dela.
- Para cada arquivo, pergunta se você quer remover as colunas cujo nome contém **"Data de conclusão"**.
- Se já existir um `.xlsx` com o mesmo nome, pula o arquivo para não sobrescrever nada.
- Se um arquivo der erro, mostra a mensagem e continua com os demais.
- No final, mostra o caminho completo de cada arquivo gerado.

## Requisitos

- Python 3.8 ou superior
- Bibliotecas:
  - `pandas`
  - `openpyxl` (usada pelo pandas para gravar o Excel, mesmo sem aparecer nos imports)
  - `termcolor`

Instalação:

```bash
pip install pandas openpyxl termcolor
```

Ou, usando um `requirements.txt` com essas três linhas:

```bash
pip install -r requirements.txt
```

Se o pip reclamar de "externally-managed-environment" (comum em Linux recente), use um ambiente virtual:

```bash
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Como usar

1. Coloque o script na mesma pasta dos CSVs que quer converter.
2. Execute:

   ```bash
   python nome_do_script.py
   ```

3. Aguarde o carregamento (na primeira execução o `pandas` demora um pouco para iniciar).
4. Responda `S` (ou `Y`) para remover as colunas de data ou `N` para mantê-las. A pergunta é feita **uma vez por arquivo**.
5. Ao final, aperte Enter para fechar.

### Exemplo

```
Programa de Conversão CSV

Encontrado(s) 2 arquivos para conversão.

convertendo alunos.csv para alunos.xlsx..
deseja remover as colunas com datas? S/N
```

## Observações e limitações

- O CSV é lido com separador `,` e codificação `utf-8`. Arquivos exportados com `;` ou em outra codificação (como `latin-1`) podem dar erro ou sair com os dados desalinhados.
- Arquivos `.xlsx` já existentes nunca são sobrescritos. Para reconverter, apague o `.xlsx` antes.
- A busca por colunas de data é feita pelo trecho `Data de conclusão` no nome da coluna. Para outros nomes, altere o texto na função `remover_data`.

## Estrutura do código

| Função | Para que serve |
|---|---|
| `remover_data(df)` | Pergunta ao usuário e remove as colunas de data, se ele aceitar |
| `formatacao_planilha(ws)` | Ajusta a largura das colunas da aba |
| `conversao(arquivo_excel, df)` | Grava o DataFrame no Excel já formatado, em uma única gravação |
| `main()` | Localiza os CSVs e controla o fluxo de conversão |
