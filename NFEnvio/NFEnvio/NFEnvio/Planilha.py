import pandas as pd
import os

# Caminho do arquivo da planilha
planilha = r'C:\Python\Planilha\2025 - EFD-REINF - Rendimentos Pagos_Creditados (Trivalle).xlsx'

# Função auxiliar para carregar o DataFrame da última aba
def carregar_ultima_aba():
    try:
        print(f'caminho da planilha: {planilha}')
        # Ler o arquivo Excel
        excel_file = pd.ExcelFile(planilha)

        # Obter a lista de abas disponíveis
        abas = excel_file.sheet_names

        # Selecionar a última aba
        ultima_aba = abas[-1]

        # Ler a planilha da última aba para um DataFrame
        data_frame = pd.read_excel(planilha, sheet_name = ultima_aba)

        return data_frame, ultima_aba
    except FileNotFoundError:
        print(f"Arquivo '{planilha}' não encontrado. Verifique o caminho e o nome do arquivo.")
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")

def buscaPeriodo():
    try:
        data_frame, ultima_aba = carregar_ultima_aba()

        if data_frame is not None:
            # Obter o valor da linha 2 da coluna B e atribuir a 'periodoEnvio'
            periodoEnvio = data_frame.iloc[0, 1]  # Assumindo que B representa a coluna 1

            print(f"Período de apuração na aba {ultima_aba}: {periodoEnvio}")

            return periodoEnvio
    except Exception as e:
        print(f"Erro ao buscar período: {e}")

def buscaCNPJ(): # Estabelecimento
    try:
        data_frame, ultima_aba = carregar_ultima_aba()

        if data_frame is not None:
            # Obter o valor da linha 2 da coluna B e atribuir a 'cnpjUser'
            cnpjUser = data_frame.iloc[1, 1]  # Assumindo que B representa a coluna 1 (Estabelecimento)

            print(f"CNPJ Goinfra na aba {ultima_aba}: {cnpjUser}")

            return cnpjUser
    except Exception as e:
        print(f"Erro ao buscar CNPJ: {e}")

def formatar_valor(valor):
    if isinstance(valor, float):
        return "{:.2f}".format(valor)
    return valor

def converter_aliquota(valor):
    if isinstance(valor, float):
        return "{:.2%}".format(valor)
    return valor

def buscaDados():
    try:
        data_frame, ultima_aba = carregar_ultima_aba()

        if data_frame is not None:
            # Ler a planilha da última aba para um DataFrame com cabeçalho na linha 5
            data_frame = pd.read_excel(
                planilha, sheet_name=ultima_aba,
                header=4,
                dtype={'CNPJ do beneficiário': str,
                       'Valor bruto': float,
                       'Valor da base de retenção do IR': float,
                       'Valor do Imposto de Renda IRRF': float}
            )

        if data_frame is not None:
            # Exemplo de como contar as linhas
            num_linhas = data_frame.shape[0]  # Aqui, df.shape[0] retorna o número de linhas
            print(f'Quantidade de linhas: {num_linhas}')

             # Listar as colunas do DataFrame
            # print("Colunas do DataFrame:")
            # print(data_frame.columns)

            # Exemplo de contagem de linhas usando .shape
            # num_linhas = data_frame.shape[0]  # O primeiro valor da tupla é o número de linhas
            # print(f'Quantidade de linhas: {num_linhas}')


            # Aplicar a formatação para manter zeros à direita em todas as colunas, exceto 'Alíquota'
            for coluna in data_frame.columns:
                if coluna != 'Alíquota':
                    data_frame[coluna] = data_frame[coluna].apply(formatar_valor)
                else:
                    # Aplicar a formatação para 'Alíquota'
                    data_frame['Alíquota'] = data_frame['Alíquota'].apply(converter_aliquota)

            # Agora você pode usar 'data_frame' para trabalhar com os dados
            print(data_frame.head())  # Exibe as primeiras linhas do DataFrame

            return data_frame
    except Exception as e:
        print(f"Erro ao buscar dados: {e}")
if __name__ == "__main__":
    buscaPeriodo()
    buscaCNPJ()
    buscaDados()
