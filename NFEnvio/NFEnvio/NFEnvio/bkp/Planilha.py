import pandas as pd
import os


# Caminho do arquivo da planilha
# planilha = os.getcwd() + r'\Rendimentos Pagos_Creditados (envio).xlsx'
planilha = r'C:\Automacao\Rendimentos Pagos_Creditados (envio).xlsx'


def buscaPeriodo():
    try:
        print(f'caminho da planilha: {planilha}')
        # Ler o arquivo Excel
        excel_file = pd.ExcelFile(planilha)

        # Obter a lista de abas disponíveis
        abas = excel_file.sheet_names

        # Selecionar a última aba
        ultima_aba = abas[-1]

        # Ler a planilha da última aba para um DataFrame
        data_frame = pd.read_excel(planilha, sheet_name=ultima_aba)

        # Agora você pode usar 'data_frame' para trabalhar com os dados
        print(data_frame.head())  # Exibe as primeiras linhas do DataFrame

        # Obter o valor da linha 2 da coluna B e atribuir a 'periodoEnvio'
        periodoEnvio = data_frame.iloc[0, 1]  # Assumindo que B representa a coluna 1

        print(f"Valor da linha 2 da coluna B na aba {ultima_aba}: {periodoEnvio}")

        return periodoEnvio

    except FileNotFoundError:
        print(f"Arquivo '{planilha}' não encontrado. Verifique o caminho e o nome do arquivo.")

    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")


def buscaCNPJ():
    try:
        # Ler o arquivo Excel
        excel_file = pd.ExcelFile(planilha)

        # Obter a lista de abas disponíveis
        abas = excel_file.sheet_names

        # Selecionar a última aba
        ultima_aba = abas[-1]

        # Ler a planilha da última aba para um DataFrame
        data_frame = pd.read_excel(planilha, sheet_name=ultima_aba)

        # Agora você pode usar 'data_frame' para trabalhar com os dados
        print(data_frame.head())  # Exibe as primeiras linhas do DataFrame

        # Obter o valor da linha 2 da coluna B e atribuir a 'periodoEnvio'
        cnpjUser = data_frame.iloc[1, 1]  # Assumindo que B representa a coluna 1

        print(f"Valor da linha 3 da coluna B na aba {ultima_aba}: {cnpjUser}")

        return cnpjUser

    except FileNotFoundError:
        print(f"Arquivo '{planilha}' não encontrado. Verifique o caminho e o nome do arquivo.")

    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")


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
        # Ler o arquivo Excel
        excel_file = pd.ExcelFile(planilha)

        # Obter a lista de abas disponíveis
        abas = excel_file.sheet_names

        # Selecionar a última aba
        ultima_aba = abas[-1]

        # Ler a planilha da última aba para um DataFrame
        data_frame = pd.read_excel(
            planilha, sheet_name=ultima_aba,
            header=4,
            dtype={'CNPJ do beneficiário': str,
                   'Valor bruto': float,
                   'Valor da base de retenção do IR': float,
                   'Valor do Imposto de Renda IRRF': float}
        )

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

    except FileNotFoundError:
        print(f"Arquivo '{planilha}' não encontrado. Verifique o caminho e o nome do arquivo.")

    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
