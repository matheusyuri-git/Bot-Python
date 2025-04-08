import pandas as pd

# Caminho do arquivo da planilha
Planilha = r'C:\Users\fernando.jfernandes\Desktop\Python\Planilha\Rendimentos Pagos_Creditados (envio).xlsx'

def buscaPeriodo():
    # Carregar a planilha no DataFrame
    df = pd.read_excel(Planilha, sheet_name= -1)

    # Função para buscar o valor do período
    Competencia = df.iloc[0, 1]
    return Competencia

def buscaCNPJ():
     # Carregar a planilha no DataFrame
    df = pd.read_excel(Planilha, sheet_name= -1)
    return df

    # Acessar o dado da segunda linha e segunda coluna (índices 1 e 1) = CNPJ Estabelecimento
    CNPJ = df.iloc[1, 1]
    return CNPJ

resultado = buscaPeriodo()
print(f"competência: {resultado}")

resultado = buscaCNPJ()
print(f"CNPJ: {resultado}")

    