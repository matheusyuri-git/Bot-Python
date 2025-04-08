from botcity.core import DesktopBot
import Planilha
import time
from datetime import datetime
import os
import pandas as pd

# Caminho do arquivo da planilha
planilha = r'C:\Python\Planilha\2025 - EFD-REINF - Rendimentos Pagos_Creditados (Trivalle).xlsx'
print(planilha)

class Bot(DesktopBot):
    def action(self, execution=None):
        CNPJ = Planilha.buscaCNPJ()  # Estabelecimento

        # Abre o navegador para envio de NF
        self.browse("https://cav.receita.fazenda.gov.br/autenticacao/login")
        
        # Localiza Serviço
        if not self.find("LOGIN_GOV", matching=0.97, waiting_time=10000):
            self.not_found("LOGIN_GOV")
        self.click()

        if not self.find("CERTIFICADO_DIGITAL", matching=0.97, waiting_time=10000):
            self.not_found("CERTIFICADO_DIGITAL")
        self.click()

        if not self.find("CERTIFICADO_GOINFRA", matching=0.97, waiting_time=10000):
            self.not_found("CERTIFICADO_GOINFRA")
        self.click()

        if not self.find("VALIDA_CERTIFICADO_GOINFRA", matching=0.97, waiting_time=10000):
            self.not_found("VALIDA_CERTIFICADO_GOINFRA")
        self.click()
# -------------------------------------------------------------------------------------------------------------------------------
        # LOCALIZA SERVIÇO
        if not self.find("LOCALIZAR_SERVICO", matching=0.90, waiting_time=10000):
            self.not_found("LOCALIZAR_SERVICO")
        self.click()
        self.paste('reinf')
        
        # Acessando o menu REINF EFD
        if not self.find("ACESSO_REINF", matching=0.97, waiting_time=10000):
            self.not_found("ACESSO_REINF")
        self.click()

        # Acessando a opção rendimentos pagos e creditados 
        if not self.find("RendimentosPAGOSeCreditados", matching=0.97, waiting_time=20000):
            self.not_found("RendimentosPAGOSeCreditados")
        self.click()

        # Clicando no icone Incluir Pagamento Credito
        self.click_at(601, 440)
        
        # Procurando o icone beneficiario PJ 
        if not self.find("Beneficiario_PJ", matching=0.97, waiting_time=10000):
            self.not_found("Beneficiario_PJ")
        self.click()

        periodoEnvio = Planilha.buscaPeriodo()
        df = Planilha.buscaDados()
        cnpj_anterior = None

        for index, row in df.iloc[0:].iterrows():
            cnpj = row['CNPJ do beneficiário']
            print(cnpj)

            if pd.isna(cnpj):
                print("Encontrou um valor NaN na coluna 'CNPJ do beneficiário'. Saindo do loop.")
                break

            if cnpj == cnpj_anterior:
                print('fazer fluxo de cnpj igual')
                print(f"Processando CNPJ anterior: {cnpj}")

            # Adicionar o Periodo de apuração e CNPJ do Estabelecimento
            if not self.find("PERIODO_APURACAO", matching=0.97, waiting_time=10000):
                self.not_found("PERIODO_APURACAO")
            self.click()
            self.paste(periodoEnvio)  # Período de Apuração
            self.tab()
            self.paste(CNPJ)  # CNPJ do Estabelecimento
            self.tab()
            self.paste(cnpj)  # CNPJ do beneficiário
            self.tab()
            self.tab()
            self.enter()
            time.sleep(1)
            
            # Incluir um novo beneficiário
            if not self.find("Novo_Beneficiario", matching=0.97, waiting_time=10000):
                self.not_found("Novo_Beneficiario")
            self.click()
            
            # Informando natureza do rendimento (Grupo de Rendimento)               
            if not self.find("Grupo_de_Rendimento", matching=0.97, waiting_time=10000):
                self.not_found("Grupo_de_Rendimento")
            self.click()
            
            # Searching for element 'Informando_o_Grupo_17 '
            if not self.find("Informando o Grupo 17", matching=0.97, waiting_time=10000):
                self.not_found("Informando o Grupo 17")
            self.click()
            
            # Searching for element 'Natureza_Rendimento '
            if not self.find("Natureza_Rendimento", matching=0.97, waiting_time=10000):
                self.not_found("Natureza_Rendimento")
            self.click_relative(28, 34)
            
            # Searching for element 'Informando_17013 '
            if not self.find("Informando_17013", matching=0.97, waiting_time=10000):
                self.not_found("Informando_17013")
            self.click()
            
            # Searching for element 'Salvar_Natureza '
            if not self.find("Salvar_Natureza", matching=0.97, waiting_time=10000):
                self.not_found("Salvar_Natureza")
            self.click()
            
            # Searching for element 'Novo_Detalhamento_Pagamento '
            if not self.find("Novo_Detalhamento_Pagamento", matching=0.97, waiting_time=10000):
                self.not_found("Novo_Detalhamento_Pagamento")
            self.click()
            time.sleep(5)
            
            #Tratando a data da coluna "Data do fato Gerador"
            data_original = row['Data do fato gerador']
            data_formatada = datetime.strptime(str(data_original), '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
            print(data_formatada)
        
            # Searching for element 'data_fato_gerador '
            if not self.find("data_fato_gerador", matching=0.97, waiting_time=10000):
                self.not_found("data_fato_gerador")
            time.sleep(3)
            self.click()
            self.paste(data_formatada)
            
            # Tratando o valor da coluna "Valor Bruto"
            Valor_bruto = str(row['Valor bruto'])
            Valor_Retenção_IR = str(row['Valor da base de retenção do IR'])
            Valor_Imposto_Renda_IRRF = str(row['Valor do Imposto de Renda IRRF']) # Valor do Imposto de Renda IRRF

            # Searching for element 'Valor_Bruto '
            if not self.find("Valor_Bruto", matching=0.97, waiting_time=10000):
              self.not_found("Valor_Bruto")
            self.click_relative(21, 32)
            self.paste(Valor_bruto)
            
            # Searching for element 'Retencao_IR '
            if not self.find("Retencao_IR", matching=0.97, waiting_time=10000):
                self.not_found("Retencao_IR")
            self.click_relative(27, 29)
            self.paste(Valor_Retenção_IR)
            self.tab() 
            self.paste(Valor_Imposto_Renda_IRRF)
            
            # Searching for element 'Salvar_Detalhamento '
            if not self.find("Salvar_Detalhamento", matching=0.97, waiting_time=10000):
                self.not_found("Salvar_Detalhamento")
            self.click()
            self.page_down()

            # Searching for element 'Salvar_Rascunho '
            if not self.find("Salvar_Rascunho", matching=0.97, waiting_time=10000):
                self.not_found("Salvar_Rascunho")
            self.click()

            # Chama a função para continuar o fluxo após salvar o rascunho
            self.continuar_fluxo_apos_rascunho()

    def continuar_fluxo_apos_rascunho(self):
        # Acessando a opção rendimentos pagos e creditados novamente
        if not self.find("RendimentosPAGOSeCreditados", matching=0.97, waiting_time=20000):
            self.not_found("RendimentosPAGOSeCreditados")
        self.click()
        
        # Searching for element 'Incluir_Pagamento_Credito '
        if not self.find("Incluir_Pagamento_Credito", matching=0.97, waiting_time=10000):
            self.not_found("Incluir_Pagamento_Credito")
        self.click()
        
        # Procurando o ícone beneficiário PJ novamente
        if not self.find("Beneficiario_PJ", matching=0.97, waiting_time=10000):
            self.not_found("Beneficiario_PJ")
        self.click()

    @staticmethod
    def not_found(label):
        print(f"Elemento não encontrado: {label}")

if __name__ == '__main__':
    Bot.main()
