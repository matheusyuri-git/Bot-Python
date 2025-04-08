from botcity.core import DesktopBot
import Planilha
import time
from datetime import datetime
import os
import pandas as pd

# Caminho do arquivo da planilha
# planilha = os.getcwd() + r'\Rendimentos Pagos_Creditados (envio).xlsx'
planilha = r'C:\Users\fernando.jfernandes\Desktop\Python\Planilha\Rendimentos Pagos_Creditados (envio).xlsx'
print(planilha)
#df = pd.read_excel(r'C:\Users\fernando.jfernandes\Desktop\Python\Planilha\Rendimentos Pagos_Creditados (envio).xlsx')
#cnpj = df['CNPJ do beneficiário']
#print(cnpj)


class Bot(DesktopBot):
    def action(self, execution=None):
        CNPJ = Planilha.buscaCNPJ()

        # Abre o navegador para envio de NF
        self.browse("https://cav.receita.fazenda.gov.br/autenticacao/login")
        self.wait(5)
        
        # Alteração
        # Searching for element 'LOGIN_GOV '
        if not self.find("LOGIN_GOV", matching=0.97, waiting_time=10000):
            self.not_found("LOGIN_GOV")
        self.click()
        self.wait(5)
        
        # Searching for element 'CERTIFICADO_DIGITAL '
        if not self.find("CERTIFICADO_DIGITAL", matching=0.97, waiting_time=10000):
            self.not_found("CERTIFICADO_DIGITAL")
        self.click()
         
        # Searching for element 'CERTIFICADO_GOINFRA '
        if not self.find("CERTIFICADO_GOINFRA", matching=0.97, waiting_time=10000):
            self.not_found("CERTIFICADO_GOINFRA")
        self.click()     
        
        # Searching for element 'VALIDA_CERTIFICADO_GOINFRA '
        if not self.find("VALIDA_CERTIFICADO_GOINFRA", matching=0.97, waiting_time=10000):
            self.not_found("VALIDA_CERTIFICADO_GOINFRA")
        self.click()
        
        # LOCALIZA SERVICO
        # Searching for element 'LOCALIZAR_SERVICO '
        if not self.find("LOCALIZAR_SERVICO", matching=0.97, waiting_time=10000):
            self.not_found("LOCALIZAR_SERVICO")
        self.click()
        self.paste('reinf')
     
        # Searching for element 'ACESSO_REINF '
        if not self.find("ACESSO_REINF", matching=0.97, waiting_time=10000):
            self.not_found("ACESSO_REINF")
        self.click()
       
       # Searching for element 'RendimentosPAGOSeCreditados '
        if not self.find("RendimentosPAGOSeCreditados", matching=0.97, waiting_time=20000):
           self.not_found("RendimentosPAGOSeCreditados")
        self.click()
        time.sleep(2)

        self.click_at(601, 440)
        
        # Searching for element 'Beneficiario_PJ '
        if not self.find("Beneficiario_PJ", matching=0.97, waiting_time=10000):
            self.not_found("Beneficiario_PJ")
        self.click()

        periodoEnvio = Planilha.buscaPeriodo()
        # --------------------------------------------------------
        df = Planilha.buscaDados()
        cnpj_anterior = None

        for index, row in df.iloc[0:].iterrows():
            cnpj = row['CNPJ do beneficiário']

            # Verifica se o valor da coluna 'CNPJ do beneficiário' é NaN
            if pd.isna(cnpj):
                print("Encontrou um valor NaN na coluna 'CNPJ do beneficiário'. Saindo do loop.")
                break

            # Faça a comparação com o CPF anterior
            if cnpj == cnpj_anterior:
                print('fazer fluxo de cnpj igual')
                print(f"Processando CNPJ anterior: {cnpj}")

        # Adicionar o Periodo
        if not self.find("PERIODO_APURACAO", matching=0.97, waiting_time=10000):
            self.not_found("PERIODO_APURACAO")
        self.click()        
        self.paste(periodoEnvio) # Período de Apuração
        self.tab()
        self.paste(CNPJ) # CNPJ do Estabelecimento
        self.tab()
        self.paste(cnpj) # cnpj do beneficiario
        self.tab()
        self.tab()
        self.enter()
        time.sleep(1)

        # Add Novo
        if not self.find("novoBene", matching=0.97, waiting_time=100000):
            self.not_found("novoBene")
            self.click()

            dig01gup = str(row['Grupo do rendimento'])[0]
            dig02gup = str(row['Grupo do rendimento'])[1]

            # Adicionar Natureza Rendimento
            dig01nat = str(row['Natureza do Rendimento'])[0]
            dig02nat = str(row['Natureza do Rendimento'])[1]
            dig03nat = str(row['Natureza do Rendimento'])[2]
            dig04nat = str(row['Natureza do Rendimento'])[3]
            dig05nat = str(row['Natureza do Rendimento'])[4]
            digito_total = dig01nat + dig02nat + dig03nat + dig04nat + dig05nat

            # Adicionar Grupo Rendimentos
            if not self.find("grupoRendimento", matching=0.97, waiting_time=100000):
                self.not_found("grupoRendimento")
            self.click()
            self.type_key(dig01gup)
            self.type_key(dig02gup)
            self.tab()
            self.type_key(dig01nat)
            self.type_key(dig02nat)
            self.type_key(dig03nat)
            self.type_key(dig04nat)
            self.type_key(dig05nat)
            self.tab()
            self.tab()
            self.enter()

            time.sleep(1)

                # Add Valores
            if not self.find("novoCred", matching=0.97, waiting_time=100000):
                self.not_found("novoCred")
            self.click()

            time.sleep(2)

            data_original = row['Data do fato gerador']
            data = datetime.strftime(data_original, '%d/%m/%Y')
            print(f'data fato: {data}')

            valorB = str(row['Valor bruto'])
            ir = str(row['Valor da base de retenção do IR'])
            irrf = str(row['Valor do Imposto de Renda IRRF'])

            if digito_total == '17013':
                    # Data
                    if not self.find("dataFato", matching=0.97, waiting_time=100000):
                        self.not_found("dataFato")
                    self.click()
                    self.paste(data)
                    self.tab()
                    self.paste(valorB)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.paste(ir)
                    self.tab()
                    self.paste(irrf)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.enter()
            else:
                    # Data
                    if not self.find("dataFato", matching=0.97, waiting_time=100000):
                        self.not_found("dataFato")
                    self.click()
                    self.paste(data)
                    self.tab()
                    self.paste(valorB)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.paste(ir)
                    self.tab()
                    self.paste(irrf)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.enter()

            time.sleep(0.5)

                # Desse tela
            if not self.find("redmentoNatu", matching=0.97, waiting_time=100000):
                    self.not_found("redmentoNatu")
            self.click()
            self.page_down()

                # Salva Rascunho
            if not self.find("rascunho", matching=0.97, waiting_time=100000):
                    self.not_found("rascunho")
            self.click()
            self.tab()
            self.tab()
            self.enter()

        else:
                # Acessa Rendimento
            if not self.find("rendimento", matching=0.97, waiting_time=1000000):
                    self.not_found("rendimento")
            self.move()

            time.sleep(2)
                # Incluir Pagamentos
            if not self.find("incluirPagamentos", matching=0.97, waiting_time=1000000):
                    self.not_found("incluirPagamentos")
            self.click()

            time.sleep(2)

                # Beneficiarios
            if not self.find("beneficiario", matching=0.97, waiting_time=1000000):
                    self.not_found("beneficiario")
            self.click()

            time.sleep(2)

            print('fazer fluxo de cnpj diferente')
            print(f"Processando CNPJ: {cnpj}")



                # Incluir NF
            if not self.find("incluirNF", matching=0.97, waiting_time=100000):
                    self.not_found("incluirNF")
            self.click()

            time.sleep(1)

            dig01gup = str(row['Grupo do rendimento'])[0]
            dig02gup = str(row['Grupo do rendimento'])[1]

                # Adicionar Natureza Rendimento
            dig01nat = str(row['Natureza do Rendimento'])[0]
            dig02nat = str(row['Natureza do Rendimento'])[1]
            dig03nat = str(row['Natureza do Rendimento'])[2]
            dig04nat = str(row['Natureza do Rendimento'])[3]
            dig05nat = str(row['Natureza do Rendimento'])[4]
            divito_total = dig01nat + dig02nat + dig03nat + dig04nat + dig05nat

                # Adicionar Grupo Rendimentos
            if not self.find("grupoRendimento", matching=0.97, waiting_time=100000):
                    self.not_found("grupoRendimento")
            self.click()
            self.type_key(dig01gup)
            self.type_key(dig02gup)
            self.tab()
            self.type_key(dig01nat)
            self.type_key(dig02nat)
            self.type_key(dig03nat)
            self.type_key(dig04nat)
            self.type_key(dig05nat)
            self.tab()
            self.tab()
            self.enter()

            time.sleep(1)

                # Add informacoes da NF
            if not self.find("credNovaInclusao", matching=0.97, waiting_time=100000):
                    self.not_found("credNovaInclusao")
            self.click()

            time.sleep(2)

            data_original = row['Data do fato gerador']
            data = datetime.strftime(data_original, '%d/%m/%Y')
            print(f'data fato: {data}')

            valorB = str(row['Valor bruto'])
            ir = str(row['Valor da base de retenção do IR'])
            irrf = str(row['Valor do Imposto de Renda IRRF'])

            if divito_total == '17013':
                    # Data
                    if not self.find("dataFato", matching=0.97, waiting_time=100000):
                        self.not_found("dataFato")
                    self.click()
                    self.paste(data)
                    self.tab()
                    self.paste(valorB)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.paste(ir)
                    self.tab()
                    self.paste(irrf)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.enter()
            else:
                    # Data
                    if not self.find("dataFato", matching=0.97, waiting_time=100000):
                        self.not_found("dataFato")
                    self.click()
                    self.paste(data)
                    self.tab()
                    self.paste(valorB)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.paste(ir)
                    self.tab()
                    self.paste(irrf)
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.tab()
                    self.enter()

            time.sleep(0.5)

            proximo_cnpj = df.iloc[index + 1]['CNPJ do beneficiário']

            if proximo_cnpj != cnpj:
                    # Salva Rascunho
                    if not self.find("rascunho", matching=0.97, waiting_time=100000):
                        self.not_found("rascunho")
                    self.click()
                    self.tab()
                    self.tab()
                    self.enter()

            cnpj_anterior = cnpj
            time.sleep(1)

    @staticmethod
    def not_found(label):
        print(f"Elemento não encontrado: {label}")


if __name__ == '__main__':
    Bot.main()
    # pyinstaller --collect-data palettable --onefile bot.py



