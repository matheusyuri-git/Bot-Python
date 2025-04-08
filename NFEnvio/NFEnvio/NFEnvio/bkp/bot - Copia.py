from botcity.core import DesktopBot
import Planilha
import time
from datetime import datetime
import os
import pandas as pd


# Caminho do arquivo da planilha
# planilha = os.getcwd() + r'\Rendimentos Pagos_Creditados (envio).xlsx'
planilha = r'C:\Automacao\Rendimentos Pagos_Creditados (envio).xlsx'


class Bot(DesktopBot):
    def action(self, execution=None):
        CNPJ = Planilha.buscaCNPJ()

        # Abre o navegador para envio de NF
        self.browse("https://cav.receita.fazenda.gov.br/autenticacao/login")

        time.sleep(10)

        # Entrar na pagina de login
        if not self.find( "entrarGov", matching=0.97, waiting_time=5):
            self.not_found("entrarGov")
        self.click()

        time.sleep(5)

        # Ativa o Certificado Digital
        if not self.find( "certificadoDigital", matching=0.97, waiting_time=5):
            self.not_found("certificadoDigital")
        self.click()

        time.sleep(5)

        # Valida o certificado
        if not self.find( "certificadoOK", matching=0.97, waiting_time=5):
            self.not_found("certificadoOK")
        self.click()

        time.sleep(5)

     # Troca perfil de envio
        if not self.find( "trocaPerfil", matching=0.90, waiting_time=2):
            self.not_found("trocaPerfil")
        self.click()

        time.sleep(2)


        time.sleep(2)

        # Add CNPJ
        if not self.find( "addCnpj", matching=0.97, waiting_time=20):
            self.not_found("addCnpj")
        self.click()
        self.paste(CNPJ)
        self.tab()
        self.enter()

        time.sleep(5)

        # Localiza Servico
        if not self.find( "localizaServico", matching=0.97, waiting_time=1000000):
            self.not_found("localizaServico")
        self.click()
        self.paste('reinf')

        time.sleep(2)

        # Acessa o Reinf
        if not self.find( "acessaReinf", matching=0.97, waiting_time=1000000):
            self.not_found("acessaReinf")
        self.click()

        time.sleep(2)

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

                # Add Novo
                if not self.find( "novoBene", matching=0.97, waiting_time=100000):
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
                if not self.find( "grupoRendimento", matching=0.97, waiting_time=100000):
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
                if not self.find( "novoCred", matching=0.97, waiting_time=100000):
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
                    if not self.find( "dataFato", matching=0.97, waiting_time=100000):
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
                    if not self.find( "dataFato", matching=0.97, waiting_time=100000):
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
                if not self.find( "redmentoNatu", matching=0.97, waiting_time=100000):
                    self.not_found("redmentoNatu")
                self.click()
                self.page_down()

                # Salva Rascunho
                if not self.find( "rascunho", matching=0.97, waiting_time=100000):
                    self.not_found("rascunho")
                self.click()
                self.tab()
                self.tab()
                self.enter()

            else:
                # Acessa Rendimento
                if not self.find( "rendimento", matching=0.97, waiting_time=1000000):
                    self.not_found("rendimento")
                self.move()

                time.sleep(2)

                # Incluir Pagamentos
                if not self.find( "incluirPagamentos", matching=0.97, waiting_time=1000000):
                    self.not_found("incluirPagamentos")
                self.click()

                time.sleep(2)

                # Beneficiarios
                if not self.find( "beneficiario", matching=0.97, waiting_time=1000000):
                    self.not_found("beneficiario")
                self.click()

                time.sleep(2)

                print('fazer fluxo de cnpj diferente')
                print(f"Processando CNPJ: {cnpj}")

                # Adicionar o Periodo
                if not self.find("periodoEnvio", matching=0.97, waiting_time=100000):
                    self.not_found("periodoEnvio")
                self.click()
                self.paste(periodoEnvio)
                self.tab()
                self.paste(CNPJ)
                self.tab()
                self.paste(cnpj)
                self.tab()
                self.tab()
                self.tab()
                self.enter()

                time.sleep(1)

                # Incluir NF
                if not self.find( "incluirNF", matching=0.97, waiting_time=100000):
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
                if not self.find( "grupoRendimento", matching=0.97, waiting_time=100000):
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
                if not self.find( "credNovaInclusao", matching=0.97, waiting_time=100000):
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
                    if not self.find( "dataFato", matching=0.97, waiting_time=100000):
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
                    if not self.find( "dataFato", matching=0.97, waiting_time=100000):
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
                    if not self.find( "rascunho", matching=0.97, waiting_time=100000):
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

