import time
import pandas as pd
import pdfkit
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from winotify import Notification, audio


class ExecutandoChrome:
    def __init__(self, login_mercante, senha_mercante, url_da_marinha, pasta_termos_pdf):

        self.URL_MARINHA = url_da_marinha
        self.LOGIN_MERCANTE = login_mercante
        self.SENHA_MERCANTE = senha_mercante
        self.PASTA_TERMOS_PDF = pasta_termos_pdf

        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--disable-infobars")
        self.options.add_argument("--enable-print-browser")
        self.options.add_argument("--kiosk-printing")
        self.options.add_argument("--disable-extensions")
        self.options.add_argument("--disable-popup-blocking")

        driver_manager = ChromeDriverManager()
        servico = ChromeService(driver_manager.install())

        self.chrome = webdriver.Chrome(options=self.options)
        self.wait = WebDriverWait(self.chrome, 10)

        self.URL_MARINHA = url_da_marinha

    def realizar_entrada_login(self):
        try:
            TELA_ERRO_PRIVACIDADE = self.chrome.find_element(
                By.ID, "details-button")
            TELA_ERRO_PRIVACIDADE.click()

            LINK_LOGIN = self.chrome.find_element(By.ID, "proceed-link")
            LINK_LOGIN.click()
        except:
            print("seguindo sem a tela do erro de privacidade")

        finally:
            time.sleep(3)
            elemento_login_mercante = self.chrome.find_element(
                By.NAME, "cpfTemp")

            elemento_senha_mercante = self.chrome.find_element(
                By.NAME, "senhaTemp")
            elemento_login_mercante.click()
            elemento_login_mercante.send_keys(self.LOGIN_MERCANTE)
            elemento_senha_mercante.click()
            elemento_senha_mercante.send_keys(self.SENHA_MERCANTE)

            elemento_avancar_botao = self.chrome.find_element(
                By.XPATH, "/html/body/form/div[1]/table[1]/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/table/tbody/tr[7]/td[2]/div")
            elemento_avancar_botao.click()

    def selecionando_termo_mercante(self):
        self.chrome.switch_to.frame("header")
        element = self.chrome.find_element(
            By.CSS_SELECTOR, '[href*="executaAcao(\'menu|MENU-CONHECIMENTO\')"]')
        element.click()

        select_element = self.chrome.find_element(By.CLASS_NAME, "cb2")
        option_to_select = select_element.find_element(
            By.XPATH, '//option[@value="menu|MENU-CONHECIMENTO-BL"]')
        option_to_select.click()

        select_element = self.chrome.find_element(By.CLASS_NAME, "cb2")
        option_to_select = select_element.find_element(
            By.XPATH, '//option[@value="KCE-LIB"]')
        option_to_select.click()

        self.chrome.switch_to.default_content()

    def botao_enviar_termo(self, ce_mercante):
        self.chrome.switch_to.frame("Principal")
        input_element = self.wait.until(
            EC.element_to_be_clickable((By.NAME, "nrCeMercante")))
        input_element.send_keys(ce_mercante)

        input_element = self.chrome.find_element(
            By.CSS_SELECTOR, 'input.bt1[name="Enviar"]')
        input_element.click()

    def processar_ce_mercante(self, df, ce_mercante, local_termos_pdf):
        referencia_despachante = df.loc[df['Numero CE'] ==
                                  ce_mercante, 'Sua Referencia'].iloc[0]
        NOME_DO_ARQUIVO = f"{referencia_despachante}_TermoMarinhaMercante.pdf".upper()

        PASTA_TERMOS_MARINHA = local_termos_pdf
        time.sleep(2)

        caminho_completo_arquivo = PASTA_TERMOS_MARINHA + "\\" + NOME_DO_ARQUIVO

        html_da_pagina = self.chrome.page_source

        # Construindo o caminho para o arquivo 'wkhtmltopdf.exe' na subpasta 'wkhtmltox/bin'
        script_directory = os.path.dirname(os.path.abspath(__file__))
        wkhtmltopdf_path = os.path.join(
            script_directory, 'wkhtmltox', 'bin', 'wkhtmltopdf.exe')

        # Configuração do pdfkit com o caminho dinâmico para wkhtmltopdf
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

        codificacao = 'utf-8'
        caminho_arquivo_html = os.path.join(
            script_directory, 'data', 'termomarinha.html')

        with open(caminho_arquivo_html, 'w', encoding='utf-8') as arquivo_html:
            arquivo_html.write(html_da_pagina)

        try:
            pdfkit.from_file(caminho_arquivo_html, caminho_completo_arquivo,
                             configuration=config, options={'encoding': codificacao})
        except Exception as e:
            print(f"erro no {e}")

        time.sleep(1)

        elemento_voltar = self.chrome.find_element(
            By.XPATH, '//a[text()="Voltar"]')
        elemento_voltar.click()

    def configurando_relatorio(self, relatorio_afrmm_xlsx):
        df = pd.read_excel(relatorio_afrmm_xlsx)

        df['Numero CE'] = df['Numero CE'].astype(str)

        # As cargas de Ceará o CE mercante começa com 042, essa função garante que o "0" fique na frente do 42.
        def adicionar_zero(numero_ce):
            if numero_ce.startswith('42'):
                return '0' + numero_ce
            return numero_ce

        df['Numero CE'] = df['Numero CE'].apply(adicionar_zero)

        # verificação se tem processo apto para puxar os termos
        self.verificar_e_encerrar_se_dataframe_vazio(df)

        return df

    def verificar_e_encerrar_se_dataframe_vazio(self, df):
        if df.empty:
            df_vazio_mensagem = Notification(app_id="Relatório AFRMM", title="Sem processos para puxar o termo!",
                                             msg="Não há novos processos para baixar o termo da AFRMM",
                                             duration="long")
            df_vazio_mensagem.set_audio(audio.SMS, loop=False)
            df_vazio_mensagem.show()
            sys.exit()

    def main(self, pasta_termos_pdf, relatorio_afrmm_xlsx):

        df = self.configurando_relatorio(relatorio_afrmm_xlsx)

        # inicio do contador do tempo
        tempo_inicio = time.time()

        self.chrome.maximize_window()
        self.chrome.get(self.URL_MARINHA)
        self.realizar_entrada_login()
        self.selecionando_termo_mercante()
        time.sleep(0.2)

        # inicio de contador de termos extraídos
        contador_linhas = 0

        # inicio da iteração para percorrer todos os processos do df_resultado
        for ce_mercante in df['Numero CE']:
            self.botao_enviar_termo(ce_mercante)
            self.processar_ce_mercante(df, ce_mercante, pasta_termos_pdf)
            self.chrome.switch_to.default_content()
            time.sleep(1)

            # acresce +1 na quantidade de termos extraídos a cada iteração
            contador_linhas += 1

        # finaliza o chrome e fornece o tempo de execução da tarefa
        self.chrome.quit()

        # fim do contador do tempo
        tempo_total = time.time() - tempo_inicio

        # Função que registra o log da execução das tarefas em excel
        self.log(contador_linhas, tempo_total)

        # carregando caminho absoluto do icone de forma dinamica para exibir na notificação de finalização
        script_directory = os.path.dirname(os.path.abspath(__file__))
        icone = os.path.join(script_directory, 'images', 'mercante-logo.png')

        AUTOMACAO_MERCANTE = Notification(app_id="TERMOS AFRMM", title="Termos da marinha extraídos!",
                                          msg=f"Foram extraídos {contador_linhas} termos de liberação AFRMM.\nTempo de execução: {
                                              tempo_total: .2f} segundos.",
                                          duration="long", icon=icone)

        AUTOMACAO_MERCANTE.set_audio(audio.SMS, loop=False)
        AUTOMACAO_MERCANTE.show()
