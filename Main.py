import pandas as pd
import win32com.client

from datetime import datetime
import os
import time
import re

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By

class Automacao_Moedas():

    def __init__(self):
        try:
            app_excel = win32com.client.Dispatch('Excel.Application')
            for arquivo_aberto in app_excel.Workbooks:
                arquivo_aberto.Close(SaveChanges=False)
            app_excel.Quit()
            del app_excel
            print('Instancias Fechadas')
        except Exception as erro:
            print(erro)

    def Configuracao_Driver(self):
        #Configuracao do Driver do Navegador
        configuracao_driver = Options()

        configuracao_driver.add_experimental_option('detach', True)
        configuracao_driver.add_argument('--headless')

        driver_chrome = Service(
            ChromeDriverManager().install()
        )
        driver_path = ChromeDriverManager().install()
        print("Caminho do Driver - ", driver_path)

        return configuracao_driver, driver_chrome

    def Receber_Dados_Moedas(self, configuracao_chrome, driver):
        navegador = webdriver.Chrome(
            service=driver, options=configuracao_chrome
        )
        navegador.maximize_window()
        navegador.get('https://cuex.com/pt/usd-brl')
        time.sleep(5)

        #Conversao Dolar
        texto_conversao = navegador.find_element(
            By.XPATH, '//*[@id="section-content"]/cx-converter/cx-calculator/div[2]/div[2]/div[1]/div/cx-calculator-exchange-result/div/div[2]/div/button'
        )
        texto_conversao = texto_conversao.text
        dolar = texto_conversao.replace('BRL', '').replace(',', '.')
        dolar = float(dolar)
        dolar = f'{dolar: ,.2f}'
        print(f'Dolar - {dolar}')

        # Euro
        navegador.find_element(
            By.XPATH, '//*[@id="section-content"]/cx-converter/cx-calculator/div[1]/div/cx-calculator-form/form/div/div[2]/div/cx-picker/div/div/div[1]'
        ).click()

        navegador.find_element(
            By.XPATH, '//*[@id="panelOptions"]/div[2]/div[1]/div[2]/div/div[2]/div'
        ).click()

        texto_conversao = navegador.find_element(
            By.XPATH,
            '//*[@id="section-content"]/cx-converter/cx-calculator/div[2]/div[2]/div[1]/div/cx-calculator-exchange-result/div/div[2]/div/button'
        )
        texto_conversao = texto_conversao.text
        euro = texto_conversao.replace('BRL', '').replace(',', '.')
        euro = float(euro)
        euro = f'{euro: ,.2f}'
        print(f'Euro - {euro}')

        # Iene
        navegador.find_element(
            By.XPATH,
            '//*[@id="section-content"]/cx-converter/cx-calculator/div[1]/div/cx-calculator-form/form/div/div[2]/div/cx-picker/div/div/div[1]'
        ).click()
        navegador.find_element(
            By.XPATH,
            '//*[@id="panelOptions"]/div[2]/div[1]/div[2]/div/div[3]'
        ).click()
        texto_conversao = navegador.find_element(
            By.XPATH,
            '//*[@id="section-content"]/cx-converter/cx-calculator/div[2]/div[2]/div[1]/div/cx-calculator-exchange-result/div/div[2]/div/button'
        )
        texto_conversao = texto_conversao.text
        iene = texto_conversao.replace('BRL', '').replace(',', '.')
        iene = float(iene)
        iene = f'{iene: ,.2f}'
        print(f'Iene - {iene}')

        #Peso
        navegador.find_element(
            By.XPATH,
            '//*[@id="section-content"]/cx-converter/cx-calculator/div[1]/div/cx-calculator-form/form/div/div[2]/div/cx-picker/div/div/div[1]'
        ).click()
        navegador.find_element(
            By.XPATH,
            '//*[@id="panelOptions"]/div[2]/div[1]/div[2]/div/div[5]'
        ).click()
        texto_conversao = navegador.find_element(
            By.XPATH,
            '//*[@id="section-content"]/cx-converter/cx-calculator/div[2]/div[2]/div[1]/div/cx-calculator-exchange-result/div/div[2]/div/button'
        )
        texto_conversao = texto_conversao.text
        peso = texto_conversao.replace('BRL', '').replace(',', '.')
        peso = float(peso)
        peso = f'{peso: ,.2f}'
        print(f'Peso - {peso}')

        navegador.close()

        #Devolver os Dados Recebidos
        moedas_dic = {
            'Dolar': dolar, 'Euro': euro, 'Iene': iene, 'Peso': peso
        }
        return moedas_dic

    def Salvar_Dados_Moedas(self, dolar, euro, iene, peso):
        df = pd.read_excel(
            'C:/Users/Gustavo Roldam/Desktop/Python/MiniCursoPython/Moedas.xlsx'
        )
        print(df)

        proxima_linha = len(df)

        data_atual = datetime.now()
        data_atual = data_atual.strftime('%d-%m-%Y')

        df.loc[proxima_linha, 'Data'] = data_atual
        df.loc[proxima_linha, 'Dolar'] = dolar
        df.loc[proxima_linha, 'Euro'] = euro
        df.loc[proxima_linha, 'Iene'] = iene
        df.loc[proxima_linha, 'Peso'] = peso

        df.to_excel(
            'C:/Users/Gustavo Roldam/Desktop/Python/MiniCursoPython/Moedas.xlsx',
            index=False
        )
        self.Abrir_Excel()

    def Abrir_Excel(self):
        time.sleep(2)

        app_excel = win32com.client.Dispatch('Excel.Application')
        app_excel.Visible = True
        planilha = app_excel.Workbooks.Open(
            'C:/Users/Gustavo Roldam/Desktop/Python/MiniCursoPython/Moedas.xlsx'
        )


obj_automacao = Automacao_Moedas()
configuracao, driver = obj_automacao.Configuracao_Driver()
dicionario_moedas = obj_automacao.Receber_Dados_Moedas(configuracao, driver)
obj_automacao.Salvar_Dados_Moedas(
    dicionario_moedas['Dolar'],
    dicionario_moedas['Euro'],
    dicionario_moedas['Iene'],
    dicionario_moedas['Peso']
)
