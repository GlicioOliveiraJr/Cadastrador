from botcity.core import DesktopBot
from telnetlib import STATUS
import gspread
from h11 import SEND_RESPONSE
from oauth2client.service_account import ServiceAccountCredentials
from pyparsing import col
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import pyautogui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyperclip
from PySimpleGUI import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook

class Bot(DesktopBot):
    def action(self, execution=None):

        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)
        navegador.get('https://app.omie.com.br/gestao/omiexperience-3zc3ae/')
        navegador.maximize_window()
        time.sleep(5)

        #Logando
        pyautogui.alert(' Quando ver a tela de login, pressione "OK" ')
        navegador.find_element(By.XPATH,'//*[@id="email"]').send_keys('junior.oliveira@omie.com.vc')
        time.sleep(1)
        navegador.find_element(By.XPATH,'//*[@id="email"]').send_keys(Keys.ENTER)
        time.sleep (3)
        navegador.find_element(By.XPATH, '//*[@id="current-password"]').send_keys('Glicio@03080308')
        navegador.find_element(By.XPATH, '//*[@id="current-password"]').send_keys(Keys.ENTER)
        pyautogui.alert("Coloque o autenticador e precione enter, depois clique em Ok.")
        pyautogui.alert('Quando abrir a tela de cadastro, pressione "OK" ')

        #Integrando com a planilha
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\glici\Meu Drive\Omie\Junior\Python\Cadastrador\credenciais.json', scope)
        gc = gspread.authorize(credentials)
        wks = gc.open_by_key('1nmj3ij21U0cSY5L1q76Oy0hBBG8z7aYT7hMiJjo1SoU')
        planilha = wks.get_worksheet(0)
        linha = 2

        #Inicio do loop ok
        while True:    
            dados = planilha.row_values(linha)

            #Verificando se o contato esta com "Cadastrado"
            while True:
                status = planilha.cell(linha,12)
                status = str(status)
                if status == "<Cell R2C12 'Cadastrado'>":
                    linha += 1
                    dados = planilha.row_values(linha)
                if status != "<Cell R2C12 'Cadastrado'>":
                    break
            
            #Colocando dados da planilha nas variaveis      
            sdr = dados[0]
            print(sdr)
            cnpj_da_contabilidade = dados[1]
            print(cnpj_da_contabilidade)
            nome_da_contabilidade = dados[2]
            print(nome_da_contabilidade)
            segmento = dados[3]
            print(segmento)
            cnpj_da_empresa = dados[4]
            print(cnpj_da_empresa)
            razao_social = dados[5]
            print(razao_social)
            regime_tributario = dados[6]
            print(regime_tributario)
            faturamento = dados[7]
            print(faturamento)
            contato_responsavel = dados[8]
            print(contato_responsavel)
            telefone = dados[9]
            print(telefone)
            obs = dados[10]
            print(obs)
            linha += 1

            #Cadastrando contato            
            navegador.find_element(By.XPATH, '//*[@id="dialogToolbar-50370"]/a[2]/div[3]').click()
            time.sleep(2)
            navegador.find_element(By.XPATH, '//*[@id="d50602c9"]').click()  
            navegador.find_element(By.XPATH,'//*[@id="d50602c9"]').send_keys(cnpj_da_empresa)
            time.sleep(2)
            navegador.find_element(By.XPATH, '//*[@id="dialogContent-50602"]/div/button').click()
            time.sleep(3)
            
            #O CNPJ ja foi cadastrado     
            try:
                if not self.find( "Ok", matching=0.97, waiting_time=10000):
                    self.not_found("Ok")
                if not self.find( "Ok", matching=0.97, waiting_time=10000):
                    self.not_found("Ok")
                self.click()    
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="dialog-50602"]/div[1]/button').click()
                time.sleep(3)
                navegador.find_element(By.XPATH, '//*[@id="dialog-50370"]/div[1]/button').click()
                time.sleep(3)
                planilha.update_acell(f'L{linha-1}', 'Cadastrado')
            except:
                pass
                
            #O CNPJ não esta cadastrado
            try:
                time.sleep(5) 
                navegador.find_element(By.XPATH, '//*[@id="d50874c3g"]/tbody/tr[1]').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="dialogToolbar-50370"]/a/div[3]').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="navbar-collapse-50370"]/ul/li[1]/a').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="dialogToolbar-50370"]/a[1]/div[3]').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="d50370c32"]/a/span[2]/div').click()
                time.sleep(5)
                navegador.find_element(By.XPATH,'//*[@id="d50369c6"]').send_keys(contato_responsavel)
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="navbar-collapse-50369"]/ul/li[3]/a').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="d50369c62"]').click()
                time.sleep(2)        
                navegador.find_element(By.XPATH,'//*[@id="d50369c62"]').send_keys(telefone)
                print(telefone)
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="navbar-collapse-50369"]/ul/li[1]/a').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="dialogToolbar-50369"]/a[1]/div[3]').click()
                time.sleep(5)

                #Incluindo Oportunidade
                navegador.find_element(By.XPATH, '//*[@id="d50369c32"]/a/span[2]/div').click()
                time.sleep(5)        
                navegador.find_element(By.XPATH,'//*[@id="d50377c12"]/span[1]/input').send_keys('Omie')
                time.sleep(2)        
                navegador.find_element(By.XPATH, '//*[@id="d50377c15"]/span[1]/input').click()
                time.sleep(1)
                navegador.find_element(By.XPATH,'//*[@id="d50377c15"]/span[1]/input').send_keys('Indicação Contador')
                time.sleep(2)
                pyautogui.press('enter')
                time.sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="navbar-collapse-50377"]/ul/li[5]/a').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="d50377c140"]').click()
                time.sleep(1)
                navegador.find_element(By.XPATH,'//*[@id="d50377c140"]').send_keys('Segmento: ')
                navegador.find_element(By.XPATH,'//*[@id="d50377c140"]').send_keys(segmento)
                time.sleep(1)
                pyautogui.press('enter')
                navegador.find_element(By.XPATH,'//*[@id="d50377c140"]').send_keys('Faturamento: ')
                navegador.find_element(By.XPATH,'//*[@id="d50377c140"]').send_keys(faturamento)
                pyautogui.press('enter')
                time.sleep(1)
                navegador.find_element(By.XPATH,'//*[@id="d50377c140"]').send_keys('Observação: ')
                navegador.find_element(By.XPATH,'//*[@id="d50377c140"]').send_keys(obs)  
                time.sleep(1)      
                navegador.find_element(By.XPATH, '//*[@id="navbar-collapse-50377"]/ul/li[7]/a').click()
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="d50377c156"]').click()        
                navegador.find_element(By.XPATH,'//*[@id="d50377c156"]').send_keys(nome_da_contabilidade)
                time.sleep(1)
                navegador.find_element(By.XPATH, '//*[@id="d50377c149"]').click() 
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="d50377c163"]').click()
                navegador.find_element(By.XPATH,'//*[@id="d50377c163"]').send_keys(sdr)  
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="d50377c149"]').click() 
                time.sleep(2)
                navegador.find_element(By.XPATH, '//*[@id="dialogToolbar-50377"]/a/div[3]').click()
                time.sleep(5)

                #Coloca "Ok" no status
                planilha.update_acell(f'L{linha-1}', 'Cadastrado')

                #Voltando para o inicio
                navegador.find_element(By.XPATH, '//*[@id="dialog-50377"]/div[1]/button').click()
                time.sleep(2)        
                navegador.find_element(By.XPATH, '//*[@id="dialog-50369"]/div[1]/button').click()
                time.sleep(2)
            except:
                pass
            
            def not_found(self, label):
                print(f"Element not found: {label}")

while True:
    if __name__ == '__main__':
        Bot.main()
