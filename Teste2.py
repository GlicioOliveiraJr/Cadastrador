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


   
#Integrando com a planilha
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\glici\Meu Drive\Omie\Junior\Python\Cadastrador\credenciais.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open_by_key('1nmj3ij21U0cSY5L1q76Oy0hBBG8z7aYT7hMiJjo1SoU')
planilha = wks.get_worksheet(0)
linha = 2

#Inicio do loop        
#while True:    
dados = planilha.row_values(linha)

#Verificando se o contato esta com "Cadastrado na planilha"
#while True:
status = planilha.cell(linha,14)
print(status)
status = str(status)
if status != "<Cell R2C14 None>":
    #linha += 1
    print('TEM')
    dados = planilha.row_values(linha)
if status == "<Cell R2C14 None>":# or status != "<Cell R2C14 'Este CNPJ já havia sido cadastrado'>":
    print('NAO TEM')
    break
#<Cell R2C14 'Este CNPJ já havia sido cadastrado'>