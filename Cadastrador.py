import gspread
from h11 import SEND_RESPONSE
from oauth2client.service_account import ServiceAccountCredentials
from pyparsing import col

scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
gc = gspread.authorize(credentials)
wks = gc.open_by_key('1nmj3ij21U0cSY5L1q76Oy0hBBG8z7aYT7hMiJjo1SoU')
planilha = wks.get_worksheet(0)



while True:
    linha = 2

    dados = planilha.row_values(linha)
    #print(dados)

    if dados[10] != 'Ok':
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
        
    linha = linha+1
    if linha > 3:
        break



   