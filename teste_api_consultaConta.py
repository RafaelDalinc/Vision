import requests
import json
import openpyxl
import pyautogui
from time import sleep
from datetime import date
from datetime import datetime


#carregando a planilha
try:       
    workbook = openpyxl.load_workbook('planilha_contas.xlsx')
    pagina_contas = workbook['Planilha1']   
except FileNotFoundError:
    print('arquivo não encontrado')
else:
    pass


def consulta_conta(cartao):
    """Simples tentativa de consultar conta via API da PIC-PAY"""
    
    token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJJRF9BUExJQ0FDQU8iOiI0MDUwIiwiQ0RfT1JHIjoiMjExIiwiaXNzIjoiaHR0cDovL2JhbS5mbmlzLmNvbS5iciIsImlhdCI6MTUyMTQ4MDA4MX0.toR2fV-uMoJKJMbBjqPpQ1mwCFStO4hDDrAxHgczl5I'

    endpoint = 'https://picpayintegration-homolog.fnis.com.br'

    servico = '/PICPAY/VPService/ConsultaInformacoesCartao/V1'

    headers = {
    'Content-Type': 'application/json',
    'x-authorization': token    
        }

    data = {
    'nrCartao' : cartao
        }
    
    max_tentativas = 3
    
    for tentativa in range(1, max_tentativas + 1):
        try:     
            response = requests.post(endpoint + servico, headers=headers,json=data )
            print(response)
            if response.status_code == 200:
                if response.text:
                    response_data = response.json()
                    pagto_min = response_data['DadosConta']['vlPagamentoMinimo'] 
                    sldo_ult_fat = response_data['DadosConta']['saldoUltimaFatura']    
                    dt_ult_pag = response_data['DadosConta'].get('dtUltimoPagto', 0) 
                    dt_ult_compra = response_data['DadosConta'].get('dtUltimaCompradtUltimaCompra', 0)
                    ultimo_vcto = response_data['DadosConta'].get('dtVctoUltimaFatura', 0)
                    return ultimo_vcto, pagto_min, sldo_ult_fat, dt_ult_pag, dt_ult_compra
        except Exception as e:
            print(f'tentativa{tentativa}: falha ao chamar a API: {e}')  
                           
    
def entrar_vision(pagina):
    sleep(0.3)
    janelas = pyautogui.getWindowsWithTitle('CQ - Extra')
    if janelas:
        janela = janelas[0]
        janela.restore()
        janela.activate()
        janela.maximize()
    else:
        print("nenhuma janela encontrada com o título: CQ - Extra")  
    sleep(0.1)
    pyautogui.press('HOME')
    sleep(0.1)
    pyautogui.write(pagina)
    sleep(0.1)
    pyautogui.press('Enter')    
    sleep(0.1)    
    

def manter_foco_vision():
    sleep(0.1)
    janelas = pyautogui.getWindowsWithTitle('CQ - Extra')
    if janelas:
        janela = janelas[0]
        janela.restore()
        janela.activate()
        janela.maximize()
    else:
        print("nenhuma janela encontrada com o título: CQ - Extra")  
        
    
def realizar_pgto(org,cartao,valor_pag):
    #simples definição para criar uma função de pagamento
    #abrindo o lote
    pyautogui.write(str(org))
    sleep(0.3)
    pyautogui.press("ENTER")
    sleep(0.3)
    pyautogui.write('001')
    sleep(0.3)
    pyautogui.write('0')
    sleep(0.3)
    pyautogui.press("TAB")
    sleep(0.3)
    pyautogui.write(str(valor_pag))
    sleep(0.3)    
    pyautogui.press("TAB")
    sleep(0.3)
    pyautogui.write('986')
    sleep(0.3)
    pyautogui.press("ENTER")
    sleep(0.3)    
    #entrando com as informações de pagamento.
    pyautogui.write(str(cartao))
    sleep(0.3)
    pyautogui.write(str(valor_pag))
    sleep(0.3)
    pyautogui.press("TAB")
    sleep(0.3)
    pyautogui.write('500')
    sleep(0.3)
    pyautogui.press("ENTER")
    sleep(1)  
    pyautogui.press("F1")
    sleep(1)  
    pyautogui.press("F1")
    sleep(1)  
    manter_foco_vision()
    pyautogui.press("F1")    
    

def realizar_db(org,cartao,valor):
    #simples definição para criar uma função de pagamento    
    #abrindo o lote
    pyautogui.write(str(org))
    sleep(0.3)
    pyautogui.press("ENTER")
    sleep(0.3)
    pyautogui.write('001')
    sleep(0.3)
    pyautogui.write(str(valor))
    sleep(0.3)
    pyautogui.press("TAB")
    sleep(0.3)
    pyautogui.write('0')
    sleep(0.3)
    pyautogui.press("TAB")
    sleep(0.3)
    pyautogui.write('986')
    sleep(0.3)
    pyautogui.press("ENTER")
    sleep(0.3)
    #entrando com as informações de pagamento.
    pyautogui.write(str(cartao))
    sleep(0.3)
    pyautogui.write(str(valor))
    sleep(0.3)
    pyautogui.press("TAB")
    sleep(0.3)
    pyautogui.write('111')
    sleep(0.3)
    pyautogui.press("TAB")
    sleep(0.3)
    pyautogui.write('999999998')
    sleep(0.3)
    pyautogui.write('00002')
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.press("TAB")
    pyautogui.write('Compra - robo aut')
    pyautogui.press("ENTER")
    sleep(1)  
    pyautogui.press("F1")
    sleep(1)  
    pyautogui.press("F1")
    sleep(1)  
    manter_foco_vision()
    pyautogui.press("F1")
    
    
def ajustar_valor(valor):
    valor_aj = round(valor, 2)
    valor_aj = str(valor_aj).replace("-", "")
    valor_aj = str(valor_aj).replace(",", "")
    valor_aj = str(valor_aj).replace(".", "")
    return valor_aj


def converter_datas(data):
    if not data:
        ano = 1111
        mes = 11
        dia = 11
        data_dt = datetime(ano, mes, dia)
        return data_dt.date()
    else:  
        ano = int(data[0:4])
        mes = int(data[5:7])
        dia = int(data[8:10])
        data_dt = datetime(ano, mes, dia)
        return data_dt.date()
    
    
for numero_linha, linha in enumerate(pagina_contas.iter_rows(min_row=2)):
    cartao = str(linha[0].value)
    acao = str(linha[1].value)
    infos = consulta_conta(cartao)              
    #realizando o unpack das informações enviadas pela função consulta_conta
    ultimo_vcto, pagto_min, sldo_ult_fat, dt_ult_pag, dt_ult_compra = infos
    #convertendo informações
    data_atual = date.today()
    ultimo_vcto_aj = converter_datas(ultimo_vcto)
    dt_ult_pag_aj = converter_datas(dt_ult_pag)   
    dt_ult_compra_aj = converter_datas(dt_ult_compra) 
    print (cartao, ultimo_vcto_aj, dt_ult_pag_aj, sldo_ult_fat, pagto_min, dt_ult_compra)
     
 #lógica para pagamento e debito
    if data_atual > ultimo_vcto_aj:
        if  ultimo_vcto_aj > dt_ult_pag_aj:
            if acao == 'rotativo':
                if pagto_min == 0:
                    valor = '5000'
                    entrar_vision('ARAT')
                    realizar_pgto('211', cartao, valor )
                    entrar_vision('ARAT')
                    realizar_db('211', cartao, '15000' )
                elif pagto_min > 0 and pagto_min != sldo_ult_fat:
                    valor_aj1 = ajustar_valor(pagto_min)
                    entrar_vision('ARAT')
                    realizar_pgto('211', cartao, valor_aj1 )
                    entrar_vision('ARAT')
                    realizar_db('211', cartao, '15000' )
                elif pagto_min > 0 and pagto_min == sldo_ult_fat:
                    valor = (sldo_ult_fat * 0.20)                     
                    valor_aj5 = ajustar_valor(valor)
                    entrar_vision('ARAT')
                    realizar_pgto('211', cartao, valor_aj5)
                    entrar_vision('ARAT')
                    realizar_db('211', cartao, '5000')
            elif acao == 'em dia':
                if sldo_ult_fat > 0:
                    valor_aj2 = ajustar_valor(sldo_ult_fat)
                    entrar_vision('ARAT')
                    realizar_pgto('211', cartao, valor_aj2 )
                    entrar_vision('ARAT')
                    realizar_db('211', cartao, '15000' )
                elif sldo_ult_fat == 0:
                    entrar_vision('ARAT')
                    realizar_db('211', cartao, '15000' )
                elif sldo_ult_fat < 0:
                    valor_aj4 = ajustar_valor(sldo_ult_fat)
                    if dt_ult_compra_aj < ultimo_vcto_aj:                  
                        entrar_vision('ARAT')
                        realizar_pgto('211', cartao, '1000' )
                        entrar_vision('ARAT')
                        realizar_db('211', cartao, valor_aj4)    
                    else:
                        entrar_vision('ARAT')
                        realizar_db('211', cartao, '5000')               
        else:
            entrar_vision('ARAT')
            realizar_db('211', cartao, '5000')    
    else:
        print(f' O {cartao} a data atual é menor que o vencimento')
        entrar_vision('ARAT')
        realizar_db('211', cartao, '5000')  
    sleep(0.3)

