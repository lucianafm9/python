# -------------------------------------------------
# Pré-requisitos: o CAPS LOCK NÃO pode estar ativado
# Instalados o módulos: pip install xlrd
# -------------------------------------------------

from automagica import *
import subprocess
from datetime import datetime, timedelta
#import imaplib
#import email
import os
import time
from datetime import date, timedelta
import xlrd
import win32clipboard
#import xml.etree.ElementTree as ET

CAMINHO_PORTAL = "https://www5.citrosuco.com.br:50001/irj/portal"
CAMINHO_EXCEL = "C:/Portal da Citrosuco/"
CAMINHO_EXCEL_PROCESSADOS= "C:/Citrosuco processados/"
USUARIO_PORTAL = "ct02003231"
SENHA_PORTAL = "tmt2019"
DATA_ATUAL = date.today()
DATA_ATUAL_TEXTO = '{}/{}/{}'.format(DATA_ATUAL.day, DATA_ATUAL.month, DATA_ATUAL.year)
NUMERO_DIA_SEMANA = DATA_ATUAL.isoweekday()
DATA_FIM_FILTRO = date.today() - timedelta(NUMERO_DIA_SEMANA)
DATA_INICIO_FILTRO = date.today() - timedelta(NUMERO_DIA_SEMANA + 6)
DATA_INCIO_FILTRO_TEXTO = DATA_INICIO_FILTRO.strftime("%d/%m/%Y")
DATA_FIM_FILTRO_TEXTO = DATA_FIM_FILTRO.strftime("%d/%m/%Y")
NOME_ARQUIVO = DATA_ATUAL.strftime("%d%m%Y") + ".xlsx"
CAMINHO_E_NOME_ARQUIVO = CAMINHO_EXCEL + NOME_ARQUIVO


# -------------------------------------------------
# Acessando o portal
# -------------------------------------------------
def access():
    browser = subprocess.Popen(r'"C:\Program Files\Internet Explorer\IEXPLORE.EXE" https://www5.citrosuco.com.br:50001/irj/portal')
    time.sleep(25)
    return browser

# -------------------------------------------------
# Login no portal
# -------------------------------------------------
def login():
    Type(USUARIO_PORTAL, interval_seconds=0.01)
    PressKey("tab")
    time.sleep(0.4)
    Type(SENHA_PORTAL, interval_seconds=0.01)
    time.sleep(5)
    PressKey("tab")
    time.sleep(0.4)
    Enter()

# -------------------------------------------------
# Filtro e realizando a busca no portal
# -------------------------------------------------
def filtro():
    time.sleep(4)
    tab(12)
    PressKey("down")
    PressKey("down")
    time.sleep(5)
    tab(2)
    Type(DATA_INCIO_FILTRO_TEXTO, interval_seconds=0.01)
    time.sleep(0.4)
    tab(2)
    Type(DATA_FIM_FILTRO_TEXTO, interval_seconds=0.01)
    tab(6)
    Enter()

# -------------------------------------------------
# Função genérica para realizar N tabs consecutivos - N deve ser passado como parâmetro
# -------------------------------------------------
def tab(qtd__repetition):
    for i in range(qtd__repetition):
        Tab()

# -------------------------------------------------
# Função genérica para realizar N shift+tabs (voltar cursor) consecutivos - N deve ser passado como parâmetro
# -------------------------------------------------
def shifttab(qtd__repetition):
    for i in range(qtd__repetition):
        PressHotkey("shift", "tab")

# -------------------------------------------------
# procurar por um arquivo existente
# -------------------------------------------------
def existearq():
    return FileExists(CAMINHO_E_NOME_ARQUIVO)

# -------------------------------------------------
# Abrir arquivo excel
# -------------------------------------------------
def abrirarq():
    loc = (CAMINHO_E_NOME_ARQUIVO)
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    return sheet

def workbook():
    loc = (CAMINHO_E_NOME_ARQUIVO)
    wb = xlrd.open_workbook(loc)
    return wb

# -------------------------------------------------
# Abrir arquivo excel e ler a célula
# -------------------------------------------------
def lercelula(sheet, row, column):
    return sheet.cell_value(row, column)

def lerceluladata(wrongValue, wb):
    return xlrd.xldate_as_tuple(wrongValue, wb.datemode)

# -------------------------------------------------
# Quantidade total de linhas
# -------------------------------------------------
def qttotallinhas(sheet):
    return sheet.nrows

# -------------------------------------------------
# Seleciona todos
# -------------------------------------------------
def selectodos():
    shifttab(9)  # cair no selecionar todos
    PressKey("space")

# -------------------------------------------------
# Verifica se retornou resultado na busca. Se retornou dados na busca, vai retornar 1, caso contrário retornará 2
# -------------------------------------------------
def existeresult():
    time.sleep(100)
    #time.sleep(300) #aguarda 5min para o resultado da pesquisa aparecer
    tab(26)
    PressHotkey("ctrl", "c")
    win32clipboard.OpenClipboard()
    dado_copiado = win32clipboard.GetClipboardData()
    if "https://www5.citrosuco.com.br:50001/irj/portal" not in dado_copiado:
        return 1
    else:
        return 2

def escrevelinha(valor_totalctrc, numero_ctrc, data_ctrc, chave_acesso):
    Type(str(valor_totalctrc), interval_seconds=0.01)
    tab(1)
    Type(str(numero_ctrc), interval_seconds=0.01)
    tab(1)
    Type(str(data_ctrc), interval_seconds=0.01)
    tab(2)
    Type(str(chave_acesso), interval_seconds=0.01)

def mudarpagina(qtdpagina):
    if(qtdpagina == 1):
        tab(2)#2 ou 3
    else:
        tab(5)
    PressKey("enter")

def confirmar():
    tab(8)
    PressKey("enter")

def moverarq():
    for oldname in os.listdir(CAMINHO_EXCEL):
        if(oldname == NOME_ARQUIVO):
            newname = 'PROC_' + oldname
            oldname = os.path.join(CAMINHO_EXCEL, oldname)
            newname = os.path.join(CAMINHO_EXCEL_PROCESSADOS, newname)
            os.rename(oldname, newname)

def arquivo():
    while True:
        try:
            browser = access()
            if browser:
                login()
                if(existearq()):
                    filtro()
                    if(existeresult() == 1):
                        selectodos()
                        sheet = abrirarq()
                        wb = workbook()
                        qtdtotalrows = qttotallinhas(sheet)
                        qtdlinhaspag = 1
                        qtdpagina = 1
                        for i in range(qtdtotalrows):
                            if(qtdlinhaspag == 1):
                                tab(19)
                            else:
                                tab(1)
                            valor_totalctrc = 0
                            numero_ctrc = ""
                            data_ctrc = ""
                            chave_acesso = ""
                            if(i < (qtdtotalrows - 1)):
                                valor_totalctrc = lercelula(sheet,i+1,0)
                                numero_ctrc = lercelula(sheet,i+1,1)
                                data_ctrc = lercelula(sheet, i + 1, 2)
                                chave_acesso = lercelula(sheet, i + 1, 3)
                            else:
                                valor_totalctrc = lercelula(sheet,i,0)
                                numero_ctrc = lercelula(sheet, i, 1)
                                data_ctrc = lercelula(sheet, i, 2)
                                chave_acesso = lercelula(sheet, i, 3)
                            year, month, day, hour, minute, second = lerceluladata(data_ctrc, wb)
                            data_ctrc = ("0" + str(day))[-2:] + '/' + ("0" + str(month))[-2:] + '/' + str(year)
                            escrevelinha(valor_totalctrc, numero_ctrc, data_ctrc, chave_acesso)
                            if(qtdlinhaspag == 10):
                                if (i < (qtdtotalrows - 1)):
                                    mudarpagina(qtdpagina)
                                    qtdpagina = qtdpagina + 1
                                else:
                                    teste = "na homologação, comentar essa linha e descomentar a linha confirmar()"
                                    #confirmar()
                                qtdlinhaspag = 1
                            else:
                                qtdlinhaspag = qtdlinhaspag + 1
                    moverarq()
                browser.kill()
                time.sleep(1800) #repete a cada 30min para verificar se possui
        except Exception as e:
            print("Erro durante o processo: " + str(e))

arquivo()