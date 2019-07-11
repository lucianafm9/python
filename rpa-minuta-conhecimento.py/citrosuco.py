# -------------------------------------------------
# Pré-requisitos: o CAPS LOCK NÃO pode estar ativado
# Instalados o módulos: pip install xlrd
# -------------------------------------------------

from automagica import *
import subprocess
from datetime import datetime, timedelta
import imaplib
import email
import time
from datetime import date, timedelta
import xlrd
import win32clipboard
import xml.etree.ElementTree as ET

CAMINHO_PORTAL = "https://www5.citrosuco.com.br:50001/irj/portal"
CAMINHO_EXCEL = "C:/Portal da Citrosuco/"
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
    #Type("24/06/2019", interval_seconds=0.01)
    time.sleep(0.4)
    tab(2)
    Type(DATA_FIM_FILTRO_TEXTO, interval_seconds=0.01)
    #Type("30/06/2019", interval_seconds=0.01)
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

def principal():
    browser = access()
    if browser:
        login()
        filtro()
        if(existeresult() == 1):
            selectodos()
            #tab(19) #fazer laço para preencher os dados
            tab(70)
            PressKey("enter")
            tab(73)
            PressKey("enter")
        else:
            print("nao existe resultado")

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
    time.sleep(120)
    tab(26)
    PressHotkey("ctrl", "c")
    win32clipboard.OpenClipboard()
    dado_copiado = win32clipboard.GetClipboardData()
    if "https://www5.citrosuco.com.br:50001/irj/portal" not in dado_copiado:
        return 1
    else:
        return 2

def arquivo():
    if(existearq()):
        browser = access()
        if browser:
            login()
            filtro()
            if(existeresult() == 1):
                sheet = abrirarq()
                wb = workbook()
                qtdtotalrows = qttotallinhas(sheet)
                for i in range(qtdtotalrows):
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
                    #print(str(valor_totalctrc) + "_" + numero_ctrc + "_" + data_ctrc + "_" + str(chave_acesso))

principal()

#arquivo()
#