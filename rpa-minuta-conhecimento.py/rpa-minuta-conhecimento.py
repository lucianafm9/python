from automagica import *
from datetime import datetime, timedelta
import imaplib
import email
import time
import xml.etree.ElementTree as ET

ORG_EMAIL = "@tmtlog.com"
FROM_EMAIL = "integracao.tmt" + ORG_EMAIL
FROM_PWD = "tmtconsultoria@2019"
SMTP_SERVER = "imap.tmtlog.com"
SMTP_PORT = 993


# -------------------------------------------------
# Access the app web site
# -------------------------------------------------
def access():
    # Acesso BSoft
    browser = ChromeBrowser()
    browser.get('https://sis.bsoft.com.br/')
    Type("TTL318", interval_seconds=0.01)
    Enter()
    time.sleep(4)
    Enter()
    time.sleep(25)
    return browser


# -------------------------------------------------
# Fill login and password fields
# -------------------------------------------------
def login():
    Type("INTEGRACAO.TMT", interval_seconds=0.01)
    Enter()
    time.sleep(0.4)
    Type("tmtconsultoria", interval_seconds=0.01)
    Enter()
    Enter()
    time.sleep(5)


# -------------------------------------------------
# Chamada tela Minuta da Conhecimento
# -------------------------------------------------
def access_menu():
    PressHotkey("alt", "f")
    PressKey("down")
    PressKey("down")
    PressKey("down")
    PressKey("down")
    Enter()
    time.sleep(1)


# -------------------------------------------------
# Parte 1 - Identificação
# -------------------------------------------------
def populate_tab_identificacao(ticket):
    # Definindo remetente
    PressKey("f1")
    tab(5)
    PressHotkey("ctrl", "enter")
    time.sleep(0.5)
    Type(ticket.imovel, interval_seconds=0.01)  # Nome da fazenda deve ser substituido pelo imovel no xml
    PressHotkey("alt", "o")
    time.sleep(0.3)
    Enter()
    PressHotkey("ctrl", "c")  # copiando codigo da fazenda
    # Definindo Destinatario - Sempre 34 - Fabrica de Matao
    Enter()
    time.sleep(0.3)
    Type("34", interval_seconds=0.01)
    # Definindo Consignatario - Sempre 34 - Fabrica de Matao
    Enter()
    time.sleep(0.3)
    Type("34", interval_seconds=0.01)
    # Definindo Redespacho - Fazenda
    Enter()
    time.sleep(0.3)
    PressHotkey("ctrl", "v")  # Colando codigo da Fazenda
    Enter()


# -------------------------------------------------
# Parte 2 - Dados do Modal
# -------------------------------------------------
def populate_tab_dados_modal(ticket):
    # Acionando segunda aba
    PressHotkey("alt", "2")
    # Inserindo Placa
    enter(9)
    time.sleep(0.3)
    Type(format_placa(ticket.placa), interval_seconds=0.1)  # Placa deve ser substituido pela placa no xml
    Enter()


# -------------------------------------------------
# Parte 3 - Documentos
# -------------------------------------------------
def populate_tab_documentos(ticket):
    # Acionando terceira aba
    PressHotkey("alt", "3")
    PressHotkey("ctrl", "i")
    time.sleep(0.3)
    sum_caixas = sum_strings(ticket.caixasrefugo, ticket.caixasliquido)
    print(sum_caixas)
    Type(sum_caixas, interval_seconds=0.01)  # Somatorio (caixasrefugo + caixasliquido) no XML
    Enter()
    time.sleep(0.3)
    Type(sum_caixas, interval_seconds=0.01)  # Mesmo valor Anterior: Somatorio (caixasrefugo + caixasliquido) no XML
    enter(2)
    time.sleep(0.3)
    invoice = getInvoice(ticket.lote)
    Type(ticket.lote, interval_seconds=0.01)  # Campo lote no XML
    enter(2)
    time.sleep(0.3)
    num_caixas = calc_num_caixas(sum_caixas)
    print(num_caixas)
    Type(num_caixas, interval_seconds=0.01)  # Nº de caixas(valor do campo Peso) * 19,2 (valor da caixa)
    enter(3)
    PressHotkey("ctrl", "enter")
    time.sleep(0.5)
    Type("LARANJA", interval_seconds=0.01)
    PressHotkey("alt", "o")
    time.sleep(0.3)
    Enter()
    tab(7)
    time.sleep(0.3)
    Type("TICKET Nº " + str(ticket.lote), interval_seconds=0.01)  # Campo lote no XML
    PressHotkey("ctrl", "g")  # gravar notas fiscais


# -------------------------------------------------
# Parte 4 - Cálculo do frete
# -------------------------------------------------
def populate_tab_calculo_frete(ticket):
    # Acionando quarta aba
    PressHotkey("alt", "4")
    Enter()
    PressHotkey("ctrl", "enter")
    time.sleep(0.5)
    Enter()
    time.sleep(0.5)
    Type(ticket.imovel + " x FABRICA DE MATAO",
         interval_seconds=0.01)  # Substituir FAZENDA TUBUNAS pelo campo imovel no XML - Manter restante: x FAZENDA MATAO
    PressHotkey("alt", "o")
    time.sleep(0.3)
    Enter()


# -------------------------------------------------
# Parte 5 - Dados do CIOT
# -------------------------------------------------
def populate_dados_ciot(ticket):
    # Gravar Minuta de Conhecimento
    PressKey("f4")
    Enter()
    time.sleep(0.3)
    Tab()
    Enter()
    time.sleep(0.3)
    tab(2)
    Enter()
    time.sleep(1)
    tab(2)
    # Data atual
    data_hora_saida = datetime.now()
    data_hora_entrada = datetime.now() + timedelta(hours=1)
    print(data_hora_entrada)
    print(data_hora_saida)
    Type(data_hora_saida.strftime('%d%m%Y'), interval_seconds=0.1)  # Data Saida - Data Atual
    Enter()
    time.sleep(0.2)
    Type(data_hora_saida.strftime('%H%M'), interval_seconds=0.1)  # Hora Saida - Hora Atual
    Enter()
    time.sleep(0.2)
    Type(data_hora_entrada.strftime('%d%m%Y'), interval_seconds=0.1)  # Data Entrada - Data Atual
    Enter()
    time.sleep(0.2)
    Type(data_hora_entrada.strftime('%H%M'), interval_seconds=0.1)  # Hora Entrada - Hora Atual + 1 hora
    tab(5)
    time.sleep(0.2)
    PressHotkey("ctrl", "enter")
    time.sleep(0.5)
    Enter()
    Type(ticket.imovel + " x FABRICA DE MATAO",
         interval_seconds=0.01)  # Substituir FAZENDA TUBUNAS pelo campo imovel no XML - Manter restante: x FAZENDA MATAO
    PressHotkey("alt", "o")
    time.sleep(0.3)
    Enter()
    tab(15)
    Enter() #rpa
    PressKey("f1")

# -------------------------------------------------
# Aux function - Incluir LOTE (número ticket), substituindo as letras com a seguinte regra:
# •	Matão: MC, MM para 1
# •	Catanduva: CC, CM para 2
# •	Araras: AC, AM para 3
# -------------------------------------------------
def getInvoice(lote):

    nota_fiscal = str(lote).replace("MC", "1")
    nota_fiscal = nota_fiscal.replace("MM", "1")
    nota_fiscal = nota_fiscal.replace("CC", "2")
    nota_fiscal = nota_fiscal.replace("CM", "2")
    nota_fiscal = nota_fiscal.replace("AC", "3")
    nota_fiscal = nota_fiscal.replace("AM", "3")

    return nota_fiscal


# -------------------------------------------------
# Aux function - Sum strings like numbers
# -------------------------------------------------
def sum_strings(str_one, str_two):
    val1 = 0.0
    val2 = 0.0

    if str_one != '':
        val1 = float(str(str_one).replace(",", "."))
    if str_two != '':
        val2 = float(str(str_two).replace(",", "."))

    return str(val1 + val2).replace(".", ",")


# -------------------------------------------------
# Aux function - Calc num box and returns in string format
# -------------------------------------------------
def calc_num_caixas(sum_caixas):
    # Nº de caixas(valor do campo Peso) * 19,2 (valor da caixa)
    val_caixa = 19.2
    qtd_caixa = 0.0

    if sum_caixas != '':
        qtd_caixa = float(str(sum_caixas).replace(",", "."))

    total = qtd_caixa * val_caixa

    return str(total).replace(".", ",")


# -------------------------------------------------
# Aux function - Calc num box and returns in string format
# -------------------------------------------------
def format_placa(placa):

    if len(placa) == 7:
        prefix = placa[:3]
        sufix = placa[3:]

        return prefix + '-' + sufix
    return placa


# -------------------------------------------------
# Aux function - Call tab command repeated times
# -------------------------------------------------
def tab(qtd__repetition):
    for i in range(qtd__repetition):
        Tab()


# -------------------------------------------------
# Aux function - Call enter command repeated times
# -------------------------------------------------
def enter(qtd__repetition):
    for i in range(qtd__repetition):
        Enter()


# -------------------------------------------------
# Aux function - Pass the object to app
# -------------------------------------------------
def send_object_to_bsoft(ticket):
    browser = access()
    if browser:
        login()
        access_menu();
        populate_tab_identificacao(ticket)
        populate_tab_dados_modal(ticket)
        populate_tab_documentos(ticket)
        populate_tab_calculo_frete(ticket)
        populate_dados_ciot(ticket)
        browser.close()


# -------------------------------------------------
# Read email from imap server
# ------------------------------------------------
def read_email():
    while True:
        try:
            mail = imaplib.IMAP4_SSL(SMTP_SERVER)
            mail.login(FROM_EMAIL, FROM_PWD)
            mail.select()

            typ, data = mail.search(None, '(UNSEEN FROM alopes.java@gmail.com SUBJECT "Ticket Balan")')
            mail_ids = data[0]

            id_list = mail_ids.split()
            print(str(len(id_list)) + ' email(s) encontrados.')

            for i in id_list:
                typ, data = mail.fetch(i, '(RFC822)')
                # Mark as unread
                mail.store(i, '-FLAGS', '(\\SEEN)')

                try:
                    for response_part in data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            email_subject = msg['subject']
                            email_from = msg['from']
                            print('Subject : ' + email_subject + '\n')

                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    body = get_email_body(part)
                                    ticket = convert_body_to_object(str(body))
                                    if ticket and ticket.imovel and str(ticket.imovel).__contains__('RIO PARDO') :
                                        # Mark as read
                                        mail.store(i, '+FLAGS', '(\\SEEN)')
                                        # pass the object to other app
                                        send_object_to_bsoft(ticket)
                                        continue
                                    else:
                                        continue
                except Exception as e:
                    print("Erro: " + str(e))
                    continue

            # Run every 60 seconds
            time.sleep(60)
        except Exception as e:
            print("Error on method read_email: " + str(e))


# -------------------------------------------------
# Converts the email body to an object
# -------------------------------------------------
def convert_body_to_object(body):
    try:
        root = ET.ElementTree(ET.fromstring(body)).getroot()
        ticket = Ticket()
        for elem in root:
            if elem.tag == 'transp':
                ticket.transp = elem.text
            elif elem.tag == 'placa':
                ticket.placa = elem.text
            elif elem.tag == 'codprodutor':
                ticket.codprodutor = elem.text
            elif elem.tag == 'produtor':
                ticket.produtor = elem.text
            elif elem.tag == 'imovel':
                ticket.imovel = elem.text
            elif elem.tag == 'nfprodutor':
                ticket.nfprodutor = elem.text
            elif elem.tag == 'pesobruto':
                ticket.pesobruto = elem.text
            elif elem.tag == 'pesotara':
                ticket.pesotara = elem.text
            elif elem.tag == 'pesorefugo':
                ticket.pesorefugo = elem.text
            elif elem.tag == 'caixasrefugo':
                ticket.caixasrefugo = elem.text
            elif elem.tag == 'pesoliquido':
                ticket.pesoliquido = elem.text
            elif elem.tag == 'caixasliquido':
                ticket.caixasliquido = elem.text
            elif elem.tag == 'balentdata':
                ticket.balentdata = elem.text
            elif elem.tag == 'balenthora':
                ticket.balenthora = elem.text
            elif elem.tag == 'balsaidata':
                ticket.balsaidata = elem.text
            elif elem.tag == 'balsaihora':
                ticket.balsaihora = elem.text
            elif elem.tag == 'lote':
                ticket.lote = elem.text
            elif elem.tag == 'acoplado':
                ticket.acoplado = elem.text
            elif elem.tag == 'lote1':
                ticket.lote1 = elem.text
            elif elem.tag == 'caixasliquidolote1':
                ticket.caixasliquidolote1 = elem.text
            elif elem.tag == 'lote2':
                ticket.lote2 = elem.text
            elif elem.tag == 'caixasliquidolote2':
                ticket.caixasliquidolote2 = elem.text
            elif elem.tag == 'lote3':
                ticket.lote3 = elem.text
            elif elem.tag == 'caixasliquidolote3':
                ticket.caixasliquidolote3 = elem.text
            elif elem.tag == 'lote4':
                ticket.lote4 = elem.text
            elif elem.tag == 'caixasliquidolote4':
                ticket.caixasliquidolote4 = elem.text
            else:
                print(elem.tag, elem.text)

        return ticket

    except Exception as e:
        print("Error on method convert_body_to_object: " + str(e))
        return None


# -------------------------------------------------
# Get the body of email
# ------------------------------------------------
def get_email_body(part):
    body = part.get_payload(decode=True)
    # treatments in text
    body = str(body).replace('\\r', '').replace('\\n', '').replace('b"', '').replace('"', '').replace('\'', '"')
    return body


# -------------------------------------------------
# Ticket class
# ------------------------------------------------
class Ticket:
    transp = ''
    placa = ''
    codprodutor = ''
    produtor = ''
    imovel = ''
    nfprodutor = ''
    pesobruto = ''
    pesotara = ''
    pesorefugo = ''
    caixasrefugo = ''
    pesoliquido = ''
    caixasliquido = ''
    balentdata = ''
    pesobruto = ''
    balenthora = ''
    balsaidata = ''
    balsaihora = ''
    lote = ''
    acoplado = ''
    balenthora = ''
    lote1 = ''
    caixasliquidolote1 = ''
    lote2 = ''
    caixasliquidolote2 = ''
    lote3 = ''
    caixasliquidolote3 = ''
    lote4 = ''
    caixasliquidolote4 = ''


# starts processing
read_email()