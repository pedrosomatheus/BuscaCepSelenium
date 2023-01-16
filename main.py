import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

import pyodbc

from validaArquivo import EscreveLog
from validaArquivo import ValidaArquivo



#Função para abrir o navegador
def abreNavegador(Url):

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    driver.get(Url)

    driver.maximize_window()

    return driver

#Função para executar comando JS
def ExecutaJs(Script):

    driver.execute_script(Script)

#Fução para clicar pelo Id
def ClickId(Id):

    driver.execute_script("document.getElementById('"+Id+"').click()")



#Função para setar o elemento na pagina
def SetaElementoId(Id, Valor):

    driver.execute_script("document.getElementById('"+Id+"').value='"+Valor+"'")


#Função para capturar o texto por Js e fazer a validação de carragemento da tela
def WebValidaTextJs(Id,TextoElemento,tempo):

    i = 0

    for i in range(tempo):

        #Capturando o valor do texto
        try:
            ValidaCarragamento = driver.execute_script("return document.getElementById('"+Id+"').innerText")
        except:
            pass


        #Validando se o elemento foi encontrado
        if ValidaCarragamento == TextoElemento:
            break

        #Validando se já deu o tempo de validação
        if i >=9:
            
            break

        time.sleep(1)

    return WebValidaTextJs
        





#Definindo variavel de tempo
tempoCurto = 10
tempoMedio = 30
tempoLongo = 60
wait = time.sleep(1)



EscreveLog("=========================== INICIO - Navegação Busca Cep ================================")



#Chamando a função que faz a validação das pasta e arquivo de log para retornar o nome do arquivo excel
mensagem = "Chamando a função que faz a validação das pasta e arquivo de log para retornar o nome do arquivo excel"
EscreveLog(mensagem)

CaminhoArquivoExcel = ValidaArquivo()


# Capturando driver
mensagem = "Capturando driver"
EscreveLog(mensagem)


for driver in pyodbc.drivers():
   
    # Pegando o nome apenas para o driver .xlsx
    mensagem = "Pegando o nome apenas para o driver .xlsx"
    EscreveLog(mensagem)
    if '.xlsx' in driver:
        myDriver = driver

# Definindo connection string
mensagem = "Definindo connection string"
EscreveLog(mensagem)

conn_str = (r'DRIVER={'+ myDriver +'};'
            f'DBQ={CaminhoArquivoExcel};'
            r'ReadOnly=1') # para leitura setar como 0

# definir nossa conexão, autocommit DEVE SER CONFIGURADO PARA TRUE, também podemos editar dados.
cnxn = pyodbc.connect(conn_str, autocommit=True)
crsr = cnxn.cursor()


for worksheet in crsr.tables():

    # Pegando worksheet
    mensagem = "Pegando worksheet"
    EscreveLog(mensagem)
    tableName = worksheet[2]
    
    
#"SELECT * FROM [Planilha1$]"
mensagem = "Query executada: SELECT * FROM [Planilha1$]"
EscreveLog(mensagem)
crsr.execute("SELECT * FROM [{}]".format(tableName))


#Loop na minha tabela
for row in crsr:
    
    #Setando variaveis
    Cep = row.CEP
    lougradouro = ""
    bairro = ""
    localidade= ""


    #Abrindo navegador
    mensagem = "Abrindo navegador"
    EscreveLog(mensagem)


    Url = "https://buscacepinter.correios.com.br/app/endereco/index.php"

    driver = abreNavegador(Url)



    #Chamando função para validar se a tela carregou
    mensagem = "Chamando função para validar se a tela carregou"
    EscreveLog(mensagem)

    Id = "titulo_tela"
    TextoElemento = "Busca CEP"

    WebValidaTextJs(Id,TextoElemento,tempoCurto)



    #Setando o valor no site
    mensagem = "Setando o valor no site"
    EscreveLog(mensagem)

    Id = "endereco"

    SetaElementoId(Id, Cep)



    #Clicando no botão pesquisar
    mensagem = "Clicando no botão pesquisar"
    EscreveLog(mensagem)

    Id = "btn_pesquisar"

    ClickId(Id)




    

    input("teste")

EscreveLog("=========================== FIM - Navegação Busca Cep ================================")





input("teste")




















