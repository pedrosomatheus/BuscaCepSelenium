import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

import pyodbc

from validaArquivo import EscreveLog
from validaArquivo import ValidaArquivo




#Chamando a função que faz a validação das pasta e arquivo de log
ValidaArquivo()



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



# we may be interested in finding all the drivers we have access to
for driver in pyodbc.drivers():
   
    # print the driver name
    print(driver)
    
    # if the driver name has '.xlsx' in it we found it!
    if '.xlsx' in driver:
        myDriver = driver

# define our connection string
conn_str = (r'DRIVER={'+ myDriver +'};'
            r'DBQ=C:\Users\Alex\Desktop\Video List.xlsx;'
            r'ReadOnly=1') # ReadOnly set to 0 means we can edit the data.

# define our connection, autocommit MUST BE SET TO TRUE, also we can edit data.
cnxn = pyodbc.connect(conn_str, autocommit=True)
crsr = cnxn.cursor()

# loop through all the tables
for worksheet in crsr.tables():
    
    # display the worksheet
    print(worksheet)
    
    # grab the table name.
    tableName = worksheet[2]
    
# define our query to grab the data.
# we want this "SELECT Topic FROM [Sheet1$]"
crsr.execute("SELECT Topic FROM [{}]".format(tableName))

# print each row of data.
for row in crsr:
    print(row.Topic)



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

SetaElementoId(Id, "06680103")





#Clicando no botão pesquisar
mensagem = "Clicando no botão pesquisar"
EscreveLog(mensagem)

Id = "btn_pesquisar"

ClickId(Id)




EscreveLog("=========================== FIM - Navegação Busca Cep ================================")

input("teste")



















