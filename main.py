import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import os

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
def WebValidaTextJs(Id,tempo,TextoElemento,TextoErro = "Não há dados a serem exibidos"):

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

                #Validando se o elemento foi encontrado
        if ValidaCarragamento == TextoErro:
            break

        #Validando se já deu o tempo de validação
        if i >=tempo:
            
            break

        time.sleep(1)

    return ValidaCarragamento
        





#Definindo variavel de tempo
tempoCurto = 10
tempoMedio = 30
tempoLongo = 60
wait = time.sleep(1)



EscreveLog("=========================== INICIO - Navegação Busca Cep ================================")

# Matando processo do excel
os.system('taskkill /F  /IM EXCEL.EXE')

# Chamando a função que faz a validação das pasta e arquivo de log para retornar o nome do arquivo excel
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
            r'ReadOnly=0') # O padrão do Excel é uma conexão somente leitura, portanto, se você quiser atualizar a planilha, inclua ReadOnly=0

# Definir nossa conexão, autocommit DEVE SER CONFIGURADO PARA TRUE, também podemos editar dados.
cnxn = pyodbc.connect(conn_str, autocommit=True)
crsr = cnxn.cursor()


for worksheet in crsr.tables():

    # Pegando worksheet
    mensagem = "Pegando worksheet"
    EscreveLog(mensagem)
    tableName = worksheet[2]
    
    
# "SELECT * FROM [Planilha1$]"
mensagem = "Query executada: SELECT * FROM [Planilha1$]"
EscreveLog(mensagem)

Query = "SELECT * FROM [{}] WHERE [Status] IS NULL".format(tableName)
crsr.execute(Query)

# Loop na minha tabela
for row in crsr:
    
    # Setando variaveis
    Cep = row.CEP
    lougradouro = ""
    bairro = ""
    localidade= ""


    #Abrindo navegador
    mensagem = "Abrindo navegador"
    EscreveLog(mensagem)


    Url = "https://buscacepinter.correios.com.br/app/endereco/index.php"

    driver = abreNavegador(Url)



    # Chamando função para validar se a tela carregou
    mensagem = "Chamando função para validar se a tela carregou"
    EscreveLog(mensagem)

    Id = "titulo_tela"
    TextoElemento = "Busca CEP"

    WebValidaTextJs(Id,tempoCurto,TextoElemento)



    # Setando o valor no site
    mensagem = "Setando o valor no site"
    EscreveLog(mensagem)

    Id = "endereco"

    SetaElementoId(Id, Cep)


    # Clicando no botão pesquisar
    mensagem = "Clicando no botão pesquisar"
    EscreveLog(mensagem)

    Id = "btn_pesquisar"

    ClickId(Id)

   

    # Chamando função para validar se o CEP carregou
    mensagem = "Chamando função para validar se o CEP carregou"
    EscreveLog(mensagem)

    Id = "mensagem-resultado"
    TextoElemento = "Resultado da Busca por Endereço ou CEP"

    if WebValidaTextJs(Id,tempoCurto,TextoElemento) != TextoElemento:

     # Elemento de erro encontrado: "Não há dados a serem exibidos"

            
        Query = f"UPDATE [{format(tableName)}] SET [Status] = 'Não há dados a serem exibidos' WHERE [CEP] = '{Cep}'"
        crsr.execute(Query)
        cnxn.commit()


        
        






    input("teste")

    


   


EscreveLog("=========================== FIM - Navegação Busca Cep ================================")























