import time
from datetime import datetime
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

#Função para executar um JS e ter o retorno do valor
def WebRetornaJs(Script):

    WebRetornaJs = driver.execute_script(Script)

    return WebRetornaJs

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

#Função para capturar o valor do texto por Js
def WebGetTextJs(Id):

    WebGetTextJs = driver.execute_script("return document.getElementById('"+Id+"').innerText")

    return WebGetTextJs





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


if len(CaminhoArquivoExcel) > 0:

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

    Query = f"SELECT * FROM [{format(tableName)}]"
    crsr.execute(Query)

    #Abrindo navegador
    mensagem = "Abrindo navegador"
    EscreveLog(mensagem)


    Url = "https://buscacepinter.correios.com.br/app/endereco/index.php"

    driver = abreNavegador(Url)


    # Loop na minha tabela
    for row in crsr:
        
        print(row)

        input("teste")
        # Setando variaveis
        Cep = row.CEP
        lougradouro = ""
        bairro = ""
        localidade= ""


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

        RetornoTexto = WebValidaTextJs(Id,tempoCurto,TextoElemento)


        mensagem = f"Validando elemento: {RetornoTexto} != {TextoElemento}"
        EscreveLog(mensagem)


        if  RetornoTexto != TextoElemento:

        # Elemento de erro encontrado: "Não há dados a serem exibidos"
            mensagem = "Elemento de erro encontrado: 'Não há dados a serem exibidos'"
            EscreveLog(mensagem)

                
            Query = f"UPDATE [{format(tableName)}] SET [Status] = 'Não há dados a serem exibidos' WHERE [CEP] = '{Cep}'"
            mensagem = f"Realizando query: {Query}"
            EscreveLog(mensagem) 

            crsr.execute(Query)
            cnxn.commit()

        else:

            #Texto encontrado capturando informações do CEP
            mensagem = "Texto encontrado capturando informações do CEP"
            EscreveLog(mensagem)

            try:

                Script = "return document.getElementsByTagName('td')[0].innerText"
                logradouro = WebRetornaJs(Script)

                Script = "return document.getElementsByTagName('td')[1].innerText"
                bairro = WebRetornaJs(Script)
                    
                Script = "return document.getElementsByTagName('td')[2].innerText"
                localidade = WebRetornaJs(Script)

                now = datetime.now()

                dataHora = now.strftime("%Y-%m-%d %H:%M:%S")

            except:
                #A elementos que dependendo do CEP que não aparecem
                pass


            Query = f"UPDATE [{format(tableName)}] SET [Logradouro] = '{logradouro}', [Bairro] = '{bairro}', [Localidade] = '{localidade}', [Data da Consulta] = '{dataHora}', [Status] = 'OK' WHERE [CEP] = '{Cep}'"
            crsr.execute(Query)
            cnxn.commit()



            # Clicando no botão pesquisar
            mensagem = "Clicando no botão Nova Busca"
            EscreveLog(mensagem)

            Id = "btn_nbusca"

            ClickId(Id)



    EscreveLog("=========================== FIM - Navegação Busca Cep ================================")



input("teste")



















