import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By



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


#Função para capturar o texto por Js
def WebGetTextJs(Id):

    teste = driver.execute_script("return document.getElementById('"+Id+"').innerText")

    return teste



#Definindo variavel de tempo
tempoCurto = 10
tempoMedio = 30
tempoLongo = 60
wait = time.sleep(1)




#Abrindo navegador

Url = "https://buscacepinter.correios.com.br/app/endereco/index.php"

driver = abreNavegador(Url)


#fazer ele escrever log em uma pasta

i = 0

for i in range(tempoCurto):

    Id = "titulo_tela"

    #validando se a pagina foi carregada
    try:
        ValidaCarragamento = WebGetTextJs(Id)
    except:
        print("Elemento ainda não encontrado")
        pass


    if ValidaCarragamento == "Busca CEP":
        print("Elemento encontrado")
        break

    #Validando se já deu o tempo de validação
    if i >=9:
        
        print("Erro ao localizar o Elemento")
        break

    wait

    

#Setando o valor no site

Id = "endereco"

SetaElementoId(Id, "06680103")


time.sleep(1)

#Clicando no botão pesquisar

Id = "btn_pesquisar"

ClickId(Id)



input("teste")






















