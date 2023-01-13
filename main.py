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


def ClickId(Id):

    driver.execute_script("document.getElementById('"+Id+"').click()")






#Abrindo navegador

Url = "https://buscacepinter.correios.com.br/app/endereco/index.php"

driver = abreNavegador(Url)

time.sleep(2)


#Clicando no botão pesquisar

Id = "btn_pesquisar"

ClickId(Id)



input("teste")






















