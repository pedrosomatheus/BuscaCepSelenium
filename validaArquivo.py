import os
from pathlib import Path
from datetime import datetime

#Função criada para escrever o log
def EscreveLog(mensagem):

    #Passando fixo para não ter que passar na chamada da função e já valido a criação dessas pastas
    ArquivoLog = "C:\\Users\\edvan\\BPA001 - BuscaCep\\1. LOG\\2023\\01\\14\\LOG.txt"

    my_file = open(ArquivoLog, 'a', encoding='utf-8')

    my_file.write(f'{mensagem}' + '\n')

    my_file.close()


CaminhoProjeto = f'{Path.home()}\\BPA001 - BuscaCep'

ArquivoLog = f'{CaminhoProjeto}\\1. LOG'

CaminhoInput = f'{CaminhoProjeto}\\2. INPUT'

CaminhoProcessamento = f'{CaminhoProjeto}\\3. PROCESSAMENTO'

CaminhoFinalizado = f'{CaminhoProjeto}\\4. FINALIZADO'

#Capturando a data de hoje
now = datetime.now()


#Validando se a pasta do caminho do projeto existe
if os.path.isdir(CaminhoProjeto) == False:

    os.mkdir(CaminhoProjeto)

#Validando se a pasta de Log existe
if os.path.isdir(ArquivoLog) == False:

    os.mkdir(ArquivoLog)

ArquivoLog = f'{ArquivoLog}\\{now.strftime("%Y")}'

#Criando pasta de acordo com ano, mes e dia
if os.path.isdir(ArquivoLog) == False:

    os.mkdir(ArquivoLog)


ArquivoLog = f'{ArquivoLog}\\{now.strftime("%m")}'

#Criando pasta de acordo com ano, mes e dia
if os.path.isdir(ArquivoLog) == False:

    os.mkdir(ArquivoLog)


ArquivoLog = f'{ArquivoLog}\\{now.strftime("%d")}'

#Criando pasta de acordo com ano, mes e dia
if os.path.isdir(ArquivoLog) == False:

    os.mkdir(ArquivoLog)   

ArquivoLog = f'{ArquivoLog}\\LOG.txt'


#Validando se a pasta  Input existe
if os.path.isdir(CaminhoInput) == False:

    os.mkdir(CaminhoInput)
    

#Validando se a pasta Processamento existe
if os.path.isdir(CaminhoProcessamento) == False:

    os.mkdir(CaminhoProcessamento)


#Validando se a pasta Finalizado existe
if os.path.isdir(CaminhoFinalizado) == False:

    os.mkdir(CaminhoFinalizado)


EscreveLog("=========================== INICIO - Valida Arquivo ================================")




















    
EscreveLog("=========================== FIM - Valida Arquivo ================================")