import os
import shutil
from pathlib import Path
from datetime import datetime

#Função criada para escrever o log
def EscreveLog(mensagem):

    now = datetime.now()

    #Passando fixo para não ter que passar na chamada da função e já valido a criação dessas pastas
    ArquivoLog = f"{Path.home()}\\BPA001 - BuscaCep\\1. LOG\\{now.strftime('%Y')}\\{now.strftime('%m')}\\{now.strftime('%d')}\\LOG.txt"


    dataHora = now.strftime("%Y-%m-%d %H:%M:%S")

    my_file = open(ArquivoLog, 'a', encoding='utf-8')

    my_file.write(f'{dataHora} - '+ f'{mensagem}' + '\n')

    my_file.close()


#Função que faz a validação se as pasta já estão criadas
def ValidaArquivo():



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


    CaminhoArquivoExcel = ""

    #Validando se já não contem arquivo na pasta PROCESSAMENTO
    mensagem = "Validando se já não contem arquivo na pasta PROCESSAMENTO"
    EscreveLog(mensagem)

    caminhosArquivo = [
    os.path.join(CaminhoProcessamento, nome) 
    for nome in os.listdir(CaminhoProcessamento)
    ]

    for arq in caminhosArquivo:
        if arq.lower().endswith(".xlsx"):
            CaminhoArquivoExcel = arq
            mensagem = f"Arquivo encontrado na pasta PROCESSAMENTO: {CaminhoArquivoExcel}"
            EscreveLog(mensagem)

    

    if len(CaminhoArquivoExcel) == 0:


        #Listando os arquivos dentro da pasta INPUT
        mensagem = "Listando os arquivos dentro da pasta INPUT"
        EscreveLog(mensagem)

        caminhosArquivo = [
        os.path.join(CaminhoInput, nome) 
        for nome in os.listdir(CaminhoInput)
        ]

        #Capturando o nome do arquivo excel e movendo para pasta de processamento
        mensagem = "Capturando o nome do arquivo excel e movendo para pasta de processamento"
        EscreveLog(mensagem)

        for arq in caminhosArquivo:
            if arq.lower().endswith(".xlsx"):
                CaminhoArquivoExcel = arq
                shutil.move(CaminhoArquivoExcel, CaminhoProcessamento)

                #Movendo arquivo de INPUT para PROCESSAMENTO
                mensagem = f"Movendo arquivo de INPUT: {CaminhoArquivoExcel} para PROCESSAMENTO: {CaminhoProcessamento}"
                EscreveLog(mensagem)

                CaminhoArquivoExcel = arq.replace("2. INPUT", "3. PROCESSAMENTO")
                break
                        


        EscreveLog("=========================== FIM - Valida Arquivo ================================")

    return CaminhoArquivoExcel



        

caminhoArquivo = ValidaArquivo()







    

