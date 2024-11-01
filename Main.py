from pandas import read_excel, isna

from isRunningWhere import acharCaminhoExecutado
from shutil import copy, copytree
from os import path, makedirs

file = read_excel(acharCaminhoExecutado('./Excel/AutoTransfer.xlsx'))

listaOrigem = file['Origem']
listaDestino = file['Destino']


def autoTransfer():
    print("="*30)
    print("INICIANDO AUTO TRANSFER")
    print("="*30)

    linha = 2
    for i in range(len(listaDestino)):
        #Verificando se um dos campos está vazio
        if isna(listaOrigem[i]) or isna(listaDestino[i]):

            #o pandas faz a leitura mínima das primeiras 8 linhas, a linha vai até a 7 pois tem o cabeçalho
            #Se for até a linha 8
            if linha <= 7:
                if((isna(listaOrigem[i]) == False and isna(listaDestino[i]) == True) or (isna(listaOrigem[i]) == True and isna(listaDestino[i]) == False)):
                    print(f"Valores vazio na linha {linha}")
                    print(f"Origem: {listaOrigem[i]}")
                    print(f"Destino: {listaDestino[i]}")
                    print('-'*50)

            #Acima da linha 7
            if linha > 7:
                print(f"Valores vazio na linha {linha}")
                print(f"Origem: {listaOrigem[i]}")
                print(f"Destino: {listaDestino[i]}")
                print('-'*50)

            linha += 1
            continue

        try:
            #Verificando se a origem é pasta
            if(path.isdir(listaOrigem[i]) == True):

                #Verificando se o destino existe
                if not path.isdir(listaDestino[i]):

                    #Criando as pastas não existentes
                    makedirs(listaDestino[i])

                #Copiando e Colando a pasta
                copytree(listaOrigem[i], listaDestino[i], dirs_exist_ok=True)

            #Verificando se a origem é arquivo
            elif(path.isfile(listaOrigem[i]) == True):

                #Verificando se o Destino existe
                if not path.isdir(listaDestino[i]):

                    #Criando as pastas não existentes
                    makedirs(listaDestino[i], exist_ok=True)

                #Copiando e Colando o arquivo
                copy(listaOrigem[i], listaDestino[i])

            #Condição para caso a origem não exista
            else:
                print(f"Origem Inexistente na linha {linha}")
                print(f"Origem: {listaOrigem[i]}")
                print('-' * 50)


        except Exception as e:
            print(f"Erro na execução Origem: {listaOrigem[i]} e Destino: {listaDestino[i]}")
            print(f"Provavelmente na linha {linha} do excel")
            print(e)
            print('-'*50)
        linha +=1

    input("Finalizado!!!")

autoTransfer()


