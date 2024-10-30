from os.path import isdir

from pandas import read_excel
from pandas.io.clipboard import paste

from isRunningWhere import acharCaminhoExecutado
from shutil import copy, copytree
from os import path, makedirs

file = read_excel(acharCaminhoExecutado('./Excel/AutoTransfer.xlsx'))

listaOrigem = file['Origem']
listaDestino = file['Destino']

def autoTransfer():

    if(len(listaDestino) != len(listaOrigem)):
        input("A quantidade de Orige está diferente da quantidade de Destino\nVerifique se cada origem tem seu destino")
        return
    for i in range(len(listaDestino)):
        try:
            #Verificando se a origem é pasta
            if(path.isdir(listaOrigem[i]) == True):
                #Verificando se o destino existe
                if(path.isdir(listaDestino[i]) == True):
                    copytree(listaOrigem[i], listaDestino[i])
                else:
                    #Criando as pastas não existentes
                    makedirs(listaDestino[i])
                    copytree(listaOrigem[i], listaDestino[i])

            #Verificando se a origem é arquivo
            elif(path.isfile(listaOrigem[i]) == True):
                #Verificando se o Destino existe
                if(path.isdir(listaDestino[i]) == True):
                    copy(listaOrigem[i], listaDestino[i])
                else:
                    #Criando as pastas não existentes
                    makedirs(listaDestino[i])
                    copy(listaOrigem[i], listaDestino[i])

            #Condição para caso a origem não exista
            else:
                print(f"Origem Inexistente: {listaOrigem[i]}")

        except Exception as e:
            print(f"Erro na execução Origem: {listaOrigem[i]} e Destino: {listaDestino[i]}")
            print(e)
            print('-'*30)

    input("Finalizado!!!")



autoTransfer()


