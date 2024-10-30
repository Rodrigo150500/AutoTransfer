import sys
import os

def acharCaminhoExecutado(path):
    # Verifica se o programa está empacotado com PyInstaller ou rodando na IDE
    if getattr(sys, 'frozen', False):
        # Empacotado: a pasta de destino está no mesmo diretório do executável

        return os.path.abspath(os.path.join(os.path.dirname(sys.executable), path))
    else:
        # Executado na IDE: subindo dois níveis e acessando a pasta de destino

        return os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), f'{path}'))