import sys
import os
import logging
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QColor
from interface import MainWindow
from database import db

# Configuração do diretório de logs
def setup_logging():
    try:
        # Usa o diretório AppData no Windows para os logs
        if getattr(sys, 'frozen', False):
            log_dir = os.path.join(os.environ['APPDATA'], 'MeuApp')
        else:
            log_dir = 'logs'
            
        # Cria o diretório se não existir
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        log_file = os.path.join(log_dir, 'app_log.txt')
        
        logging.basicConfig(
            filename=log_file,
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        return True
    except Exception as e:
        print(f"Erro ao configurar logging: {str(e)}")
        return False

def resource_path(relative_path):
    """ Obtém o caminho absoluto para recursos empacotados """
    try:
        # PyInstaller cria um temp folder e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    try:
        if not setup_logging():
            sys.exit(1)
            
        logging.info('Iniciando aplicação')
        app = QApplication(sys.argv)
        
        # Define o diretório de trabalho atual
        if getattr(sys, 'frozen', False):
            # Se estiver rodando como executável
            executable_dir = os.path.dirname(sys.executable)
            logging.info(f'Diretório do executável: {executable_dir}')
            os.chdir(executable_dir)
        
        logging.info('Criando banco de dados')
        if not db.criar_banco_dados():
            logging.error("Erro ao criar banco de dados!")
            sys.exit(1)
            
        logging.info('Iniciando janela principal')
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.exception("Erro não tratado:")
        raise