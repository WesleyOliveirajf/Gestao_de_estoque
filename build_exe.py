import PyInstaller.__main__
import os

# Obter o diretório atual
current_dir = os.path.dirname(os.path.abspath(__file__))

# Configurar os caminhos dos arquivos
main_script = os.path.join(current_dir, 'main.py')

PyInstaller.__main__.run([
    'main.py',
    '--name=TorpControl',
    '--onefile',
    '--windowed',
    '--add-data=assets;assets',  # Incluir pasta assets
    '--hidden-import=pandas',
    '--hidden-import=sqlite3',
    '--hidden-import=reportlab',
    '--clean',
    '--noconfirm',
]) 