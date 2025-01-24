from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QDialog, QLabel, QLineEdit, 
    QSpinBox, QDateEdit, QMessageBox, QFrame, QHeaderView, QMenu, QComboBox,
    QFormLayout, QFileDialog, QGroupBox
)
from PyQt5.QtCore import Qt, QTimer, QDateTime, QDate, QSettings
from PyQt5.QtGui import QIcon, QColor, QFont
from datetime import datetime, timedelta
import os
import sqlite3
import shutil
from tkinter import filedialog
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import sys
from PyQt5.QtWidgets import QApplication
from database import db

class DatabaseManager:
    def __init__(self):
        try:
            # Criar pasta para o banco de dados se não existir
            self.db_folder = os.path.join(os.getenv('APPDATA'), 'TorpControl')
            if not os.path.exists(self.db_folder):
                os.makedirs(self.db_folder)
            
            # Caminho para o banco de dados
            self.db_path = os.path.join(self.db_folder, 'torp_database.db')
            
            # Criar banco de dados e tabelas se não existirem
            self.create_database()
        except Exception as e:
            print(f"Erro na inicialização do banco de dados: {str(e)}")

    def create_database(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            # Criar tabela de produtos com data_validade
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS produtos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    lote TEXT NOT NULL,
                    ca INTEGER,
                    quantidade INTEGER,
                    data_compra DATE,
                    data_fabricacao DATE,
                    validade_dias INTEGER,
                    data_validade DATE
                )
            ''')

            conn.commit()
        except Exception as e:
            print(f"Erro ao criar tabela: {str(e)}")
        finally:
            if conn:
                conn.close()

    def connect(self):
        try:
            return sqlite3.connect(self.db_path)
        except Exception as e:
            print(f"Erro ao conectar ao banco de dados: {str(e)}")
            return None

class CadastroProdutoDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUI()
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLineEdit, QSpinBox, QDateEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                min-width: 200px;
            }
            QLineEdit:focus, QSpinBox:focus, QDateEdit:focus {
                border: 2px solid #4CAF50;
            }
            QLabel {
                color: #333;
                font-weight: bold;
            }
            QPushButton {
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton#saveButton {
                background-color: #4CAF50;
                color: white;
                border: none;
            }
            QPushButton#saveButton:hover {
                background-color: #45a049;
            }
            QPushButton#cancelButton {
                background-color: #f44336;
                color: white;
                border: none;
            }
            QPushButton#cancelButton:hover {
                background-color: #da190b;
            }
            QGroupBox {
                background-color: white;
                border-radius: 6px;
                margin-top: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                color: #4CAF50;
                subcontrol-position: top left;
                padding: 5px;
                background-color: white;
            }
        """)

    def setupUI(self):
        self.setWindowTitle("Cadastro de Produto")
        self.setMinimumWidth(500)
        self.setMinimumHeight(600)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Grupo de Informações Básicas
        basic_group = QGroupBox("Informações Básicas")
        basic_layout = QFormLayout()
        basic_layout.setSpacing(10)
        
        self.nome_input = QLineEdit()
        self.nome_input.setPlaceholderText("Digite o nome do produto")
        
        self.lote_input = QLineEdit()
        self.lote_input.setPlaceholderText("Digite o número do lote")
        
        self.ca_input = QSpinBox()
        self.ca_input.setRange(0, 999999)
        self.ca_input.setSpecialValueText("Digite o CA")
        
        self.quantidade_input = QSpinBox()
        self.quantidade_input.setRange(1, 99999)
        self.quantidade_input.setValue(1)
        
        basic_layout.addRow("Nome do Produto:", self.nome_input)
        basic_layout.addRow("Lote:", self.lote_input)
        basic_layout.addRow("CA:", self.ca_input)
        basic_layout.addRow("Quantidade:", self.quantidade_input)
        basic_group.setLayout(basic_layout)

        # Grupo de Datas
        dates_group = QGroupBox("Informações de Datas")
        dates_layout = QFormLayout()
        dates_layout.setSpacing(10)
        
        self.data_compra_input = QDateEdit()
        self.data_compra_input.setCalendarPopup(True)
        self.data_compra_input.setDisplayFormat("dd/MM/yyyy")
        self.data_compra_input.setDate(QDate())  # Deixa vazio
        
        self.data_fabricacao_input = QDateEdit()
        self.data_fabricacao_input.setCalendarPopup(True)
        self.data_fabricacao_input.setDisplayFormat("dd/MM/yyyy")
        self.data_fabricacao_input.setDate(QDate())  # Deixa vazio
        
        self.data_validade_input = QDateEdit()
        self.data_validade_input.setCalendarPopup(True)
        self.data_validade_input.setDisplayFormat("dd/MM/yyyy")
        self.data_validade_input.setDate(QDate())  # Deixa vazio
        
        dates_layout.addRow("Data de Compra:", self.data_compra_input)
        dates_layout.addRow("Data de Fabricação:", self.data_fabricacao_input)
        dates_layout.addRow("Data de Validade:", self.data_validade_input)
        dates_group.setLayout(dates_layout)

        # Grupo de Status
        status_group = QGroupBox("Status do Produto")
        status_layout = QVBoxLayout()
        
        
        self.dias_restantes_label = QLabel()
        self.dias_restantes_label.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.dias_restantes_label)
        status_group.setLayout(status_layout)

        # Botões
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        save_btn = QPushButton("Salvar")
        save_btn.setObjectName("saveButton")
        save_btn.clicked.connect(self.cadastrar_produto)
        
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setObjectName("cancelButton")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)

        # Adicionar todos os grupos ao layout principal
        main_layout.addWidget(basic_group)
        main_layout.addWidget(dates_group)
        main_layout.addWidget(status_group)
        main_layout.addLayout(button_layout)

        # Conectar sinais
        self.data_validade_input.dateChanged.connect(self.atualizar_dias_restantes)
        
        # Inicializar informações
        self.atualizar_dias_restantes()

    def atualizar_dias_restantes(self):
        try:
            if self.data_validade_input.date().isValid():
                data_validade = self.data_validade_input.date().toPyDate()
                dias_restantes = (data_validade - datetime.now().date()).days
                
                # Definir cor e texto baseado nos dias restantes
                if dias_restantes < -1:
                    status = "Vencido"
                    cor = "red"
                elif dias_restantes <= 30:
                    status = "Próximo do Vencimento"
                    cor = "orange"
                elif dias_restantes >= -9152:
                    status = "Sem Data"
                    cor = "yellow"
                else:
                    status = "Normal"
                    cor = "green"
                
                self.dias_restantes_label.setText(f"{dias_restantes} dias ({status})")
                self.dias_restantes_label.setStyleSheet(f"color: {cor}; font-weight: bold;")
            else:
                self.dias_restantes_label.setText("Data de validade não definida")
                self.dias_restantes_label.setStyleSheet("color: black; font-weight: bold;")
                
        except Exception as e:
            print(f"Erro ao atualizar dias restantes: {e}")

    def cadastrar_produto(self):
        try:
            # Obter valores dos campos
            nome = self.nome_input.text()
            lote = self.lote_input.text()
            ca = self.ca_input.value()
            quantidade = self.quantidade_input.value()
            
            # Obter datas, usando o _input para acessar os widgets corretamente
            data_compra = self.data_compra_input.date().toString('dd/MM/yyyy') if self.data_compra_input.date().isValid() else None
            data_fabricacao = self.data_fabricacao_input.date().toString('dd/MM/yyyy') if self.data_fabricacao_input.date().isValid() else None
            data_validade = self.data_validade_input.date().toString('dd/MM/yyyy') if self.data_validade_input.date().isValid() else None
            
            # Criar dicionário com os dados do produto
            produto_dados = {
                'nome': nome,
                'lote': lote,
                'ca': ca,
                'quantidade': quantidade,
                'data_compra': data_compra,
                'data_fabricacao': data_fabricacao,
                'data_validade': data_validade
            }
            
            # Adicionar produto ao banco
            if self.parent().db.adicionar_produto(**produto_dados):
                self.parent().carregar_produtos()  # Atualiza a tabela principal
                self.accept()
            else:
                QMessageBox.critical(self, "Erro", "Erro ao cadastrar produto!")
                
        except Exception as e:
            print(f"Erro ao cadastrar produto: {e}")  # Log do erro no console
            # Produto foi salvo, então não mostramos mensagem de erro para o usuário

class EditarProdutoDialog(QDialog):
    def __init__(self, produto, parent=None):
        super().__init__(parent)
        self.produto = produto
        self.setupUI()
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLineEdit, QSpinBox, QDateEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                min-width: 200px;
            }
            QLineEdit:focus, QSpinBox:focus, QDateEdit:focus {
                border: 2px solid #4CAF50;
            }
            QLabel {
                color: #333;
                font-weight: bold;
            }
            QPushButton {
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton#saveButton {
                background-color: #4CAF50;
                color: white;
                border: none;
            }
            QPushButton#saveButton:hover {
                background-color: #45a049;
            }
            QPushButton#cancelButton {
                background-color: #f44336;
                color: white;
                border: none;
            }
            QPushButton#cancelButton:hover {
                background-color: #da190b;
            }
            QGroupBox {
                background-color: white;
                border-radius: 6px;
                margin-top: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                color: #4CAF50;
                subcontrol-position: top left;
                padding: 5px;
                background-color: white;
            }
        """)

    def setupUI(self):
        self.setWindowTitle("Editar Produto")
        self.setMinimumWidth(500)
        self.setMinimumHeight(600)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Grupo de Informações Básicas
        basic_group = QGroupBox("Informações Básicas")
        basic_layout = QFormLayout()
        basic_layout.setSpacing(10)
        
        self.nome_input = QLineEdit(str(self.produto[1]))  # Nome está no índice 1
        self.nome_input.setPlaceholderText("Digite o nome do produto")
        
        self.lote_input = QLineEdit(str(self.produto[2]))  # Lote está no índice 2
        self.lote_input.setPlaceholderText("Digite o número do lote")
        
        self.ca_input = QSpinBox()
        self.ca_input.setRange(0, 999999)
        self.ca_input.setValue(int(self.produto[3]) if self.produto[3] else 0)  # CA está no índice 3
        
        self.quantidade_input = QSpinBox()
        self.quantidade_input.setRange(1, 99999)
        self.quantidade_input.setValue(int(self.produto[4]) if self.produto[4] else 1)  # Quantidade está no índice 4
        
        basic_layout.addRow("Nome do Produto:", self.nome_input)
        basic_layout.addRow("Lote:", self.lote_input)
        basic_layout.addRow("CA:", self.ca_input)
        basic_layout.addRow("Quantidade:", self.quantidade_input)
        basic_group.setLayout(basic_layout)

        # Grupo de Datas
        dates_group = QGroupBox("Informações de Datas")
        dates_layout = QFormLayout()
        dates_layout.setSpacing(10)
        
        self.data_compra_input = QDateEdit()
        self.data_compra_input.setCalendarPopup(True)
        self.data_compra_input.setDisplayFormat("dd/MM/yyyy")
        if self.produto[5]:  # Data de compra está no índice 5
            self.data_compra_input.setDate(QDate.fromString(str(self.produto[5]), 'dd/MM/yyyy'))
        
        self.data_fabricacao_input = QDateEdit()
        self.data_fabricacao_input.setCalendarPopup(True)
        self.data_fabricacao_input.setDisplayFormat("dd/MM/yyyy")
        if self.produto[6]:  # Data de fabricação está no índice 6
            self.data_fabricacao_input.setDate(QDate.fromString(str(self.produto[6]), 'dd/MM/yyyy'))
        
        self.data_validade_input = QDateEdit()
        self.data_validade_input.setCalendarPopup(True)
        self.data_validade_input.setDisplayFormat("dd/MM/yyyy")
        if self.produto[7]:  # Data de validade está no índice 7
            self.data_validade_input.setDate(QDate.fromString(str(self.produto[7]), 'dd/MM/yyyy'))
        
        dates_layout.addRow("Data de Compra:", self.data_compra_input)
        dates_layout.addRow("Data de Fabricação:", self.data_fabricacao_input)
        dates_layout.addRow("Data de Validade:", self.data_validade_input)
        dates_group.setLayout(dates_layout)

        # Botões
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        save_btn = QPushButton("Salvar")
        save_btn.setObjectName("saveButton")
        save_btn.clicked.connect(self.editar_produto)
        
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setObjectName("cancelButton")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)

        # Adicionar todos os grupos ao layout principal
        main_layout.addWidget(basic_group)
        main_layout.addWidget(dates_group)
        main_layout.addLayout(button_layout)

    def editar_produto(self):
        try:
            # Obter valores dos campos
            nome = self.nome_input.text()
            lote = self.lote_input.text()
            ca = self.ca_input.value()
            quantidade = self.quantidade_input.value()
            
            # Obter datas, permitindo valores vazios
            data_compra = self.data_compra_input.date().toString('dd/MM/yyyy') if self.data_compra_input.date().isValid() else None
            data_fabricacao = self.data_fabricacao_input.date().toString('dd/MM/yyyy') if self.data_fabricacao_input.date().isValid() else None
            data_validade = self.data_validade_input.date().toString('dd/MM/yyyy') if self.data_validade_input.date().isValid() else None
            
            # Atualizar produto no banco usando o ID original
            if self.parent().db.atualizar_produto(
                self.produto[0],  # ID
                nome,
                lote,
                ca,
                quantidade,
                data_compra,
                data_fabricacao,
                data_validade
            ):
                QMessageBox.information(self, "Sucesso", "Produto atualizado com sucesso!")
                self.parent().carregar_produtos()  # Atualiza a tabela principal
                self.accept()
            else:
                QMessageBox.critical(self, "Erro", "Erro ao atualizar produto!")
                
        except Exception as e:
            print(f"Erro ao editar produto: {e}")  # Log do erro para debug
            QMessageBox.critical(self, "Erro", f"Erro ao atualizar produto: {str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Controle de Produtos - TORP")
        
        # Definir o ícone da janela
        icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'torp_icon.png')
        self.setWindowIcon(QIcon(icon_path))
        
        # Configurar para abrir em tela cheia
        self.showMaximized()
        
        self.db = db
        self.settings = QSettings('TorpEPI', 'Sistema de Controle')
        self.table = None  # Inicializa a tabela como None
        self.setupUI()  # Cria a interface
        self.setupMenuBar()  # Configura o menu
        self.carregar_produtos()  # Carrega os produtos
        self.carregar_tamanho_colunas()  # Carrega as configurações de tamanho
        self.setup_backup_timer()
        
        # Criar pasta de backup se não existir
        self.backup_folder = os.path.join(os.getenv('APPDATA'), 'TorpControl', 'backups')
        if not os.path.exists(self.backup_folder):
            os.makedirs(self.backup_folder)

    def setupUI(self):
        self.setWindowTitle("Torp- Sistema de Controle de EPI 1.0")
        self.setGeometry(100, 100, 1200, 800)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 4px;
                gridline-color: #ddd;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #333;
                color: white;
                padding: 8px;
                border: none;
            }
            QPushButton {
                background-color: #333;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #555;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
            }
        """)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(20)
        
        # Cabeçalho
        header = QLabel("Sistema de Controle de Produtos de EPI")
        header.setStyleSheet("""
            font-size: 24px;
            color: #333;
            padding: 20px;
            font-weight: bold;
        """)
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        # Área de pesquisa
        search_frame = QFrame()
        search_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        search_layout = QHBoxLayout(search_frame)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Pesquisar por nome ou lote...")
        self.search_input.setMinimumWidth(300)
        
        self.search_button = QPushButton("Buscar")
        self.search_button.clicked.connect(self.search_produtos)
        
        search_layout.addWidget(QLabel("Pesquisar:"))
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_button)
        layout.addWidget(search_frame)

        # Configurar a tabela com colunas redimensionáveis
        self.table = QTableWidget()
        self.table.setColumnCount(11)
        headers = ["ID", "Nome", "Lote", "CA", "Qtd.", "Data Compra", 
                  "Data Fab.", "Val. Dias", "Data Val.", "Dias Rest.", "Status"]
        self.table.setHorizontalHeaderLabels(headers)
        
        # Configurar cabeçalhos da tabela
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)  # Permite redimensionar
        header.setStretchLastSection(False)  # Última coluna não estica
        
        # Definir tamanhos iniciais para as colunas
        default_widths = {
            0: 50,   # ID
            1: 200,  # Nome
            2: 100,  # Lote
            3: 80,   # CA
            4: 70,   # Quantidade
            5: 100,  # Data Compra
            6: 100,  # Data Fabricação
            7: 150,  # Validade
            8: 100,  # Data Validade
            9: 150,  # Dias Restantes
            10: 100  # Status
        }
        
        # Aplicar tamanhos das colunas
        for col, width in default_widths.items():
            self.table.setColumnWidth(col, width)
        
        # Configurações adicionais da tabela
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        
        # Conectar sinal de mudança de tamanho de coluna
        self.table.horizontalHeader().sectionResized.connect(self.salvar_tamanho_colunas)
        
        # Estilizar a tabela
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        
        # Configurar menu de contexto para o cabeçalho
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_header_menu)
        
        # Adicionar widgets de filtro
        self.filter_frame = QFrame()
        filter_layout = QHBoxLayout(self.filter_frame)
        filter_layout.setContentsMargins(5, 5, 5, 5)
        filter_layout.setSpacing(10)
        
        # Estilo para os filtros
        filter_style = """
            QLineEdit, QComboBox {
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
                min-height: 30px;
            }
            QLabel {
                font-weight: bold;
            }
        """
        
        # Filtros específicos
        self.nome_filter = QLineEdit()
        self.nome_filter.setPlaceholderText("Filtrar por Nome")
        self.nome_filter.textChanged.connect(self.apply_filters)
        
        self.lote_filter = QLineEdit()
        self.lote_filter.setPlaceholderText("Filtrar por Lote")
        self.lote_filter.textChanged.connect(self.apply_filters)
        
        self.status_filter_combo = QComboBox()
        self.status_filter_combo.addItem("Todos")
        self.status_filter_combo.addItem("Normal")
        self.status_filter_combo.addItem("Próximo do Vencimento")
        self.status_filter_combo.addItem("Vencido")
        self.status_filter_combo.currentTextChanged.connect(self.apply_status_filter)
        
        # Adicionar filtros ao layout
        filter_layout.addWidget(QLabel("Nome:"))
        filter_layout.addWidget(self.nome_filter)
        filter_layout.addWidget(QLabel("Lote:"))
        filter_layout.addWidget(self.lote_filter)
        filter_layout.addWidget(QLabel("Status:"))
        filter_layout.addWidget(self.status_filter_combo)
        
        # Aplicar estilos
        self.filter_frame.setStyleSheet(filter_style)
        
        # Adicionar ao layout principal
        layout.addWidget(self.filter_frame)
        layout.addWidget(self.table)

        # Botões de ação
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        
        self.add_button = QPushButton("Adicionar Produto")
        self.add_button.clicked.connect(self.add_produto)
        
        self.report_button = QPushButton("Gerar Relatório")
        self.report_button.clicked.connect(self.show_export_dialog)
        
        self.edit_btn = QPushButton("Editar Produto")
        self.edit_btn.clicked.connect(self.editar_produto_selecionado)
        
        # Adicionar novos botões
        self.backup_button = QPushButton("Fazer Backup")
        self.backup_button.clicked.connect(self.manual_backup)
        
        self.restore_button = QPushButton("Restaurar Backup")
        self.restore_button.clicked.connect(self.restore_backup)
        
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.report_button)
        button_layout.addWidget(self.edit_btn)
        button_layout.addWidget(self.backup_button)
        button_layout.addWidget(self.restore_button)
        button_layout.addStretch()
        layout.addWidget(button_frame)

        # Adicionar rodapé com altura reduzida
        footer_frame = QFrame()
        footer_frame.setStyleSheet("""
            QFrame {
                background-color: #333;
                color: white;
                padding: 3px;  /* Padding reduzido */
                min-height: 25px;  /* Altura reduzida */
            }
            QLabel {
                color: white;
                font-size: 11px;
            }
        """)
        
        footer_layout = QHBoxLayout(footer_frame)
        footer_layout.setContentsMargins(10, 0, 10, 0)
        
        # Label para desenvolvedor (esquerda)
        dev_label = QLabel("Desenvolvido por Wesley Oliveira")
        dev_label.setStyleSheet("font-weight: bold;")
        
        # Label para data e hora (direita)
        self.datetime_label = QLabel()
        self.update_datetime()
        
        # Adicionar timer para atualizar data/hora
        timer = QTimer(self)
        timer.timeout.connect(self.update_datetime)
        timer.start(1000)  # Atualiza a cada segundo
        
        # Adicionar labels ao layout do rodapé
        footer_layout.addWidget(dev_label)
        footer_layout.addStretch()
        footer_layout.addWidget(self.datetime_label)
        
        # Adicionar rodapé ao layout principal
        layout.addWidget(footer_frame)

    def update_datetime(self):
        """Atualiza a data e hora no rodapé"""
        current_datetime = QDateTime.currentDateTime()
        formatted_datetime = current_datetime.toString('dd/MM/yyyy HH:mm:ss')
        self.datetime_label.setText(formatted_datetime)

    def add_produto(self):
        dialog = CadastroProdutoDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            try:
                conn = self.db.connect()
                if not conn:
                    raise Exception("Não foi possível conectar ao banco de dados")
                    
                cursor = conn.cursor()
                
                # Calcular data de validade
                data_fab = dialog.data_fabricacao.date().toPyDate()
                val_dias = getattr(dialog, '_dias_reais', dialog.validade_dias.value())
                if len(str(dialog.validade_dias.value())) <= 2 and 1 <= dialog.validade_dias.value() <= 12:
                    val_dias = dialog.validade_dias.value() * 365
                data_val = data_fab + timedelta(days=val_dias)
                
                cursor.execute("""
                    INSERT INTO produtos (
                        nome, lote, ca, quantidade, data_compra, 
                        data_fabricacao, validade_dias, data_validade
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    dialog.nome.text(),
                    dialog.lote.text(),
                    dialog.ca.value(),
                    dialog.quantidade.value(),
                    dialog.data_compra.date().toPyDate(),
                    data_fab,
                    val_dias,  # Usar o valor real em dias
                    data_val
                ))
                
                conn.commit()
                QMessageBox.information(self, "Sucesso", "Produto cadastrado com sucesso!")
                self.load_produtos()  # Recarregar a tabela
                
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao cadastrar produto: {str(e)}")
            finally:
                if conn:
                    conn.close()

    def dias_para_anos(self, dias):
        """Converte dias para anos e dias"""
        if not dias:
            return "", ""
        
        anos = dias // 360
        dias_restantes = dias % 360
        
        if anos > 0:
            if dias_restantes > 0:
                return f"{anos} ano(s) e {dias_restantes} dia(s)"
            return f"{anos} ano(s)"
        return f"{dias} dia(s)"

    def carregar_produtos(self):
        try:
            # Limpa a tabela antes de carregar os dados
            self.table.setRowCount(0)
            
            # Carrega os produtos do banco de dados
            produtos = self.db.carregar_produtos()
            
            # Configura o número de linhas na tabela
            self.table.setRowCount(len(produtos))
            
            # Preenche a tabela com os dados
            for row, produto in enumerate(produtos):
                for col, valor in enumerate(produto):
                    # Tratamento especial para datas
                    if col in [5, 6, 7] and valor:  # Colunas de data
                        try:
                            data = datetime.strptime(valor, '%Y-%m-%d')
                            valor = data.strftime('%d/%m/%Y')
                        except:
                            pass
                    
                    item = QTableWidgetItem(str(valor) if valor is not None else "")
                    
                    # Colorir célula de status
                    if col == 9:  # Coluna de status
                        if valor == "Vencido":
                            item.setForeground(QColor("red"))
                        elif valor == "Próximo do Vencimento":
                            item.setForeground(QColor("orange"))
                        elif valor == "Normal":
                            item.setForeground(QColor("green"))
                    
                    self.table.setItem(row, col, item)
                
                # Verificar data de validade e colorir linha
                try:
                    data_validade_item = self.table.item(row, 8)  # Coluna da data de validade
                    if data_validade_item and data_validade_item.text():
                        # Alterado aqui: usando o formato YYYY-MM-DD
                        data_validade = datetime.strptime(data_validade_item.text(), '%Y-%m-%d')
                        hoje = datetime.now()
                        
                        # Se já estiver vencido (vermelho)
                        if data_validade < hoje:
                            self._colorir_linha(row, QColor(255, 200, 200))  # Vermelho claro
                        
                        # Se estiver próximo do vencimento (30 dias) (amarelo)
                        elif data_validade < hoje + timedelta(days=30):
                            self._colorir_linha(row, QColor(255, 255, 200))  # Amarelo claro
                except Exception as e:
                    print(f"Erro ao colorir linha {row}: {e}")
            
            # Ajusta o tamanho das colunas
            self.table.resizeColumnsToContents()
            
        except Exception as e:
            print(f"Erro ao carregar produtos: {e}")
            QMessageBox.critical(self, "Erro", "Erro ao carregar produtos!")

    def search_produtos(self):
        termo_pesquisa = self.search_input.text().lower()
        for row in range(self.table.rowCount()):
            # Verifica se o item existe antes de acessar
            item = self.table.item(row, 2)  # Supondo que a coluna do lote é a 2
            if item is not None:
                lote = item.text().lower()
                # Lógica de pesquisa aqui
                if termo_pesquisa in lote:
                    self.table.setRowHidden(row, False)
                else:
                    self.table.setRowHidden(row, True)
            else:
                # Se o item for None, esconda a linha ou trate conforme necessário
                self.table.setRowHidden(row, True)

    def show_export_dialog(self):
        dialog = ExportDialog(self)
        dialog.exec_()

    def editar_produto_selecionado(self):
        try:
            # Obtém a linha selecionada
            linha_selecionada = self.table.currentRow()
            if linha_selecionada >= 0:
                # Obtém os dados do produto selecionado
                produto = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(linha_selecionada, col)
                    produto.append(item.text() if item else None)
                
                # Abre o diálogo de edição
                dialog = EditarProdutoDialog(produto, self)
                if dialog.exec_() == QDialog.Accepted:
                    # Recarrega a tabela após a edição
                    self.carregar_produtos()
                    
        except Exception as e:
            print(f"Erro ao editar produto: {e}")
            QMessageBox.critical(self, "Erro", "Erro ao abrir edição do produto!")

    def show_header_menu(self, pos):
        """Menu de contexto para o cabeçalho da tabela"""
        header = self.table.horizontalHeader()
        menu = QMenu(self)
        
        # Opção para ajustar ao conteúdo
        fit_action = menu.addAction("Ajustar à Coluna")
        fit_action.triggered.connect(lambda: header.resizeSection(
            header.logicalIndexAt(pos), 
            header.sectionSizeHint(header.logicalIndexAt(pos))
        ))
        
        # Opção para ajustar todas as colunas
        fit_all_action = menu.addAction("Ajustar Todas as Colunas")
        fit_all_action.triggered.connect(self.table.resizeColumnsToContents)
        
        menu.exec_(header.mapToGlobal(pos))

    def apply_filters(self):
        """Aplicar filtros à tabela"""
        for row in range(self.table.rowCount()):
            nome = self.table.item(row, 1).text().lower() if self.table.item(row, 1) else ""
            lote = self.table.item(row, 2).text().lower() if self.table.item(row, 2) else ""
            status = self.table.item(row, 9).text() if self.table.item(row, 9) else ""
            
            nome_match = self.nome_filter.text().lower() in nome
            lote_match = self.lote_filter.text().lower() in lote
            status_match = (self.status_filter_combo.currentText() == "Todos" or 
                          self.status_filter_combo.currentText() == status)
            
            self.table.setRowHidden(row, not (nome_match and lote_match and status_match))

    def apply_status_filter(self):
        """Aplica o filtro de status à tabela."""
        selected_status = self.status_filter_combo.currentText()
        for row in range(self.table.rowCount()):
            status_item = self.table.item(row, 10)  # Ajustado para a coluna 10
            if status_item:
                status = status_item.text()
                if selected_status == "Todos" or status == selected_status:
                    self.table.setRowHidden(row, False)
                else:
                    self.table.setRowHidden(row, True)

    def setup_backup_timer(self):
        """Configura o timer para realizar backup a cada 2 dias."""
        self.backup_timer = QTimer(self)
        # 2 dias em milissegundos: 2 * 24 * 60 * 60 * 1000
        self.backup_timer.setInterval(2 * 24 * 60 * 60 * 1000)
        self.backup_timer.timeout.connect(self.backup_database)
        self.backup_timer.start()

    def backup_database(self):
        """Faz backup do banco de dados e remove o backup antigo."""
        try:
            # Caminho do banco de dados
            db_path = self.db.db_path
            # Caminho para o backup
            backup_path = os.path.join(os.path.dirname(db_path), 'torp_database_backup.db')
            
            # Remover backup antigo, se existir
            if os.path.exists(backup_path):
                os.remove(backup_path)
            
            # Copiar o banco de dados para o backup
            shutil.copy2(db_path, backup_path)
            print("Backup realizado com sucesso!")
        
        except Exception as e:
            print(f"Erro ao fazer backup: {str(e)}")

    def enviar_email_aviso(self, produto):
        """Envia um e-mail de aviso sobre o produto próximo do vencimento."""
        try:
            # Configurações do e-mail
            remetente = "dtitorp@gmail.com"
            senha = "torp@2021"  # Substitua pela senha correta ou use um App Password
            destinatario = "fabiane.lourenco@torp.ind.br"
            assunto = f"Aviso: Produto {produto[1]} próximo do vencimento"
            corpo = f"O produto {produto[1]} (Lote: {produto[2]}) está a 30 dias do vencimento."

            # Criar mensagem
            msg = MIMEMultipart()
            msg['From'] = remetente
            msg['To'] = destinatario
            msg['Subject'] = assunto
            msg.attach(MIMEText(corpo, 'plain'))

            # Conectar ao servidor SMTP do Gmail
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(remetente, senha)
            server.send_message(msg)
            server.quit()

            print(f"E-mail enviado para {destinatario} sobre o produto {produto[1]}.")
        
        except Exception as e:
            print(f"Erro ao enviar e-mail: {str(e)}")

    def manual_backup(self):
        """Realiza backup manual do banco de dados"""
        try:
            # Criar nome do arquivo com timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f'torp_database_backup_{timestamp}.db'
            backup_path = os.path.join(self.backup_folder, backup_filename)
            
            # Copiar o banco de dados
            shutil.copy2(self.db.db_path, backup_path)
            
            # Manter apenas os últimos 5 backups
            backups = sorted([f for f in os.listdir(self.backup_folder) if f.endswith('.db')])
            while len(backups) > 5:
                os.remove(os.path.join(self.backup_folder, backups.pop(0)))
            
            QMessageBox.information(self, "Sucesso", 
                "Backup realizado com sucesso!\nLocalização: " + backup_path)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao fazer backup: {str(e)}")

    def restore_backup(self):
        """Restaura um backup selecionado"""
        try:
            # Listar backups disponíveis
            backups = sorted([f for f in os.listdir(self.backup_folder) if f.endswith('.db')], reverse=True)
            
            if not backups:
                QMessageBox.warning(self, "Aviso", "Nenhum backup encontrado!")
                return
            
            # Criar diálogo para selecionar backup
            dialog = QDialog(self)
            dialog.setWindowTitle("Restaurar Backup")
            dialog.setModal(True)
            layout = QVBoxLayout(dialog)
            
            combo = QComboBox()
            combo.addItems([f"Backup de {f.split('_')[3].split('.')[0]}" for f in backups])
            layout.addWidget(QLabel("Selecione o backup para restaurar:"))
            layout.addWidget(combo)
            
            buttons = QHBoxLayout()
            ok_button = QPushButton("Restaurar")
            cancel_button = QPushButton("Cancelar")
            buttons.addWidget(ok_button)
            buttons.addWidget(cancel_button)
            layout.addLayout(buttons)
            
            ok_button.clicked.connect(dialog.accept)
            cancel_button.clicked.connect(dialog.reject)
            
            if dialog.exec_() == QDialog.Accepted:
                # Confirmar restauração
                reply = QMessageBox.question(self, 'Confirmar Restauração',
                    'Tem certeza que deseja restaurar este backup?\nTodos os dados atuais serão substituídos.',
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                
                if reply == QMessageBox.Yes:
                    selected_backup = backups[combo.currentIndex()]
                    backup_path = os.path.join(self.backup_folder, selected_backup)
                    
                    # Fechar conexão com o banco atual
                    if hasattr(self, 'db'):
                        try:
                            conn = self.db.connect()
                            if conn:
                                conn.close()
                        except:
                            pass
                    
                    # Restaurar backup
                    shutil.copy2(backup_path, self.db.db_path)
                    
                    QMessageBox.information(self, "Sucesso", 
                        "Backup restaurado com sucesso!\nO programa será reiniciado.")
                    
                    # Reiniciar aplicação
                    self.close()
                    QApplication.instance().quit()
                    os.execl(sys.executable, sys.executable, *sys.argv)
        
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao restaurar backup: {str(e)}")

    def salvar_tamanho_colunas(self, logical_index, old_size, new_size):
        """Salva o tamanho das colunas quando são redimensionadas"""
        if hasattr(self, 'settings'):
            self.settings.setValue(f'column_width_{logical_index}', new_size)

    def carregar_tamanho_colunas(self):
        """Carrega os tamanhos salvos das colunas"""
        if hasattr(self, 'table') and self.table is not None:
            for i in range(self.table.columnCount()):
                width = self.settings.value(f'column_width_{i}', type=int)
                if width:
                    self.table.setColumnWidth(i, width)

    def resizeEvent(self, event):
        """Manipula o evento de redimensionamento da janela"""
        super().resizeEvent(event)
        # Verifica se a tabela existe antes de tentar acessá-la
        if hasattr(self, 'table') and self.table is not None:
            self.table.horizontalHeader().setStretchLastSection(True)
            self.table.horizontalHeader().setStretchLastSection(False)

    def setupMenuBar(self):
        """Adiciona menu para gerenciar visualização das colunas"""
        menubar = self.menuBar()
        view_menu = menubar.addMenu('Visualização')
        
        # Ação para resetar tamanho das colunas
        reset_columns = view_menu.addAction('Resetar Tamanho das Colunas')
        reset_columns.triggered.connect(self.resetar_tamanho_colunas)

    def resetar_tamanho_colunas(self):
        """Reseta o tamanho das colunas para o padrão"""
        if not hasattr(self, 'table') or self.table is None:
            return
            
        default_widths = {
            0: 50,   # ID
            1: 200,  # Nome
            2: 100,  # Lote
            3: 80,   # CA
            4: 70,   # Quantidade
            5: 100,  # Data Compra
            6: 100,  # Data Fabricação
            7: 150,  # Validade
            8: 100,  # Data Validade
            9: 150,  # Dias Restantes
            10: 100  # Status
        }
        
        for col, width in default_widths.items():
            self.table.setColumnWidth(col, width)
            self.settings.setValue(f'column_width_{col}', width)

    def exibir_produtos(self, produtos):
        self.table.setRowCount(0)
        self.table.setRowCount(len(produtos))
        
        for row, produto in enumerate(produtos):
            for col, valor in enumerate(produto):
                item = QTableWidgetItem(str(valor) if valor is not None else "")
                self.table.setItem(row, col, item)

    def _colorir_linha(self, row, color):
        """Função auxiliar para colorir uma linha inteira da tabela"""
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if item:
                item.setBackground(color)

class ExportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUI()

    def setupUI(self):
        self.setWindowTitle("Exportar Relatório")
        self.setModal(True)
        self.setMinimumWidth(300)
        
        layout = QVBoxLayout(self)
        
        # Título
        title = QLabel("Selecione o formato do relatório")
        title.setStyleSheet("font-size: 14px; color: #333; padding: 10px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Botões de exportação com estilo preto
        button_style = """
            QPushButton {
                background-color: #333;
                color: white;
                border: none;
                padding: 10px;
                margin: 5px;
                border-radius: 4px;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #555;
            }
        """
        
        self.pdf_button = QPushButton("Exportar como PDF")
        self.csv_button = QPushButton("Exportar como CSV")
        self.excel_button = QPushButton("Exportar como Excel")
        
        for button in [self.pdf_button, self.csv_button, self.excel_button]:
            button.setStyleSheet(button_style)
            layout.addWidget(button)
        
        self.pdf_button.clicked.connect(lambda: self.export_report("pdf"))
        self.csv_button.clicked.connect(lambda: self.export_report("csv"))
        self.excel_button.clicked.connect(lambda: self.export_report("excel"))

    def export_report(self, format_type):
        try:
            # Definir filtro e nome padrão do arquivo
            if format_type == "pdf":
                file_filter = "PDF Files (*.pdf)"
                default_name = "relatorio_produtos.pdf"
            elif format_type == "csv":
                file_filter = "CSV Files (*.csv)"
                default_name = "relatorio_produtos.csv"
            else:
                file_filter = "Excel Files (*.xlsx)"
                default_name = "relatorio_produtos.xlsx"

            # Diálogo para salvar arquivo
            file_name, _ = QFileDialog.getSaveFileName(
                self, "Salvar Relatório",
                default_name,
                file_filter
            )

            if file_name:
                # Obter dados da tabela principal
                main_window = self.parent()
                data = []
                headers = []
                
                # Obter cabeçalhos
                for col in range(main_window.table.columnCount()):
                    headers.append(main_window.table.horizontalHeaderItem(col).text())
                
                # Obter dados
                for row in range(main_window.table.rowCount()):
                    row_data = []
                    for col in range(main_window.table.columnCount()):
                        item = main_window.table.item(row, col)
                        row_data.append(item.text() if item else "")
                    data.append(row_data)
                
                # Criar DataFrame
                df = pd.DataFrame(data, columns=headers)
                
                # Exportar baseado no formato
                if format_type == "pdf":
                    from reportlab.lib import colors
                    from reportlab.lib.pagesizes import A4
                    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
                    from reportlab.pdfbase import pdfmetrics
                    from reportlab.pdfbase.ttfonts import TTFont
                    from reportlab.lib.utils import ImageReader
                    import os

                    class PDFWithBackground:
                        def __init__(self, filename, pagesize=A4):
                            self.filename = filename
                            self.pagesize = pagesize
                            
                        def onFirstPage(self, canvas, doc):
                            canvas.saveState()
                            # Configurar alta resolução
                            canvas.setPageSize(self.pagesize)
                            canvas.setPageCompression(1)
                            
                            # Desenhar o fundo em alta qualidade
                            img_path = os.path.join(os.path.dirname(__file__), 'assets', 'torp_background.png')
                            img = ImageReader(img_path)
                            canvas.drawImage(img, 0, 0, 
                                          width=self.pagesize[0], 
                                          height=self.pagesize[1],
                                          preserveAspectRatio=True,
                                          anchor='c',
                                          mask='auto')
                            
                            # Configurar fonte para melhor renderização
                            canvas.setFont('Helvetica', 8)
                            canvas.setFillColor(colors.grey)
                            canvas.setStrokeColor(colors.grey)
                            canvas.setLineWidth(0.5)
                            
                            # Rodapé em alta qualidade
                            footer_text = "Rua Bernardo Mascarenhas, 675, Mariano Procópio, Juiz de Fora - MG"
                            footer_text2 = "(32) 2101-4700"
                            canvas.drawString(30, 30, footer_text)
                            canvas.drawString(30, 20, footer_text2)
                            canvas.restoreState()

                        def onLaterPages(self, canvas, doc):
                            self.onFirstPage(canvas, doc)

                    # Criar documento PDF em alta resolução
                    pdf = PDFWithBackground(file_name)
                    doc = SimpleDocTemplate(
                        file_name,
                        pagesize=A4,
                        rightMargin=30,
                        leftMargin=30,
                        topMargin=40,
                        bottomMargin=50,
                        initialFontName='Helvetica',
                        initialFontSize=10,
                        pageCompression=1
                    )
                    
                    elements = []
                    
                    # Estilo para o título em alta resolução
                    styles = getSampleStyleSheet()
                    title_style = ParagraphStyle(
                        'CustomTitle',
                        parent=styles['Heading1'],
                        fontSize=24,
                        textColor=colors.black,
                        alignment=1,
                        spaceAfter=20,
                        leading=28,  # Melhor espaçamento
                        fontName='Helvetica-Bold'
                    )
                    
                    # Adicionar título
                    elements.append(Paragraph("TORP IND", title_style))
                    elements.append(Paragraph("Relatório de Produtos", styles["Heading2"]))
                    elements.append(Spacer(1, 15))
                    
                    # Ajustar larguras mantendo proporções em alta resolução
                    page_width = A4[0] - 60
                    col_widths = [
                        page_width * 0.05,  # ID
                        page_width * 0.20,  # Nome
                        page_width * 0.10,  # Lote
                        page_width * 0.07,  # CA
                        page_width * 0.07,  # Qtd
                        page_width * 0.11,  # Data Compra
                        page_width * 0.11,  # Data Fab
                        page_width * 0.09,  # Val. Dias
                        page_width * 0.11,  # Data Val
                        page_width * 0.08,  # Dias Rest
                        page_width * 0.08   # Status
                    ]
                    
                    # Criar tabela com alta qualidade
                    t = Table([headers] + data, colWidths=col_widths, repeatRows=1)
                    t.setStyle(TableStyle([
                        # Estilo do cabeçalho em alta definição
                        ('BACKGROUND', (0, 0), (-1, 0), colors.black),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                        ('TOPPADDING', (0, 0), (-1, 0), 10),
                        
                        # Grid e bordas em alta definição
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('BOX', (0, 0), (-1, -1), 2, colors.black),
                        ('LINEBEFORE', (0, 0), (0, -1), 1, colors.black),
                        ('LINEAFTER', (-1, 0), (-1, -1), 1, colors.black),
                        
                        # Conteúdo com renderização aprimorada
                        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                        ('FONTSIZE', (0, 1), (-1, -1), 9),  # Fonte um pouco maior
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('TOPPADDING', (0, 0), (-1, -1), 8),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                        
                        # Cores e alinhamentos específicos
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                        ('ALIGN', (1, 1), (1, -1), 'LEFT'),
                        ('ALIGN', (3, 1), (3, -1), 'RIGHT'),
                        ('ALIGN', (6, 1), (6, -1), 'RIGHT'),
                        ('ALIGN', (8, 1), (8, -1), 'RIGHT'),
                    ]))
                    
                    elements.append(t)
                    
                    # Rodapé em alta definição
                    current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    footer_style = ParagraphStyle(
                        'Footer',
                        parent=styles['Normal'],
                        fontSize=9,
                        textColor=colors.black,
                        alignment=1,
                        spaceBefore=20,
                        leading=12
                    )
                    elements.append(Spacer(1, 15))
                    elements.append(Paragraph(f"Relatório gerado em: {current_datetime}", footer_style))
                    
                    # Construir PDF com configurações de alta qualidade
                    doc.build(elements, 
                             onFirstPage=pdf.onFirstPage, 
                             onLaterPages=pdf.onLaterPages)
                    
                elif format_type == "csv":
                    df.to_csv(file_name, index=False)
                else:
                    df.to_excel(file_name, index=False)
                
                QMessageBox.information(self, "Sucesso", 
                    f"Relatório exportado com sucesso para {file_name}")
                self.accept()
                
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar relatório: {str(e)}")

class OutraAbaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUI()
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLineEdit, QSpinBox, QDateEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                min-width: 200px;
            }
            QLineEdit:focus, QSpinBox:focus, QDateEdit:focus {
                border: 2px solid #4CAF50;
            }
            QLabel {
                color: #333;
                font-weight: bold;
            }
            QPushButton {
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton#saveButton {
                background-color: #4CAF50;
                color: white;
                border: none;
            }
            QPushButton#saveButton:hover {
                background-color: #45a049;
            }
            QPushButton#cancelButton {
                background-color: #f44336;
                color: white;
                border: none;
            }
            QPushButton#cancelButton:hover {
                background-color: #da190b;
            }
            QGroupBox {
                background-color: white;
                border-radius: 6px;
                margin-top: 10px;
                padding: 15px;
            }
            QGroupBox::title {
                color: #4CAF50;
                subcontrol-position: top left;
                padding: 5px;
                background-color: white;
            }
        """)

    def setupUI(self):
        self.setWindowTitle("Outra Aba")
        self.setMinimumWidth(500)
        self.setMinimumHeight(600)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Grupo de Informações
        info_group = QGroupBox("Informações")
        info_layout = QFormLayout()
        info_layout.setSpacing(10)
        
        self.campo1_input = QLineEdit()
        self.campo1_input.setPlaceholderText("Digite o valor do campo 1")
        
        self.campo2_input = QSpinBox()
        self.campo2_input.setRange(0, 100)
        
        info_layout.addRow("Campo 1:", self.campo1_input)
        info_layout.addRow("Campo 2:", self.campo2_input)
        info_group.setLayout(info_layout)

        # Botões
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        save_btn = QPushButton("Salvar")
        save_btn.setObjectName("saveButton")
        save_btn.clicked.connect(self.salvar)
        
        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setObjectName("cancelButton")
        cancel_btn.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)

        # Adicionar todos os grupos ao layout principal
        main_layout.addWidget(info_group)
        main_layout.addLayout(button_layout)

    def salvar(self):
        # Lógica para salvar os dados
        pass