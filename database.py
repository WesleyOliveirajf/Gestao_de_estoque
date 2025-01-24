import os
import sqlite3
from datetime import datetime, timedelta
import sys

class DatabaseManager:
    def __init__(self):
        self.db_path = self.get_database_path()
        self.criar_banco_dados()

    def get_database_path(self):
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(application_path, 'database.db')

    def connect(self):
        try:
            return sqlite3.connect(self.db_path)
        except Exception as e:
            print(f"Erro ao conectar ao banco de dados: {e}")
            return None

    def criar_banco_dados(self):
        try:
            conn = self.connect()
            if not conn:
                return False
            
            cursor = conn.cursor()
            
            # Verifica se a tabela existe
            cursor.execute('''
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name='produtos'
            ''')
            
            if cursor.fetchone() is None:
                # Cria a tabela com a nova estrutura
                cursor.execute('''
                    CREATE TABLE produtos (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nome TEXT NOT NULL,
                        lote TEXT NOT NULL,
                        ca INTEGER,
                        quantidade INTEGER NOT NULL,
                        data_compra DATE,
                        data_fabricacao DATE,
                        data_validade DATE,
                        dias_restantes INTEGER,
                        status TEXT
                    )
                ''')
                conn.commit()
                return True
            return True
            
        except Exception as e:
            print(f"Erro ao criar banco de dados: {e}")
            return False
        finally:
            if conn:
                conn.close()

    def adicionar_produto(self, nome, lote, ca, quantidade, data_compra, data_fabricacao, data_validade):
        try:
            conn = self.connect()
            if not conn:
                return False
            
            cursor = conn.cursor()
            
            # Converter datas, permitindo valores vazios
            data_compra_obj = datetime.strptime(data_compra, '%d/%m/%Y') if data_compra else None
            data_fabricacao_obj = datetime.strptime(data_fabricacao, '%d/%m/%Y') if data_fabricacao else None
            data_validade_obj = datetime.strptime(data_validade, '%d/%m/%Y') if data_validade else None
            
            # Calcular dias restantes, se a data de validade for fornecida
            dias_restantes = (data_validade_obj.date() - datetime.now().date()).days if data_validade_obj else None
            
            # Determinar status
            status = "Normal"
            if dias_restantes is not None:
                if dias_restantes < 0:
                    status = "Vencido"
                elif dias_restantes <= 30:
                    status = "Próximo do Vencimento"
            
            cursor.execute('''
                INSERT INTO produtos (
                    nome, lote, ca, quantidade, data_compra, 
                    data_fabricacao, data_validade, dias_restantes, status
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                nome, lote, ca, quantidade,
                data_compra_obj.strftime('%Y-%m-%d') if data_compra_obj else None,
                data_fabricacao_obj.strftime('%Y-%m-%d') if data_fabricacao_obj else None,
                data_validade_obj.strftime('%Y-%m-%d') if data_validade_obj else None,
                dias_restantes,
                status
            ))
            
            conn.commit()
            return True
            
        except Exception as e:
            print(f"Erro ao adicionar produto: {e}")
            return False
        finally:
            if conn:
                conn.close()

    def atualizar_produto(self, id, nome, lote, ca, quantidade, data_compra, data_fabricacao, data_validade):
        try:
            conn = self.connect()
            if not conn:
                return False
            
            cursor = conn.cursor()
            
            # Converter datas para o formato do banco de dados
            data_compra_obj = datetime.strptime(data_compra, '%d/%m/%Y').strftime('%Y-%m-%d') if data_compra else None
            data_fabricacao_obj = datetime.strptime(data_fabricacao, '%d/%m/%Y').strftime('%Y-%m-%d') if data_fabricacao else None
            data_validade_obj = datetime.strptime(data_validade, '%d/%m/%Y').strftime('%Y-%m-%d') if data_validade else None
            
            # Calcular dias restantes e status
            dias_restantes = None
            status = "Normal"
            if data_validade_obj:
                data_validade_date = datetime.strptime(data_validade_obj, '%Y-%m-%d').date()
                dias_restantes = (data_validade_date - datetime.now().date()).days
                if dias_restantes < -9152:
                    status = "Vencido"
                elif dias_restantes <= 30:
                    status = "Próximo do Vencimento"
            
            cursor.execute('''
                UPDATE produtos 
                SET nome = ?, 
                    lote = ?, 
                    ca = ?, 
                    quantidade = ?, 
                    data_compra = ?, 
                    data_fabricacao = ?, 
                    data_validade = ?,
                    dias_restantes = ?,
                    status = ?
                WHERE id = ?
            ''', (nome, lote, ca, quantidade, data_compra_obj, data_fabricacao_obj, data_validade_obj, 
                  dias_restantes, status, id))
            
            conn.commit()
            return True
            
        except Exception as e:
            print(f"Erro ao atualizar produto: {e}")
            return False
        finally:
            if conn:
                conn.close()

    def carregar_produtos(self):
        try:
            conn = self.connect()
            if not conn:
                return []
            
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM produtos')
            return cursor.fetchall()
            
        except Exception as e:
            print(f"Erro ao carregar produtos: {e}")
            return []
        finally:
            if conn:
                conn.close()

    def get_produto(self, produto_id):
        try:
            conn = self.connect()
            if not conn:
                return None
            
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM produtos WHERE id = ?', (produto_id,))
            return cursor.fetchone()
            
        except Exception as e:
            print(f"Erro ao buscar produto: {e}")
            return None
        finally:
            if conn:
                conn.close()

    def excluir_produto(self, produto_id):
        try:
            conn = self.connect()
            if not conn:
                return False
            
            cursor = conn.cursor()
            cursor.execute('DELETE FROM produtos WHERE id = ?', (produto_id,))
            conn.commit()
            return True
            
        except Exception as e:
            print(f"Erro ao excluir produto: {e}")
            return False
        finally:
            if conn:
                conn.close()

    def filtrar_por_status(self, status):
        try:
            conn = self.connect()
            if not conn:
                return []
            
            cursor = conn.cursor()
            # Query SQL para filtrar por status
            cursor.execute('''
                SELECT id, nome, lote, ca, quantidade, 
                       data_compra, data_fabricacao, data_validade, 
                       val_dias, dias_rest, status
                FROM produtos 
                WHERE status = ?
            ''', (status,))
            
            return cursor.fetchall()
            
        except Exception as e:
            print(f"Erro ao filtrar por status: {e}")
            return []
        finally:
            if conn:
                conn.close()

# Criar instância global do DatabaseManager
db = DatabaseManager() 