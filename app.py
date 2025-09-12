from flask import Flask, Blueprint, session, request, flash, redirect, url_for, render_template, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
import os
from dotenv import load_dotenv
import logging
from datetime import datetime, timedelta
import chardet
import re
import sqlite3
import uuid
from werkzeug.utils import secure_filename
import pdfkit
from flask import send_from_directory
from sqlalchemy.exc import SQLAlchemyError
import calendar
from collections import Counter
import random
import shutil
from sqlalchemy import or_

# Carregar variáveis de ambiente
load_dotenv()

# Inicializar aplicação Flask
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY')


# Configurações de segurança da sessão
app.config.update(
    SESSION_COOKIE_SECURE=False,     # Só envia cookies via HTTPS - Desativado para teste (originalmente True)
    SESSION_COOKIE_HTTPONLY=True,   # Impede acesso via JavaScript
    SESSION_COOKIE_SAMESITE='Lax',  # Proteção contra CSRF
    PERMANENT_SESSION_LIFETIME=timedelta(minutes=30),  # Tempo de expiração
    SESSION_REFRESH_EACH_REQUEST=True,        # Renova o tempo de vida a cada requisição
    TEMPLATES_AUTO_RELOAD=True                # Recarrega templates em desenvolvimento

)

if not app.secret_key:
    raise ValueError("Defina FLASK_SECRET_KEY no arquivo .env")

# Configuração do banco de dados SQLite
DATABASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ComparasDB.db')
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DATABASE}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
DB_PATH_FORNECEDORES = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fornecedores.db')

# Configuração para upload de arquivos
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # Limite de 5MB para uploads
ALLOWED_EXTENSIONS = {'pdf'}

# Criar pasta de uploads se não existir
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
# Ensure ALLOWED_EXTENSIONS is defined correctly
ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png', 'pdf'}

def allowed_file(filename, allowed_extensions=None):
    if not filename or '.' not in filename:
        logging.warning(f"Invalid filename: {filename}")
        return False
    try:
        extension = filename.rsplit('.', 1)[1].lower()
        allowed = allowed_extensions or ALLOWED_EXTENSIONS
        logging.info(f"Checking extension: {extension} against allowed: {allowed}")
        return extension in allowed
    except IndexError:
        logging.warning(f"Filename has no valid extension: {filename}")
        return False

# Configuração do logger
log_file = 'app_errors.log'
log_dir = os.path.dirname(log_file)
if log_dir and not os.path.exists(log_dir):
    os.makedirs(log_dir)

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    encoding='utf-8'
)

# Definir o caminho do arquivo de senhas
SENHAS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "senhas.txt")

def log_error(message):
    logging.error(message)

# Inicializar SQLAlchemy
db = SQLAlchemy(app)

# Modelo para Kanban
class Kanban(db.Model):
    __tablename__ = 'Kanban'
    id = db.Column(db.Integer, primary_key=True)
    task_name = db.Column(db.Text, nullable=False)
    status = db.Column(db.Text, nullable=False)

# Modelo para Materiais
class Materiais(db.Model):
    __tablename__ = 'Materiais'
    CodMaterial = db.Column(db.Integer, primary_key=True, autoincrement=True)
    DescricaoMaterial = db.Column(db.Text, nullable=False)
    Empresa = db.Column(db.Text, nullable=False)
    Aplicacao = db.Column(db.Text, nullable=False)
    QuantidadeEstoque = db.Column(db.Integer, default=0)
    Fornecedor = db.Column(db.Text)
    NumeroNF = db.Column(db.Text)
    FatorConsumo = db.Column(db.Float, nullable=True, default=0.0)
    Ativo = db.Column(db.Boolean, default=False)

    def to_dict(self):
        return {
            'CodMaterial': self.CodMaterial,
            'DescricaoMaterial': self.DescricaoMaterial,
            'Empresa': self.Empresa,
            'Aplicacao': self.Aplicacao,
            'QuantidadeEstoque': self.QuantidadeEstoque,
            'Fornecedor': self.Fornecedor,
            'NumeroNF': self.NumeroNF,
            'FatorConsumo': self.FatorConsumo,
            'Ativo': self.Ativo
        }

# Modelo para Solicitações de Compra
class SolicitacoesCompra(db.Model):
    __tablename__ = 'SolicitacoesCompra'
    id = db.Column(db.Integer, primary_key=True)
    cod_material = db.Column(db.Integer, db.ForeignKey('Materiais.CodMaterial'), nullable=False)
    especificacao = db.Column(db.Text, nullable=False)
    quantidade = db.Column(db.Integer, nullable=False)
    unidade_medida = db.Column(db.String(20), nullable=False, default='Unidade')
    aplicacao = db.Column(db.Text, nullable=True)
    empresa = db.Column(db.Text, nullable=False)
    data_solicitacao = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    usuario = db.Column(db.Text, nullable=False)
    foto_path = db.Column(db.Text, nullable=True)
    marca = db.Column(db.Text, nullable=True)
    ativo = db.Column(db.String(3), nullable=False)
    nome_ativo = db.Column(db.Text, nullable=True)
    prioridade = db.Column(db.String(20), nullable=False, default='Programado')
    status_aprovacao = db.Column(db.String(20), nullable=True, default=None)
    observacoes_col = db.Column(db.Text, nullable=True)  # NOVA COLUNA AQUI
    material = db.relationship('Materiais', backref='solicitacoes')

    def to_dict(self):
        return {
            'id': self.id,
            'cod_material': self.cod_material,
            'especificacao': self.especificacao,
            'quantidade': self.quantidade,
            'unidade_medida': self.unidade_medida,
            'aplicacao': self.aplicacao,
            'empresa': self.empresa,
            'data_solicitacao': self.data_solicitacao.isoformat() if self.data_solicitacao else None,
            'usuario': self.usuario,
            'foto_path': self.foto_path,
            'marca': self.marca,
            'ativo': self.ativo,
            'nome_ativo': self.nome_ativo,
            'prioridade': self.prioridade,
            'status_aprovacao': self.status_aprovacao
        }

# Modelo para Solicitações Preenchidas
class SolicitacoesPreenchidas(db.Model):
    __tablename__ = 'SolicitacoesPreenchidas'
    id = db.Column(db.Integer, primary_key=True)
    solicitacao_id = db.Column(db.Integer, db.ForeignKey('SolicitacoesCompra.id'), nullable=False)
    fornecedor_id = db.Column(db.Integer, nullable=False)
    valor_unitario = db.Column(db.Float, nullable=False)
    valor_total = db.Column(db.Float, nullable=False)
    valor_frete = db.Column(db.Float, nullable=True)
    prazo_entrega = db.Column(db.Text, nullable=False)
    condicao_pagamento = db.Column(db.Text, nullable=False)
    data_preenchimento = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    usuario = db.Column(db.Text, nullable=False)
    status = db.Column(db.Text, nullable=True, default='Aguardando Aprovação')
    pdf_path = db.Column(db.Text, nullable=True)
    observacoes = db.Column(db.Text, nullable=True)  # Novo campo para observações
    # REMOVER ESTA LINHA: aprovacao_pdf_path = db.Column(db.Text, nullable=True)
    
    solicitacao = db.relationship('SolicitacoesCompra', backref='preenchimentos_fornecidos')
    
    historico_descontos = db.relationship(
        "HistoricoDescontos", 
        backref="solicitacao_preenchida",
        lazy="dynamic",
        cascade="all, delete-orphan"
    )

    def to_dict(self):
        return {
            'id': self.id,
            'solicitacao_id': self.solicitacao_id,
            'fornecedor_id': self.fornecedor_id,
            'valor_unitario': self.valor_unitario,
            'valor_total': self.valor_total,
            'valor_frete': self.valor_frete,
            'prazo_entrega': self.prazo_entrega,
            'condicao_pagamento': self.condicao_pagamento,
            'data_preenchimento': self.data_preenchimento.isoformat() if self.data_preenchimento else None,
            'usuario': self.usuario,
            'status': self.status,
            'pdf_path': self.pdf_path,
            'observacoes': self.observacoes,  # Adicionado ao dict
            'historico_descontos': [h.to_dict() for h in self.historico_descontos]
        }
    
def migrate_observacoes_col():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Verificar se a coluna já existe
        cursor.execute("PRAGMA table_info(SolicitacoesCompra)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'observacoes_col' not in columns:
            cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN observacoes_col TEXT")
            conn.commit()
            logging.info("Coluna observacoes_col adicionada à tabela SolicitacoesCompra")
            print("✓ Coluna observacoes_col adicionada com sucesso!")
        else:
            print("✓ Coluna observacoes_col já existe na tabela")
        
        conn.close()
        return True
    except sqlite3.Error as e:
        logging.error(f"Erro na migração observacoes_col: {str(e)}")
        print(f"✗ Erro na migração: {str(e)}")
        return False
    
# Modelo para Estoque
class Estoque(db.Model):
    __tablename__ = 'Estoque'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    preenchimento_id = db.Column(db.Integer, db.ForeignKey('SolicitacoesPreenchidas.id'), nullable=False)
    cod_material = db.Column(db.Integer, db.ForeignKey('Materiais.CodMaterial'), nullable=False)
    quantidade = db.Column(db.Integer, nullable=False)
    fornecedor = db.Column(db.Text, nullable=False)
    numero_nf = db.Column(db.Text, nullable=False)
    data_entrada = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    usuario = db.Column(db.Text, nullable=False)
    preenchimento = db.relationship('SolicitacoesPreenchidas', backref='estoque')
    material = db.relationship('Materiais', backref='estoque')

# Modelo para Auditoria
class Auditoria(db.Model):
    __tablename__ = 'Auditoria'
    id = db.Column(db.Integer, primary_key=True)
    solicitacao_id = db.Column(db.Integer, db.ForeignKey('SolicitacoesCompra.id'), nullable=False)
    data_validacao = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    colaborador_1 = db.Column(db.Text, nullable=False)
    colaborador_2 = db.Column(db.Text, nullable=True)
    status = db.Column(db.Text, nullable=False)  # 'Conforme' ou 'Não Conforme'
    observacao = db.Column(db.Text, nullable=True)
    solicitacao = db.relationship('SolicitacoesCompra', backref='auditorias')

# Modelo para Requisições
class Requisicoes(db.Model):
    __tablename__ = 'Requisicoes'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    preenchimento_id = db.Column(db.Integer, db.ForeignKey('SolicitacoesPreenchidas.id'), nullable=True)  # Alterado para nullable=True
    cod_material = db.Column(db.Integer, db.ForeignKey('Materiais.CodMaterial'), nullable=False)
    quantidade = db.Column(db.Integer, nullable=False)
    ticket = db.Column(db.Text, nullable=False)
    data_requisicao = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    usuario = db.Column(db.Text, nullable=False)
    preenchimento = db.relationship('SolicitacoesPreenchidas', backref='requisicoes')
    material = db.relationship('Materiais', backref='requisicoes')

# Modelo para Pedidos de Compra
class PedidosCompra(db.Model):
    __tablename__ = 'PedidosCompra'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero_pedido = db.Column(db.Text, nullable=False, unique=True)
    data_criacao = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    usuario = db.Column(db.Text, nullable=False)
    status = db.Column(db.Text, nullable=False, default='Gerado')
    pdf_path = db.Column(db.Text, nullable=True)
    valor_total = db.Column(db.Float, nullable=False)
    valor_frete = db.Column(db.Float, nullable=True, default=0.0)
    valor_liquido = db.Column(db.Float, nullable=False)
    forma_pagamento = db.Column(db.Text, nullable=True)
    # ADICIONE ESTA LINHA
    observacoes = db.Column(db.Text, nullable=True)
    comprovante_pagamento = db.Column(db.Text, nullable=True)
    preenchimentos = db.relationship('SolicitacoesPreenchidas', secondary='pedido_preenchimento_associacao')
# Tabela de associação para relacionar PedidosCompra e SolicitacoesPreenchidas
pedido_preenchimento_associacao = db.Table('pedido_preenchimento_associacao',
    db.Column('pedido_id', db.Integer, db.ForeignKey('PedidosCompra.id'), primary_key=True),
    db.Column('preenchimento_id', db.Integer, db.ForeignKey('SolicitacoesPreenchidas.id'), primary_key=True)
)

# Modelo para Histórico de Descontos
class HistoricoDescontos(db.Model):
    __tablename__ = 'HistoricoDescontos'
    id = db.Column(db.Integer, primary_key=True)
    preenchimento_id = db.Column(db.Integer, db.ForeignKey('SolicitacoesPreenchidas.id'), nullable=False)
    valor_unitario_anterior = db.Column(db.Float, nullable=False)
    valor_unitario_novo = db.Column(db.Float, nullable=False)
    valor_frete_anterior = db.Column(db.Float, nullable=True)
    valor_frete_novo = db.Column(db.Float, nullable=True)
    data_alteracao = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    usuario = db.Column(db.Text, nullable=False)

    def to_dict(self):
        return {
            'id': self.id,
            'cod_material': self.cod_material,
            'especificacao': self.especificacao,
            'quantidade': self.quantidade,
            'unidade_medida': self.unidade_medida,
            'aplicacao': self.aplicacao,
            'empresa': self.empresa,
            'data_solicitacao': self.data_solicitacao.isoformat() if self.data_solicitacao else None,
            'usuario': self.usuario,
            'foto_path': self.foto_path,
            'marca': self.marca,
            'ativo': self.ativo,
            'nome_ativo': self.nome_ativo,
            'prioridade': self.prioridade,
            'status_aprovacao': self.status_aprovacao,
            'observacoes_col': self.observacoes_col  # NOVO CAMPO AQUI
        }

def check_historico_descontos_schema():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(HistoricoDescontos)")
        columns = [col[1] for col in cursor.fetchall()]
        print("Colunas em HistoricoDescontos:", columns)
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao verificar esquema: {str(e)}")

def add_comprovante_pagamento_column():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("ALTER TABLE PedidosCompra ADD COLUMN comprovante_pagamento TEXT")
        conn.commit()
        conn.close()
        logging.info("Added comprovante_pagamento column to PedidosCompra")
    except sqlite3.Error as e:
        logging.error(f"Error adding comprovante_pagamento column: {str(e)}")

# Funções de Banco de Dados
def init_db():
    with app.app_context():
        db.create_all()

def create_database():
    try:
        conn = sqlite3.connect(DATABASE)
        conn.close()
        print(f"Database '{DATABASE}' criada ou já existe.")
        return True
    except sqlite3.Error as e:
        print(f"Erro ao criar a database: {str(e)}")
        return False

def create_materiais_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Materiais';")
        if cursor.fetchone():
            print("Tabela 'Materiais' já existe.")
            cursor.close()
            conn.close()
            return True
        
        create_table_query = """
        CREATE TABLE Materiais (
            CodMaterial INTEGER PRIMARY KEY AUTOINCREMENT,
            DescricaoMaterial TEXT NOT NULL,
            Empresa TEXT NOT NULL,
            Aplicacao TEXT NOT NULL,
            QuantidadeEstoque INTEGER DEFAULT 0,
            Fornecedor TEXT,
            NumeroNF TEXT,
            FatorConsumo REAL DEFAULT 0.0
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'Materiais' criada com sucesso.")
        cursor.close()
        conn.close()
        return True
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela Materiais: {str(e)}")
        log_error(f"Erro ao criar a tabela Materiais: {str(e)}")
        return False

def create_kanban_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS Kanban (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_name TEXT NOT NULL,
            status TEXT NOT NULL
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'Kanban' criada com sucesso.")
        cursor.close()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela Kanban: {str(e)}")
        return False
    return True

def create_solicitacoes_compra_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS SolicitacoesCompra (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cod_material INTEGER NOT NULL,
            especificacao TEXT NOT NULL,
            quantidade INTEGER NOT NULL,
            unidade_medida TEXT NOT NULL DEFAULT 'Unidade',
            aplicacao TEXT,
            empresa TEXT NOT NULL,
            data_solicitacao DATETIME NOT NULL,
            usuario TEXT NOT NULL,
            foto_path TEXT,
            marca TEXT,
            ativo TEXT NOT NULL,
            nome_ativo TEXT,
            prioridade TEXT NOT NULL DEFAULT 'Programado',  -- Novo campo
            FOREIGN KEY (cod_material) REFERENCES Materiais(CodMaterial)
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'SolicitacoesCompra' criada com sucesso.")
        cursor.close()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela SolicitacoesCompra: {str(e)}")
        return False
    return True

def create_solicitacoes_preenchidas_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='SolicitacoesPreenchidas';")
        if cursor.fetchone():
            # Verificar se as colunas necessárias existem
            cursor.execute("PRAGMA table_info(SolicitacoesPreenchidas)")
            columns = [col[1] for col in cursor.fetchall()]
            
            # Migrações para colunas existentes
            if 'fornecedor_id' not in columns:
                cursor.execute("ALTER TABLE SolicitacoesPreenchidas ADD COLUMN fornecedor_id INTEGER NOT NULL DEFAULT 0")
                logging.info("Added fornecedor_id column to SolicitacoesPreenchidas")
            
            if 'valor_frete' not in columns:
                cursor.execute("ALTER TABLE SolicitacoesPreenchidas ADD COLUMN valor_frete REAL")
                logging.info("Added valor_frete column to SolicitacoesPreenchidas")
            
            # Não precisamos adicionar marca/ativo aqui pois são propriedades derivadas
            cursor.close()
            conn.close()
            return True
        
        # Criação da tabela se não existir
        create_table_query = """
        CREATE TABLE SolicitacoesPreenchidas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            solicitacao_id INTEGER NOT NULL,
            fornecedor_id INTEGER NOT NULL,
            valor_unitario REAL NOT NULL,
            valor_total REAL NOT NULL,
            valor_frete REAL,
            prazo_entrega TEXT NOT NULL,
            condicao_pagamento TEXT NOT NULL,
            data_preenchimento DATETIME NOT NULL,
            usuario TEXT NOT NULL,
            status TEXT DEFAULT 'Aguardando Aprovacao',
            pdf_path TEXT,
            FOREIGN KEY (solicitacao_id) REFERENCES SolicitacoesCompra(id)
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'SolicitacoesPreenchidas' criada com sucesso.")
        cursor.close()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela SolicitacoesPreenchidas: {str(e)}")
        return False
    return True

def create_estoque_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS Estoque (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            preenchimento_id INTEGER NOT NULL,
            cod_material INTEGER NOT NULL,
            quantidade INTEGER NOT NULL,
            fornecedor TEXT NOT NULL,
            numero_nf TEXT NOT NULL,
            data_entrada DATETIME NOT NULL,
            usuario TEXT NOT NULL,
            FOREIGN KEY (preenchimento_id) REFERENCES SolicitacoesPreenchidas(id),
            FOREIGN KEY (cod_material) REFERENCES Materiais(CodMaterial)
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'Estoque' criada com sucesso.")
        cursor.close()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela Estoque: {str(e)}")
        return False
    return True

def create_requisicoes_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS Requisicoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            preenchimento_id INTEGER NOT NULL,
            cod_material INTEGER NOT NULL,
            quantidade INTEGER NOT NULL,
            ticket TEXT NOT NULL,
            data_requisicao DATETIME NOT NULL,
            usuario TEXT NOT NULL,
            FOREIGN KEY (preenchimento_id) REFERENCES SolicitacoesPreenchidas(id),
            FOREIGN KEY (cod_material) REFERENCES Materiais(CodMaterial)
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'Requisicoes' criada com sucesso.")
        cursor.close()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela Requisicoes: {str(e)}")
        return False
    return True

def create_pedidos_compra_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS PedidosCompra (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_pedido TEXT NOT NULL UNIQUE,
            data_criacao DATETIME NOT NULL,
            usuario TEXT NOT NULL,
            status TEXT NOT NULL,
            pdf_path TEXT,
            valor_total REAL NOT NULL
        )
        """
        cursor.execute(create_table_query)
        create_associacao_table_query = """
        CREATE TABLE IF NOT EXISTS pedido_preenchimento_associacao (
            pedido_id INTEGER,
            preenchimento_id INTEGER,
            PRIMARY KEY (pedido_id, preenchimento_id),
            FOREIGN KEY (pedido_id) REFERENCES PedidosCompra(id),
            FOREIGN KEY (preenchimento_id) REFERENCES SolicitacoesPreenchidas(id)
        )
        """
        cursor.execute(create_associacao_table_query)
        conn.commit()
        print("Tabelas 'PedidosCompra' e 'pedido_preenchimento_associacao' criadas com sucesso.")
        cursor.close()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar tabelas PedidosCompra ou associação: {str(e)}")
        return False
    return True

def create_fornecedores_db():
    try:
        conn = sqlite3.connect(DB_PATH_FORNECEDORES)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS fornecedores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome_fantasia TEXT NOT NULL,
                cnpj TEXT NOT NULL UNIQUE,
                telefone TEXT NOT NULL,
                email TEXT NOT NULL,
                endereco TEXT NOT NULL,
                bairro TEXT NOT NULL,
                cidade TEXT NOT NULL,
                estado TEXT NOT NULL,
                contato TEXT NOT NULL,
                materiais TEXT NOT NULL
            )
        ''')
        conn.commit()
        conn.close()
        app.logger.info("Tabela 'fornecedores' criada com sucesso.")
        return True
    except sqlite3.Error as e:
        app.logger.error(f"Erro ao criar tabela fornecedores: {str(e)}")
        return False

class HistoricoEstoque(db.Model):
    __tablename__ = 'HistoricoEstoque'
    id = db.Column(db.Integer, primary_key=True)
    cod_material = db.Column(db.Integer, db.ForeignKey('Materiais.CodMaterial'))
    usuario = db.Column(db.Text, nullable=False)
    quantidade_anterior = db.Column(db.Integer, nullable=False)
    quantidade_nova = db.Column(db.Integer, nullable=False)
    data_alteracao = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    motivo = db.Column(db.Text, nullable=True)
    
    material = db.relationship('Materiais', backref='historico_estoque')

def get_db_connection(db_path):
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        return conn
    except sqlite3.Error as e:
        app.logger.error(f"Erro ao conectar ao banco de fornecedores: {str(e)}")
        return None

def validate_cnpj(cnpj):
    cnpj = ''.join(filter(str.isdigit, cnpj))
    if len(cnpj) != 14:
        return False
    return True

def validate_email(email):
    return re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email) is not None

@app.template_filter('format_cnpj')
def format_cnpj(cnpj):
    if not cnpj:
        return ''
    cnpj = ''.join(filter(str.isdigit, cnpj))
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

# Funções de Autenticação
def ler_senhas():
    senhas = {}
    try:
        caminho_absoluto = os.path.abspath(SENHAS_FILE)
        print(f"Tentando ler arquivo de senhas em: {caminho_absoluto}")
        
        if not os.path.exists(caminho_absoluto):
            print("Arquivo de senhas não encontrado, criando arquivo vazio")
            with open(caminho_absoluto, "w", encoding='utf-8') as f:
                f.write("")
        
        # Tentar ler com encoding UTF-8 e fallback para latin-1
        try:
            with open(caminho_absoluto, "r", encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(caminho_absoluto, "r", encoding='latin-1') as f:
                content = f.read()
        
        # Processar conteúdo
        for line in content.splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                partes = line.split("%")
                if len(partes) >= 3:
                    usuario = partes[0]
                    senha = partes[1]
                    pagina = partes[2]
                    empresa = partes[3] if len(partes) >= 4 else ""
                    senhas[usuario] = {
                        "senha": senha,
                        "pagina": pagina,
                        "empresa": empresa
                    }
        
        print(f"Total de usuários carregados: {len(senhas)}")
    except Exception as e:
        print(f"Erro crítico ao ler senhas.txt: {str(e)}")
        logging.error(f"Erro ao ler senhas.txt: {str(e)}", exc_info=True)
    
    return senhas


def salvar_senhas(senhas):
    try:
        with open(SENHAS_FILE, "w", encoding='utf-8') as f:
            for usuario, dados in senhas.items():
                empresa = dados.get('empresa', '')
                f.write(f"{usuario}%{dados['senha']}%{dados['pagina']}%{empresa}\n")
    except Exception as e:
        logging.error(f"Erro ao salvar senhas.txt: {str(e)}")


def registrar_log(usuario, tipo_acao, ip=None):
    acao = 'Acesso' if tipo_acao == 'login' else 'Logout'
    ip_info = f" - IP: {ip}" if ip else ''
    try:
        with open("arquivo.log", "a", encoding="utf-8") as log_file:
            log_file.write(f"{datetime.now()} - {acao} do usuário: {usuario}{ip_info}\n")
    except Exception as e:
        logging.error(f"Erro ao registrar log: {str(e)}")

def ler_logs():
    try:
        if not os.path.exists('arquivo.log'):
            with open('arquivo.log', 'w', encoding='utf-8') as log_file:
                log_file.write("Log de início\n")
        
        with open('arquivo.log', 'rb') as log_file:
            raw_data = log_file.read()
            encoding = chardet.detect(raw_data)['encoding'] or 'utf-8'
        
        with open('arquivo.log', 'r', encoding=encoding) as log_file:
            return log_file.readlines()
    except Exception as e:
        logging.error(f"Erro ao ler logs: {str(e)}")
        return []

def validar_usuario(usuario):
    if not usuario or len(usuario) < 3:
        return False, "O usuário deve ter pelo menos 3 caracteres."
    if not re.match(r'^[a-zA-Z0-9._-]+$', usuario):
        return False, "O usuário só pode conter letras, números, pontos, hífens ou sublinhados."
    if '%' in usuario:
        return False, "O usuário não pode conter o caractere '%'."
    return True, ""

def validar_senha(senha):
    if len(senha) < 6:
        return False, "A senha deve ter pelo menos 6 caracteres."
    if not re.search(r"[A-Z]", senha):
        return False, "A senha deve conter pelo menos uma letra maiúscula."
    if not re.search(r"[0-9]", senha):
        return False, "A senha deve conter pelo menos um número."
    if not re.search(r"[\W_]", senha):
        return False, "A senha deve conter pelo menos um caractere especial."
    return True, ""

# Registrar funções e objetos no jinja_env.globals
app.jinja_env.globals.update(
    ler_senhas=ler_senhas,
    salvar_senhas=salvar_senhas,
    registrar_log=registrar_log,
    ler_logs=ler_logs,
    validar_usuario=validar_usuario,
    validar_senha=validar_senha,
    Kanban=Kanban,
    Materiais=Materiais,
    SolicitacoesCompra=SolicitacoesCompra,
    SolicitacoesPreenchidas=SolicitacoesPreenchidas,
    Estoque=Estoque,
    Requisicoes=Requisicoes,
    PedidosCompra=PedidosCompra,
    db=db
)

# Blueprint para rotas
routes_bp = Blueprint('routes_bp', __name__)

@routes_bp.before_request
def verificar_tempo_inatividade():
    if request.endpoint in ['routes_bp.login', 'static']:
        return
    
    print(f"Verificando sessão para endpoint: {request.endpoint}")
    print(f"Sessão atual: {dict(session)}")
    print(f"Cookies recebidos: {request.cookies}")
    if 'usuario' not in session:
        print(f"Usuário não autenticado tentando acessar {request.endpoint}")
        return redirect(url_for('routes_bp.login'))
    
    agora = datetime.now()
    ultima_atividade = session.get('ultima_atividade')
    
    if ultima_atividade:
        ultima_atividade = datetime.strptime(ultima_atividade, '%Y-%m-%d %H:%M:%S.%f')
        if agora - ultima_atividade > timedelta(minutes=30):
            usuario = session.get('usuario')
            registrar_log(usuario, 'logout (inatividade)', request.remote_addr)
            session.clear()
            flash('Você foi desconectado por inatividade.', 'warning')
            return redirect(url_for('routes_bp.login'))
    
    session['ultima_atividade'] = str(agora)

@routes_bp.route('/', methods=['GET'])
def home():
    return redirect(url_for('routes_bp.login'))

@routes_bp.route('/uploads/<filename>')
def uploads(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@routes_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        print(f"Dados recebidos - Usuário: {request.form.get('usuario')}")
        usuario_completo = request.form.get('usuario', '').strip()
        senha = request.form.get('senha', '').strip()
        
        senhas = ler_senhas()
        print(f"Usuários no sistema: {list(senhas.keys())}")
        
        if usuario_completo in senhas and senha == senhas[usuario_completo]["senha"]:
            usuario_real = usuario_completo.split('%')[0]
            session['usuario'] = usuario_real
            session['_fresh'] = True
            print(f"Sessão criada para {usuario_real}. Conteúdo: {dict(session)}")
            print(f"Cookie de sessão antes de redirecionar: {request.cookies.get('__Secure-sessionid', 'N/A')}")
            pagina_destino = senhas[usuario_completo]["pagina"].replace('.html', '')
            print(f"Redirecionando para: {pagina_destino}")
            response = redirect(url_for(f'routes_bp.{pagina_destino}'))
            print(f"Headers da resposta: {response.headers}")
            return response
        
        flash('Credenciais inválidas', 'error')
        print("Falha na autenticação")
    
    print(f"Cookies recebidos na requisição GET: {request.cookies}")
    return render_template('login.html')

@routes_bp.route('/logout', methods=['GET'])
def logout():
    usuario = session.get('usuario')
    if usuario:
        app.jinja_env.globals['registrar_log'](usuario, 'logout', request.remote_addr)
        session.clear()
    return redirect(url_for('routes_bp.login'))

@routes_bp.route('/senhas', methods=['GET'])
def listar_senhas():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    senhas = app.jinja_env.globals['ler_senhas']()
    return render_template('senhas.html', senhas=senhas)

@routes_bp.route('/adicionar_senha', methods=['POST'])
def adicionar_senha():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))

    try:
        usuario = request.form['usuario'].strip()
        senha = request.form['senha'].strip()
        pagina = request.form['pagina']
        pagina_personalizada = request.form.get('pagina_personalizada', '').strip()
        empresa = request.form.get('empresa', '').strip()

        # Verifica se algum campo contém o caractere proibido
        for campo_nome, campo_valor in [("Usuário", usuario), ("Senha", senha), ("Empresa", empresa), ("Página personalizada", pagina_personalizada)]:
            if "%" in campo_valor:
                flash(f'O campo "{campo_nome}" não pode conter o caractere "%".', 'error')
                return redirect(url_for('routes_bp.listar_senhas'))

        # Validação básica do usuário
        is_valid_usuario, msg_usuario = validar_usuario(usuario)
        if not is_valid_usuario:
            flash(msg_usuario, 'error')
            return redirect(url_for('routes_bp.listar_senhas'))

        # Validação básica da senha
        is_valid_senha, msg_senha = validar_senha(senha)
        if not is_valid_senha:
            flash(msg_senha, 'error')
            return redirect(url_for('routes_bp.listar_senhas'))

        # Validação da página
        if not pagina:
            flash('O campo página é obrigatório.', 'error')
            return redirect(url_for('routes_bp.listar_senhas'))

        if pagina == 'outro':
            if not pagina_personalizada:
                flash('Página personalizada é obrigatória quando "outro" é selecionado.', 'error')
                return redirect(url_for('routes_bp.listar_senhas'))
            if not pagina_personalizada.endswith('.html'):
                flash('A página personalizada deve terminar com ".html".', 'error')
                return redirect(url_for('routes_bp.listar_senhas'))
            pagina = pagina_personalizada

        if not empresa:
            flash('O campo Empresa é obrigatório.', 'error')
            return redirect(url_for('routes_bp.listar_senhas'))

        senhas = ler_senhas()
        if usuario in senhas:
            flash('Usuário já existe.', 'error')
            return redirect(url_for('routes_bp.listar_senhas'))

        senhas[usuario] = {
            "senha": senha,
            "pagina": pagina,
            "empresa": empresa
        }

        salvar_senhas(senhas)
        flash('Senha adicionada com sucesso.', 'success')
        return redirect(url_for('routes_bp.listar_senhas'))

    except Exception as e:
        flash('Erro ao adicionar senha.', 'error')
        return redirect(url_for('routes_bp.listar_senhas'))



@routes_bp.route('/excluir_senha/<usuario>', methods=['POST'])
def excluir_senha(usuario):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        senhas = app.jinja_env.globals['ler_senhas']()
        if usuario in senhas:
            del senhas[usuario]
            app.jinja_env.globals['salvar_senhas'](senhas)
            flash(f'Senha para {usuario} excluída com sucesso.', 'success')
        else:
            flash(f'Usuário {usuario} não encontrado.', 'error')
        return redirect(url_for('routes_bp.listar_senhas'))
    except Exception as e:
        flash('Erro ao excluir senha.', 'error')
        return redirect(url_for('routes_bp.listar_senhas'))

@routes_bp.route('/kanban', methods=['GET'])
def kanban():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('kanban.html')

@routes_bp.route('/tasks', methods=['GET'])
def get_tasks():
    tasks = app.jinja_env.globals['Kanban'].query.all()
    return jsonify({
        'tasks': [{'id': t.id, 'task_name': t.task_name, 'status': t.status} for t in tasks]
    })

@routes_bp.route('/tasks', methods=['POST'])
def add_task():
    data = request.get_json()
    if not data or 'task_name' not in data or 'status' not in data:
        return jsonify({'message': 'Task name and status are required'}), 400
    
    valid_statuses = ['A Fazer', 'Em Progresso', 'Concluído']
    if data['status'] not in valid_statuses:
        return jsonify({'message': 'Invalid status'}), 400
    
    task = app.jinja_env.globals['Kanban'](task_name=data['task_name'], status=data['status'])
    app.jinja_env.globals['db'].session.add(task)
    app.jinja_env.globals['db'].session.commit()
    return jsonify({'message': 'Task added successfully'}), 201

@routes_bp.route('/tasks/<int:id>', methods=['PUT'])
def update_task(id):
    data = request.get_json()
    if not data or 'status' not in data:
        return jsonify({'message': 'Status is required'}), 400
    
    valid_statuses = ['A Fazer', 'Em Progresso', 'Concluído']
    if data['status'] not in valid_statuses:
        return jsonify({'message': 'Invalid status'}), 400
    
    task = app.jinja_env.globals['Kanban'].query.get_or_404(id)
    task.status = data['status']
    app.jinja_env.globals['db'].session.commit()
    return jsonify({'message': 'Task updated successfully'})

@routes_bp.route('/tasks/<int:id>', methods=['DELETE'])
def delete_task(id):
    task = app.jinja_env.globals['Kanban'].query.get_or_404(id)
    app.jinja_env.globals['db'].session.delete(task)
    app.jinja_env.globals['db'].session.commit()
    return jsonify({'message': 'Task deleted successfully'})

#Menu Master
@routes_bp.route('/menu_master', methods=['GET'])
def menu_master():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_master.html')

#Menu Master
@routes_bp.route('/menu_aprovacao', methods=['GET'])
def menu_aprovacao():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_aprovacao.html')

#Menu Estoquista
@routes_bp.route('/menu_estoquista', methods=['GET'])
def menu_estoquista():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_estoquista.html')

#Menu Solicitante
@routes_bp.route('/menu_solicitante', methods=['GET'])
def menu_solicitante():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_solicitante.html')

#Menu Comprado
@routes_bp.route('/menu_comprador', methods=['GET'])
def menu_comprador():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_comprador.html')

#Menu Supervisor
@routes_bp.route('/menu_supervisor', methods=['GET'])
def menu_supervisor():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_supervisor.html')

#Menu Auditoria
@routes_bp.route('/menu_auditoria', methods=['GET'])
def menu_auditoria():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_auditoria.html')

#Menu Financeiro
@routes_bp.route('/menu_financeiro', methods=['GET'])
def menu_financeiro():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_financeiro.html')

#Menu Cadastro
@routes_bp.route('/menu_cadastro', methods=['GET'])
def menu_cadastro():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_cadastro.html')

@routes_bp.route('/menu_reenvio', methods=['GET'])
def menu_reenvio():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_reenvio.html')

@routes_bp.route('/formulario_custodia', methods=['GET'])
def formulario_custodia():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('formulario_custodia.html')


@routes_bp.route('/menu_logistica_auditoria', methods=['GET'])
def menu_logistica_auditoria():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_logistica_auditoria.html')

@routes_bp.route('/menu_logistica_supervisor', methods=['GET'])
def menu_logistica_supervisor():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_logistica_supervisor.html')

@routes_bp.route('/menu_logistica_analista1', methods=['GET'])
def menu_logistica_analista1():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_logistica_analista1.html')

@routes_bp.route('/menu_logistica_analista2', methods=['GET'])
def menu_logistica_analista2():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_logistica_analista2.html')

@routes_bp.route('/menu_logistica_analista3', methods=['GET'])
def menu_logistica_analista3():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('menu_logistica_analista3.html')

@routes_bp.route('/logs', methods=['GET'])
def logs():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    logs = app.jinja_env.globals['ler_logs']()
    return render_template('logs.html', logs=logs)

@routes_bp.route('/cadastrar_material', methods=['GET'])
def cadastrar_material():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('cadastrar_material.html')

@routes_bp.route('/materiais', methods=['GET'])
def listar_materiais():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    try:
        materiais = app.jinja_env.globals['Materiais'].query.all()
        return render_template('materiais.html', materiais=materiais)
    except Exception as e:
        flash(f'Erro ao carregar materiais: {str(e)}', 'error')
        return render_template('materiais.html', materiais=[])

@routes_bp.route('/adicionar_material', methods=['POST'])
def adicionar_material():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    try:
        descricao = request.form.get('descricao', '').strip()
        empresa = request.form.get('empresa', '').strip()
        aplicacao = request.form.get('aplicacao', '').strip()

        if not descricao or not empresa or not aplicacao:
            flash('Todos os campos são obrigatórios.', 'error')
            return redirect(url_for('routes_bp.cadastrar_material'))

        if len(descricao) < 3:
            flash('A descrição deve ter pelo menos 3 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastrar_material'))

        if len(empresa) < 2:
            flash('A empresa deve ter pelo menos 2 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastrar_material'))

        if len(aplicacao) < 3:
            flash('A aplicação deve ter pelo menos 3 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastrar_material'))

        material = app.jinja_env.globals['Materiais'](
            DescricaoMaterial=descricao,
            Empresa=empresa,
            Aplicacao=aplicacao,
            QuantidadeEstoque=0,
            Fornecedor=None,
            NumeroNF=None,
            FatorConsumo=0.0
        )
        app.jinja_env.globals['db'].session.add(material)
        app.jinja_env.globals['db'].session.commit()

        flash('Material adicionado com sucesso.', 'success')
        return redirect(url_for('routes_bp.listar_materiais'))
    except Exception as e:
        app.jinja_env.globals['db'].session.rollback()
        flash(f'Erro ao adicionar material: {str(e)}', 'error')
        return redirect(url_for('routes_bp.cadastrar_material'))

@routes_bp.route('/excluir_material/<int:cod>', methods=['POST'])
def excluir_material(cod):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    try:
        material = app.jinja_env.globals['Materiais'].query.get_or_404(cod)
        app.jinja_env.globals['db'].session.delete(material)
        app.jinja_env.globals['db'].session.commit()
        flash(f'Material {material.DescricaoMaterial} excluído com sucesso.', 'success')
        return redirect(url_for('routes_bp.listar_materiais'))
    except Exception as e:
        flash(f'Erro ao excluir material: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_materiais'))

@routes_bp.route('/buscar_material', methods=['GET', 'POST'])
def buscar_material():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    materiais = []
    empresas_unicas = []
    aplicacoes_unicas = []
    
    try:
        # Obter valores únicos para os filtros
        empresas_unicas = db.session.query(
            Materiais.Empresa
        ).distinct().order_by(
            Materiais.Empresa
        ).all()
        
        aplicacoes_unicas = db.session.query(
            Materiais.Aplicacao
        ).distinct().order_by(
            Materiais.Aplicacao
        ).all()
        
        empresas_unicas = [e[0] for e in empresas_unicas if e[0]]
        aplicacoes_unicas = [a[0] for a in aplicacoes_unicas if a[0]]

        if request.method == 'POST':
            termo = request.form.get('termo', '').strip()
            empresa = request.form.get('empresa', '').strip()
            aplicacao = request.form.get('aplicacao', '').strip()
            status = request.form.get('status', '').strip()

            query = Materiais.query

            if termo:
                if termo.isdigit():
                    query = query.filter_by(CodMaterial=int(termo))
                else:
                    query = query.filter(
                        Materiais.DescricaoMaterial.ilike(f'%{termo}%')
                    )
            
            if empresa:
                query = query.filter_by(Empresa=empresa)
            
            if aplicacao:
                query = query.filter_by(Aplicacao=aplicacao)
            
            if status:
                query = query.filter_by(Ativo=(status == 'ativo'))

            materiais = query.all()
            
            if not materiais:
                flash('Nenhum material encontrado com os filtros aplicados.', 'info')
        
    except Exception as e:
        flash(f'Erro ao buscar materiais: {str(e)}', 'error')
    
    return render_template('buscar_material.html', 
                         materiais=materiais,
                         empresas_unicas=empresas_unicas,
                         aplicacoes_unicas=aplicacoes_unicas)

@routes_bp.route('/solicitar_compra/<int:cod>', methods=['GET'])
def solicitar_compra(cod):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Busca o material
        material = Materiais.query.get_or_404(cod)
        
        # Carrega os dados do usuário
        senhas = ler_senhas()
        usuario = session.get('usuario')
        
        # Verifica se o usuário existe
        if usuario not in senhas:
            flash('Usuário não encontrado.', 'error')
            return redirect(url_for('routes_bp.login'))
        
        empresa_usuario = senhas[usuario].get('empresa', '')
        
        return render_template('solicitar_compra.html', material=material, empresa_usuario=empresa_usuario)

    except SQLAlchemyError as e:
        flash(f'Erro ao carregar material: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))
    except Exception as e:
        flash(f'Erro inesperado: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))

@routes_bp.route('/abrir_solicitacao', methods=['POST'])
def abrir_solicitacao():
    if 'usuario' not in session:
        flash('Você precisa estar logado para fazer uma solicitação.', 'error')
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Extração dos campos do formulário
        cod_material = request.form.get('cod_material')
        especificacao = request.form.get('especificacao', '').strip()
        quantidade = request.form.get('quantidade', '').strip()
        unidade_medida = request.form.get('unidade_medida', '').strip()
        aplicacao = request.form.get('aplicacao', '').strip()
        empresa = request.form.get('empresa', '').strip()
        marca = request.form.get('marca', '').strip()
        ativo = request.form.get('ativo', '').strip()
        nome_ativo = request.form.get('nome_ativo', '').strip()
        prioridade = request.form.get('prioridade', 'Programado').strip()  # Novo campo
        foto = request.files.get('foto')

        # Validação do cod_material
        if not cod_material:
            flash('Código do material é obrigatório.', 'error')
            return redirect(url_for('routes_bp.buscar_material'))
        
        try:
            cod_material = int(cod_material)
            if cod_material <= 0:
                raise ValueError
        except ValueError:
            flash('Código do material deve ser um número positivo.', 'error')
            return redirect(url_for('routes_bp.buscar_material'))

        # Verificar se o material existe
        material = db.session.get(Materiais, cod_material)
        if not material:
            flash(f'Material com código {cod_material} não encontrado.', 'error')
            return redirect(url_for('routes_bp.buscar_material'))

        # Validação dos campos obrigatórios
        required_fields = {
            'especificacao': especificacao,
            'quantidade': quantidade,
            'unidade_medida': unidade_medida,
            'empresa': empresa,
            'ativo': ativo,
            'prioridade': prioridade  # Novo campo obrigatório
        }
        
        for field_name, field_value in required_fields.items():
            if not field_value:
                flash(f'O campo {field_name} é obrigatório.', 'error')
                return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))

        # Validação específica para ativo e nome_ativo
        if ativo not in ['Sim', 'Não']:
            flash('O campo "Material é ativo?" deve ser "Sim" ou "Não".', 'error')
            return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))
        
        if ativo == 'Sim' and not nome_ativo:
            flash('O nome do ativo é obrigatório quando o material é ativo.', 'error')
            return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))

        # Validação da prioridade
        if prioridade not in ['Programado', 'Urgente']:
            flash('Prioridade inválida. Deve ser "Programado" ou "Urgente".', 'error')
            return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))

        # Validação de quantidade
        if not quantidade.isdigit() or int(quantidade) <= 0:
            flash('Quantidade deve ser um número positivo.', 'error')
            return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))

        # Validação de unidade de medida
        valid_unidades = [
            'Peça', 'Unidade', 'Metro', 'Metro quadrado', 'Metro cúbico', 
            'Litro', 'Quilograma', 'Grama', 'Caixa', 'Pacote', 'Par'
        ]
        if unidade_medida not in valid_unidades:
            flash('Unidade de medida inválida.', 'error')
            return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))

        # Processamento do upload da foto
        foto_path = None
        if foto and foto.filename:
            logging.info(f"Uploaded file: {foto.filename}")
            if allowed_file(foto.filename, {'jpg', 'jpeg', 'png', 'pdf'}):
                filename = f"{uuid.uuid4()}_{secure_filename(foto.filename)}"
                foto_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                try:
                    foto.save(foto_path)
                    logging.info(f"File saved to: {foto_path}")
                except Exception as e:
                    logging.error(f"Failed to save file: {str(e)}")
                    flash(f'Erro ao salvar a foto: {str(e)}', 'error')
                    return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))
            else:
                extension = foto.filename.rsplit('.', 1)[1].lower() if '.' in foto.filename else 'N/A'
                logging.error(f"Invalid file extension: {extension}")
                flash(f'Arquivo de foto inválido. Apenas JPG, JPEG, PNG e PDF são permitidos. Extensão detectada: {extension}', 'error')
                return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))
        elif foto:
            logging.warning("Empty file uploaded")
            flash('Nenhum arquivo selecionado ou arquivo vazio.', 'error')
            return redirect(url_for('routes_bp.solicitar_compra', cod=cod_material))

        # Criar a solicitação de compra
        solicitacao = SolicitacoesCompra(
            cod_material=cod_material,
            especificacao=especificacao,
            quantidade=int(quantidade),
            unidade_medida=unidade_medida,
            aplicacao=aplicacao or None,
            empresa=empresa,
            usuario=session['usuario'],
            foto_path=foto_path,
            marca=marca or None,
            ativo=ativo,
            nome_ativo=nome_ativo or None if ativo == 'Sim' else None,
            prioridade=prioridade  # Novo campo
        )
        db.session.add(solicitacao)
        db.session.commit()

        flash('Solicitação de compra aberta com sucesso.', 'success')
        return redirect(url_for('routes_bp.buscar_material'))
    except Exception as e:
        db.session.rollback()
        logging.error(f"Error in abrir_solicitacao: {str(e)}")
        flash(f'Erro ao abrir solicitação: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))
    
# Registrar filtro basename no Jinja2
def basename_filter(filepath):
    return os.path.basename(filepath) if filepath else ''

app.jinja_env.filters['basename'] = basename_filter
       
@routes_bp.route('/listar_solicitacoes', methods=['GET'])
def listar_solicitacoes():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Consulta para obter apenas solicitações sem nenhum preenchimento
        solicitacoes = db.session.query(SolicitacoesCompra).outerjoin(
            SolicitacoesPreenchidas,
            SolicitacoesCompra.id == SolicitacoesPreenchidas.solicitacao_id
        ).filter(
            SolicitacoesPreenchidas.id.is_(None)
        ).all()
        
        # Função para ler usuários e empresas do arquivo senhas.txt
        def get_usuarios_empresas():
            usuarios = set()
            empresas = set()
            try:
                with open('senhas.txt', 'r', encoding='utf-8') as f:
                    for line in f:
                        partes = line.strip().split('%')
                        if len(partes) >= 4:  # Verifica se tem pelo menos 4 partes
                            usuario = partes[0]
                            empresa = partes[3]
                            usuarios.add(usuario)
                            empresas.add(empresa)
            except Exception as e:
                logging.error(f"Erro ao ler senhas.txt: {str(e)}")
            return sorted(usuarios), sorted(empresas)
        
        # Obter listas únicas de empresas e usuários do arquivo
        usuarios, empresas = get_usuarios_empresas()
        
        if not solicitacoes:
            flash('Nenhuma solicitação aberta encontrada.', 'info')
        return render_template('listar_solicitacoes.html', 
                            solicitacoes=solicitacoes,
                            empresas=empresas,
                            usuarios=usuarios)
    except Exception as e:
        logging.error(f"Error in listar_solicitacoes: {str(e)}")
        flash(f'Erro ao carregar solicitações: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))

@routes_bp.route('/aprovar_solicitacao', methods=['GET'])
def aprovar_solicitacao():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Consulta para obter apenas solicitações sem nenhum preenchimento
        solicitacoes = db.session.query(SolicitacoesCompra).outerjoin(
            SolicitacoesPreenchidas,
            SolicitacoesCompra.id == SolicitacoesPreenchidas.solicitacao_id
        ).filter(
            SolicitacoesPreenchidas.id.is_(None)
        ).all()
        
        # Função para ler usuários e empresas do arquivo senhas.txt
        def get_usuarios_empresas():
            usuarios = set()
            empresas = set()
            try:
                with open('senhas.txt', 'r', encoding='utf-8') as f:
                    for line in f:
                        partes = line.strip().split('%')
                        if len(partes) >= 4:  # Verifica se tem pelo menos 4 partes
                            usuario = partes[0]
                            empresa = partes[3]
                            usuarios.add(usuario)
                            empresas.add(empresa)
            except Exception as e:
                logging.error(f"Erro ao ler senhas.txt: {str(e)}")
            return sorted(usuarios), sorted(empresas)
        
        # Obter listas únicas de empresas e usuários do arquivo
        usuarios, empresas = get_usuarios_empresas()
        
        if not solicitacoes:
            flash('Nenhuma solicitação aberta encontrada.', 'info')
        return render_template('aprovar_solicitacao.html', 
                            solicitacoes=solicitacoes,
                            empresas=empresas,
                            usuarios=usuarios)
    except Exception as e:
        logging.error(f"Error in aprovar_solicitacao: {str(e)}")
        flash(f'Erro ao carregar solicitações: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))
    
def migrate_solicitacoes_compra():
    try:
        # Usar o mesmo banco de dados do Flask-SQLAlchemy
        DATABASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ComparasDB.db')
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()

        # Verificar se a tabela SolicitacoesCompra existe
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='SolicitacoesCompra'")
        table_exists = cursor.fetchone()

        if not table_exists:
            # Criar a tabela com todas as colunas do modelo
            cursor.execute("""
                CREATE TABLE SolicitacoesCompra (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cod_material INTEGER NOT NULL,
                    especificacao TEXT NOT NULL,
                    quantidade INTEGER NOT NULL,
                    unidade_medida TEXT NOT NULL DEFAULT 'Unidade',
                    aplicacao TEXT,
                    empresa TEXT NOT NULL,
                    data_solicitacao DATETIME NOT NULL,
                    usuario TEXT NOT NULL,
                    foto_path TEXT,
                    FOREIGN KEY (cod_material) REFERENCES Materiais(CodMaterial)
                )
            """)
            logging.info("Created SolicitacoesCompra table with all columns")
        else:
            # Verificar colunas atuais
            cursor.execute("PRAGMA table_info(SolicitacoesCompra)")
            columns = [col[1] for col in cursor.fetchall()]
            logging.info(f"Current columns in SolicitacoesCompra: {columns}")

            # Adicionar colunas faltantes, se necessário
            if 'unidade_medida' not in columns:
                cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN unidade_medida TEXT NOT NULL DEFAULT 'Unidade'")
                logging.info("Added unidade_medida column to SolicitacoesCompra")
            
            if 'aplicacao' not in columns:
                cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN aplicacao TEXT")
                logging.info("Added aplicacao column to SolicitacoesCompra")
            
            if 'foto_path' not in columns:
                cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN foto_path TEXT")
                logging.info("Added foto_path column to SolicitacoesCompra")

        conn.commit()
        logging.info("Database migration for SolicitacoesCompra completed successfully")

        # Verificar esquema atualizado
        cursor.execute("PRAGMA table_info(SolicitacoesCompra)")
        updated_columns = [col[1] for col in cursor.fetchall()]
        logging.info(f"Updated columns in SolicitacoesCompra: {updated_columns}")
    except sqlite3.Error as e:
        logging.error(f"Error during migration: {str(e)}")
        raise
    finally:
        conn.close()

@routes_bp.route('/preencher_solicitacao/<int:id>', methods=['GET', 'POST'])
def preencher_solicitacao(id):
    if 'usuario' not in session:
        flash('Por favor, faça login para continuar.', 'error')
        return redirect(url_for('routes_bp.login'))
    
    try:
        solicitacao = SolicitacoesCompra.query.get_or_404(id)
        
        # Obter lista de fornecedores com CNPJ
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('SELECT id, nome_fantasia, cnpj FROM fornecedores ORDER BY nome_fantasia')
            fornecedores = [dict(row) for row in cursor.fetchall()]
            conn.close()
        else:
            fornecedores = []
            flash('Erro ao conectar ao banco de fornecedores', 'warning')

        if request.method == 'POST':
            # Verificar limite de cotações
            cotacoes_existentes = SolicitacoesPreenchidas.query.filter_by(solicitacao_id=id).count()
            if cotacoes_existentes >= 9:
                flash('Limite de 9 cotações atingido para esta solicitação.', 'error')
                return render_template('preencher_solicitacao.html', 
                                      solicitacao=solicitacao, 
                                      fornecedores=fornecedores)

            # Obter listas de dados do formulário
            fornecedor_ids = request.form.getlist('fornecedor_id[]')
            valor_unitario_list = request.form.getlist('valor_unitario[]')
            valor_total_list = request.form.getlist('valor_total[]')
            valor_frete_list = request.form.getlist('valor_frete[]')
            prazo_entrega_list = request.form.getlist('prazo_entrega[]')
            condicao_pagamento_list = request.form.getlist('condicao_pagamento[]')
            observacao_list = request.form.getlist('observacao[]')  # Adicionado: capturar observações
            pdf_files = request.files.getlist('pdf_file[]')

            # Log para depuração
            logging.info(f"Dados recebidos do formulário: {request.form}")
            logging.info(f"Arquivos recebidos: {[f.filename for f in pdf_files if f]}")
            logging.info(f"Observações recebidas: {observacao_list}")

            # Verificar consistência das listas
            expected_length = len(fornecedor_ids)
            if not all(len(lst) == expected_length for lst in [
                valor_unitario_list, valor_total_list, prazo_entrega_list, 
                condicao_pagamento_list, observacao_list, pdf_files]):
                flash('Erro: Número inconsistente de campos nas cotações.', 'error')
                return render_template('preencher_solicitacao.html', 
                                      solicitacao=solicitacao, 
                                      fornecedores=fornecedores)

            if not fornecedor_ids or not valor_unitario_list:
                flash('Nenhuma cotação fornecida.', 'error')
                return render_template('preencher_solicitacao.html', 
                                      solicitacao=solicitacao, 
                                      fornecedores=fornecedores)

            if len(fornecedor_ids) > (9 - cotacoes_existentes):
                flash(f'Você pode adicionar no máximo {9 - cotacoes_existentes} cotações.', 'error')
                return render_template('preencher_solicitacao.html', 
                                      solicitacao=solicitacao, 
                                      fornecedores=fornecedores)

            # Função auxiliar para parse de valores monetários
            def parse_br_currency(value):
                if not value:
                    return None
                try:
                    cleaned = value.replace('.', '').replace(',', '.')
                    return float(cleaned)
                except (ValueError, AttributeError) as e:
                    logging.error(f"Erro ao parsear valor monetário '{value}': {str(e)}")
                    return None

            for i in range(len(fornecedor_ids)):
                try:
                    fornecedor_id = fornecedor_ids[i]
                    valor_unitario = valor_unitario_list[i].strip()
                    valor_total = valor_total_list[i].strip()
                    valor_frete = valor_frete_list[i].strip() if i < len(valor_frete_list) else ''
                    prazo_entrega = prazo_entrega_list[i].strip()
                    condicao_pagamento = condicao_pagamento_list[i].strip()
                    observacao = observacao_list[i].strip() if i < len(observacao_list) else None  # Adicionado: processar observação
                    pdf_file = pdf_files[i] if i < len(pdf_files) else None

                    if not all([fornecedor_id, valor_unitario, valor_total, prazo_entrega, condicao_pagamento]):
                        flash(f'Todos os campos obrigatórios devem ser preenchidos para a cotação {i+1}.', 'error')
                        continue
                    
                    # Verificar se o fornecedor existe
                    conn = get_db_connection(DB_PATH_FORNECEDORES)
                    if conn:
                        conn.row_factory = sqlite3.Row
                        cursor = conn.cursor()
                        cursor.execute('SELECT id FROM fornecedores WHERE id = ?', (fornecedor_id,))
                        if not cursor.fetchone():
                            flash(f'Fornecedor inválido na cotação {i+1}.', 'error')
                            conn.close()
                            continue
                        conn.close()

                    # Processamento dos valores monetários
                    valor_unitario_parsed = parse_br_currency(valor_unitario)
                    valor_total_parsed = parse_br_currency(valor_total)
                    valor_frete_parsed = parse_br_currency(valor_frete) if valor_frete else None

                    if valor_unitario_parsed is None or valor_total_parsed is None:
                        flash(f'Valor unitário e total devem ser números válidos na cotação {i+1}.', 'error')
                        continue
                    if valor_unitario_parsed <= 0 or valor_total_parsed <= 0:
                        flash(f'Valores unitário e total devem ser positivos na cotação {i+1}.', 'error')
                        continue
                    if valor_frete_parsed is not None and valor_frete_parsed < 0:
                        flash(f'O valor do frete deve ser positivo ou zero na cotação {i+1}.', 'error')
                        continue

                    # Processar o upload do PDF
                    pdf_path = None
                    if pdf_file and pdf_file.filename:
                        if allowed_file(pdf_file.filename):
                            filename = f"{uuid.uuid4()}_{secure_filename(pdf_file.filename)}"
                            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                            try:
                                pdf_file.save(pdf_path)
                                logging.info(f"PDF salvo em: {pdf_path}")
                            except Exception as e:
                                logging.error(f"Erro ao salvar PDF na cotação {i+1}: {str(e)}")
                                flash(f'Erro ao salvar o PDF na cotação {i+1}.', 'error')
                                continue
                        else:
                            extension = pdf_file.filename.rsplit('.', 1)[1].lower() if '.' in pdf_file.filename else 'N/A'
                            flash(f'Arquivo inválido na cotação {i+1}. Apenas PDFs são permitidos. Extensão detectada: {extension}', 'error')
                            continue

                    # Criar nova instância de SolicitacoesPreenchidas
                    preenchimento = SolicitacoesPreenchidas(
                        solicitacao_id=id,
                        fornecedor_id=int(fornecedor_id),
                        valor_unitario=valor_unitario_parsed,
                        valor_frete=valor_frete_parsed,
                        valor_total=valor_total_parsed,
                        prazo_entrega=prazo_entrega,
                        condicao_pagamento=condicao_pagamento,
                        usuario=session['usuario'],
                        data_preenchimento=datetime.utcnow(),
                        status='Aguardando Aprovação',
                        pdf_path=pdf_path,
                        observacoes=observacao if observacao else None  # Adicionado: salvar observação
                    )
                    db.session.add(preenchimento)
                    logging.info(f"Cotação {i+1} adicionada: Fornecedor ID {fornecedor_id}, Observações: {observacao}")

                except Exception as e:
                    logging.error(f"Erro ao processar cotação {i+1}: {str(e)}")
                    flash(f'Erro ao processar cotação {i+1}: {str(e)}', 'error')
                    continue

            try:
                db.session.commit()
                flash('Cotações preenchidas com sucesso.', 'success')
                return redirect(url_for('routes_bp.listar_solicitacoes'))
            except Exception as e:
                db.session.rollback()
                flash(f'Erro ao salvar as cotações: {str(e)}', 'error')
                logging.error(f"Erro no commit do banco: {str(e)}")
                return render_template('preencher_solicitacao.html', 
                                      solicitacao=solicitacao, 
                                      fornecedores=fornecedores)
        
        return render_template('preencher_solicitacao.html', 
                              solicitacao=solicitacao, 
                              fornecedores=fornecedores)
    
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao processar cotações: {str(e)}', 'error')
        logging.error(f"Erro geral em preencher_solicitacao: {str(e)}")
        return render_template('preencher_solicitacao.html', 
                              solicitacao=solicitacao, 
                              fornecedores=fornecedores)
def add_observacoes_column_to_solicitacoes_preenchidas():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Verificar se a coluna já existe
        cursor.execute("PRAGMA table_info(SolicitacoesPreenchidas)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'observacoes' not in columns:
            cursor.execute("ALTER TABLE SolicitacoesPreenchidas ADD COLUMN observacoes TEXT")
            conn.commit()
            logging.info("Coluna observacoes adicionada à tabela SolicitacoesPreenchidas")
            print("✓ Coluna observacoes adicionada com sucesso!")
        else:
            print("✓ Coluna observacoes já existe na tabela")
        
        conn.close()
        return True
    except sqlite3.Error as e:
        logging.error(f"Erro ao adicionar coluna observacoes: {str(e)}")
        print(f"✗ Erro ao adicionar coluna: {str(e)}")
        return False
    
@routes_bp.route('/listar_solicitacoes_preenchidas', methods=['GET'])
def listar_solicitacoes_preenchidas():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        empresa = request.args.get('empresa')
        usuario = request.args.get('usuario')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')
        
        query = SolicitacoesPreenchidas.query.join(SolicitacoesCompra)
        
        if empresa:
            query = query.filter(SolicitacoesCompra.empresa == empresa)
        if usuario:
            query = query.filter(SolicitacoesPreenchidas.usuario == usuario)
        if data_inicio:
            query = query.filter(SolicitacoesPreenchidas.data_preenchimento >= data_inicio)
        if data_fim:
            data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
            query = query.filter(SolicitacoesPreenchidas.data_preenchimento <= data_fim_ajustada)
        
        preenchimentos = query.order_by(SolicitacoesPreenchidas.data_preenchimento.desc()).all()
        
        preenchimentos_por_material = {}
        for p in preenchimentos:
            material_nome = p.solicitacao.material.DescricaoMaterial if p.solicitacao.material else 'N/A'
            if material_nome not in preenchimentos_por_material:
                preenchimentos_por_material[material_nome] = []
            preenchimentos_por_material[material_nome].append({
                'id': p.id,
                'fornecedor_nome': get_fornecedor_nome(p.fornecedor_id),
                'solicitacao': p.solicitacao,
                'valor_unitario': p.valor_unitario,
                'valor_frete': p.valor_frete,
                'valor_total': p.valor_total,
                'prazo_entrega': p.prazo_entrega,
                'condicao_pagamento': p.condicao_pagamento,
                'status': p.status,
                'usuario': p.usuario,
                'pdf_path': p.pdf_path,
                'historico_descontos': [h.to_dict() for h in p.historico_descontos],
                'observacoes': p.observacoes  # Adicione esta linha
            })
        
        empresas = db.session.query(SolicitacoesCompra.empresa).distinct().order_by(SolicitacoesCompra.empresa).all()
        usuarios = db.session.query(SolicitacoesPreenchidas.usuario).distinct().order_by(SolicitacoesPreenchidas.usuario).all()
        
        return render_template(
            'listar_solicitacoes_preenchidas.html',
            preenchimentos_por_material=preenchimentos_por_material,
            empresas=[e[0] for e in empresas if e[0]],
            usuarios=[u[0] for u in usuarios if u[0]],
            filtros={
                'empresa': empresa,
                'usuario': usuario,
                'data_inicio': data_inicio,
                'data_fim': data_fim
            }
        )
    
    except Exception as e:
        flash(f'Erro ao carregar solicitações preenchidas: {str(e)}', 'error')
        app.logger.error(f'Erro em listar_solicitacoes_preenchidas: {str(e)}', exc_info=True)
        return render_template(
            'listar_solicitacoes_preenchidas.html',
            preenchimentos_por_material={},
            empresas=[],
            usuarios=[],
            filtros={}
        )

def get_fornecedor_nome(fornecedor_id):
    conn = get_db_connection(DB_PATH_FORNECEDORES)
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute('SELECT nome_fantasia FROM fornecedores WHERE id = ?', (fornecedor_id,))
            result = cursor.fetchone()
            return result['nome_fantasia'] if result else 'Fornecedor não encontrado'
        finally:
            conn.close()
    return 'Fornecedor não encontrado'

@routes_bp.route('/download_pdf/<int:preenchimento_id>', methods=['GET'])
def download_pdf(preenchimento_id):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        preenchimento = app.jinja_env.globals['SolicitacoesPreenchidas'].query.get_or_404(preenchimento_id)
        if not preenchimento.pdf_path or not os.path.exists(preenchimento.pdf_path):
            flash('Arquivo PDF não encontrado.', 'error')
            return redirect(url_for('routes_bp.listar_solicitacoes_preenchidas'))
        
        return send_file(preenchimento.pdf_path, as_attachment=True)
    except Exception as e:
        flash(f'Erro ao baixar PDF: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_solicitacoes_preenchidas'))

@routes_bp.route('/atualizar_status_preenchimento/<int:id>', methods=['POST'])
def atualizar_status_preenchimento(id):
    if 'usuario' not in session:
        return jsonify({'success': False, 'message': 'Usuário não autenticado'}), 401

    try:
        preenchimento = SolicitacoesPreenchidas.query.get_or_404(id)
        novo_status = request.form.get('status')

        print(f"DEBUG >> Atualizando status para: {novo_status}")

        # Lógica para reverter status (status especial 'Reaberto')
        if novo_status == 'Reaberto':
            # Verificar se o status atual permite reverter
            if preenchimento.status in ['Aprovado', 'Reprovado']:
                preenchimento.status = 'Aguardando Aprovacao'  # Volta para o status inicial
                db.session.commit()
                print(f"DEBUG >> Status revertido para: Aguardando Aprovacao")
                
                return jsonify({
                    'success': True, 
                    'message': 'Status revertido com sucesso!'
                })
            else:
                return jsonify({
                    'success': False, 
                    'message': 'Não é possível reverter este status'
                }), 400

        # Lógica normal de atualização de status
        if novo_status:
            preenchimento.status = novo_status

        db.session.commit()
        
        # Log para debug
        print(f"DEBUG >> Status atualizado: {preenchimento.status}")
        
        return jsonify({
            'success': True, 
            'message': 'Status atualizado com sucesso!'
        })

    except Exception as e:
        db.session.rollback()
        print(f"ERRO >> {str(e)}")
        return jsonify({
            'success': False, 
            'message': f'Erro ao atualizar status: {str(e)}'
        }), 500
    
@routes_bp.route('/gerar_pedido_compra', methods=['GET', 'POST'])
def gerar_pedido_compra():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        if request.method == 'POST':
            # Obter dados do formulário (já existe)
            preenchimento_ids = request.form.getlist('preenchimento_ids')
            forma_pagamento = request.form.get('forma_pagamento', '').strip()
            condicao_pagamento = request.form.get('condicao_pagamento', '').strip()
            # ADICIONE ESTA LINHA PARA CAPTURAR AS OBSERVAÇÕES
            observacoes = request.form.get('observacoes', '').strip()

            # Validações básicas
            if not preenchimento_ids:
                flash('Nenhum preenchimento selecionado.', 'error')
                return redirect(url_for('routes_bp.gerar_pedido_compra'))

            if not forma_pagamento or not condicao_pagamento:
                flash('Forma e condição de pagamento são obrigatórias.', 'error')
                return redirect(url_for('routes_bp.gerar_pedido_compra'))

            # Obter preenchimentos e validar status
            preenchimentos = SolicitacoesPreenchidas.query.filter(
                SolicitacoesPreenchidas.id.in_(preenchimento_ids)
            ).all()

            for preenchimento in preenchimentos:
                if preenchimento.status != 'Aprovado':
                    flash(f'O preenchimento ID {preenchimento.id} não está aprovado.', 'error')
                    return redirect(url_for('routes_bp.gerar_pedido_compra'))

            # Gerar número do pedido sequencial
            ultimo_pedido = PedidosCompra.query.order_by(PedidosCompra.id.desc()).first()
            proximo_numero = (ultimo_pedido.id + 1) if ultimo_pedido else 1
            numero_pedido = f"PC{datetime.now().year}{proximo_numero:04d}"

            # Calcular valores totais
            valor_total = sum(p.valor_total for p in preenchimentos)
            valor_frete_total = sum(p.valor_frete if p.valor_frete is not None else 0 for p in preenchimentos)
            valor_liquido = valor_total - valor_frete_total

            # Verificar/atualizar estrutura da tabela
            try:
                conn = sqlite3.connect(DATABASE)
                cursor = conn.cursor()
                cursor.execute("PRAGMA table_info(PedidosCompra)")
                columns = [col[1] for col in cursor.fetchall()]
                
                if 'valor_liquido' not in columns:
                    cursor.execute("ALTER TABLE PedidosCompra ADD COLUMN valor_liquido REAL NOT NULL DEFAULT 0")
                    conn.commit()
                    logging.info("Coluna valor_liquido adicionada à tabela PedidosCompra")
                
                conn.close()
            except Exception as e:
                logging.error(f"Erro ao verificar tabela: {str(e)}")

            # Obter informações dos fornecedores
            fornecedor_ids = {p.fornecedor_id for p in preenchimentos}
            fornecedores_info = {}
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute(f'''
                        SELECT id, nome_fantasia, cnpj, telefone, email, endereco, cidade, estado 
                        FROM fornecedores 
                        WHERE id IN ({",".join("?"*len(fornecedor_ids))})
                    ''', list(fornecedor_ids))
                    
                    for row in cursor.fetchall():
                        fornecedores_info[row[0]] = {
                            'nome': row[1],
                            'cnpj': format_cnpj(row[2]) if row[2] else 'N/A',
                            'telefone': row[3],
                            'email': row[4],
                            'endereco': f"{row[5]}, {row[6]}/{row[7]}"
                        }
                finally:
                    conn.close()

            # Criar o pedido de compra
            try:
                pedido = PedidosCompra(
                    numero_pedido=numero_pedido,
                    usuario=session['usuario'],
                    status='Gerado',
                    valor_total=valor_total,
                    valor_frete=valor_frete_total if valor_frete_total != 0 else None,
                    valor_liquido=valor_liquido,
                    forma_pagamento=f"{forma_pagamento} - {condicao_pagamento}",
                    # ADICIONE O CAMPO OBSERVAÇÕES
                    observacoes=observacoes,
                    data_criacao=datetime.now()
                )
                
                pedido.preenchimentos = preenchimentos
                db.session.add(pedido)
                
                for preenchimento in preenchimentos:
                    preenchimento.status = 'Em Processamento'
                
                db.session.commit()
            except Exception as e:
                db.session.rollback()
                flash(f'Erro ao criar pedido: {str(e)}', 'error')
                return redirect(url_for('routes_bp.gerar_pedido_compra'))

            # Preparar dados para o PDF
            materiais_por_fornecedor = {}
            for preenchimento in preenchimentos:
                fornecedor_id = preenchimento.fornecedor_id
                if fornecedor_id not in materiais_por_fornecedor:
                    materiais_por_fornecedor[fornecedor_id] = {
                        'info': fornecedores_info.get(fornecedor_id, {
                            'nome': 'Fornecedor não encontrado',
                            'cnpj': 'N/A',
                            'telefone': 'N/A',
                            'email': 'N/A',
                            'endereco': 'N/A'
                        }),
                        'itens': []
                    }
                
                materiais_por_fornecedor[fornecedor_id]['itens'].append({
                    'material': preenchimento.solicitacao.material.DescricaoMaterial,
                    'marca': preenchimento.solicitacao.marca or 'Não especificado',
                    'especificacao': preenchimento.solicitacao.especificacao,
                    'quantidade': preenchimento.solicitacao.quantidade,
                    'unidade': preenchimento.solicitacao.unidade_medida,
                    'valor_unitario': preenchimento.valor_unitario,
                    'valor_total': preenchimento.valor_total,
                    'prazo_entrega': preenchimento.prazo_entrega,
                    'prioridade': preenchimento.solicitacao.prioridade
                })

            # Geração do PDF - Versão robusta
            try:
                # 1. Verificar/Criar diretório de upload
                upload_dir = app.config['UPLOAD_FOLDER']
                if not os.path.exists(upload_dir):
                    os.makedirs(upload_dir)
                    logging.info(f"Diretório de upload criado: {upload_dir}")

                # 2. Gerar nome único para o PDF
                pdf_filename = f"pedido_{numero_pedido}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                pdf_path = os.path.join(upload_dir, pdf_filename)

                # 3. Renderizar HTML
                html_content = render_template(
                    'pedido_compra_pdf.html',
                    pedido=pedido,
                    materiais_por_fornecedor=materiais_por_fornecedor,
                    data_criacao=datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                    usuario=session['usuario'],
                    total_itens=len(preenchimento_ids),
                    fornecedores_count=len(materiais_por_fornecedor),
                    # ADICIONE ESTA LINHA PARA PASSAR AS OBSERVAÇÕES PARA O TEMPLATE
                    observacoes=observacoes
                )

               # 4. Configurar wkhtmltopdf (Windows e Linux)
                wkhtmltopdf_paths = [
                    r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe',
                    r'C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe',
                    '/usr/local/bin/wkhtmltopdf',
                    '/usr/bin/wkhtmltopdf'
                ]

                # Verifica se o executável existe em um dos caminhos fixos
                config = None
                for path in wkhtmltopdf_paths:
                    if os.path.exists(path):
                        config = pdfkit.configuration(wkhtmltopdf=path)
                        break

                # Fallback: tenta localizar no PATH do sistema
                if config is None:
                    found_path = shutil.which('wkhtmltopdf')
                    if found_path:
                        config = pdfkit.configuration(wkhtmltopdf=found_path)

                # Se ainda não encontrado, lança erro
                if config is None:
                    raise FileNotFoundError("wkhtmltopdf não encontrado. Verifique se está instalado e no PATH do sistema.")

                # 5. Configurações de conversão
                options = {
                    'encoding': 'UTF-8',
                    'quiet': '',
                    'enable-local-file-access': '',
                    'margin-top': '10mm',
                    'margin-right': '10mm',
                    'margin-bottom': '10mm',
                    'margin-left': '10mm',
                    'footer-center': f'Página [page] de [topage] - {numero_pedido}',
                    'footer-font-size': '8'
                }

                # 6. Gerar PDF
                pdfkit.from_string(html_content, pdf_path, configuration=config, options=options)

                # 7. Verificar se o PDF foi criado
                if not os.path.exists(pdf_path):
                    raise Exception("Arquivo PDF não foi gerado")

                # 8. Atualizar pedido com caminho do PDF
                pedido.pdf_path = pdf_path
                db.session.commit()

                flash(f'Pedido {numero_pedido} gerado com sucesso!', 'success')
                return redirect(url_for('routes_bp.listar_pedidos_compra'))

            except Exception as e:
                logging.error(f"Erro ao gerar PDF: {str(e)}", exc_info=True)
                
                # Salvar HTML para debug
                try:
                    debug_path = os.path.join(upload_dir, f"debug_{numero_pedido}.html")
                    with open(debug_path, 'w', encoding='utf-8') as f:
                        f.write(html_content)
                    logging.info(f"HTML de debug salvo em: {debug_path}")
                except Exception as debug_error:
                    logging.error(f"Erro ao salvar HTML: {str(debug_error)}")

                flash('Pedido criado, mas ocorreu um erro ao gerar o PDF. Consulte os logs.', 'warning')
                return redirect(url_for('routes_bp.listar_pedidos_compra'))

        # Método GET - mostrar formulário
        preenchimentos = SolicitacoesPreenchidas.query.filter_by(status='Aprovado')\
            .join(SolicitacoesCompra)\
            .join(Materiais)\
            .order_by(Materiais.DescricaoMaterial)\
            .all()

        # Obter informações dos fornecedores
        fornecedor_ids = {p.fornecedor_id for p in preenchimentos}
        fornecedores = {}
        if fornecedor_ids:
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute(f'''
                        SELECT id, nome_fantasia, cnpj, telefone, email 
                        FROM fornecedores 
                        WHERE id IN ({",".join("?"*len(fornecedor_ids))})
                    ''', list(fornecedor_ids))
                    
                    for row in cursor.fetchall():
                        fornecedores[row[0]] = {
                            'nome': row[1],
                            'cnpj': format_cnpj(row[2]) if row[2] else 'N/A',
                            'telefone': row[3],
                            'email': row[4]
                        }
                finally:
                    conn.close()

        # Organizar preenchimentos por material
        preenchimentos_por_material = {}
        for preenchimento in preenchimentos:
            fornecedor_info = fornecedores.get(preenchimento.fornecedor_id, {})
            preenchimento.fornecedor_nome = fornecedor_info.get('nome', 'Fornecedor não encontrado')
            preenchimento.fornecedor_cnpj = fornecedor_info.get('cnpj', 'N/A')
            preenchimento.fornecedor_telefone = fornecedor_info.get('telefone', 'N/A')
            
            material_nome = preenchimento.solicitacao.material.DescricaoMaterial
            if material_nome not in preenchimentos_por_material:
                preenchimentos_por_material[material_nome] = []
            preenchimentos_por_material[material_nome].append(preenchimento)

        return render_template(
            'gerar_pedido_compra.html',
            preenchimentos_por_material=preenchimentos_por_material,
            formas_pagamento=['À Vista', 'A Prazo', 'Boleto', 'Cartão de Crédito', 'Transferência Bancária']
        )

    except Exception as e:
        db.session.rollback()
        logging.error(f"Erro em gerar_pedido_compra: {str(e)}", exc_info=True)
        flash(f'Erro ao gerar pedido: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_solicitacoes_preenchidas'))
    
@routes_bp.route('/download_comprovante/<int:pedido_id>', methods=['GET'])
def download_comprovante(pedido_id):
    """Endpoint para download seguro de comprovantes de pedidos"""
    
    # Verificação de autenticação
    if 'usuario' not in session:
        flash('🔒 Acesso restrito. Por favor, faça login.', 'error')
        return redirect(url_for('routes_bp.login'))

    try:
        # Busca o pedido com tratamento de 404
        pedido = PedidosCompra.query.get_or_404(pedido_id)
        
        # Verifica existência do comprovante
        if not pedido.comprovante_pagamento:
            flash('📄 Este pedido não possui comprovante cadastrado.', 'info')
            return redirect(url_for('routes_bp.listar_pedidos_compra'))

        # Configurações de caminho
        uploads_dir = os.path.abspath(app.config['UPLOAD_FOLDER'])
        file_to_download = None

        # Tentativa 1: Caminho absoluto/relativo do banco
        stored_path = os.path.normpath(pedido.comprovante_pagamento)
        if os.path.isabs(stored_path):
            candidate_path = stored_path
        else:
            candidate_path = os.path.join(uploads_dir, stored_path)

        if os.path.isfile(candidate_path):
            file_to_download = candidate_path

        # Tentativa 2: Busca por padrão conhecido (se necessário)
        if not file_to_download:
            from glob import glob
            pattern = os.path.join(uploads_dir, f'pedido_{pedido_id}_*')
            matching_files = glob(pattern)
            
            if matching_files:
                file_to_download = matching_files[0]

        # Validação final do arquivo
        if not file_to_download or not os.path.isfile(file_to_download):
            flash('⚠️ Comprovante não encontrado no servidor.', 'warning')
            return redirect(url_for('routes_bp.listar_pedidos_compra'))

        # Verificação de segurança
        if not os.path.abspath(file_to_download).startswith(uploads_dir):
            logging.warning(f'Tentativa de acesso a caminho não autorizado: {file_to_download}')
            flash('⛔ Caminho do arquivo inválido.', 'danger')
            return redirect(url_for('routes_bp.listar_pedidos_compra'))

        # Preparação do download
        filename = f"COMPROVANTE_{pedido.numero_pedido}_{os.path.basename(file_to_download)}"
        
        return send_file(
            file_to_download,
            as_attachment=True,
            download_name=filename,
            mimetype='application/pdf',
            conditional=True
        )
        
    except Exception as e:
        logging.error(f'❌ Erro ao baixar comprovante {pedido_id}: {str(e)}', exc_info=True)
        flash('⛔ Erro inesperado ao baixar comprovante.', 'danger')
        return redirect(url_for('routes_bp.listar_pedidos_compra'))
    
@routes_bp.route('/listar_pedidos_compra', methods=['GET'])
def listar_pedidos_compra():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Obter parâmetros de filtro
        status = request.args.get('status')
        empresa_filtro = request.args.get('empresa')
        usuario_filtro = request.args.get('usuario')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')

        # Ler usuários e empresas do arquivo senhas.txt
        usuarios_empresas = {}
        empresas_unicas = set()
        usuarios_unicos = set()
        
        try:
            with open('senhas.txt', 'r', encoding='utf-8') as f:
                for line in f:
                    partes = line.strip().split('%')
                    if len(partes) >= 4:  # Verifica se tem pelo menos 4 partes
                        usuario = partes[0]
                        empresa = partes[3]
                        usuarios_empresas[usuario] = empresa
                        empresas_unicas.add(empresa)
                        usuarios_unicos.add(usuario)
        except Exception as e:
            logging.error(f"Erro ao ler senhas.txt: {str(e)}")
            usuarios_empresas = {}

        # Consulta base
        query = db.session.query(PedidosCompra).join(
            pedido_preenchimento_associacao,
            PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
        ).join(
            SolicitacoesPreenchidas,
            pedido_preenchimento_associacao.c.preenchimento_id == SolicitacoesPreenchidas.id
        ).join(
            SolicitacoesCompra,
            SolicitacoesPreenchidas.solicitacao_id == SolicitacoesCompra.id
        )

        # Aplicar filtros
        if status:
            query = query.filter(PedidosCompra.status == status)
        
        if empresa_filtro:
            query = query.filter(
                (SolicitacoesCompra.empresa == empresa_filtro) |
                (PedidosCompra.usuario.in_(
                    [usuario for usuario, empresa in usuarios_empresas.items() if empresa == empresa_filtro]
                ))
            )
        
        if usuario_filtro:
            query = query.filter(PedidosCompra.usuario == usuario_filtro)
        
        if data_inicio:
            query = query.filter(PedidosCompra.data_criacao >= data_inicio)
        
        if data_fim:
            # Adiciona 1 dia para incluir todo o dia final
            data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
            query = query.filter(PedidosCompra.data_criacao <= data_fim_ajustada)

        # Ordenar e executar a consulta
        pedidos = query.order_by(
            PedidosCompra.data_criacao.desc()
        ).distinct().all()

        # Obter informações de fornecedores e marcas
        pedidos_completos = []
        fornecedor_ids = set()
        
        for pedido in pedidos:
            for preenchimento in pedido.preenchimentos:
                fornecedor_ids.add(preenchimento.fornecedor_id)
        
        # Buscar fornecedores
        fornecedores = {}
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute(f'SELECT id, nome_fantasia, cnpj FROM fornecedores WHERE id IN ({",".join("?"*len(fornecedor_ids))})', 
                             list(fornecedor_ids))
                for row in cursor.fetchall():
                    fornecedores[row[0]] = {
                        'nome_fantasia': row[1],
                        'cnpj': format_cnpj(row[2]) if row[2] else 'N/A'
                    }
            finally:
                conn.close()

        # Estruturar dados para o template
        for pedido in pedidos:
            preenchimentos_info = []
            for preenchimento in pedido.preenchimentos:
                fornecedor_info = fornecedores.get(preenchimento.fornecedor_id, {
                    'nome_fantasia': 'Fornecedor não encontrado',
                    'cnpj': 'N/A'
                })
                
                # Obter empresa do usuário do arquivo senhas.txt ou usar a da solicitação como fallback
                empresa_usuario = usuarios_empresas.get(pedido.usuario, preenchimento.solicitacao.empresa)
                
                preenchimentos_info.append({
                    'id': preenchimento.id,
                    'marca': preenchimento.marca or 'Não informado',
                    'fornecedor_nome': fornecedor_info['nome_fantasia'],
                    'fornecedor_cnpj': fornecedor_info.get('cnpj', 'N/A'),
                    'material': preenchimento.solicitacao.material.DescricaoMaterial if preenchimento.solicitacao.material else 'N/A',
                    'empresa': empresa_usuario
                })
            pedidos_completos.append({
                'pedido': pedido,
                'preenchimentos': preenchimentos_info,
                'observacoes': pedido.observacoes
            })

        return render_template(
            'listar_pedidos_compra.html', 
            pedidos_completos=pedidos_completos,
            empresas=sorted(empresas_unicas),
            usuarios=sorted(usuarios_unicos),
            filtros={
                'empresa': empresa_filtro,
                'usuario': usuario_filtro,
                'data_inicio': data_inicio,
                'data_fim': data_fim,
                'status': status
            }
        )
    except Exception as e:
        flash(f'Erro ao listar pedidos de compra: {str(e)}', 'error')
        return render_template(
            'listar_pedidos_compra.html', 
            pedidos_completos=[],
            empresas=[],
            usuarios=[],
            filtros={}
        )
    
#Nova Pagina pagamento
@routes_bp.route('/listar_pedidos_compras_pg', methods=['GET'])
def listar_pedidos_compra_pg():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Obter parâmetros de filtro (exceto status, pois queremos apenas AgPagamento)
        empresa_filtro = request.args.get('empresa')
        usuario_filtro = request.args.get('usuario')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')

        # Ler usuários e empresas do arquivo senhas.txt
        usuarios_empresas = {}
        empresas_unicas = set()
        usuarios_unicos = set()
        
        try:
            with open('senhas.txt', 'r', encoding='utf-8') as f:
                for line in f:
                    partes = line.strip().split('%')
                    if len(partes) >= 4:
                        usuario = partes[0]
                        empresa = partes[3]
                        usuarios_empresas[usuario] = empresa
                        empresas_unicas.add(empresa)
                        usuarios_unicos.add(usuario)
        except Exception as e:
            logging.error(f"Erro ao ler senhas.txt: {str(e)}")
            usuarios_empresas = {}

        # Consulta base - FILTRANDO APENAS POR AgPagamento
        query = db.session.query(PedidosCompra).filter(
            PedidosCompra.status == 'AgPagamento'  # Filtro fixo para AgPagamento
        ).join(
            pedido_preenchimento_associacao,
            PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
        ).join(
            SolicitacoesPreenchidas,
            pedido_preenchimento_associacao.c.preenchimento_id == SolicitacoesPreenchidas.id
        ).join(
            SolicitacoesCompra,
            SolicitacoesPreenchidas.solicitacao_id == SolicitacoesCompra.id
        )

        # Aplicar outros filtros (exceto status)
        if empresa_filtro:
            query = query.filter(
                (SolicitacoesCompra.empresa == empresa_filtro) |
                (PedidosCompra.usuario.in_(
                    [usuario for usuario, empresa in usuarios_empresas.items() if empresa == empresa_filtro]
                ))
            )
        
        if usuario_filtro:
            query = query.filter(PedidosCompra.usuario == usuario_filtro)
        
        if data_inicio:
            query = query.filter(PedidosCompra.data_criacao >= data_inicio)
        
        if data_fim:
            data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
            query = query.filter(PedidosCompra.data_criacao <= data_fim_ajustada)

        # Ordenar e executar a consulta
        pedidos = query.order_by(
            PedidosCompra.data_criacao.desc()
        ).distinct().all()

        # Restante do código permanece igual...
        fornecedor_ids = set()
        for pedido in pedidos:
            for preenchimento in pedido.preenchimentos:
                fornecedor_ids.add(preenchimento.fornecedor_id)
        
        fornecedores = {}
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute(f'SELECT id, nome_fantasia, cnpj FROM fornecedores WHERE id IN ({",".join("?"*len(fornecedor_ids))})', 
                             list(fornecedor_ids))
                for row in cursor.fetchall():
                    fornecedores[row[0]] = {
                        'nome_fantasia': row[1],
                        'cnpj': format_cnpj(row[2]) if row[2] else 'N/A'
                    }
            finally:
                conn.close()

        pedidos_completos = []
        for pedido in pedidos:
            preenchimentos_info = []
            for preenchimento in pedido.preenchimentos:
                fornecedor_info = fornecedores.get(preenchimento.fornecedor_id, {
                    'nome_fantasia': 'Fornecedor não encontrado',
                    'cnpj': 'N/A'
                })
                empresa_usuario = usuarios_empresas.get(pedido.usuario, preenchimento.solicitacao.empresa)
                
                preenchimentos_info.append({
                    'id': preenchimento.id,
                    'marca': preenchimento.marca or 'Não informado',
                    'fornecedor_nome': fornecedor_info['nome_fantasia'],
                    'fornecedor_cnpj': fornecedor_info.get('cnpj', 'N/A'),
                    'material': preenchimento.solicitacao.material.DescricaoMaterial if preenchimento.solicitacao.material else 'N/A',
                    'empresa': empresa_usuario
                })
            pedidos_completos.append({
                'pedido': pedido,
                'preenchimentos': preenchimentos_info
            })

        return render_template(
            'listar_pedidos_compras_pg.html', 
            pedidos_completos=pedidos_completos,
            empresas=sorted(empresas_unicas),
            usuarios=sorted(usuarios_unicos),
            filtros={
                'empresa': empresa_filtro,
                'usuario': usuario_filtro,
                'data_inicio': data_inicio,
                'data_fim': data_fim,
                'status': 'AgPagamento'  # Definindo o status fixo para o template
            }
        )
    except Exception as e:
        flash(f'Erro ao listar pedidos de compra: {str(e)}', 'error')
        return render_template(
            'listar_pedidos_compras_pg.html', 
            pedidos_completos=[],
            empresas=[],
            usuarios=[],
            filtros={}
        )
    
@routes_bp.route('/download_pedido_pdf/<int:pedido_id>', methods=['GET'])
def download_pedido_pdf(pedido_id):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        pedido = app.jinja_env.globals['PedidosCompra'].query.get_or_404(pedido_id)
        if not pedido.pdf_path or not os.path.exists(pedido.pdf_path):
            flash('Arquivo PDF não encontrado.', 'error')
            return redirect(url_for('routes_bp.listar_pedidos_compra'))
        
        return send_file(pedido.pdf_path, as_attachment=True)
    except Exception as e:
        flash(f'Erro ao baixar PDF: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_pedidos_compra'))

@routes_bp.route('/atualizar_status_pedido/<int:id>', methods=['POST'])
def atualizar_status_pedido(id):
    if 'usuario' not in session:
        flash('Usuário não autenticado.', 'error')
        return jsonify({'flash': [['Usuário não autenticado.', 'error']]}), 401

    try:
        pedido = app.jinja_env.globals['PedidosCompra'].query.get_or_404(id)
        status = request.form.get('status', '').strip()
        fornecedor_id = request.form.get('fornecedor_id', '').strip()
        numero_nf = request.form.get('numero_nf', '').strip()
        comprovante_pagamento = request.form.get('comprovante', '').strip()
        pdf_file = request.files.get('pdf_file')

        valid_statuses = ['Gerado', 'Aprovado', 'AgPagamento', 'AgEntrega', 'Entregue', 'Reprovado']
        if not status or status not in valid_statuses:
            flash(f'Status inválido: {status}', 'error')
            return jsonify({'flash': [['Status inválido.', 'error']]}), 400

        # Lógica para pedidos à vista
        if pedido.forma_pagamento and 'À Vista' in pedido.forma_pagamento:
            if status == 'Aprovado':
                # Para pedidos à vista, aprovar muda automaticamente para AgPagamento
                pedido.status = 'AgPagamento'
                flash('Pedido aprovado e movido para Aguardando Pagamento.', 'success')
            elif status == 'AgPagamento':
                # Validar comprovante e PDF
                if not comprovante_pagamento:
                    flash('Número do comprovante é obrigatório para Aguardando Pagamento.', 'error')
                    return jsonify({'flash': [['Número do comprovante é obrigatório.', 'error']]}), 400
                if not pdf_file or not pdf_file.filename.endswith('.pdf'):
                    flash('Anexe um arquivo PDF válido para Aguardando Pagamento.', 'error')
                    return jsonify({'flash': [['Anexe um arquivo PDF válido.', 'error']]}), 400
                if pdf_file.content_length > 5 * 1024 * 1024:  # Limite de 5MB
                    flash('O arquivo PDF deve ter no máximo 5MB.', 'error')
                    return jsonify({'flash': [['O arquivo PDF deve ter no máximo 5MB.', 'error']]}), 400

                # Salvar PDF
                filename = f"pedido_{id}_{uuid.uuid4().hex[:8]}_{secure_filename(pdf_file.filename)}"
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                try:
                    pdf_file.save(pdf_path)
                except Exception as e:
                    flash(f'Erro ao salvar o PDF: {str(e)}', 'error')
                    return jsonify({'flash': [['Erro ao salvar o PDF.', 'error']]}), 500

                pedido.pdf_path = pdf_path
                pedido.comprovante_pagamento = comprovante_pagamento
                pedido.status = 'AgEntrega'
                flash('Comprovante salvo. Pedido movido para Aguardando Entrega.', 'success')

        # Lógica para pedidos faturados a prazo
        else:
            if status == 'Aprovado':
                # Para pedidos a prazo, aprovar muda diretamente para AgEntrega
                pedido.status = 'AgEntrega'
                flash('Pedido aprovado e movido para Aguardando Entrega.', 'success')

        # Lógica para status Entregue (comum a ambos os tipos de pagamento)
        if status == 'Entregue':
            if not fornecedor_id or not numero_nf:
                flash('Fornecedor e número da NF são obrigatórios para o status Entregue.', 'error')
                return jsonify({'flash': [['Fornecedor e número da NF são obrigatórios.', 'error']]}), 400

            # Validar fornecedor
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute('SELECT id FROM fornecedores WHERE id = ?', (fornecedor_id,))
                    if not cursor.fetchone():
                        flash('Fornecedor inválido.', 'error')
                        return jsonify({'flash': [['Fornecedor inválido.', 'error']]}), 400
                finally:
                    conn.close()

            # Atualizar preenchimentos e estoque
            for preenchimento in pedido.preenchimentos:
                preenchimento.status = 'Entregue'
                preenchimento.fornecedor_id = fornecedor_id
                preenchimento.numero_nf = numero_nf

                material = app.jinja_env.globals['Materiais'].query.filter_by(
                    CodMaterial=preenchimento.solicitacao.cod_material
                ).first()

                if material:
                    material.QuantidadeEstoque += preenchimento.solicitacao.quantidade
                    material.Fornecedor = fornecedor_id
                    material.NumeroNF = numero_nf

                    estoque_existente = app.jinja_env.globals['Estoque'].query.filter_by(
                        preenchimento_id=preenchimento.id
                    ).first()

                    if not estoque_existente:
                        estoque = app.jinja_env.globals['Estoque'](
                            preenchimento_id=preenchimento.id,
                            cod_material=preenchimento.solicitacao.cod_material,
                            quantidade=preenchimento.solicitacao.quantidade,
                            fornecedor=fornecedor_id,
                            numero_nf=numero_nf,
                            usuario=session.get('usuario')
                        )
                        app.jinja_env.globals['db'].session.add(estoque)
                    else:
                        estoque_existente.fornecedor = fornecedor_id
                        estoque_existente.numero_nf = numero_nf
                        estoque_existente.quantidade = preenchimento.solicitacao.quantidade

            pedido.status = status
            flash('Pedido marcado como Entregue. Estoque atualizado.', 'success')

        # Atualizar status para outros casos
        if status not in ['Aprovado', 'AgPagamento', 'Entregue']:
            pedido.status = status
            flash('Status atualizado com sucesso.', 'success')

        app.jinja_env.globals['db'].session.commit()

        # Retornar mensagens flash como JSON
        messages = []
        for message, category in session.get('_flashes', []):
            messages.append([message, category])
        session['_flashes'] = []  # Limpar flashes após envio
        return jsonify({'flash': messages}), 204

    except Exception as e:
        app.jinja_env.globals['db'].session.rollback()
        flash(f'Erro interno ao atualizar status: {str(e)}', 'error')
        return jsonify({'flash': [['Erro interno ao atualizar status.', 'error']]}), 500
    
@routes_bp.route('/adicionar_estoque/<int:preenchimento_id>', methods=['GET', 'POST'])
def adicionar_estoque(preenchimento_id):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        preenchimento = app.jinja_env.globals['SolicitacoesPreenchidas'].query.get_or_404(preenchimento_id)
        if preenchimento.status != 'Entregue':
            flash('O status deve ser "Entregue" para adicionar ao estoque.', 'error')
            return redirect(url_for('routes_bp.listar_solicitacoes_preenchidas'))
        
        if request.method == 'POST':
            fornecedor = request.form.get('fornecedor', '').strip()
            numero_nf = request.form.get('numero_nf', '').strip()

            if not fornecedor or not numero_nf:
                flash('Fornecedor e número da NF são obrigatórios.', 'error')
                return render_template('adicionar_estoque.html', preenchimento=preenchimento)
            
            material = app.jinja_env.globals['Materiais'].query.filter_by(CodMaterial=preenchimento.solicitacao.cod_material).first()
            if material:
                material.QuantidadeEstoque += preenchimento.solicitacao.quantidade
                material.Fornecedor = fornecedor
                material.NumeroNF = numero_nf
            else:
                material = app.jinja_env.globals['Materiais'](
                    CodMaterial=preenchimento.solicitacao.cod_material,
                    DescricaoMaterial=preenchimento.solicitacao.material.DescricaoMaterial,
                    Empresa=preenchimento.solicitacao.empresa,
                    Aplicacao=preenchimento.solicitacao.material.Aplicacao,
                    QuantidadeEstoque=preenchimento.solicitacao.quantidade,
                    Fornecedor=fornecedor,
                    NumeroNF=numero_nf,
                    FatorConsumo=0.0
                )
                app.jinja_env.globals['db'].session.add(material)
            
            estoque = app.jinja_env.globals['Estoque'](
                preenchimento_id=preenchimento_id,
                cod_material=preenchimento.solicitacao.cod_material,
                quantidade=preenchimento.solicitacao.quantidade,
                fornecedor=fornecedor,
                numero_nf=numero_nf,
                usuario=session['usuario']
            )
            app.jinja_env.globals['db'].session.add(estoque)
            app.jinja_env.globals['db'].session.commit()

            flash('Material adicionado ao estoque com sucesso.', 'success')
            return redirect(url_for('routes_bp.listar_estoque'))
        
        return render_template('adicionar_estoque.html', preenchimento=preenchimento)
    except Exception as e:
        flash(f'Erro ao adicionar ao estoque: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_solicitacoes_preenchidas'))

@routes_bp.route('/requisitar_material/<int:preenchimento_id>', methods=['GET', 'POST'])
def requisitar_material(preenchimento_id):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        preenchimento = app.jinja_env.globals['SolicitacoesPreenchidas'].query.get_or_404(preenchimento_id)
        estoque = app.jinja_env.globals['Estoque'].query.filter_by(preenchimento_id=preenchimento_id).first()
        
        if not estoque:
            flash('Material não está no estoque.', 'error')
            return redirect(url_for('routes_bp.listar_estoque'))
        
        if request.method == 'POST':
            quantidade = request.form.get('quantidade', '').strip()

            if not quantidade:
                flash('Quantidade é obrigatória.', 'error')
                return render_template('requisitar_material.html', preenchimento=preenchimento, estoque=estoque)
            
            try:
                quantidade = int(quantidade)
                if quantidade <= 0:
                    flash('Quantidade deve ser um número positivo.', 'error')
                    return render_template('requisitar_material.html', preenchimento=preenchimento, estoque=estoque)
                
                if quantidade > estoque.quantidade:
                    flash(f'Quantidade requisitada ({quantidade}) excede o estoque disponível ({estoque.quantidade}).', 'error')
                    return render_template('requisitar_material.html', preenchimento=preenchimento, estoque=estoque)
            except ValueError:
                flash('Quantidade deve ser um número válido.', 'error')
                return render_template('requisitar_material.html', preenchimento=preenchimento, estoque=estoque)
            
            material = app.jinja_env.globals['Materiais'].query.filter_by(CodMaterial=preenchimento.solicitacao.cod_material).first()
            if not material:
                flash(f'Material com código {preenchimento.solicitacao.cod_material} não encontrado.', 'error')
                return render_template('requisitar_material.html', preenchimento=preenchimento, estoque=estoque)
            
            if quantidade > material.QuantidadeEstoque:
                flash(f'Quantidade requisitada ({quantidade}) excede o estoque total do material ({material.QuantidadeEstoque}).', 'error')
                return render_template('requisitar_material.html', preenchimento=preenchimento, estoque=estoque)
            
            ticket = str(uuid.uuid4())[:8]
            requisicao = app.jinja_env.globals['Requisicoes'](
                preenchimento_id=preenchimento_id,
                cod_material=preenchimento.solicitacao.cod_material,
                quantidade=quantidade,
                ticket=ticket,
                usuario=session['usuario']
            )
            
            estoque.quantidade -= quantidade
            material.QuantidadeEstoque -= quantidade
            
            app.jinja_env.globals['db'].session.add(requisicao)
            app.jinja_env.globals['db'].session.commit()

            flash(f'Requisição realizada com sucesso. Ticket: {ticket}', 'success')
            return redirect(url_for('routes_bp.listar_requisicoes'))
        
        return render_template('requisitar_material.html', preenchimento=preenchimento, estoque=estoque)
    except Exception as e:
        app.jinja_env.globals['db'].session.rollback()
        flash(f'Erro ao requisitar material: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_estoque'))

@routes_bp.route('/listar_estoque', methods=['GET'])
def listar_estoque():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Busca todos os materiais com estoque positivo
        materiais = Materiais.query.filter(Materiais.QuantidadeEstoque > 0).all()
        
        estoques = []
        for material in materiais:
            # Tenta encontrar um registro de estoque
            estoque = Estoque.query.filter_by(cod_material=material.CodMaterial).first()
            
            estoques.append({
                'material': material,
                'preenchimento_id': estoque.preenchimento_id if estoque else None
            })
            
        return render_template('listar_estoque.html', estoques=estoques)
    except Exception as e:
        flash(f'Erro ao carregar estoque: {str(e)}', 'error')
        return render_template('listar_estoque.html', estoques=[])
    
#Auditoria tela de listar estoque
@routes_bp.route('/listar_estoque_auditor', methods=['GET'])
def listar_estoque_auditor():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Busca todos os materiais com estoque positivo
        materiais = Materiais.query.filter(Materiais.QuantidadeEstoque > 0).all()
        
        estoques = []
        for material in materiais:
            # Tenta encontrar um registro de estoque
            estoque = Estoque.query.filter_by(cod_material=material.CodMaterial).first()
            
            estoques.append({
                'material': material,
                'preenchimento_id': estoque.preenchimento_id if estoque else None
            })
            
        return render_template('listar_estoque_auditor.html', estoques=estoques)
    except Exception as e:
        flash(f'Erro ao carregar estoque: {str(e)}', 'error')
        return render_template('listar_estoque_auditor.html', estoques=[])

@routes_bp.route('/requisitar_manual/<int:cod_material>', methods=['GET', 'POST'])
def requisitar_manual(cod_material):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        material = Materiais.query.get_or_404(cod_material)
        
        if request.method == 'POST':
            quantidade = request.form.get('quantidade', '').strip()
            
            if not quantidade:
                flash('Quantidade é obrigatória.', 'error')
                return render_template('requisitar_manual.html', material=material)
            
            try:
                quantidade = int(quantidade)
                if quantidade <= 0:
                    flash('Quantidade deve ser um número positivo.', 'error')
                    return render_template('requisitar_manual.html', material=material)
                
                if quantidade > material.QuantidadeEstoque:
                    flash(f'Quantidade requisitada ({quantidade}) excede o estoque disponível ({material.QuantidadeEstoque}).', 'error')
                    return render_template('requisitar_manual.html', material=material)
                
                # Cria um ticket único
                ticket = str(uuid.uuid4())[:8]
                
                # Cria a requisição com preenchimento_id como NULL
                requisicao = Requisicoes(
                    preenchimento_id=None,  # Explicitamente definido como NULL
                    cod_material=material.CodMaterial,
                    quantidade=quantidade,
                    ticket=ticket,
                    usuario=session['usuario']
                )
                
                # Atualiza o estoque
                material.QuantidadeEstoque -= quantidade
                
                db.session.add(requisicao)
                db.session.commit()

                flash(f'Requisição realizada com sucesso. Ticket: {ticket}', 'success')
                return redirect(url_for('routes_bp.listar_requisicoes'))
                
            except ValueError:
                flash('Quantidade deve ser um número válido.', 'error')
                return render_template('requisitar_manual.html', material=material)
        
        return render_template('requisitar_manual.html', material=material)
        
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao requisitar material: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_estoque'))

@routes_bp.route('/listar_requisicoes', methods=['GET'])
def listar_requisicoes():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        requisicoes = app.jinja_env.globals['Requisicoes'].query.all()
        for requisicao in requisicoes:
            estoque = app.jinja_env.globals['Estoque'].query.filter_by(cod_material=requisicao.cod_material).first()
            if estoque:
                requisicao.fornecedor = estoque.fornecedor
                requisicao.numero_nf = estoque.numero_nf
            else:
                requisicao.fornecedor = 'Não disponível'
                requisicao.numero_nf = 'Não disponível'
        return render_template('listar_requisicoes.html', requisicoes=requisicoes)
    except Exception as e:
        flash(f'Erro ao carregar requisições: {str(e)}', 'error')
        return render_template('listar_requisicoes.html', requisicoes=[])
    
def atualizar_estrutura_requisicoes():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Verifica se a coluna já permite NULL
        cursor.execute("PRAGMA table_info(Requisicoes)")
        columns = cursor.fetchall()
        preenchimento_id_col = next((col for col in columns if col[1] == 'preenchimento_id'), None)
        
        if preenchimento_id_col and preenchimento_id_col[3] == 1:  # 1 significa NOT NULL
            print("Atualizando estrutura da tabela Requisicoes...")
            
            # Executa cada comando separadamente
            try:
                # 1. Criar tabela temporária
                cursor.execute("CREATE TABLE temp_Requisicoes AS SELECT * FROM Requisicoes")
                
                # 2. Dropar tabela original
                cursor.execute("DROP TABLE Requisicoes")
                
                # 3. Criar nova tabela com estrutura atualizada
                cursor.execute("""
                    CREATE TABLE Requisicoes (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        preenchimento_id INTEGER,
                        cod_material INTEGER NOT NULL,
                        quantidade INTEGER NOT NULL,
                        ticket TEXT NOT NULL,
                        data_requisicao DATETIME NOT NULL,
                        usuario TEXT NOT NULL,
                        FOREIGN KEY (preenchimento_id) REFERENCES SolicitacoesPreenchidas(id),
                        FOREIGN KEY (cod_material) REFERENCES Materiais(CodMaterial)
                    )
                """)
                
                # 4. Copiar dados da tabela temporária
                cursor.execute("""
                    INSERT INTO Requisicoes (id, preenchimento_id, cod_material, quantidade, ticket, data_requisicao, usuario)
                    SELECT id, preenchimento_id, cod_material, quantidade, ticket, data_requisicao, usuario 
                    FROM temp_Requisicoes
                """)
                
                # 5. Dropar tabela temporária
                cursor.execute("DROP TABLE temp_Requisicoes")
                
                conn.commit()
                print("Estrutura da tabela Requisicoes atualizada com sucesso.")
            except sqlite3.Error as e:
                conn.rollback()
                print(f"Erro durante a migração: {str(e)}")
                return False
        else:
            print("A coluna preenchimento_id já permite valores NULL ou não foi encontrada.")
            
        conn.close()
        return True
    except Exception as e:
        print(f"Erro ao atualizar estrutura: {str(e)}")
        return False
    
@routes_bp.route('/pcp_material/<int:cod>', methods=['GET', 'POST'])
def pcp_material(cod):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))

    material = Materiais.query.get_or_404(cod)

    if request.method == 'POST':
        status = request.form.get('status')  # 'pcp', 'pcm' ou 'inativo'
        fator_consumo = request.form.get('fator_consumo', '0').replace(',', '.')

        try:
            fator_consumo_float = float(fator_consumo) if fator_consumo else 0.0
            
            # Lógica das opções
            if status == 'pcp':
                material.Ativo = True
                material.FatorConsumo = fator_consumo_float
                redirect_to = 'listar_pcp'
            elif status == 'pcm':
                material.Ativo = False
                material.FatorConsumo = fator_consumo_float
                redirect_to = 'listar_ativos'
            else:  # inativo
                material.Ativo = False
                material.FatorConsumo = 0.0
                redirect_to = 'listar_estoque'
            
            db.session.commit()
            flash('Configurações salvas com sucesso!', 'success')
            return redirect(url_for(f'routes_bp.{redirect_to}'))

        except ValueError:
            flash('Valor do fator de consumo inválido', 'error')
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao salvar: {str(e)}', 'error')

    # Calcula dias de estoque
    dias_estoque = material.QuantidadeEstoque / material.FatorConsumo if material.FatorConsumo > 0 else 0
    
    # Determina o status atual para seleção no formulário
    status_atual = 'inativo'
    if material.FatorConsumo > 0:
        status_atual = 'pcp' if material.Ativo else 'pcm'
    
    return render_template('pcp_material.html', 
                        material=material,
                        dias_estoque=dias_estoque,
                        status_atual=status_atual)


@routes_bp.route('/listar_pcp', methods=['GET'])
def listar_pcp():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Mostrar apenas materiais PCP (Ativo=True + FatorConsumo>0)
        materiais = Materiais.query.filter(
            Materiais.Ativo == True,
            Materiais.FatorConsumo > 0
        ).all()
        
        for material in materiais:
            material.dias_estoque = material.QuantidadeEstoque / material.FatorConsumo
            
        return render_template('listar_pcp.html', 
                            materiais=materiais,
                            total_materiais=len(materiais))
    except Exception as e:
        flash(f'Erro ao carregar PCP: {str(e)}', 'error')
        return render_template('listar_pcp.html', 
                            materiais=[], 
                            total_materiais=0)

@routes_bp.route('/listar_ativos', methods=['GET'])
def listar_ativos():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Mostrar apenas materiais PCM (Ativo=False + FatorConsumo>0)
        materiais = Materiais.query.filter(
            Materiais.Ativo == False,
            Materiais.FatorConsumo > 0
        ).all()
        
        for material in materiais:
            material.dias_estoque = material.QuantidadeEstoque / material.FatorConsumo
            
        return render_template('ativos.html', 
                            materiais=materiais,
                            total_materiais=len(materiais))
    except Exception as e:
        flash(f'Erro ao carregar PCM: {str(e)}', 'error')
        return render_template('ativos.html', 
                            materiais=[], 
                            total_materiais=0)
    
@routes_bp.route('/cadastro-fornecedores.html', methods=['GET'])
def cadastro_fornecedores():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('cadastro-fornecedores.html')

@routes_bp.route('/lista_fornecedores.html', methods=['GET'])
def lista_fornecedores():
    if 'usuario' not in session:
        flash('Você precisa estar logado para acessar esta página.', 'warning')
        return redirect(url_for('routes_bp.login'))
    
    try:
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if not conn:
            flash('Erro ao conectar ao banco de dados de fornecedores.', 'error')
            app.logger.error('Falha ao conectar ao banco de fornecedores')
            return render_template('lista_fornecedores.html', fornecedores=[])
        
        try:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, nome_fantasia, cnpj, telefone, email, endereco, bairro, cidade, estado, contato, materiais
                FROM fornecedores
                ORDER BY nome_fantasia
            ''')
            fornecedores = [dict(row) for row in cursor.fetchall()]
            
            # Adicionar log para depuração
            app.logger.info(f"Fornecedores recuperados: {fornecedores}")
            
            for fornecedor in fornecedores:
                cnpj = fornecedor['cnpj']
                if cnpj and len(cnpj) == 14:
                    fornecedor['cnpj_formatado'] = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
                else:
                    fornecedor['cnpj_formatado'] = cnpj
            
            return render_template('lista_fornecedores.html', 
                                 fornecedores=fornecedores,
                                 total_fornecedores=len(fornecedores))
            
        except sqlite3.Error as db_error:
            app.logger.error(f'Erro no banco de dados: {str(db_error)}')
            flash('Erro ao acessar os dados dos fornecedores.', 'error')
            return render_template('lista_fornecedores.html', fornecedores=[])
            
        finally:
            if conn:
                conn.close()
                
    except Exception as e:
        app.logger.error(f'Erro inesperado em lista_fornecedores: {str(e)}', exc_info=True)
        flash('Ocorreu um erro inesperado ao listar os fornecedores.', 'error')
        return render_template('lista_fornecedores.html', fornecedores=[])
    
@routes_bp.route('/submit_form', methods=['POST'])
def submit_form():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        nome_fantasia = request.form.get('nomeFantasia', '').strip()
        cnpj = request.form.get('cnpj', '').strip()
        telefone = request.form.get('telefone', '').strip()
        email = request.form.get('email', '').strip()
        endereco = request.form.get('endereco', '').strip()
        bairro = request.form.get('bairro', '').strip()
        cidade = request.form.get('cidade', '').strip()
        estado = request.form.get('estado', '').strip()
        contato = request.form.get('contato', '').strip()
        materiais = request.form.get('materiais', '').strip()

        required_fields = {
            'Nome Fantasia': nome_fantasia,
            'CNPJ': cnpj,
            'Telefone': telefone,
            'Email': email,
            'Endereço': endereco,
            'Bairro': bairro,
            'Cidade': cidade,
            'Estado': estado,
            'Contato': contato,
            'Materiais': materiais
        }
        
        for field_name, field_value in required_fields.items():
            if not field_value:
                flash(f'O campo {field_name} é obrigatório.', 'error')
                return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(nome_fantasia) < 3:
            flash('Nome Fantasia deve ter pelo menos 3 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if not validate_cnpj(cnpj):
            flash('CNPJ inválido. Deve conter 14 dígitos.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if not validate_email(email):
            flash('Email inválido. Insira um email válido.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(telefone) < 10:
            flash('Telefone inválido. Deve conter pelo menos 10 dígitos.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(endereco) < 5:
            flash('Endereço deve ter pelo menos 5 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(bairro) < 3:
            flash('Bairro deve ter pelo menos 3 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(cidade) < 3:
            flash('Cidade deve ter pelo menos 3 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(estado) != 2:
            flash('Estado deve ter exatamente 2 caracteres (UF).', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(contato) < 3:
            flash('Contato deve ter pelo menos 3 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        if len(materiais) < 3:
            flash('Materiais deve ter pelo menos 3 caracteres.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        # Verificar duplicidade de CNPJ
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if not conn:
            flash('Erro ao conectar ao banco de dados.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        cursor = conn.cursor()
        cursor.execute('SELECT cnpj FROM fornecedores WHERE cnpj = ?', (cnpj,))
        if cursor.fetchone():
            conn.close()
            flash('CNPJ já cadastrado.', 'error')
            return redirect(url_for('routes_bp.cadastro_fornecedores'))

        # Inserir fornecedor no banco
        cursor.execute('''
            INSERT INTO fornecedores (nome_fantasia, cnpj, telefone, email, endereco, bairro, cidade, estado, contato, materiais)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (nome_fantasia, cnpj, telefone, email, endereco, bairro, cidade, estado, contato, materiais))
        conn.commit()
        conn.close()
        flash('Fornecedor cadastrado com sucesso.', 'success')
        return redirect(url_for('routes_bp.lista_fornecedores'))

    except sqlite3.Error as e:
        if 'conn' in locals() and conn:
            conn.close()
        flash(f'Erro ao cadastrar fornecedor: {str(e)}', 'error')
        return redirect(url_for('routes_bp.cadastro_fornecedores'))
    except Exception as e:
        if 'conn' in locals() and conn:
            conn.close()
        flash(f'Erro ao cadastrar fornecedor: {str(e)}', 'error')
        return redirect(url_for('routes_bp.cadastro_fornecedores'))

# Definição da rota /search_fornecedores
@routes_bp.route('/search_fornecedores', methods=['GET'])
def search_fornecedores():
    try:
        query = request.args.get('q', '').strip()
        if not query:
            return jsonify([]), 400
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if not conn:
            app.logger.error('Falha ao conectar ao banco de fornecedores')
            return jsonify([]), 500
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, nome_fantasia, cnpj
            FROM fornecedores
            WHERE nome_fantasia LIKE ? OR cnpj LIKE ?
            ORDER BY nome_fantasia
        ''', (f'%{query}%', f'%{query}%'))
        fornecedores = [
            {
                'id': row['id'],
                'nome_fantasia': row['nome_fantasia'],
                'cnpj': format_cnpj(row['cnpj'])
            } for row in cursor.fetchall()
        ]
        conn.close()
        return jsonify(fornecedores)
    except sqlite3.Error as e:
        app.logger.error(f'Erro ao buscar fornecedores: {str(e)}')
        return jsonify([]), 500
    except Exception as e:
        app.logger.error(f'Erro inesperado em search_fornecedores: {str(e)}')
        return jsonify([]), 500
    
    
def verificar_coluna_observacoes():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(PedidosCompra)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'observacoes' not in columns:
            cursor.execute("ALTER TABLE PedidosCompra ADD COLUMN observacoes TEXT")
            conn.commit()
            logging.info("Coluna observacoes adicionada à tabela PedidosCompra")
            print("✓ Coluna observacoes adicionada com sucesso!")
        
        conn.close()
    except sqlite3.Error as e:
        logging.error(f"Erro ao verificar coluna observacoes: {str(e)}")

# Rota temporária para diagnóstico - pode remover depois
@routes_bp.route('/verificar_campo_ativo')
def verificar_campo_ativo():
    try:
        # Verifica se a coluna Ativo existe na tabela Materiais
        result = db.engine.execute("PRAGMA table_info(Materiais)").fetchall()
        colunas = [col[1] for col in result]
        existe = 'Ativo' in colunas
        
        return jsonify({
            'campo_ativo_existe': existe,
            'colunas': colunas,
            'materiais_ativos': Materiais.query.filter_by(Ativo=True).count()
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

#------------------------------------------Ativos 
# Adicione no início do arquivo, com as outras constantes
REGISTRO_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'registro.txt')

# Função para ler registros
def ler_registros():
    registros = []
    try:
        if not os.path.exists(REGISTRO_FILE):
            return registros
            
        with open(REGISTRO_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:  # Ignora linhas vazias
                    partes = line.split('|')
                    if len(partes) >= 3:
                        registros.append({
                            'nome': partes[0],
                            'descricao': partes[1],
                            'empresa': partes[2],
                            'data': partes[3] if len(partes) > 3 else ''
                        })
    except Exception as e:
        logging.error(f"Erro ao ler registro.txt: {str(e)}")
    return registros

# Função para adicionar registro
def adicionar_registro(nome, descricao, empresa):
    try:
        data_atual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(REGISTRO_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{nome}|{descricao}|{empresa}|{data_atual}\n")
        return True
    except Exception as e:
        logging.error(f"Erro ao adicionar registro: {str(e)}")
        return False

# Adicione ao jinja_env.globals
app.jinja_env.globals.update(
    ler_registros=ler_registros,
    adicionar_registro=adicionar_registro
)

@routes_bp.route('/registros', methods=['GET'], endpoint='registros')
def listar_registros():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    registros = ler_registros()
    return render_template('registros.html', registros=registros)


@routes_bp.route('/adicionar_registro', methods=['POST'])
def adicionar_registro_route():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    nome = request.form.get('nome', '').strip()
    descricao = request.form.get('descricao', '').strip()
    empresa = request.form.get('empresa', '').strip()
    
    # Validações básicas
    if not all([nome, descricao, empresa]):
        flash('Todos os campos são obrigatórios.', 'error')
        return redirect(url_for('routes_bp.registros'))
    
    # Verifica duplicados
    registros = ler_registros()
    for registro in registros:
        if (registro['nome'].lower() == nome.lower() and 
            registro['empresa'].lower() == empresa.lower()):
            flash('Já existe um registro com este nome e empresa!', 'error')
            return redirect(url_for('routes_bp.registros'))
    
    # Adiciona se não for duplicado
    if adicionar_registro(nome, descricao, empresa):
        flash('Registro adicionado com sucesso!', 'success')
    else:
        flash('Erro ao adicionar registro.', 'error')
    
    return redirect(url_for('routes_bp.registros'))


# Função para remover um registro específico
def remover_registro(linha_para_remover):
    try:
        # Lê todos os registros
        with open(REGISTRO_FILE, 'r', encoding='utf-8') as f:
            linhas = f.readlines()
        
        # Remove a linha específica
        linhas = [linha for linha in linhas if linha.strip() != linha_para_remover.strip()]
        
        # Reescreve o arquivo sem a linha removida
        with open(REGISTRO_FILE, 'w', encoding='utf-8') as f:
            f.writelines(linhas)
        
        return True
    except Exception as e:
        logging.error(f"Erro ao remover registro: {str(e)}")
        return False
    
@routes_bp.route('/excluir_registro', methods=['POST'])
def excluir_registro_route():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    # Obtém os dados do formulário
    nome = request.form.get('nome', '').strip()
    descricao = request.form.get('descricao', '').strip()
    empresa = request.form.get('empresa', '').strip()
    data = request.form.get('data', '').strip()
    
    # Reconstroi a linha no formato do arquivo
    linha_para_remover = f"{nome}|{descricao}|{empresa}|{data}\n"
    
    if remover_registro(linha_para_remover):
        flash('Registro excluído com sucesso!', 'success')
    else:
        flash('Erro ao excluir registro.', 'error')
    
    return redirect(url_for('routes_bp.registros'))

@routes_bp.route('/registros')
def get_registros():
    try:
        registros = ler_registros()
        
        # Verifica se há registros válidos
        if not registros:
            return jsonify({"error": "Nenhum registro válido encontrado"}), 404
            
        return jsonify({
            "success": True,
            "data": registros
        })
        
    except Exception as e:
        logging.error(f"Erro na rota /registros: {str(e)}")
        return jsonify({
            "error": "Erro interno ao processar registros",
            "details": str(e)
        }), 500
    
# Adicione esta função (se já não existir)
def ler_nomes_registros():
    nomes = set()
    try:
        if not os.path.exists(REGISTRO_FILE):
            return sorted(nomes)
            
        with open(REGISTRO_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line:  # Ignora linhas vazias
                    partes = line.split('|')
                    if len(partes) >= 1:
                        nomes.add(partes[0])  # Adiciona o primeiro campo (nome)
    except Exception as e:
        logging.error(f"Erro ao ler nomes do registro.txt: {str(e)}")
    return sorted(nomes)

# Adicione ao jinja_env.globals (procure por app.jinja_env.globals.update e adicione)
app.jinja_env.globals.update(
    ler_nomes_registros=ler_nomes_registros,
    # ... outros globais que já existam
)

@routes_bp.route('/buscar_ativos', methods=['GET'])
def buscar_ativos():
    termo = request.args.get('q', '').lower()
    registros = ler_registros()
    resultados = [{'nome': r['nome']} for r in registros if termo in r['nome'].lower()]
    return jsonify(resultados)

#Tela auditoria -----------------------------------------------------------
@routes_bp.route('/auditoria_solicitacoes', methods=['GET'])
def auditoria_solicitacoes():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Obter parâmetros de filtro
        empresa = request.args.get('empresa')
        usuario = request.args.get('usuario')
        ativo = request.args.get('ativo')
        nome_ativo = request.args.get('nome_ativo')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')
        status = request.args.get('status')

        # Query base
        query = db.session.query(SolicitacoesCompra)
        
        # Aplicar filtros
        if empresa:
            query = query.filter(SolicitacoesCompra.empresa == empresa)
        if usuario:
            query = query.filter(SolicitacoesCompra.usuario == usuario)
        if ativo:
            query = query.filter(SolicitacoesCompra.ativo == ativo)
        if nome_ativo and ativo == 'Sim':
            query = query.filter(SolicitacoesCompra.nome_ativo.ilike(f'%{nome_ativo}%'))
        if data_inicio:
            query = query.filter(SolicitacoesCompra.data_solicitacao >= data_inicio)
        if data_fim:
            data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
            query = query.filter(SolicitacoesCompra.data_solicitacao <= data_fim_ajustada)

        # Executar query
        solicitacoes = query.order_by(SolicitacoesCompra.data_solicitacao.desc()).all()

        # Obter fornecedores
        fornecedor_ids = set()
        for solicitacao in solicitacoes:
            preenchimento = db.session.query(SolicitacoesPreenchidas).filter_by(
                solicitacao_id=solicitacao.id
            ).first()
            if preenchimento and preenchimento.fornecedor_id:
                fornecedor_ids.add(preenchimento.fornecedor_id)

        fornecedores = {}
        if fornecedor_ids:
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute(
                        f'SELECT id, nome_fantasia, cnpj FROM fornecedores WHERE id IN ({",".join("?"*len(fornecedor_ids))})',
                        list(fornecedor_ids)
                    )
                    for row in cursor.fetchall():
                        fornecedores[row[0]] = {
                            'nome_fantasia': row[1],
                            'cnpj': format_cnpj(row[2]) if row[2] else 'N/A'
                        }
                finally:
                    conn.close()

        # Preparar dados para o template
        auditoria = []
        for solicitacao in solicitacoes:
            material = db.session.get(Materiais, solicitacao.cod_material)
            preenchimento = db.session.query(SolicitacoesPreenchidas).filter_by(
                solicitacao_id=solicitacao.id
            ).first()
            pedido = None
            if preenchimento:
                pedido = db.session.query(PedidosCompra).join(
                    pedido_preenchimento_associacao,
                    PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
                ).filter(
                    pedido_preenchimento_associacao.c.preenchimento_id == preenchimento.id
                ).first()
            estoque = None
            if preenchimento:
                estoque = db.session.query(Estoque).filter_by(
                    preenchimento_id=preenchimento.id
                ).first()
            requisicoes = []
            if preenchimento:
                requisicoes = db.session.query(Requisicoes).filter_by(
                    preenchimento_id=preenchimento.id
                ).all()
            
            # Aplicar filtro de status
            if status:
                if status == 'Aberta' and preenchimento:
                    continue
                elif status != 'Aberta' and (not preenchimento or preenchimento.status != status):
                    continue
            
            # Adicionar informações do fornecedor
            fornecedor_info = {}
            if preenchimento and preenchimento.fornecedor_id:
                fornecedor_info = fornecedores.get(preenchimento.fornecedor_id, {
                    'nome_fantasia': 'Fornecedor não encontrado',
                    'cnpj': 'N/A'
                })

            auditoria.append({
                'solicitacao': solicitacao,
                'material': material,
                'preenchimento': preenchimento,
                'pedido': pedido,
                'estoque': estoque,
                'requisicoes': requisicoes,
                'fornecedor': fornecedor_info
            })

        # Obter valores para filtros
        empresas = db.session.query(
            SolicitacoesCompra.empresa
        ).distinct().order_by(
            SolicitacoesCompra.empresa
        ).all()

        usuarios = db.session.query(
            SolicitacoesCompra.usuario
        ).distinct().order_by(
            SolicitacoesCompra.usuario
        ).all()

        nomes_ativos = db.session.query(
            SolicitacoesCompra.nome_ativo
        ).filter(
            SolicitacoesCompra.ativo == 'Sim',
            SolicitacoesCompra.nome_ativo.isnot(None)
        ).distinct().order_by(
            SolicitacoesCompra.nome_ativo
        ).all()

        return render_template(
            'auditoria_solicitacoes.html',
            auditoria=auditoria,
            empresas=[e[0] for e in empresas if e[0]],
            usuarios=[u[0] for u in usuarios if u[0]],
            nomes_ativos=[n[0] for n in nomes_ativos if n[0]],
            filtros={
                'empresa': empresa,
                'usuario': usuario,
                'ativo': ativo,
                'nome_ativo': nome_ativo,
                'data_inicio': data_inicio,
                'data_fim': data_fim,
                'status': status
            }
        )

    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao carregar auditoria: {str(e)}', 'error')
        app.logger.error(f'Erro em auditoria_solicitacoes: {str(e)}', exc_info=True)
        return render_template(
            'auditoria_solicitacoes.html',
            auditoria=[],
            empresas=[],
            usuarios=[],
            nomes_ativos=[],
            filtros={}
        )
    
# 2° Tela Auditoria
@routes_bp.route('/tela_auditoria', methods=['GET', 'POST'])
def tela_auditoria():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Processar formulário de auditoria se for POST
        if request.method == 'POST':
            solicitacao_id = request.form.get('solicitacao_id')
            colaborador_1 = request.form.get('colaborador_1', '').strip()
            colaborador_2 = request.form.get('colaborador_2', '').strip()
            status = request.form.get('status', '').strip()
            observacao = request.form.get('observacao', '').strip()
            
            # Validações
            if not all([solicitacao_id, colaborador_1, status]):
                flash('Campos obrigatórios não preenchidos', 'error')
                return redirect(url_for('routes_bp.tela_auditoria'))
            
            if status not in ['Conforme', 'Não Conforme']:
                flash('Status inválido', 'error')
                return redirect(url_for('routes_bp.tela_auditoria'))
            
            # Criar registro de auditoria
            auditoria = Auditoria(
                solicitacao_id=solicitacao_id,
                colaborador_1=colaborador_1,
                colaborador_2=colaborador_2 if colaborador_2 else None,
                status=status,
                observacao=observacao if observacao else None
            )
            db.session.add(auditoria)
            db.session.commit()
            flash('Auditoria registrada com sucesso!', 'success')
            return redirect(url_for('routes_bp.tela_auditoria'))

        # Obter parâmetros de filtro (código existente)
        empresa = request.args.get('empresa')
        usuario = request.args.get('usuario')
        ativo = request.args.get('ativo')
        nome_ativo = request.args.get('nome_ativo')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')
        status = request.args.get('status')

        # Query base (código existente)
        query = db.session.query(SolicitacoesCompra)
        
        # Aplicar filtros (código existente)
        if empresa:
            query = query.filter(SolicitacoesCompra.empresa == empresa)
        if usuario:
            query = query.filter(SolicitacoesCompra.usuario == usuario)
        if ativo:
            query = query.filter(SolicitacoesCompra.ativo == ativo)
        if nome_ativo and ativo == 'Sim':
            query = query.filter(SolicitacoesCompra.nome_ativo.ilike(f'%{nome_ativo}%'))
        if data_inicio:
            query = query.filter(SolicitacoesCompra.data_solicitacao >= data_inicio)
        if data_fim:
            data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
            query = query.filter(SolicitacoesCompra.data_solicitacao <= data_fim_ajustada)

        # Executar query (código existente)
        solicitacoes = query.order_by(SolicitacoesCompra.data_solicitacao.desc()).all()

        # Obter fornecedores (código existente)
        fornecedor_ids = set()
        for solicitacao in solicitacoes:
            preenchimento = db.session.query(SolicitacoesPreenchidas).filter_by(
                solicitacao_id=solicitacao.id
            ).first()
            if preenchimento and preenchimento.fornecedor_id:
                fornecedor_ids.add(preenchimento.fornecedor_id)

        fornecedores = {}
        if fornecedor_ids:
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute(
                        f'SELECT id, nome_fantasia, cnpj FROM fornecedores WHERE id IN ({",".join("?"*len(fornecedor_ids))})',
                        list(fornecedor_ids)
                    )
                    for row in cursor.fetchall():
                        fornecedores[row[0]] = {
                            'nome_fantasia': row[1],
                            'cnpj': format_cnpj(row[2]) if row[2] else 'N/A'
                        }
                finally:
                    conn.close()

        # Preparar dados para o template (atualizado com auditoria)
        auditoria_data = []
        for solicitacao in solicitacoes:
            material = db.session.get(Materiais, solicitacao.cod_material)
            preenchimento = db.session.query(SolicitacoesPreenchidas).filter_by(
                solicitacao_id=solicitacao.id
            ).first()
            pedido = None
            if preenchimento:
                pedido = db.session.query(PedidosCompra).join(
                    pedido_preenchimento_associacao,
                    PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
                ).filter(
                    pedido_preenchimento_associacao.c.preenchimento_id == preenchimento.id
                ).first()
            estoque = None
            if preenchimento:
                estoque = db.session.query(Estoque).filter_by(
                    preenchimento_id=preenchimento.id
                ).first()
            requisicoes = []
            if preenchimento:
                requisicoes = db.session.query(Requisicoes).filter_by(
                    preenchimento_id=preenchimento.id
                ).all()
            
            # Aplicar filtro de status
            if status:
                if status == 'Aberta' and preenchimento:
                    continue
                elif status != 'Aberta' and (not preenchimento or preenchimento.status != status):
                    continue
            
            # Adicionar informações do fornecedor
            fornecedor_info = {}
            if preenchimento and preenchimento.fornecedor_id:
                fornecedor_info = fornecedores.get(preenchimento.fornecedor_id, {
                    'nome_fantasia': 'Fornecedor não encontrado',
                    'cnpj': 'N/A'
                })

            # Obter dados de auditoria para esta solicitação
            auditorias = db.session.query(Auditoria).filter_by(
                solicitacao_id=solicitacao.id
            ).order_by(Auditoria.data_validacao.desc()).all()

            auditoria_data.append({
                'solicitacao': solicitacao,
                'material': material,
                'preenchimento': preenchimento,
                'pedido': pedido,
                'estoque': estoque,
                'requisicoes': requisicoes,
                'fornecedor': fornecedor_info,
                'auditorias': auditorias  # Adiciona os dados de auditoria
            })

        # Obter valores para filtros (código existente)
        empresas = db.session.query(
            SolicitacoesCompra.empresa
        ).distinct().order_by(
            SolicitacoesCompra.empresa
        ).all()

        usuarios = db.session.query(
            SolicitacoesCompra.usuario
        ).distinct().order_by(
            SolicitacoesCompra.usuario
        ).all()

        nomes_ativos = db.session.query(
            SolicitacoesCompra.nome_ativo
        ).filter(
            SolicitacoesCompra.ativo == 'Sim',
            SolicitacoesCompra.nome_ativo.isnot(None)
        ).distinct().order_by(
            SolicitacoesCompra.nome_ativo
        ).all()

        # Obter lista de usuários para seleção de auditores
        senhas = ler_senhas()
        usuarios_disponiveis = sorted(senhas.keys())

        return render_template(
            'tela_auditoria.html',
            auditoria=auditoria_data,
            empresas=[e[0] for e in empresas if e[0]],
            usuarios=[u[0] for u in usuarios if u[0]],
            nomes_ativos=[n[0] for n in nomes_ativos if n[0]],
            usuarios_disponiveis=usuarios_disponiveis,
            filtros={
                'empresa': empresa,
                'usuario': usuario,
                'ativo': ativo,
                'nome_ativo': nome_ativo,
                'data_inicio': data_inicio,
                'data_fim': data_fim,
                'status': status
            }
        )

    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao carregar auditoria: {str(e)}', 'error')
        app.logger.error(f'Erro em tela_auditoria: {str(e)}', exc_info=True)
        return render_template(
            'tela_auditoria.html',
            auditoria=[],
            empresas=[],
            usuarios=[],
            nomes_ativos=[],
            usuarios_disponiveis=[],
            filtros={}
        )
    
def create_auditoria_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS Auditoria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            solicitacao_id INTEGER NOT NULL,
            data_validacao DATETIME NOT NULL,
            colaborador_1 TEXT NOT NULL,
            colaborador_2 TEXT,
            status TEXT NOT NULL,
            observacao TEXT,
            FOREIGN KEY (solicitacao_id) REFERENCES SolicitacoesCompra(id)
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'Auditoria' criada com sucesso.")
        cursor.close()
        conn.close()
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela Auditoria: {str(e)}")
        return False
    return True

@routes_bp.route('/excluir_auditoria/<int:id>', methods=['POST'])
def excluir_auditoria(id):
    if 'usuario' not in session:
        flash('Você precisa estar logado para realizar esta ação.', 'error')
        return redirect(url_for('routes_bp.login'))
    
    try:
        auditoria = Auditoria.query.get_or_404(id)
        solicitacao_id = auditoria.solicitacao_id
        
        db.session.delete(auditoria)
        db.session.commit()
        
        flash('Registro de auditoria excluído com sucesso.', 'success')
        return redirect(url_for('routes_bp.tela_auditoria'))
    
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao excluir auditoria: {str(e)}', 'error')
        return redirect(url_for('routes_bp.tela_auditoria'))


#Dasboard ----------------------------------------------
@routes_bp.route('/dashboard', methods=['GET'])
def dashboard():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Dados básicos
        total_materiais = Materiais.query.count()
        
        # Materiais com estoque crítico (menos de 5 dias de consumo)
        materiais_criticos = Materiais.query.filter(
            Materiais.FatorConsumo > 0,
            (Materiais.QuantidadeEstoque / Materiais.FatorConsumo) < 5
        ).all()
        
        # Total em estoque
        total_estoque = db.session.query(
            db.func.sum(Materiais.QuantidadeEstoque)
        ).scalar() or 0
        
        # Requisições
        requisicoes_abertas = Requisicoes.query.count()
        
        # Últimas requisições com mais informações
        ultimas_requisicoes = db.session.query(
            Requisicoes,
            Materiais,
            SolicitacoesPreenchidas
        ).join(
            Materiais,
            Requisicoes.cod_material == Materiais.CodMaterial
        ).join(
            SolicitacoesPreenchidas,
            Requisicoes.preenchimento_id == SolicitacoesPreenchidas.id
        ).order_by(
            Requisicoes.data_requisicao.desc()
        ).limit(5).all()
        
        # Top materiais em estoque
        materiais_estoque = Materiais.query.order_by(
            Materiais.QuantidadeEstoque.desc()
        ).limit(5).all()
        
        # Status das requisições
        requisicoes_concluidas = Requisicoes.query.join(
            SolicitacoesPreenchidas,
            Requisicoes.preenchimento_id == SolicitacoesPreenchidas.id
        ).filter(
            SolicitacoesPreenchidas.status == 'Entregue'
        ).count()
        
        requisicoes_pendentes = requisicoes_abertas - requisicoes_concluidas
        
        # Preparar dados para os gráficos
        estoque_grafico = {
            'labels': [m.DescricaoMaterial for m in materiais_estoque],
            'data': [m.QuantidadeEstoque for m in materiais_estoque]
        }
        
        requisicoes_grafico = {
            'data': [requisicoes_concluidas, requisicoes_pendentes]
        }
        
        # Consumo mensal (dados fictícios para exemplo)
        consumo_mensal = {
            'labels': ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
            'data': [120, 190, 170, 210, 200, 230, 250, 240, 260, 280, 300, 320]
        }
        
        return render_template(
            'dashboard.html',
            total_materiais=total_materiais,
            estoque_critico=len(materiais_criticos),
            materiais_criticos=materiais_criticos,
            requisicoes_abertas=requisicoes_abertas,
            total_estoque=total_estoque,
            ultimas_requisicoes=ultimas_requisicoes,
            estoque_grafico=estoque_grafico,
            requisicoes_grafico=requisicoes_grafico,
            consumo_mensal=consumo_mensal
        )
        
    except Exception as e:
        flash(f'Erro ao carregar dashboard: {str(e)}', 'error')
        app.logger.error(f'Erro no dashboard: {str(e)}')
        return render_template(
            'dashboard.html',
            total_materiais=0,
            estoque_critico=0,
            materiais_criticos=[],
            requisicoes_abertas=0,
            total_estoque=0,
            ultimas_requisicoes=[],
            estoque_grafico={'labels': [], 'data': []},
            requisicoes_grafico={'data': [0, 0]},
            consumo_mensal={'labels': [], 'data': []}
        )
    
@routes_bp.route('/dashboard/data', methods=['GET'])
def dashboard_data():
    if 'usuario' not in session:
        return jsonify({'error': 'Não autenticado'}), 401
    
    try:
        # Dados resumidos
        total_materiais = Materiais.query.count()
        estoque_critico = Materiais.query.filter(
            Materiais.FatorConsumo > 0,
            (Materiais.QuantidadeEstoque / Materiais.FatorConsumo) < 5
        ).count()
        total_estoque = db.session.query(
            db.func.sum(Materiais.QuantidadeEstoque)
        ).scalar() or 0
        requisicoes_abertas = Requisicoes.query.count()
        
        # Status das solicitações
        status_solicitacoes = db.session.query(
            SolicitacoesPreenchidas.status,
            db.func.count(SolicitacoesPreenchidas.id)
        ).group_by(SolicitacoesPreenchidas.status).all()
        
        # Últimas solicitações
        ultimas_solicitacoes = SolicitacoesCompra.query.order_by(
            SolicitacoesCompra.data_solicitacao.desc()
        ).limit(5).all()
        
        # Materiais com maior estoque (top 5)
        materiais_estoque = Materiais.query.order_by(
            Materiais.QuantidadeEstoque.desc()
        ).limit(5).all()
        
        # Formatando os dados para o dashboard
        dados = {
            'totais': {
                'materiais': total_materiais,
                'estoque_critico': estoque_critico,
                'total_estoque': total_estoque,
                'requisicoes_abertas': requisicoes_abertas
            },
            'status_solicitacoes': [
                {'status': s[0], 'quantidade': s[1]} for s in status_solicitacoes
            ],
            'ultimas_solicitacoes': [
                {
                    'id': s.id,
                    'material': s.material.DescricaoMaterial if s.material else 'N/A',
                    'quantidade': s.quantidade,
                    'data': s.data_solicitacao.strftime('%d/%m/%Y')
                } for s in ultimas_solicitacoes
            ],
            'top_materiais': [
                {
                    'nome': m.DescricaoMaterial,
                    'quantidade': m.QuantidadeEstoque,
                    'dias_estoque': m.QuantidadeEstoque / m.FatorConsumo if m.FatorConsumo > 0 else 0
                } for m in materiais_estoque
            ]
        }
        
        return jsonify(dados)
    
    except Exception as e:
        app.logger.error(f'Erro no dashboard: {str(e)}')
        return jsonify({'error': str(e)}), 500


#-------------------------------------------------
# Updated financeiro route
@routes_bp.route('/financeiro', methods=['GET'])
def financeiro():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        empresa = request.args.get('empresa')
        usuario = request.args.get('usuario')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')

        query = db.session.query(PedidosCompra).filter(
            PedidosCompra.status == 'AgPagamento'
        ).join(
            pedido_preenchimento_associacao,
            PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
        ).join(
            SolicitacoesPreenchidas,
            pedido_preenchimento_associacao.c.preenchimento_id == SolicitacoesPreenchidas.id
        ).join(
            SolicitacoesCompra,
            SolicitacoesPreenchidas.solicitacao_id == SolicitacoesCompra.id
        )

        if empresa:
            query = query.filter(SolicitacoesCompra.empresa == empresa)
        if usuario:
            query = query.filter(PedidosCompra.usuario == usuario)
        if data_inicio:
            query = query.filter(PedidosCompra.data_criacao >= data_inicio)
        if data_fim:
            data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
            query = query.filter(PedidosCompra.data_criacao <= data_fim_ajustada)

        pedidos = query.order_by(PedidosCompra.data_criacao.desc()).distinct().all()

        fornecedor_ids = set()
        for pedido in pedidos:
            for preenchimento in pedido.preenchimentos:
                fornecedor_ids.add(preenchimento.fornecedor_id)
        
        fornecedores = {}
        if fornecedor_ids:
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if conn:
                try:
                    cursor = conn.cursor()
                    cursor.execute(
                        f'SELECT id, nome_fantasia, cnpj FROM fornecedores WHERE id IN ({",".join("?"*len(fornecedor_ids))})',
                        list(fornecedor_ids)
                    )
                    for row in cursor.fetchall():
                        fornecedores[row[0]] = {
                            'nome_fantasia': row[1],
                            'cnpj': format_cnpj(row[2]) if row[2] else 'N/A'
                        }
                finally:
                    conn.close()

        pedidos_completos = []
        for pedido in pedidos:
            preenchimentos_info = []
            for preenchimento in pedido.preenchimentos:
                fornecedor_info = fornecedores.get(preenchimento.fornecedor_id, {
                    'nome_fantasia': 'Fornecedor não encontrado',
                    'cnpj': 'N/A'
                })
                preenchimentos_info.append({
                    'id': preenchimento.id,
                    'marca': preenchimento.marca or 'Não informado',
                    'fornecedor_nome': fornecedor_info['nome_fantasia'],
                    'fornecedor_cnpj': fornecedor_info['cnpj'],
                    'material': preenchimento.solicitacao.material.DescricaoMaterial if preenchimento.solicitacao.material else 'N/A',
                    'empresa': preenchimento.solicitacao.empresa,
                    'valor_total': float(preenchimento.valor_total) if preenchimento.valor_total else 0.0
                })
            pedidos_completos.append({
                'pedido': pedido,
                'preenchimentos': preenchimentos_info
            })

        empresas = db.session.query(SolicitacoesCompra.empresa).distinct().order_by(SolicitacoesCompra.empresa).all()
        usuarios = db.session.query(PedidosCompra.usuario).distinct().order_by(PedidosCompra.usuario).all()

        return render_template(
            'financeiro.html',
            pedidos=pedidos_completos,
            empresas=[e[0] for e in empresas if e[0]],
            usuarios=[u[0] for u in usuarios if u[0]],
            filtros={
                'empresa': empresa,
                'usuario': usuario,
                'data_inicio': data_inicio,
                'data_fim': data_fim
            }
        )
    except Exception as e:
        flash(f'Erro ao carregar página financeiro: {str(e)}', 'error')
        return render_template(
            'financeiro.html',
            pedidos=[],
            empresas=[],
            usuarios=[],
            filtros={}
        )

# Updated confirmar_pagamento route with explicit file size validation
@routes_bp.route('/api/pedidos/<int:pedido_id>/confirmar_pagamento', methods=['POST'])
def confirmar_pagamento(pedido_id):
    if 'usuario' not in session:
        return jsonify({'success': False, 'error': 'Usuário não autenticado'}), 401

    pedido = PedidosCompra.query.get_or_404(pedido_id)
    if pedido.status != 'AgPagamento':
        return jsonify({'success': False, 'error': 'Este pedido não está aguardando pagamento'}), 400

    if 'comprovante' not in request.files:
        return jsonify({'success': False, 'error': 'Nenhum arquivo enviado'}), 400

    comprovante = request.files['comprovante']
    if comprovante.filename == '':
        return jsonify({'success': False, 'error': 'Nenhum arquivo selecionado'}), 400

    if comprovante and allowed_file(comprovante.filename):
        try:
            filename = secure_filename(comprovante.filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            unique_filename = f"{timestamp}_{filename}"
            comprovante_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            
            # Save file and check size
            comprovante.save(comprovante_path)
            if os.path.getsize(comprovante_path) > 5 * 1024 * 1024:
                os.remove(comprovante_path)
                return jsonify({'success': False, 'error': 'O arquivo PDF deve ter no máximo 5MB'}), 400

            pedido.comprovante_pagamento = comprovante_path
            pedido.status = 'AgEntrega'
            db.session.commit()
            return jsonify({'success': True, 'message': 'Pagamento confirmado com sucesso'})
        except Exception as e:
            db.session.rollback()
            logging.error(f"Erro ao confirmar pagamento: {e}")
            return jsonify({'success': False, 'error': f'Erro ao confirmar pagamento: {str(e)}'}), 500
    else:
        return jsonify({'success': False, 'error': 'Arquivo inválido. Apenas PDFs são permitidos'}), 400
    
#Excluir fonecedor  
@routes_bp.route('/delete_supplier', methods=['POST'])
def delete_supplier():
    if 'usuario' not in session:
        return jsonify({'status': 'error', 'message': 'Usuário não autenticado'}), 401
    
    try:
        data = request.get_json()
        supplier_id = data.get('id')
        
        if not supplier_id:
            return jsonify({'status': 'error', 'message': 'ID do fornecedor não fornecido'}), 400
        
        # Verificar se há solicitações associadas a este fornecedor no banco principal
        try:
            count = db.session.query(SolicitacoesPreenchidas).filter_by(fornecedor_id=supplier_id).count()
            
            if count > 0:
                return jsonify({
                    'status': 'error', 
                    'message': 'Não é possível excluir: fornecedor está associado a solicitações'
                }), 400
            
            # Se não houver associações, excluir o fornecedor do banco de fornecedores
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if not conn:
                return jsonify({'status': 'error', 'message': 'Erro ao conectar ao banco de dados'}), 500
                
            try:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM fornecedores WHERE id = ?', (supplier_id,))
                conn.commit()
                return jsonify({'status': 'success', 'message': 'Fornecedor excluído com sucesso'})
            except sqlite3.Error as e:
                conn.rollback()
                return jsonify({'status': 'error', 'message': f'Erro no banco de dados: {str(e)}'}), 500
            finally:
                if conn:
                    conn.close()
                    
        except SQLAlchemyError as e:
            return jsonify({'status': 'error', 'message': f'Erro ao verificar solicitações: {str(e)}'}), 500
                
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Erro inesperado: {str(e)}'}), 500
    
#Inserir estoque manual caso necessario   
@routes_bp.route('/definir-quantidade/<int:cod>', methods=['POST'])
def definir_quantidade(cod):
    if 'usuario' not in session:
        flash('Acesso não autorizado', 'error')
        return redirect(url_for('routes_bp.login'))
    
    try:
        quantidade = request.form.get('quantidade', 0)
        motivo = request.form.get('motivo', 'Ajuste manual')  # Novo campo
        
        # Validações
        try:
            quantidade = int(quantidade)
        except ValueError:
            flash('Valor inválido para quantidade. Use apenas números inteiros.', 'error')
            return redirect(url_for('routes_bp.listar_materiais'))
        
        if quantidade < 0:
            flash('Quantidade não pode ser negativa', 'error')
            return redirect(url_for('routes_bp.listar_materiais'))
        
        material = Materiais.query.get_or_404(cod)
        
        # Registrar no histórico antes de alterar
        historico = HistoricoEstoque(
            cod_material=material.CodMaterial,
            usuario=session['usuario'],
            quantidade_anterior=material.QuantidadeEstoque,
            quantidade_nova=quantidade,
            motivo=motivo
        )
        db.session.add(historico)
        
        # Atualizar estoque
        material.QuantidadeEstoque = quantidade
        db.session.commit()
        
        flash(f'Estoque de {material.DescricaoMaterial} atualizado de {historico.quantidade_anterior} para {quantidade} unidades', 'success')
    
    except Exception as e:
        db.session.rollback()
        flash(f'Erro ao atualizar quantidade: {str(e)}', 'error')
    
    return redirect(url_for('routes_bp.listar_materiais'))

@routes_bp.route('/relatorio-estoque')
def relatorio_estoque():
    if 'usuario' not in session:
        flash('Acesso não autorizado', 'error')
        return redirect(url_for('routes_bp.login'))

    # Obter parâmetros de filtro
    page = request.args.get('page', 1, type=int)
    per_page = 20  # Itens por página
    data_inicio = request.args.get('data_inicio')
    data_fim = request.args.get('data_fim')
    usuario = request.args.get('usuario')
    material = request.args.get('material')

    # Construir query base
    query = HistoricoEstoque.query.join(Materiais)

    # Aplicar filtros
    if data_inicio:
        query = query.filter(HistoricoEstoque.data_alteracao >= data_inicio)
    if data_fim:
        data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
        query = query.filter(HistoricoEstoque.data_alteracao <= data_fim_ajustada)
    if usuario:
        query = query.filter(HistoricoEstoque.usuario.ilike(f'%{usuario}%'))
    if material:
        query = query.filter(or_(
            Materiais.DescricaoMaterial.ilike(f'%{material}%'),
            Materiais.CodMaterial.ilike(f'%{material}%')
        ))

    # Paginação
    historico_paginado = query.order_by(HistoricoEstoque.data_alteracao.desc()).paginate(
        page=page, 
        per_page=per_page, 
        error_out=False
    )

    return render_template('relatorio_estoque.html', historico=historico_paginado)

@routes_bp.route('/api/solicitacoes/<int:solicitacao_id>/status', methods=['POST'])
def atualizar_status_solicitacao(solicitacao_id):
    if 'usuario' not in session:
        app.logger.error("Unauthorized access attempt to update status")
        return jsonify({'success': False, 'error': 'Usuário não autenticado'}), 401
    
    try:
        data = request.get_json()
        app.logger.debug(f"Received data: {data}")
        status = data.get('status')
        
        if not status:
            app.logger.error("No status provided in request")
            return jsonify({'success': False, 'error': 'Nenhum status fornecido'}), 400
        
        if status not in ['Aprovado', 'Reprovado']:
            app.logger.error(f"Invalid status received: {status}")
            return jsonify({'success': False, 'error': f'Status inválido: {status}'}), 400
        
        solicitacao = SolicitacoesCompra.query.get_or_404(solicitacao_id)
        app.logger.debug(f"Updating status for solicitacao {solicitacao_id} to {status}")
        solicitacao.status_aprovacao = status
        db.session.commit()
        app.logger.info(f"Status updated successfully for solicitacao {solicitacao_id}: {status}")
        
        return jsonify({
            'success': True, 
            'message': f'Solicitação {status.lower()} com sucesso!'
        })
        
    except Exception as e:
        db.session.rollback()
        app.logger.error(f"Error updating status for solicitacao {solicitacao_id}: {str(e)}", exc_info=True)
        return jsonify({'success': False, 'error': f'Erro ao atualizar status: {str(e)}'}), 500
    
def migrate_solicitacoes_compra_status():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Verificar se a coluna status_aprovacao já existe
        cursor.execute("PRAGMA table_info(SolicitacoesCompra)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'status_aprovacao' not in columns:
            cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN status_aprovacao TEXT")
            conn.commit()
            logging.info("Coluna status_aprovacao adicionada à tabela SolicitacoesCompra")
        
        conn.close()
    except sqlite3.Error as e:
        logging.error(f"Erro na migração: {str(e)}")

@routes_bp.route('/atualizar_valores_cotacao', methods=['POST'])
def atualizar_valores_cotacao():
    if 'usuario' not in session:
        return jsonify({'success': False, 'message': 'Usuário não autenticado'}), 401
    
    try:
        data = request.get_json()
        preenchimento_id = data.get('preenchimento_id')
        valor_unitario = data.get('valor_unitario')
        valor_frete = data.get('valor_frete', 0)
        valor_unitario_original = data.get('valor_unitario_original')
        observacoes = data.get('observacoes')  # Novo: Pegar observações do request
        
        if not all([preenchimento_id, valor_unitario, valor_unitario_original]):
            return jsonify({'success': False, 'message': 'Dados incompletos'}), 400
        
        # Buscar o preenchimento
        preenchimento = SolicitacoesPreenchidas.query.get_or_404(preenchimento_id)
        
        # Salvar valores anteriores para o histórico
        valor_unitario_anterior = preenchimento.valor_unitario
        valor_frete_anterior = preenchimento.valor_frete
        
        # Calcular novo valor total
        quantidade = preenchimento.solicitacao.quantidade
        novo_valor_total = (valor_unitario * quantidade) + valor_frete
        
        # Atualizar os valores
        preenchimento.valor_unitario = valor_unitario
        preenchimento.valor_frete = valor_frete if valor_frete > 0 else None
        preenchimento.valor_total = novo_valor_total
        preenchimento.observacoes = observacoes  # Novo: Atualizar observações
        
        # Registrar no histórico (mantendo o existente, sem mudança aqui)
        historico = HistoricoDescontos(
            preenchimento_id=preenchimento_id,
            valor_unitario_anterior=valor_unitario_anterior,
            valor_unitario_novo=valor_unitario,
            valor_frete_anterior=valor_frete_anterior,
            valor_frete_novo=valor_frete if valor_frete > 0 else None,
            usuario=session['usuario']
        )
        db.session.add(historico)
        
        db.session.commit()
        
        return jsonify({
            'success': True, 
            'message': 'Valores atualizados com sucesso!'
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'message': f'Erro ao atualizar valores: {str(e)}'}), 500
    
def create_historico_descontos_table():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        create_table_query = """
        CREATE TABLE IF NOT EXISTS HistoricoDescontos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            preenchimento_id INTEGER NOT NULL,
            valor_unitario_anterior REAL NOT NULL,
            valor_unitario_novo REAL NOT NULL,
            valor_frete_anterior REAL,
            valor_frete_novo REAL,
            data_alteracao DATETIME NOT NULL,
            usuario TEXT NOT NULL,
            FOREIGN KEY (preenchimento_id) REFERENCES SolicitacoesPreenchidas(id)
        )
        """
        cursor.execute(create_table_query)
        conn.commit()
        print("Tabela 'HistoricoDescontos' criada com sucesso.")
        cursor.close()
        conn.close()
        return True
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela HistoricoDescontos: {str(e)}")
        return False

@routes_bp.route('/historico_descontos', methods=['GET'])
def listar_historico_descontos():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Obter parâmetros de filtro
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')
        usuario = request.args.get('usuario')
        preenchimento_id = request.args.get('preenchimento_id')
        
        # Query base
        query = db.session.query(HistoricoDescontos).join(
            SolicitacoesPreenchidas
        ).join(
            SolicitacoesCompra
        ).join(
            Materiais
        )
        
        # Aplicar filtros
        if data_inicio:
            query = query.filter(HistoricoDescontos.data_alteracao >= data_inicio)
        if data_fim:
            data_fim_ajustada = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
            query = query.filter(HistoricoDescontos.data_alteracao <= data_fim_ajustada)
        if usuario:
            query = query.filter(HistoricoDescontos.usuario.ilike(f'%{usuario}%'))
        if preenchimento_id:
            query = query.filter(HistoricoDescontos.preenchimento_id == preenchimento_id)
        
        # Executar query
        historicos = query.order_by(HistoricoDescontos.data_alteracao.desc()).all()
        
        # Obter lista de usuários para filtro
        usuarios = db.session.query(
            HistoricoDescontos.usuario
        ).distinct().order_by(
            HistoricoDescontos.usuario
        ).all()
        
        return render_template(
            'historico_descontos.html',
            historicos=historicos,
            usuarios=[u[0] for u in usuarios if u[0]],
            filtros={
                'data_inicio': data_inicio,
                'data_fim': data_fim,
                'usuario': usuario,
                'preenchimento_id': preenchimento_id
            }
        )
    
    except Exception as e:
        flash(f'Erro ao carregar histórico de descontos: {str(e)}', 'error')
        app.logger.error(f'Erro em listar_historico_descontos: {str(e)}', exc_info=True)
        return render_template(
            'historico_descontos.html',
            historicos=[],
            usuarios=[],
            filtros={}
        )
def migrate_solicitacoes_compra_status_aprovacao():
    """Adiciona a coluna status_aprovacao na tabela SolicitacoesCompra se não existir"""
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Verificar se a coluna status_aprovacao já existe
        cursor.execute("PRAGMA table_info(SolicitacoesCompra)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'status_aprovacao' not in columns:
            cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN status_aprovacao TEXT")
            conn.commit()
            logging.info("Coluna status_aprovacao adicionada à tabela SolicitacoesCompra")
            print("✓ Coluna status_aprovacao adicionada com sucesso!")
        else:
            print("✓ Coluna status_aprovacao já existe na tabela")
        
        conn.close()
        return True
    except sqlite3.Error as e:
        logging.error(f"Erro na migração status_aprovacao: {str(e)}")
        print(f"✗ Erro na migração: {str(e)}")
        return False

  # Registro do Blueprint (apenas uma vez)
app.register_blueprint(routes_bp)  


# Inicialização do banco de dados e execução do app
if __name__ == '__main__':
    create_database()
    create_auditoria_table()
    create_fornecedores_db()
    create_solicitacoes_preenchidas_table()
    create_historico_descontos_table()
    
    # Adicione esta linha para criar a coluna status_aprovacao
    migrate_solicitacoes_compra_status_aprovacao()
    
    # Nova migração para observacoes
    migrate_observacoes_col()  # ADICIONE ESTA LINHA
    migrate_solicitacoes_compra_status_aprovacao()
    add_observacoes_column_to_solicitacoes_preenchidas()
   
    with app.app_context():
        # Cria todas as tabelas definidas nos modelos
        db.create_all()
        
        # Executa migrações manuais se necessário
        if not atualizar_estrutura_requisicoes():
            print("Falha na migração. Verifique os logs acima.")
            exit(1)
        
        migrate_solicitacoes_compra()
    
    logging.basicConfig(
        filename='app_errors.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        encoding='utf-8'
    )
    app.run(debug=True, host='0.0.0.0', port=80)