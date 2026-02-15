from flask import Flask, Blueprint, session, request, flash, redirect, url_for, render_template, jsonify, send_file, send_from_directory
from flask_sqlalchemy import SQLAlchemy
import os
from dotenv import load_dotenv
import logging
from datetime import datetime, timedelta, timezone
import chardet
import re
import sqlite3
import uuid
from werkzeug.utils import secure_filename
from os.path import basename
import pdfkit
from sqlalchemy.exc import SQLAlchemyError
from collections import Counter
import random
import shutil
from sqlalchemy import or_
from decimal import Decimal, InvalidOperation
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
from flask import make_response

# Carregar variáveis de ambiente
load_dotenv()

# Substitua a função get_local_time() existente (no início do arquivo)
def get_local_time():
    """Retorna o datetime atual no fuso horário de Brasília (UTC-3) SEM timezone"""
    from datetime import datetime, timedelta
    
    # Obter o tempo UTC e converter para Brasília (UTC-3)
    utc_now = datetime.utcnow()
    brasil_time = utc_now - timedelta(hours=3)
    
    return brasil_time

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
    """Verifica se o arquivo tem uma extensão permitida"""
    if not filename:
        logging.warning("Filename is empty or None")
        return False
    
    if '.' not in filename:
        logging.warning(f"Filename has no extension: {filename}")
        return False
    
    try:
        extension = filename.rsplit('.', 1)[1].lower()
        allowed = allowed_extensions or ALLOWED_EXTENSIONS
        
        # Log para debug
        logging.info(f"Verificando arquivo: {filename}, extensão: {extension}, permitido: {extension in allowed}")
        
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
    aplicacao = db.Column(db.Text, nullable=True)  # Aplicação específica do item
    aplicacao_geral = db.Column(db.Text, nullable=True)  # NOVA COLUNA: Aplicação geral do formulário
    empresa = db.Column(db.Text, nullable=False)
    data_solicitacao = db.Column(db.DateTime, nullable=False, default=get_local_time)
    usuario = db.Column(db.Text, nullable=False)
    foto_path = db.Column(db.Text, nullable=True)
    marca = db.Column(db.Text, nullable=True)
    ativo = db.Column(db.String(3), nullable=False)
    nome_ativo = db.Column(db.Text, nullable=True)
    prioridade = db.Column(db.String(20), nullable=False, default='Programado')
    status_aprovacao = db.Column(db.String(20), nullable=True, default=None)
    observacoes_col = db.Column(db.Text, nullable=True)
    material = db.relationship('Materiais', backref='solicitacoes')
    comprador_atribuido = db.Column(db.Text, nullable=True)  # Novo campo

    def to_dict(self):
        return {
            'id': self.id,
            'cod_material': self.cod_material,
            'especificacao': self.especificacao,
            'quantidade': self.quantidade,
            'unidade_medida': self.unidade_medida,
            'aplicacao': self.aplicacao,
            'aplicacao_geral': self.aplicacao_geral,  # Adicione esta linha
            'empresa': self.empresa,
            'data_solicitacao': self.data_solicitacao.isoformat() if self.data_solicitacao else None,
            'usuario': self.usuario,
            'foto_path': self.foto_path,
            'marca': self.marca,
            'ativo': self.ativo,
            'nome_ativo': self.nome_ativo,
            'prioridade': self.prioridade,
            'status_aprovacao': self.status_aprovacao,
            'comprador_atribuido': self.comprador_atribuido
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
    data_preenchimento = db.Column(db.DateTime, nullable=False, default=get_local_time)
    usuario = db.Column(db.Text, nullable=False)
    status = db.Column(db.Text, nullable=True, default='Rascunho')
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
            'observacoes': self.observacoes,
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
    data_entrada = db.Column(db.DateTime, nullable=False, default=get_local_time)
    usuario = db.Column(db.Text, nullable=False)
    preenchimento = db.relationship('SolicitacoesPreenchidas', backref='estoque')
    material = db.relationship('Materiais', backref='estoque')

# Modelo para Auditoria
class Auditoria(db.Model):
    __tablename__ = 'Auditoria'
    id = db.Column(db.Integer, primary_key=True)
    solicitacao_id = db.Column(db.Integer, db.ForeignKey('SolicitacoesCompra.id'), nullable=False)
    data_validacao = db.Column(db.DateTime, nullable=False, default=get_local_time)
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
    data_requisicao = db.Column(db.DateTime, nullable=False, default=get_local_time)
    usuario = db.Column(db.Text, nullable=False)
    preenchimento = db.relationship('SolicitacoesPreenchidas', backref='requisicoes')
    material = db.relationship('Materiais', backref='requisicoes')

# Modelo para Pedidos de Compra
class PedidosCompra(db.Model):
    __tablename__ = 'PedidosCompra'
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    numero_pedido = db.Column(db.Text, nullable=False, unique=True)
    data_criacao = db.Column(db.DateTime, nullable=False, default=get_local_time)
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

class HistoricoDescontos(db.Model):
    __tablename__ = 'HistoricoDescontos'
    id = db.Column(db.Integer, primary_key=True)
    preenchimento_id = db.Column(db.Integer, db.ForeignKey('SolicitacoesPreenchidas.id'), nullable=False)
    valor_unitario_anterior = db.Column(db.Float, nullable=False)
    valor_unitario_novo = db.Column(db.Float, nullable=False)
    valor_frete_anterior = db.Column(db.Float, nullable=True)
    valor_frete_novo = db.Column(db.Float, nullable=True)
    data_alteracao = db.Column(db.DateTime, nullable=False, default=get_local_time)
    usuario = db.Column(db.Text, nullable=False)

    def to_dict(self):
        return {
            'id': self.id,
            'preenchimento_id': self.preenchimento_id,
            'valor_unitario_anterior': self.valor_unitario_anterior,
            'valor_unitario_novo': self.valor_unitario_novo,
            'valor_frete_anterior': self.valor_frete_anterior,
            'valor_frete_novo': self.valor_frete_novo,
            'data_alteracao': self.data_alteracao.isoformat() if self.data_alteracao else None,
            'usuario': self.usuario
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
        create_historico_descontos_table()

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

def get_rastreabilidade_entries(solicitacao):
    """Gera entradas de rastreabilidade com hora local correta"""
    entries = []
    
    # As datas já estão em horário de Brasília (sem timezone)
    # Entrada de criação
    data_criacao = solicitacao.data_solicitacao
    
    entries.append({
        'data': data_criacao,
        'evento': 'Solicitação Criada',
        'descricao': f'Solicitação registrada no sistema por {solicitacao.usuario}'
    })
    
    # Entrada de aprovação/reprovação (se aplicável)
    if solicitacao.status_aprovacao:
        if solicitacao.status_aprovacao == 'Aprovado':
            entries.append({
                'data': data_criacao,  # Usar data de criação como fallback
                'evento': 'Aprovação',
                'descricao': f'Solicitação aprovada'
            })
        elif solicitacao.status_aprovacao == 'Reprovado':
            # Tentar obter data da auditoria
            auditoria = Auditoria.query.filter_by(solicitacao_id=solicitacao.id).first()
            data_reprovacao = auditoria.data_validacao if auditoria else data_criacao
            
            # Verificar observações para ver se foi reprovação individual
            motivo = ''
            if solicitacao.observacoes_col and 'REPROVADA INDIVIDUALMENTE' in solicitacao.observacoes_col:
                motivo = solicitacao.observacoes_col
            
            entries.append({
                'data': data_reprovacao,
                'evento': 'Reprovação',
                'descricao': f'Solicitação reprovada' + (f'. Motivo: {motivo}' if motivo else '')
            })
    
    # Entrada de cotação (se existir preenchimento)
    if solicitacao.preenchimentos_fornecidos:
        primeiro_preenchimento = solicitacao.preenchimentos_fornecidos[0]
        data_preenchimento = primeiro_preenchimento.data_preenchimento
        
        # Obter nome do fornecedor para a descrição
        fornecedor_nome = get_fornecedor_nome(primeiro_preenchimento.fornecedor_id)
        
        entries.append({
            'data': data_preenchimento,
            'evento': 'Cotação Enviada',
            'descricao': f'Cotação enviada para o fornecedor {fornecedor_nome}'
        })
        
        # Entrada de resposta do fornecedor (mesma data do preenchimento)
        entries.append({
            'data': data_preenchimento,
            'evento': 'Resposta do Fornecedor',
            'descricao': f'Fornecedor {fornecedor_nome} respondeu com os valores informados'
        })
    
    # Ordenar por data
    entries.sort(key=lambda x: x['data'] if x['data'] else datetime.min)
    return entries


@app.template_filter('format_brasil_time')
def format_brasil_time(dt):
    """Filtro Jinja2 para formatar datetimes no formato brasileiro"""
    if dt is None:
        return ""
    
    # Se é string, tentar converter para datetime
    if isinstance(dt, str):
        try:
            # Tentar diferentes formatos
            for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M:%S.%f']:
                try:
                    dt = datetime.strptime(dt, fmt)
                    break
                except ValueError:
                    continue
        except:
            return str(dt)
    
    # Se já é um objeto datetime, formatar diretamente
    if hasattr(dt, 'strftime'):
        return dt.strftime('%d/%m/%Y %H:%M')
    
    return str(dt)

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
    data_alteracao = db.Column(db.DateTime, nullable=False, default=get_local_time)
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
            return senhas
        
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
                    
                    # CORREÇÃO: Armazenar tanto o usuário completo quanto o nome base
                    senhas[usuario] = {
                        "senha": senha,
                        "pagina": pagina,
                        "empresa": empresa,
                        "usuario_base": usuario.split('%')[0]  # Adiciona o nome base do usuário
                    }
        
        print(f"Total de usuários carregados: {len(senhas)}")
        print(f"Usuários disponíveis: {list(senhas.keys())}")  # DEBUG
        
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


def registrar_log(usuario, tipo_acao, descricao=None, ip=None):
    """
    Registra ação no arquivo de log
    
    Args:
        usuario: Nome do usuário
        tipo_acao: Tipo de ação ('login', 'logout', 'reprovar_solicitacao_individual', etc.)
        descricao: Descrição adicional da ação (opcional)
        ip: Endereço IP do usuário (opcional)
    """
    acao_map = {
        'login': 'Acesso',
        'logout': 'Logout',
        'reprovar_solicitacao_individual': 'Reprovação Individual de Solicitação',
        'aprovar_solicitacao': 'Aprovação de Solicitação',
        'reprovar_solicitacao': 'Reprovação de Solicitação'
    }
    
    acao = acao_map.get(tipo_acao, tipo_acao)
    ip_info = f" - IP: {ip}" if ip else ''
    descricao_info = f" - {descricao}" if descricao else ''
    
    try:
        with open("arquivo.log", "a", encoding="utf-8") as log_file:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_file.write(f"{timestamp} - {acao} do usuário: {usuario}{ip_info}{descricao_info}\n")
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
    get_rastreabilidade_entries=get_rastreabilidade_entries,
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
    if request.endpoint in ['routes_bp.login', 'static', 'routes_bp.debug_senhas']:
        return
    
    print(f"Verificando sessão para endpoint: {request.endpoint}")
    print(f"Sessão atual: {dict(session)}")
    
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
        usuario_completo = request.form.get('usuario', '').strip()
        senha = request.form.get('senha', '').strip()
        
        senhas = ler_senhas()
        
        if usuario_completo in senhas and senha == senhas[usuario_completo]["senha"]:
            # CORREÇÃO: Usar o usuário base para a sessão
            usuario_base = senhas[usuario_completo]["usuario_base"]
            session['usuario'] = usuario_base
            session['usuario_completo'] = usuario_completo  # Guardar o completo
            session['_fresh'] = True
            
            # Redireciona para o menu correto
            pagina_destino = senhas[usuario_completo]["pagina"].replace('.html', '')
            return redirect(url_for(f'routes_bp.{pagina_destino}'))
        
        flash('Credenciais inválidas', 'error')
    
    return render_template('login.html')

@routes_bp.before_request
def verificar_acesso_menu():
    """Verifica se o usuário está acessando o menu correto"""
    
    # Lista de endpoints de menus
    endpoints_menus = {
        'routes_bp.menu_master': 'menu_master.html',
        'routes_bp.menu_supervisor': 'menu_supervisor.html',
        'routes_bp.menu_aprovacao': 'menu_aprovacao.html',
        'routes_bp.menu_comprador': 'menu_comprador.html',
        'routes_bp.menu_cadastro': 'menu_cadastro.html',
        'routes_bp.menu_solicitante': 'menu_solicitante.html',
        'routes_bp.menu_financeiro': 'menu_financeiro.html',
        'routes_bp.menu_estoquista': 'menu_estoquista.html',
        'routes_bp.menu_auditoria': 'menu_auditoria.html'
    }
    
    # Se não for um endpoint de menu, não faz nada
    if request.endpoint not in endpoints_menus:
        return
    
    # Se não estiver logado, redireciona para login
    if 'usuario' not in session:
        flash('Faça login para acessar.', 'error')
        return redirect(url_for('routes_bp.login'))
    
    # Verifica qual menu o usuário deveria acessar
    usuario_completo = session.get('usuario_completo')
    if usuario_completo:
        senhas = ler_senhas()
        if usuario_completo in senhas:
            menu_permitido = senhas[usuario_completo]["pagina"]
            
            # Se o menu que está tentando acessar não for o permitido
            if menu_permitido != endpoints_menus[request.endpoint]:
                # Limpa a sessão e pede login novamente
                session.clear()
                flash('Acesso não permitido. Faça login novamente.', 'error')
                return redirect(url_for('routes_bp.login'))

@routes_bp.route('/logout', methods=['GET'])
def logout():
    usuario = session.get('usuario')
    if usuario:
        registrar_log(usuario, 'logout', request.remote_addr)
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
        
        # Verifica se é uma atualização ou criação nova
        if usuario in senhas:
            mensagem = 'Senha atualizada com sucesso.'
        else:
            mensagem = 'Senha adicionada com sucesso.'

        senhas[usuario] = {
            "senha": senha,
            "pagina": pagina,
            "empresa": empresa
        }

        salvar_senhas(senhas)
        flash(mensagem, 'success')
        return redirect(url_for('routes_bp.listar_senhas'))

    except Exception as e:
        flash('Erro ao adicionar/atualizar senha.', 'error')
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

#Usuario Desativado
@routes_bp.route('/desativado', methods=['GET'])
def desativado():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    return render_template('desativado.html')

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

@routes_bp.route('/solicitar_compra', defaults={'cod': None}, methods=['GET'])
@routes_bp.route('/solicitar_compra/<int:cod>', methods=['GET'])
def solicitar_compra(cod):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))

    try:
        senhas = ler_senhas()
        usuario = session.get('usuario')

        if usuario not in senhas:
            flash('Usuário não encontrado.', 'error')
            return redirect(url_for('routes_bp.login'))

        empresa_usuario = senhas[usuario].get('empresa', '')
        
        # BUSCAR TODOS OS MATERIAIS - ADICIONE ESTA LINHA
        todos_materiais = Materiais.query.all()

        # Se o código foi passado (botão de material)
        if cod:
            material = Materiais.query.get_or_404(cod)
            return render_template(
                'solicitar_compra.html',
                material=material,
                empresa_usuario=empresa_usuario,
                todos_materiais=todos_materiais  # ADICIONE ESTE PARÂMETRO
            )
        else:
            # Se o acesso veio do menu (sem material específico)
            return render_template(
                'solicitar_compra.html',
                material=None,
                empresa_usuario=empresa_usuario,
                todos_materiais=todos_materiais  # ADICIONE ESTE PARÂMETRO
            )

    except SQLAlchemyError as e:
        flash(f'Erro ao carregar material: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))
    except Exception as e:
        flash(f'Erro inesperado: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))
    
@routes_bp.route('/debug_materiais')
def debug_materiais():
    """Rota temporária para debug dos materiais"""
    try:
        materiais = Materiais.query.all()
        resultado = {
            'total': len(materiais),
            'materiais': [{
                'CodMaterial': m.CodMaterial,
                'DescricaoMaterial': m.DescricaoMaterial
            } for m in materiais]
        }
        return jsonify(resultado)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@routes_bp.route('/api/materiais_search', methods=['GET'])
def api_materiais_search():
    """API para buscar materiais com paginação para Select2"""
    try:
        search_term = request.args.get('q', '').strip()
        page = request.args.get('page', 1, type=int)
        per_page = 30
        
        # Query base
        query = Materiais.query
        
        # Aplicar filtro de busca se houver termo
        if search_term:
            if search_term.isdigit():
                # Busca por código
                query = query.filter(Materiais.CodMaterial == int(search_term))
            else:
                # Busca por descrição
                query = query.filter(Materiais.DescricaoMaterial.ilike(f'%{search_term}%'))
        
        # Paginação
        materiais_paginated = query.order_by(Materiais.DescricaoMaterial).paginate(
            page=page, 
            per_page=per_page, 
            error_out=False
        )
        
        # Formatar resultados para Select2
        materiais_data = []
        for material in materiais_paginated.items:
            materiais_data.append({
                'id': material.CodMaterial,
                'text': f"{material.DescricaoMaterial} (Cód: {material.CodMaterial})",
                'CodMaterial': material.CodMaterial,
                'DescricaoMaterial': material.DescricaoMaterial
            })
        
        return jsonify({
            'success': True,
            'materiais': materiais_data,
            'total_count': materiais_paginated.total,
            'pagination': {
                'more': materiais_paginated.has_next
            }
        })
        
    except Exception as e:
        logging.error(f"Erro na API de busca de materiais: {str(e)}")
        return jsonify({
            'success': False, 
            'error': str(e),
            'materiais': []
        }), 500

@routes_bp.route('/api/materiais', methods=['GET'])
def api_materiais():
    """API para buscar todos os materiais"""
    try:
        print("🔍 Acessando API de materiais...")
        
        # Buscar todos os materiais
        materiais = Materiais.query.all()
        print(f"📦 Materiais encontrados: {len(materiais)}")
        
        # Log dos primeiros 3 materiais para debug
        for i, material in enumerate(materiais[:3]):
            print(f"  {i+1}. Cód: {material.CodMaterial}, Desc: {material.DescricaoMaterial}")
        
        if not materiais:
            print("⚠️ Nenhum material encontrado no banco de dados!")
            return jsonify({
                'success': False, 
                'error': 'Nenhum material cadastrado',
                'materiais': []
            })
        
        resultado = {
            'success': True,
            'materiais': [{
                'CodMaterial': m.CodMaterial,
                'DescricaoMaterial': m.DescricaoMaterial
            } for m in materiais]
        }
        
        print("✅ API de materiais retornando dados com sucesso")
        return jsonify(resultado)
        
    except Exception as e:
        print(f"❌ Erro na API de materiais: {str(e)}")
        logging.error(f"Erro na API de materiais: {str(e)}", exc_info=True)
        return jsonify({
            'success': False, 
            'error': str(e),
            'materiais': []
        }), 500

@routes_bp.route('/abrir_solicitacao', methods=['POST'])
def abrir_solicitacao():
    if 'usuario' not in session:
        flash('Você precisa estar logado para fazer uma solicitação.', 'error')
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Obter arrays do formulário
        cod_materiais = request.form.getlist('cod_material[]')
        especificacoes = request.form.getlist('especificacao[]')
        quantidades = request.form.getlist('quantidade[]')
        unidades_medida = request.form.getlist('unidade_medida[]')
        aplicacoes = request.form.getlist('aplicacao[]')  # Aplicações específicas por item
        empresas = request.form.getlist('empresa[]')
        marcas = request.form.getlist('marca[]')
        ativos = request.form.getlist('ativo[]')
        nomes_ativos = request.form.getlist('nome_ativo[]')
        prioridades = request.form.getlist('prioridade[]')
        fotos = request.files.getlist('foto[]')

        # DEBUG: Verificar o que está chegando
        aplicacao_geral = request.form.get('aplicacao', '').strip()
        print(f"🔍 DEBUG - Dados recebidos:")
        print(f"  Total de itens: {len(cod_materiais)}")
        print(f"  Aplicação geral: '{aplicacao_geral}'")
        print(f"  Aplicações específicas: {aplicacoes}")
        print(f"  Empresas: {empresas}")
        print(f"  Quantidades: {quantidades}")

        # Processar cada solicitação
        solicitacoes_criadas = 0
        
        for i, cod_material in enumerate(cod_materiais):
            if not cod_material:
                print(f"⚠️  Pular índice {i}: cod_material vazio")
                continue  # Pular se não tem código de material
                
            # Validação do cod_material
            try:
                cod_material = int(cod_material)
                if cod_material <= 0:
                    raise ValueError
            except ValueError:
                flash(f'Código do material inválido na solicitação {i+1}.', 'error')
                print(f"❌ Código material inválido: {cod_material}")
                continue

            # Verificar se o material existe
            material = db.session.get(Materiais, cod_material)
            if not material:
                flash(f'Material com código {cod_material} não encontrado na solicitação {i+1}.', 'error')
                print(f"❌ Material não encontrado: {cod_material}")
                continue

            # Validação dos campos obrigatórios para esta solicitação
            if (i >= len(especificacoes) or not especificacoes[i] or 
                i >= len(quantidades) or not quantidades[i] or
                i >= len(unidades_medida) or not unidades_medida[i] or
                i >= len(empresas) or not empresas[i] or
                i >= len(ativos) or not ativos[i] or
                i >= len(prioridades) or not prioridades[i]):
                flash(f'Campos obrigatórios não preenchidos na solicitação {i+1}.', 'error')
                print(f"❌ Campos obrigatórios faltando no índice {i}")
                continue

            # CORREÇÃO: Processar aplicação corretamente
            aplicacao = None

            # Primeiro tenta pegar a aplicação específica do item
            if i < len(aplicacoes) and aplicacoes[i] and aplicacoes[i].strip():
                aplicacao = aplicacoes[i].strip()
                print(f"✅ Aplicação específica encontrada para índice {i}: '{aplicacao}'")
            else:
                # Se não tem aplicação específica, tenta a aplicação geral do formulário
                if aplicacao_geral:
                    aplicacao = aplicacao_geral
                    print(f"✅ Usando aplicação geral do formulário: '{aplicacao}'")
                else:
                    # Se não tem aplicação geral, usa a do material como fallback
                    aplicacao = material.Aplicacao if material and material.Aplicacao else 'Aplicação não especificada'
                    print(f"ℹ️  Usando aplicação do material: '{aplicacao}'")

            # Processar upload da foto se existir
            foto_path = None
            if i < len(fotos) and fotos[i] and fotos[i].filename:
                foto = fotos[i]
                if allowed_file(foto.filename, {'jpg', 'jpeg', 'png', 'pdf'}):
                    filename = f"{uuid.uuid4()}_{secure_filename(foto.filename)}"
                    foto_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    foto.save(foto_path)
                    print(f"📸 Foto salva: {foto_path}")

            # Criar a solicitação de compra
            solicitacao = SolicitacoesCompra(
                cod_material=cod_material,
                especificacao=especificacoes[i],
                quantidade=int(quantidades[i]),
                unidade_medida=unidades_medida[i],
                aplicacao = aplicacao.strip().lower(),  # Normaliza para evitar duplicidade
                aplicacao_geral=aplicacao_geral,  # Salva a aplicação geral separadamente
                empresa=empresas[i],
                usuario=session['usuario'],
                foto_path=foto_path,
                marca=marcas[i] if i < len(marcas) and marcas[i] else None,
                ativo=ativos[i],
                nome_ativo=nomes_ativos[i] if i < len(nomes_ativos) and ativos[i] == 'Sim' and nomes_ativos[i] else None,
                prioridade=prioridades[i],
                status_aprovacao=None,  # Garantir que seja NULL
                comprador_atribuido=get_next_comprador(aplicacao)  # 🔹 atribuição automática
            )
            db.session.add(solicitacao)
            solicitacoes_criadas += 1
            print(f"✅ Solicitação {i+1} criada: Material {cod_material}, Qtd {quantidades[i]}, Aplicação: '{aplicacao}'")

        if solicitacoes_criadas > 0:
            db.session.commit()
            flash(f'{solicitacoes_criadas} solicitações de compra abertas com sucesso.', 'success')
            print(f"🎉 {solicitacoes_criadas} solicitações salvas no banco")
        else:
            flash('Nenhuma solicitação válida para criar.', 'error')
            print("❌ Nenhuma solicitação criada")

        return redirect(url_for('routes_bp.solicitar_compra'))
        
    except Exception as e:
        db.session.rollback()
        logging.error(f"Error in abrir_solicitacao: {str(e)}")
        print(f"💥 ERRO CRÍTICO: {str(e)}")
        flash(f'Erro ao abrir solicitações: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))

def get_compradores():
    """Retorna lista de nomes de compradores a partir do arquivo senhas.txt"""
    senhas = ler_senhas()
    compradores = []
    
    for usuario_completo, dados in senhas.items():
        # CORREÇÃO: Verificar se a página contém 'comprador' e usar o usuário base
        if 'comprador' in dados['pagina'].lower():
            compradores.append(dados['usuario_base'])
    
    return sorted(set(compradores))  # Usar set para remover duplicatasrn sorted(compradores)


@routes_bp.route('/listar_solicitacoes', methods=['GET'])
def listar_solicitacoes():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))

    try:
        # Parâmetros de paginação
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 10, type=int)
        
        # Parâmetros de filtro
        empresa = request.args.get('empresa')
        usuario = request.args.get('usuario')
        aplicacao_filter = request.args.get('aplicacao')
        status = request.args.get('status')
        comprador_filter = request.args.get('comprador')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')

        # Subquery: IDs das solicitações que já estão em pedidos de compra
        subquery_pedidos = db.session.query(
            SolicitacoesPreenchidas.solicitacao_id
        ).join(
            pedido_preenchimento_associacao,
            SolicitacoesPreenchidas.id == pedido_preenchimento_associacao.c.preenchimento_id
        ).distinct()

        # Query base
        query = SolicitacoesCompra.query.filter(
            SolicitacoesCompra.status_aprovacao == "Aprovado",
            ~SolicitacoesCompra.id.in_(subquery_pedidos)
        )

        # Aplicar filtros
        if empresa:
            query = query.filter(SolicitacoesCompra.empresa == empresa)
        if usuario:
            query = query.filter(SolicitacoesCompra.usuario == usuario)
        if aplicacao_filter:
            query = query.filter(SolicitacoesCompra.aplicacao == aplicacao_filter)
        if status:
            query = query.filter(SolicitacoesCompra.status_aprovacao == status)
        if comprador_filter:
            if comprador_filter == 'Não atribuído':
                query = query.filter(
                    (SolicitacoesCompra.comprador_atribuido == None) | 
                    (SolicitacoesCompra.comprador_atribuido == '')
                )
            else:
                query = query.filter(SolicitacoesCompra.comprador_atribuido == comprador_filter)
        if data_inicio:
            query = query.filter(SolicitacoesCompra.data_solicitacao >= datetime.strptime(data_inicio, '%Y-%m-%d'))
        if data_fim:
            query = query.filter(SolicitacoesCompra.data_solicitacao <= datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1))

        # Obter todas as solicitações
        solicitacoes = query.all()
        
        # Agrupar por aplicação E data completa (com hora)
        grupos = {}
        for s in solicitacoes:
            if s.aplicacao:
                # Criar chave única com aplicação + data completa (com hora)
                data_hora_formatada = s.data_solicitacao.strftime('%Y-%m-%d %H:%M')
                chave_grupo = f"{s.aplicacao}|{data_hora_formatada}"
                
                if chave_grupo not in grupos:
                    grupos[chave_grupo] = []
                
                # Verificar se a solicitação pode ser reprovada individualmente
                preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                    solicitacao_id=s.id
                ).filter(
                    SolicitacoesPreenchidas.status != 'Rascunho'
                ).all()
                
                s.pode_reprovar_individual = len(preenchimentos) == 0
                grupos[chave_grupo].append(s)

        # Ordenar grupos pela data mais recente
        grupos_ordenados = sorted(grupos.items(), key=lambda x: x[1][0].data_solicitacao if x[1] else datetime.min, reverse=True)
        
        # Calcular paginação
        total_grupos = len(grupos_ordenados)
        total_paginas = (total_grupos + per_page - 1) // per_page
        inicio = (page - 1) * per_page
        fim = inicio + per_page
        
        grupos_paginados = grupos_ordenados[inicio:fim]

        # Listas para filtros (usando todas as solicitações)
        empresas = sorted({s.empresa for s in solicitacoes if s.empresa})
        usuarios = sorted({s.usuario for s in solicitacoes if s.usuario})
        aplicacoes = sorted({s.aplicacao for s in solicitacoes if s.aplicacao})
        compradores = get_compradores()

        return render_template(
            'listar_solicitacoes.html',
            grupos_paginados=grupos_paginados,
            empresas=empresas,
            usuarios=usuarios,
            aplicacoes=aplicacoes,
            compradores=compradores,
            page=page,
            per_page=per_page,
            total_paginas=total_paginas,
            total_grupos=total_grupos,
            request_args=request.args
        )

    except Exception as e:
        flash(f'Erro ao carregar solicitações: {str(e)}', 'error')
        return render_template(
            'listar_solicitacoes.html',
            grupos_paginados=[],
            empresas=[],
            usuarios=[],
            aplicacoes=[],
            compradores=[],
            page=1,
            per_page=10,
            total_paginas=1,
            total_grupos=0,
            request_args={}
        )
    
#Tela Comprado
@routes_bp.route('/listar_solicitacoes_comprador', methods=['GET'])
def listar_solicitacoes_comprador():
    """Rota do menu comprador para exibir solicitações APROVADAS para preenchimento."""
    
    if 'usuario' not in session:
        flash('Acesso não autorizado', 'error')
        return redirect(url_for('routes_bp.login'))

    try:
        # Subquery: solicitações que já têm preenchimento DIFERENTE de 'Rascunho'
        # (ou seja, já foram enviadas para aprovação)
        subquery = db.session.query(SolicitacoesPreenchidas.solicitacao_id).filter(
            SolicitacoesPreenchidas.status != 'Rascunho'
        ).distinct()
        
        # Buscar solicitações APROVADAS que NÃO têm preenchimento enviado
        # (pode ter rascunho ou não ter nenhum preenchimento)
        solicitacoes = SolicitacoesCompra.query.filter(
            SolicitacoesCompra.status_aprovacao == 'Aprovado',
            ~SolicitacoesCompra.id.in_(subquery)
        ).all()

        # DEBUG: Mostrar informações das solicitações encontradas
        print(f"DEBUG - Solicitações APROVADAS sem preenchimento final: {len(solicitacoes)}")
        for s in solicitacoes[:5]:  # Mostrar apenas as 5 primeiras para debug
            print(f"  ID: {s.id}, Material: {s.material.DescricaoMaterial if s.material else 'N/A'}, Status Aprovação: {s.status_aprovacao}")
            
            # Verificar se tem preenchimentos
            preenchimentos = SolicitacoesPreenchidas.query.filter_by(solicitacao_id=s.id).all()
            if preenchimentos:
                for p in preenchimentos:
                    print(f"    Preenchimento ID: {p.id}, Status: {p.status}")

        # Coletar dados para filtros
        empresas = db.session.query(SolicitacoesCompra.empresa)\
                             .distinct().order_by(SolicitacoesCompra.empresa).all()
        empresas = [e[0] for e in empresas if e[0]]

        usuarios = db.session.query(SolicitacoesCompra.usuario)\
                             .distinct().order_by(SolicitacoesCompra.usuario).all()
        usuarios = [u[0] for u in usuarios if u[0]]

        aplicacoes = db.session.query(SolicitacoesCompra.aplicacao)\
                               .distinct().order_by(SolicitacoesCompra.aplicacao).all()
        aplicacoes = [a[0] for a in aplicacoes if a[0]]

        # Obter compradores
        compradores = get_compradores()

        return render_template(
            'listar_solicitacoes_comprador.html',
            solicitacoes=solicitacoes,
            empresas=empresas,
            usuarios=usuarios,
            aplicacoes=aplicacoes,
            compradores=compradores,
            titulo_pagina="Solicitações Aprovadas - Preencher Cotação"
        )

    except Exception as e:
        flash(f'Erro ao carregar solicitações: {str(e)}', 'error')
        print(f"ERRO: {str(e)}")
        import traceback
        traceback.print_exc()
        
        return render_template(
            'listar_solicitacoes_comprador.html',
            solicitacoes=[],
            empresas=[],
            usuarios=[],
            aplicacoes=[],
            compradores=[],
            titulo_pagina="Solicitações Aprovadas - Preencher Cotação"
        )

@routes_bp.route('/aprovar_solicitacao', methods=['GET'])
def aprovar_solicitacao():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Obter parâmetro de filtro
        filtro_status = request.args.get('filtro_status', 'pendentes')
        aplicacao_filtro = request.args.get('aplicacao', '')  # NOVO: Filtro por aplicação
        
        print(f"🚀 DEBUG - aprovar_solicitacao com filtro: {filtro_status}")
        
        if filtro_status == 'pendentes':
            # Buscar solicitações pendentes
            solicitacoes_none = SolicitacoesCompra.query.filter(
                SolicitacoesCompra.status_aprovacao.is_(None)
            ).all()
            
            solicitacoes_empty = SolicitacoesCompra.query.filter(
                SolicitacoesCompra.status_aprovacao == ''
            ).all()
            
            solicitacoes_pendente = SolicitacoesCompra.query.filter(
                SolicitacoesCompra.status_aprovacao == 'Pendente'
            ).all()
            
            # Combinar todos os resultados
            solicitacoes = solicitacoes_none + solicitacoes_empty + solicitacoes_pendente
            
            # Remover duplicatas
            seen_ids = set()
            solicitacoes_unique = []
            for sol in solicitacoes:
                if sol.id not in seen_ids:
                    seen_ids.add(sol.id)
                    solicitacoes_unique.append(sol)
            
            solicitacoes = solicitacoes_unique
            
        elif filtro_status == 'abertas':
            # Solicitações sem preenchimento FINALIZADO
            subquery = db.session.query(SolicitacoesPreenchidas.solicitacao_id).filter(
                SolicitacoesPreenchidas.status.in_(['Aprovado', 'Reprovado', 'Entregue', 'Em Processamento'])
            ).distinct()
            
            solicitacoes = SolicitacoesCompra.query.filter(
                ~SolicitacoesCompra.id.in_(subquery)
            ).all()
            
        elif filtro_status == 'aprovadas':
            solicitacoes = SolicitacoesCompra.query.filter(
                SolicitacoesCompra.status_aprovacao == 'Aprovado'
            ).all()
            
        elif filtro_status == 'reprovadas':
            solicitacoes = SolicitacoesCompra.query.filter(
                SolicitacoesCompra.status_aprovacao == 'Reprovado'
            ).all()
            
        elif filtro_status == 'preenchidas':
            subquery = db.session.query(SolicitacoesPreenchidas.solicitacao_id).filter(
                SolicitacoesPreenchidas.status.in_(['Aprovado', 'Reprovado', 'Entregue', 'Em Processamento'])
            ).distinct()
            
            solicitacoes = SolicitacoesCompra.query.filter(
                SolicitacoesCompra.id.in_(subquery)
            ).all()
            
        else:
            solicitacoes = SolicitacoesCompra.query.all()

        # CORREÇÃO COMPLETA: Obter aplicações únicas de forma mais abrangente
        aplicacoes_unicas = set()
        
        # 1. Buscar aplicações dos materiais
        materiais_com_aplicacao = Materiais.query.filter(
            Materiais.Aplicacao.isnot(None),
            Materiais.Aplicacao != ''
        ).all()
        for material in materiais_com_aplicacao:
            if material.Aplicacao and material.Aplicacao.strip():
                aplicacoes_unicas.add(material.Aplicacao.strip())
        
        # 2. Buscar aplicações das solicitações (tanto aplicacao quanto aplicacao_geral)
        solicitacoes_com_aplicacao = SolicitacoesCompra.query.filter(
            or_(
                SolicitacoesCompra.aplicacao.isnot(None),
                SolicitacoesCompra.aplicacao_geral.isnot(None)
            )
        ).all()
        
        for solicitacao in solicitacoes_com_aplicacao:
            if solicitacao.aplicacao and solicitacao.aplicacao.strip():
                aplicacoes_unicas.add(solicitacao.aplicacao.strip())
            if solicitacao.aplicacao_geral and solicitacao.aplicacao_geral.strip():
                aplicacoes_unicas.add(solicitacao.aplicacao_geral.strip())
        
        # 3. Buscar também das solicitações atuais para garantir completude
        for solicitacao in solicitacoes:
            if solicitacao.aplicacao and solicitacao.aplicacao.strip():
                aplicacoes_unicas.add(solicitacao.aplicacao.strip())
            if solicitacao.aplicacao_geral and solicitacao.aplicacao_geral.strip():
                aplicacoes_unicas.add(solicitacao.aplicacao_geral.strip())
            # Se não tem aplicação específica, usar a do material
            elif solicitacao.material and solicitacao.material.Aplicacao and solicitacao.material.Aplicacao.strip():
                aplicacoes_unicas.add(solicitacao.material.Aplicacao.strip())
        
        aplicacoes_unicas = sorted(aplicacoes_unicas)
        
        print(f"🔍 DEBUG - Total de aplicações únicas encontradas: {len(aplicacoes_unicas)}")
        for i, aplicacao in enumerate(aplicacoes_unicas[:10]):  # Mostrar apenas as primeiras 10 para debug
            print(f"  {i+1}. {aplicacao}")

        # NOVO: Aplicar filtro por aplicação se especificado
        if aplicacao_filtro:
            solicitacoes_filtradas = []
            for solicitacao in solicitacoes:
                # Verificar aplicação em várias fontes
                aplicacao_encontrada = False
                
                # 1. Verificar aplicação específica da solicitação
                if solicitacao.aplicacao and aplicacao_filtro.lower() in solicitacao.aplicacao.lower():
                    aplicacao_encontrada = True
                
                # 2. Verificar aplicação geral da solicitação
                elif solicitacao.aplicacao_geral and aplicacao_filtro.lower() in solicitacao.aplicacao_geral.lower():
                    aplicacao_encontrada = True
                
                # 3. Verificar aplicação do material
                elif solicitacao.material and solicitacao.material.Aplicacao and aplicacao_filtro.lower() in solicitacao.material.Aplicacao.lower():
                    aplicacao_encontrada = True
                
                if aplicacao_encontrada:
                    solicitacoes_filtradas.append(solicitacao)
            
            solicitacoes = solicitacoes_filtradas
            print(f"✅ Aplicado filtro por aplicação: '{aplicacao_filtro}' - {len(solicitacoes)} solicitações encontradas")

        def get_usuarios_empresas():
            usuarios = set()
            empresas = set()
            try:
                with open('senhas.txt', 'r', encoding='utf-8') as f:
                    for line in f:
                        partes = line.strip().split('%')
                        if len(partes) >= 4:
                            usuario = partes[0]
                            empresa = partes[3]
                            usuarios.add(usuario)
                            empresas.add(empresa)
            except Exception as e:
                logging.error(f"Erro ao ler senhas.txt: {str(e)}")
            return sorted(usuarios), sorted(empresas)
        
        usuarios, empresas = get_usuarios_empresas()
        
        if not solicitacoes:
            flash('Nenhuma solicitação encontrada com os filtros aplicados.', 'info')
            print("⚠️ DEBUG - Nenhuma solicitação para exibir")
        
        return render_template('aprovar_solicitacao.html', 
                            solicitacoes=solicitacoes,
                            empresas=empresas,
                            usuarios=usuarios,
                            aplicacoes_unicas=aplicacoes_unicas,  # CORRIGIDO
                            filtro_status=filtro_status,
                            aplicacao_filtro=aplicacao_filtro)
        
    except Exception as e:
        logging.error(f"Error in aprovar_solicitacao: {str(e)}")
        print(f"❌ DEBUG - Erro: {str(e)}")
        flash(f'Erro ao carregar solicitações: {str(e)}', 'error')
        return redirect(url_for('routes_bp.buscar_material'))
    
@routes_bp.route('/debug_ultimas_solicitacoes')
def debug_ultimas_solicitacoes():
    """Debug das últimas solicitações criadas"""
    if 'usuario' not in session:
        return jsonify({'error': 'Não autorizado'}), 401
    
    try:
        # Buscar as últimas 5 solicitações
        solicitacoes = SolicitacoesCompra.query.order_by(SolicitacoesCompra.id.desc()).limit(5).all()
        
        resultado = []
        for sol in solicitacoes:
            resultado.append({
                'id': sol.id,
                'cod_material': sol.cod_material,
                'descricao_material': sol.material.DescricaoMaterial if sol.material else 'N/A',
                'especificacao': sol.especificacao,
                'quantidade': sol.quantidade,
                'empresa': sol.empresa,
                'usuario': sol.usuario,
                'data_solicitacao': sol.data_solicitacao.isoformat() if sol.data_solicitacao else None,
                'status_aprovacao': str(sol.status_aprovacao),
                'ativo': sol.ativo,
                'nome_ativo': sol.nome_ativo,
                'prioridade': sol.prioridade
            })
        
        return jsonify({
            'total_solicitacoes': len(solicitacoes),
            'solicitacoes': resultado
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@routes_bp.route('/debug_solicitacoes_status')
def debug_solicitacoes_status():
    """Debug completo do status das solicitações"""
    if 'usuario' not in session:
        return jsonify({'error': 'Não autorizado'}), 401
    
    try:
        # Buscar as últimas 20 solicitações
        solicitacoes = SolicitacoesCompra.query.order_by(SolicitacoesCompra.id.desc()).limit(20).all()
        
        resultado = []
        for sol in solicitacoes:
            resultado.append({
                'id': sol.id,
                'cod_material': sol.cod_material,
                'especificacao': sol.especificacao[:50] + '...' if sol.especificacao else '',
                'status_aprovacao': str(sol.status_aprovacao),  # Converter para string para ver NULL
                'status_aprovacao_raw': sol.status_aprovacao,
                'data_solicitacao': sol.data_solicitacao.isoformat() if sol.data_solicitacao else None,
                'usuario': sol.usuario,
                'is_null': sol.status_aprovacao is None,
                'is_empty': sol.status_aprovacao == '',
                'is_pendente': sol.status_aprovacao == 'Pendente'
            })
        
        # Contar por status
        counts = {
            'total': len(solicitacoes),
            'null': len([s for s in resultado if s['is_null']]),
            'empty': len([s for s in resultado if s['is_empty']]),
            'pendente': len([s for s in resultado if s['is_pendente']]),
            'aprovado': len([s for s in resultado if s['status_aprovacao'] == 'Aprovado']),
            'reprovado': len([s for s in resultado if s['status_aprovacao'] == 'Reprovado']),
        }
        
        return jsonify({
            'counts': counts,
            'solicitacoes': resultado
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@routes_bp.route('/debug_ultima_solicitacao')
def debug_ultima_solicitacao():
    """Debug da última solicitação criada"""
    if 'usuario' not in session:
        return jsonify({'error': 'Não autorizado'}), 401
    
    try:
        ultima_solicitacao = SolicitacoesCompra.query.order_by(SolicitacoesCompra.id.desc()).first()
        
        if not ultima_solicitacao:
            return jsonify({'message': 'Nenhuma solicitação encontrada'})
        
        resultado = {
            'id': ultima_solicitacao.id,
            'cod_material': ultima_solicitacao.cod_material,
            'especificacao': ultima_solicitacao.especificacao,
            'status_aprovacao': str(ultima_solicitacao.status_aprovacao),
            'status_aprovacao_is_none': ultima_solicitacao.status_aprovacao is None,
            'data_solicitacao': ultima_solicitacao.data_solicitacao.isoformat() if ultima_solicitacao.data_solicitacao else None,
            'usuario': ultima_solicitacao.usuario,
            'empresa': ultima_solicitacao.empresa
        }
        
        return jsonify(resultado)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@routes_bp.route('/debug_solicitacoes_recentes')
def debug_solicitacoes_recentes():
    """Rota para debug - mostra as últimas 10 solicitações criadas"""
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        solicitacoes = SolicitacoesCompra.query.order_by(SolicitacoesCompra.id.desc()).limit(10).all()
        
        resultado = []
        for sol in solicitacoes:
            resultado.append({
                'id': sol.id,
                'cod_material': sol.cod_material,
                'especificacao': sol.especificacao,
                'status_aprovacao': sol.status_aprovacao,
                'data_solicitacao': sol.data_solicitacao.isoformat() if sol.data_solicitacao else None,
                'usuario': sol.usuario
            })
        
        return jsonify({
            'total': len(solicitacoes),
            'solicitacoes': resultado
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
@routes_bp.route('/debug_solicitacoes_detalhado')
def debug_solicitacoes_detalhado():
    """Debug detalhado de TODAS as solicitações"""
    if 'usuario' not in session:
        return jsonify({'error': 'Não autorizado'}), 401
    
    try:
        solicitacoes = SolicitacoesCompra.query.order_by(SolicitacoesCompra.id.desc()).limit(10).all()
        
        resultado = []
        for sol in solicitacoes:
            # Verificar valores exatos
            status_raw = sol.status_aprovacao
            resultado.append({
                'id': sol.id,
                'cod_material': sol.cod_material,
                'especificacao': sol.especificacao[:50] + '...' if sol.especificacao else '',
                'status_aprovacao_raw': status_raw,
                'status_aprovacao_repr': repr(status_raw),
                'status_aprovacao_type': type(status_raw).__name__,
                'is_none': status_raw is None,
                'is_empty_string': status_raw == '',
                'is_pendente': status_raw == 'Pendente',
                'data_solicitacao': sol.data_solicitacao.isoformat() if sol.data_solicitacao else None,
                'usuario': sol.usuario
            })
        
        return jsonify({
            'total': len(solicitacoes),
            'solicitacoes': resultado
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@routes_bp.route('/debug_verificar_coluna_status')
def debug_verificar_coluna_status():
    """Verificar se a coluna existe e seus valores"""
    try:
        # Verificar estrutura da tabela
        result = db.engine.execute("PRAGMA table_info(SolicitacoesCompra)").fetchall()
        colunas = [col[1] for col in result]
        
        # Verificar valores diretamente via SQL
        sql_result = db.engine.execute("""
            SELECT id, status_aprovacao, 
                   typeof(status_aprovacao) as tipo,
                   length(status_aprovacao) as tamanho
            FROM SolicitacoesCompra 
            ORDER BY id DESC 
            LIMIT 10
        """).fetchall()
        
        sql_data = []
        for row in sql_result:
            sql_data.append({
                'id': row[0],
                'status_aprovacao': row[1],
                'tipo': row[2],
                'tamanho': row[3] if row[3] is not None else 'NULL'
            })
        
        return jsonify({
            'colunas_tabela': colunas,
            'existe_status_aprovacao': 'status_aprovacao' in colunas,
            'dados_sql': sql_data
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
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
        flash('Acesso não autorizado', 'error')
        return redirect(url_for('routes_bp.login'))
    
    # === FUNÇÃO AUXILIAR PARA CONVERSÃO DE VALORES ===
    def parse_br_currency_final(value_str):
        if not value_str:
            return 0.0
        
        valor = str(value_str)
        valor = valor.strip()
        valor = valor.replace('R$', '').replace('r$', '').strip()
        
        import re
        valor = re.sub(r'[^\d,\.\-]', '', valor)
        
        if not valor or valor == '-':
            return 0.0
        
        negativo = False
        if valor.startswith('-'):
            negativo = True
            valor = valor[1:]
        
        valor = valor.lstrip('0')
        if valor == '' or valor.startswith(('.', ',')):
            valor = '0' + valor
        
        try:
            if ',' not in valor and '.' not in valor:
                resultado = float(valor)
            else:
                num_virgulas = valor.count(',')
                num_pontos = valor.count('.')
                
                if num_pontos >= 1 and num_virgulas == 1:
                    valor_sem_pontos = valor.replace('.', '')
                    valor_final = valor_sem_pontos.replace(',', '.')
                    resultado = float(valor_final)
                elif num_virgulas == 1 and num_pontos == 0:
                    partes = valor.split(',')
                    if len(partes) == 2:
                        resultado = float(f"{partes[0]}.{partes[1]}")
                    else:
                        resultado = float(valor.replace(',', '.'))
                elif num_pontos == 1 and num_virgulas == 0:
                    resultado = float(valor)
                elif num_virgulas >= 1 and num_pontos == 1:
                    valor_sem_virgulas = valor.replace(',', '')
                    resultado = float(valor_sem_virgulas)
                else:
                    valor_limpo = valor.replace('.', '').replace(',', '')
                    resultado = float(valor_limpo)
        except ValueError:
            resultado = 0.0
        
        if negativo:
            resultado = -resultado
        
        return resultado
    
    # === MODO GRUPO ===
    grupo_ids_param = request.args.get('grupo_ids', '')
    grupo_ids = []
    modo_grupo = False
    solicitacoes_grupo = []
    
    if grupo_ids_param:
        try:
            grupo_ids = [int(x) for x in grupo_ids_param.split(',') if x.strip().isdigit()]
            if grupo_ids:
                solicitacoes_grupo = SolicitacoesCompra.query.filter(
                    SolicitacoesCompra.id.in_(grupo_ids)
                ).all()
                modo_grupo = len(solicitacoes_grupo) > 1
        except:
            flash('IDs de grupo inválidos.', 'error')
            return redirect(url_for('routes_bp.listar_solicitacoes_comprador'))
    
    solicitacao = SolicitacoesCompra.query.get_or_404(id)
    
    if modo_grupo and id not in grupo_ids:
        if solicitacao not in solicitacoes_grupo:
            solicitacoes_grupo.append(solicitacao)
            grupo_ids.append(id)
    
    todas_solicitacoes = solicitacoes_grupo if modo_grupo else [solicitacao]
    todas_solicitacoes_ids = [s.id for s in todas_solicitacoes]
    
    if request.method == 'POST':
        action = request.form.get('action')
        
        if not action:
            if 'salvar_rascunho' in request.form:
                action = 'salvar_rascunho'
            elif 'salvar_finalizar' in request.form:
                action = 'salvar_finalizar'
        
        is_rascunho = action == 'salvar_rascunho'
        is_finalizar = action == 'salvar_finalizar'
        
        if not (is_rascunho or is_finalizar):
            flash('Ação inválida. Selecione "Salvar Rascunho" ou "Finalizar".', 'error')
            return redirect(request.url)
        
        try:
            usuario = session['usuario']
            
            # ============================================
            # PROCESSAR FORNECEDORES REMOVIDOS
            # ============================================
            fornecedores_removidos_json = request.form.get('fornecedores_removidos', '[]')
            fornecedores_ids_remover = []
            
            if fornecedores_removidos_json and fornecedores_removidos_json != '[]':
                try:
                    import json
                    fornecedores_ids_remover = json.loads(fornecedores_removidos_json)
                    
                    if fornecedores_ids_remover:
                        print(f"🚨 Processando exclusão de fornecedores: {fornecedores_ids_remover}")
                        
                        for fornecedor_id in fornecedores_ids_remover:
                            if fornecedor_id and fornecedor_id != '':
                                preenchimentos = SolicitacoesPreenchidas.query.filter(
                                    SolicitacoesPreenchidas.solicitacao_id.in_(todas_solicitacoes_ids),
                                    SolicitacoesPreenchidas.fornecedor_id == int(fornecedor_id),
                                    SolicitacoesPreenchidas.status == 'Rascunho'
                                ).all()
                                
                                for preenchimento in preenchimentos:
                                    HistoricoDescontos.query.filter_by(
                                        preenchimento_id=preenchimento.id
                                    ).delete()
                                    db.session.delete(preenchimento)
                                    print(f"✅ Preenchimento ID {preenchimento.id} do fornecedor {fornecedor_id} excluído")
                        
                        db.session.commit()
                        print(f"✅ Todos os preenchimentos dos fornecedores {fornecedores_ids_remover} foram removidos")
                        
                except Exception as e:
                    db.session.rollback()
                    print(f"❌ Erro ao remover cotações: {str(e)}")
            
            # === Captura de dados em listas ===
            fornecedores_ids = request.form.getlist('fornecedor_id[]')
            valores_frete = request.form.getlist('valor_frete[]')
            prazos = request.form.getlist('prazo_entrega[]')
            condicoes = request.form.getlist('condicao_pagamento[]')
            observacoes = request.form.getlist('observacao[]')
            preenchimento_ids = request.form.getlist('preenchimento_id[]')  # 🔴 IMPORTANTE!
            pdf_files = request.files.getlist('pdf_file[]')
            
            # === Captura de dados dos materiais ===
            todos_valores_unitarios = request.form.getlist('valor_unitario[]')
            todos_ids_solicitacao = request.form.getlist('solicitacao_id[]')
            
            print(f"📋 Preenchimento IDs recebidos: {preenchimento_ids}")
            print(f"📋 Fornecedores IDs recebidos: {fornecedores_ids}")
            print(f"📋 Valores unitários: {len(todos_valores_unitarios)}")
            
            # === Processar cada cotação ===
            cotacoes_salvas = 0
            total_cotacoes = 0
            erros_validacao = []
            
            for idx, fornecedor_id in enumerate(fornecedores_ids):
                fornecedor_id = fornecedor_id.strip()
                if not fornecedor_id:
                    continue
                
                # PULAR se este fornecedor está na lista de removidos
                if fornecedor_id in fornecedores_ids_remover:
                    print(f"⏭️ Pulando fornecedor {fornecedor_id} (marcado para exclusão)")
                    continue
                
                total_cotacoes += 1
                
                # === Campos da cotação atual ===
                vf_str = valores_frete[idx].strip() if idx < len(valores_frete) else '0'
                prazo = prazos[idx].strip() if idx < len(prazos) else ''
                condicao = condicoes[idx].strip() if idx < len(condicoes) else ''
                obs = observacoes[idx].strip() if idx < len(observacoes) else ''
                preenchimento_id = preenchimento_ids[idx] if idx < len(preenchimento_ids) and preenchimento_ids[idx] else None
                
                print(f"🔍 Processando cotação {idx+1}: fornecedor {fornecedor_id}, preenchimento_id {preenchimento_id}")
                
                # === VALIDAÇÃO EXTRA PARA FINALIZAR ===
                if is_finalizar:
                    if not prazo:
                        erros_validacao.append(f'Cotação {idx+1}: Prazo de entrega é obrigatório.')
                    if not condicao:
                        erros_validacao.append(f'Cotação {idx+1}: Condição de pagamento é obrigatória.')
                
                # === CONVERSÃO SEGURA ===
                valor_frete = parse_br_currency_final(vf_str)
                
                # === Processar cada material desta cotação ===
                num_materiais_por_cotacao = len(todas_solicitacoes)
                inicio_idx = idx * num_materiais_por_cotacao
                
                if inicio_idx >= len(todos_valores_unitarios):
                    erros_validacao.append(f'Cotação {idx+1}: Dados dos materiais incompletos.')
                    continue
                
                primeiro_material_desta_cotacao = True
                
                for i, sol in enumerate(todas_solicitacoes):
                    material_idx = inicio_idx + i
                    
                    if material_idx >= len(todos_valores_unitarios):
                        erros_validacao.append(f'Cotação {idx+1}, Material {i+1}: Valor unitário faltando.')
                        continue
                    
                    valor_unitario_str = todos_valores_unitarios[material_idx].strip()
                    
                    if not valor_unitario_str or valor_unitario_str in ['0', '0.00', '0,00']:
                        erros_validacao.append(f'Cotação {idx+1}, Material {i+1}: Valor unitário é obrigatório.')
                        continue
                    
                    valor_unitario = parse_br_currency_final(valor_unitario_str)
                    
                    if valor_unitario <= 0:
                        erros_validacao.append(f'Cotação {idx+1}, Material {i+1}: Valor unitário deve ser maior que zero.')
                        continue
                    
                    valor_total = round(valor_unitario * sol.quantidade, 2)
                    
                    # 🔴 BUSCAR PREENCHIMENTO - PRIORIDADE 1: PELO ID ESPECÍFICO
                    preenchimento = None
                    
                    if primeiro_material_desta_cotacao and preenchimento_id and preenchimento_id != '':
                        try:
                            preenchimento = SolicitacoesPreenchidas.query.get(int(preenchimento_id))
                            print(f"   🔍 Busca por ID {preenchimento_id}: {'Encontrado' if preenchimento else 'Não encontrado'}")
                        except:
                            pass
                    
                    # 🔴 BUSCAR PREENCHIMENTO - PRIORIDADE 2: PELA COMBINAÇÃO
                    if not preenchimento:
                        preenchimento = SolicitacoesPreenchidas.query.filter_by(
                            solicitacao_id=sol.id,
                            fornecedor_id=int(fornecedor_id),
                            status='Rascunho'
                        ).first()
                        print(f"   🔍 Busca por fornecedor {fornecedor_id} + solicitação {sol.id}: {'Encontrado' if preenchimento else 'Não encontrado'}")
                    
                    # 🔴 SE AINDA NÃO TEM, CRIAR NOVO
                    if not preenchimento:
                        preenchimento = SolicitacoesPreenchidas(
                            solicitacao_id=sol.id,
                            fornecedor_id=int(fornecedor_id),
                            valor_unitario=valor_unitario,
                            valor_total=valor_total,
                            valor_frete=valor_frete if valor_frete > 0 else None,
                            prazo_entrega=prazo,
                            condicao_pagamento=condicao,
                            observacoes=obs,
                            data_preenchimento=get_local_time(),
                            usuario=usuario,
                            status='Rascunho' if is_rascunho else 'Aguardando Aprovacao'
                        )
                        db.session.add(preenchimento)
                        print(f"   ✅ Novo preenchimento criado para solicitação {sol.id}")
                    else:
                        # ATUALIZAR PREENCHIMENTO EXISTENTE
                        print(f"   🔄 Atualizando preenchimento ID {preenchimento.id}")
                        
                        # Verificar mudanças para histórico
                        valor_unitario_anterior = float(preenchimento.valor_unitario) if preenchimento.valor_unitario else 0.0
                        valor_frete_anterior = float(preenchimento.valor_frete) if preenchimento.valor_frete else 0.0
                        
                        valor_unitario_mudou = abs(valor_unitario_anterior - valor_unitario) > 0.001
                        
                        valor_frete_mudou = False
                        if valor_frete > 0:
                            if valor_frete_anterior is None or abs(valor_frete_anterior - valor_frete) > 0.001:
                                valor_frete_mudou = True
                        else:
                            if valor_frete_anterior is not None and valor_frete_anterior > 0:
                                valor_frete_mudou = True
                        
                        if valor_unitario_mudou or valor_frete_mudou:
                            historico = HistoricoDescontos(
                                preenchimento_id=preenchimento.id,
                                valor_unitario_anterior=valor_unitario_anterior,
                                valor_unitario_novo=valor_unitario,
                                valor_frete_anterior=valor_frete_anterior if valor_frete_anterior != 0 else None,
                                valor_frete_novo=valor_frete if valor_frete > 0 else None,
                                data_alteracao=get_local_time(),
                                usuario=usuario
                            )
                            db.session.add(historico)
                            print(f"   📝 Histórico de desconto registrado")
                        
                        # Atualizar dados
                        preenchimento.valor_unitario = valor_unitario
                        preenchimento.valor_total = valor_total
                        preenchimento.valor_frete = valor_frete if valor_frete > 0 else None
                        preenchimento.prazo_entrega = prazo
                        preenchimento.condicao_pagamento = condicao
                        preenchimento.observacoes = obs
                        preenchimento.status = 'Rascunho' if is_rascunho else 'Aguardando Aprovacao'
                        preenchimento.data_preenchimento = get_local_time()
                        preenchimento.usuario = usuario
                    
                    # PDF apenas para o primeiro material da cotação
                    if idx < len(pdf_files) and i == 0 and pdf_files[idx] and pdf_files[idx].filename:
                        pdf_file = pdf_files[idx]
                        if allowed_file(pdf_file.filename, {'pdf'}):
                            timestamp = get_local_time().strftime('%Y%m%d_%H%M%S')
                            safe_name = secure_filename(pdf_file.filename)
                            unique_name = f"cotacao_{preenchimento.id}_{timestamp}_{safe_name}"
                            path = os.path.join(app.config['UPLOAD_FOLDER'], unique_name)
                            pdf_file.save(path)
                            preenchimento.pdf_path = unique_name
                            print(f"   📎 PDF salvo: {unique_name}")
                    
                    primeiro_material_desta_cotacao = False
                
                cotacoes_salvas += 1
                print(f"✅ Cotação {idx+1} processada com sucesso")
            
            if erros_validacao:
                for erro in erros_validacao:
                    flash(erro, 'error')
                return redirect(request.url)
            
            if total_cotacoes == 0:
                flash('É necessário preencher pelo menos uma cotação.', 'error')
                return redirect(request.url)
            
            db.session.commit()
            print(f"💾 COMMIT REALIZADO! {cotacoes_salvas} materiais em {total_cotacoes} cotações")
            
            if is_rascunho:
                flash(f'{cotacoes_salvas} material(is) em {total_cotacoes} cotação(ões) salvo(s) como rascunho!', 'success')
            else:
                flash(f'{cotacoes_salvas} material(is) em {total_cotacoes} cotação(ões) finalizado(s) e enviado(s) para aprovação!', 'success')
            
            return redirect(url_for('routes_bp.listar_solicitacoes_comprador'))
        
        except Exception as e:
            db.session.rollback()
            flash(f'Erro ao salvar: {str(e)}', 'error')
            import traceback
            print(f"❌ Erro completo: {traceback.format_exc()}")
            return redirect(request.url)
    
    # === GET: Carregar rascunhos ===
    if modo_grupo:
        cotacoes_raw = SolicitacoesPreenchidas.query.filter(
            SolicitacoesPreenchidas.solicitacao_id.in_(grupo_ids),
            SolicitacoesPreenchidas.status == 'Rascunho'
        ).order_by(SolicitacoesPreenchidas.fornecedor_id, SolicitacoesPreenchidas.solicitacao_id).all()
    else:
        cotacoes_raw = SolicitacoesPreenchidas.query.filter_by(
            solicitacao_id=id,
            status='Rascunho'
        ).order_by(SolicitacoesPreenchidas.fornecedor_id).all()
    
    # Organizar cotações por fornecedor
    cotacoes_por_fornecedor = {}
    for c in cotacoes_raw:
        if c.fornecedor_id not in cotacoes_por_fornecedor:
            cotacoes_por_fornecedor[c.fornecedor_id] = []
        cotacoes_por_fornecedor[c.fornecedor_id].append(c)
    
    # Criar estrutura para o template
    cotacoes_estruturadas = []
    
    for fornecedor_id, cotacoes_fornecedor in cotacoes_por_fornecedor.items():
        if not cotacoes_fornecedor:
            continue
        
        cotacao_ref = cotacoes_fornecedor[0]
        
        # Buscar dados do fornecedor
        fornecedor_info = {
            'nome_fantasia': '',
            'cnpj': '',
            'telefone': '',
            'endereco': ''
        }
        
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT nome_fantasia, cnpj, telefone, endereco, bairro, cidade, estado
                FROM fornecedores
                WHERE id = ?
            """, (fornecedor_id,))
            row = cursor.fetchone()
            conn.close()
            
            if row:
                fornecedor_info['nome_fantasia'] = row[0] or ''
                fornecedor_info['cnpj'] = format_cnpj(row[1]) if row[1] else ''
                fornecedor_info['telefone'] = row[2] or ''
                endereco_parts = []
                if row[3]:
                    endereco_parts.append(row[3])
                if row[4]:
                    endereco_parts.append(row[4])
                if row[5] and row[6]:
                    endereco_parts.append(f"{row[5]}/{row[6]}")
                fornecedor_info['endereco'] = ', '.join(endereco_parts) if endereco_parts else 'Não informado'
        
        cotacao_estruturada = {
            'fornecedor_id': fornecedor_id,
            'fornecedor_nome_fantasia': fornecedor_info['nome_fantasia'],
            'fornecedor_cnpj': fornecedor_info['cnpj'],
            'fornecedor_telefone': fornecedor_info['telefone'],
            'fornecedor_endereco': fornecedor_info['endereco'],
            'valor_frete': cotacao_ref.valor_frete,
            'prazo_entrega': cotacao_ref.prazo_entrega or '',
            'condicao_pagamento': cotacao_ref.condicao_pagamento or '',
            'observacoes': cotacao_ref.observacoes or '',
            'pdf_path': cotacao_ref.pdf_path,
            'materiais': []
        }
        
        for cotacao in cotacoes_fornecedor:
            solicitacao_match = next((s for s in todas_solicitacoes if s.id == cotacao.solicitacao_id), None)
            
            if solicitacao_match:
                historico_descontos = []
                try:
                    historico_query = HistoricoDescontos.query.filter_by(
                        preenchimento_id=cotacao.id
                    ).order_by(HistoricoDescontos.data_alteracao.desc()).all()
                    
                    for historico in historico_query:
                        historico_descontos.append({
                            'valor_unitario_anterior': historico.valor_unitario_anterior,
                            'valor_unitario_novo': historico.valor_unitario_novo,
                            'valor_frete_anterior': historico.valor_frete_anterior,
                            'valor_frete_novo': historico.valor_frete_novo,
                            'data_alteracao': historico.data_alteracao.strftime('%d/%m/%Y %H:%M:%S'),
                            'usuario': historico.usuario
                        })
                except Exception as e:
                    print(f"DEBUG - Erro ao buscar histórico de descontos: {str(e)}")
                
                material_info = {
                    'preenchimento_id': cotacao.id,  # 🔴 ADICIONAR ESTE CAMPO!
                    'solicitacao_id': cotacao.solicitacao_id,
                    'valor_unitario': float(cotacao.valor_unitario) if cotacao.valor_unitario else 0.0,
                    'valor_total': float(cotacao.valor_total) if cotacao.valor_total else 0.0,
                    'descricao': solicitacao_match.material.DescricaoMaterial if solicitacao_match.material else '',
                    'quantidade': float(solicitacao_match.quantidade) if solicitacao_match.quantidade else 0.0,
                    'unidade': solicitacao_match.unidade_medida or 'un',
                    'historico_descontos': historico_descontos
                }
                cotacao_estruturada['materiais'].append(material_info)
        
        cotacoes_estruturadas.append(cotacao_estruturada)
    
    if not cotacoes_estruturadas:
        cotacoes_estruturadas = [{
            'fornecedor_id': None,
            'fornecedor_nome_fantasia': '',
            'fornecedor_cnpj': '',
            'fornecedor_telefone': '',
            'fornecedor_endereco': '',
            'valor_frete': None,
            'prazo_entrega': '',
            'condicao_pagamento': '',
            'observacoes': '',
            'pdf_path': None,
            'materiais': []
        }]
    
    return render_template(
        'preencher_solicitacao.html',
        solicitacao=solicitacao,
        solicitacoes_grupo=todas_solicitacoes,
        cotacoes_salvas=cotacoes_estruturadas,
        modo_grupo=modo_grupo,
        grupo_ids=grupo_ids_param
    )

# Imprimir -------------------------------------------------------------------------
@routes_bp.route('/cotacao_imprimir/<int:id>', methods=['GET'])
def cotacao_imprimir(id):
    if 'usuario' not in session:
        flash('Acesso não autorizado', 'error')
        return redirect(url_for('routes_bp.login'))

    # === VERIFICAR SE EXISTE GRUPO ATIVO NA SESSÃO ===
    grupo_sessao = session.get('grupo_impressao', [])
    grupo_ids_param = request.args.get('grupo_ids', '')
    grupo_ids = []
    modo_grupo = False
    solicitacoes_grupo = []

    # Primeiro: verificar se veio por parâmetro
    if grupo_ids_param:
        try:
            grupo_ids = [int(x) for x in grupo_ids_param.split(',') if x.strip().isdigit()]
            if grupo_ids:
                solicitacoes_grupo = SolicitacoesCompra.query.filter(
                    SolicitacoesCompra.id.in_(grupo_ids)
                ).all()
                modo_grupo = len(solicitacoes_grupo) > 1
        except Exception as e:
            print(f"DEBUG - Erro ao processar grupo_ids: {e}")
            flash('IDs de grupo inválidos.', 'error')
    
    # Segundo: verificar se tem grupo na sessão
    elif grupo_sessao:
        try:
            # Limpar IDs inválidos
            grupo_ids = [gid for gid in grupo_sessao if isinstance(gid, int) and gid > 0]
            if grupo_ids:
                solicitacoes_grupo = SolicitacoesCompra.query.filter(
                    SolicitacoesCompra.id.in_(grupo_ids)
                ).all()
                modo_grupo = len(solicitacoes_grupo) > 1
                print(f"DEBUG - Usando grupo da sessão: {grupo_ids}")
        except Exception as e:
            print(f"DEBUG - Erro ao processar grupo da sessão: {e}")
    
    # Terceiro: se não tem grupo, buscar solicitação principal
    if not modo_grupo:
        solicitacao_principal = SolicitacoesCompra.query.get_or_404(id)
        solicitacoes_grupo = [solicitacao_principal]
        grupo_ids = [id]
    
    # Configuração final do grupo
    if modo_grupo:
        # Garantir que a solicitação principal está no grupo
        solicitacao_principal_id = id
        if solicitacao_principal_id not in grupo_ids:
            grupo_ids.append(solicitacao_principal_id)
        
        # Buscar todas as solicitações do grupo (agora atualizado)
        solicitacoes_grupo = SolicitacoesCompra.query.filter(
            SolicitacoesCompra.id.in_(grupo_ids)
        ).all()
        
        # Ordenar por ID para consistência
        solicitacoes_grupo.sort(key=lambda x: x.id)
        grupo_ids.sort()
    else:
        # Modo individual
        solicitacao_principal = SolicitacoesCompra.query.get_or_404(id)
        solicitacoes_grupo = [solicitacao_principal]
        grupo_ids = [id]

    # Garantir que temos a solicitação principal
    solicitacao = solicitacoes_grupo[0]

    # === BUSCAR TODAS AS COTAÇÕES ===
    print(f"DEBUG - Buscando cotações para IDs: {grupo_ids}")
    
    # Buscar TODAS as cotações para TODOS os IDs do grupo
    cotacoes_raw = SolicitacoesPreenchidas.query.filter(
        SolicitacoesPreenchidas.solicitacao_id.in_(grupo_ids)
    ).order_by(SolicitacoesPreenchidas.fornecedor_id, SolicitacoesPreenchidas.solicitacao_id).all()

    print(f"DEBUG - Encontradas {len(cotacoes_raw)} cotações no total")

    # Organizar cotações por fornecedor
    cotacoes_por_fornecedor = {}
    for c in cotacoes_raw:
        if c.fornecedor_id not in cotacoes_por_fornecedor:
            cotacoes_por_fornecedor[c.fornecedor_id] = []
        cotacoes_por_fornecedor[c.fornecedor_id].append(c)

    # Criar estrutura para o template
    cotacoes_estruturadas = []
    
    # Coletar TODOS os IDs de solicitação únicos de todas as cotações
    todos_ids_solicitacao = set()
    for fornecedor_id, cotacoes_fornecedor in cotacoes_por_fornecedor.items():
        for c in cotacoes_fornecedor:
            todos_ids_solicitacao.add(c.solicitacao_id)
    
    print(f"DEBUG - IDs de solicitação encontrados nas cotações: {todos_ids_solicitacao}")
    
    # Buscar informações de TODAS as solicitações encontradas
    solicitacoes_completas = {}
    if todos_ids_solicitacao:
        solicitacoes_encontradas = SolicitacoesCompra.query.filter(
            SolicitacoesCompra.id.in_(list(todos_ids_solicitacao))
        ).all()
        for sol in solicitacoes_encontradas:
            solicitacoes_completas[sol.id] = sol
    
    # Se não encontrou nenhum material nas cotações, usar as solicitações do grupo
    if not solicitacoes_completas:
        for sol in solicitacoes_grupo:
            solicitacoes_completas[sol.id] = sol
    
    print(f"DEBUG - Total de solicitações completas: {len(solicitacoes_completas)}")
    
    for fornecedor_id, cotacoes_fornecedor in cotacoes_por_fornecedor.items():
        if not cotacoes_fornecedor:
            continue
            
        # Pegar a primeira cotação como referência para dados gerais
        cotacao_ref = cotacoes_fornecedor[0]
        
        # Buscar dados do fornecedor
        fornecedor_info = {
            'nome_fantasia': '',
            'cnpj': '',
            'telefone': '',
            'endereco': ''
        }
        
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT nome_fantasia, cnpj, telefone, endereco, bairro, cidade, estado 
                FROM fornecedores 
                WHERE id = ?
            """, (fornecedor_id,))
            row = cursor.fetchone()
            conn.close()
            
            if row:
                fornecedor_info['nome_fantasia'] = row[0] or ''
                fornecedor_info['cnpj'] = row[1] or ''  # Sem formatação para evitar erros
                fornecedor_info['telefone'] = row[2] or ''
                endereco_parts = []
                if row[3]:
                    endereco_parts.append(row[3])
                if row[4]:
                    endereco_parts.append(row[4])
                if row[5] and row[6]:
                    endereco_parts.append(f"{row[5]}/{row[6]}")
                fornecedor_info['endereco'] = ', '.join(endereco_parts) if endereco_parts else 'Não informado'
        
        # Calcular descontos totais para esta cotação (soma de todos os descontos dos materiais)
        desconto_total = 0
        try:
            historico_query = HistoricoDescontos.query.filter_by(
                preenchimento_id=cotacao_ref.id
            ).all()
            for historico in historico_query:
                if historico.valor_unitario_novo and historico.valor_unitario_anterior:
                    desconto_total += (historico.valor_unitario_anterior - historico.valor_unitario_novo)
        except Exception as e:
            print(f"DEBUG - Erro ao calcular desconto total: {e}")
        
        # Criar estrutura da cotação
        cotacao_estruturada = {
            'id': cotacao_ref.id,
            'fornecedor_id': fornecedor_id,
            'fornecedor_nome_fantasia': fornecedor_info['nome_fantasia'],
            'fornecedor_cnpj': fornecedor_info['cnpj'],
            'fornecedor_telefone': fornecedor_info['telefone'],
            'fornecedor_endereco': fornecedor_info['endereco'],
            'valor_frete': cotacao_ref.valor_frete,
            'prazo_entrega': cotacao_ref.prazo_entrega or '',
            'condicao_pagamento': cotacao_ref.condicao_pagamento or '',
            'observacoes': cotacao_ref.observacoes or '',
            'desconto_valor': desconto_total,
            'pdf_path': cotacao_ref.pdf_path,
            'materiais': []
        }
        
        # Adicionar TODOS os materiais desta cotação
        for cotacao in cotacoes_fornecedor:
            # Buscar a solicitação correspondente
            solicitacao_match = solicitacoes_completas.get(cotacao.solicitacao_id)
            
            if solicitacao_match:
                # Buscar histórico de descontos para esta cotação específica
                historico_descontos = []
                try:
                    historico_query = HistoricoDescontos.query.filter_by(
                        preenchimento_id=cotacao.id
                    ).order_by(HistoricoDescontos.data_alteracao.desc()).all()
                    
                    for historico in historico_query:
                        historico_descontos.append({
                            'valor_unitario_anterior': historico.valor_unitario_anterior,
                            'valor_unitario_novo': historico.valor_unitario_novo,
                            'valor_frete_anterior': historico.valor_frete_anterior,
                            'valor_frete_novo': historico.valor_frete_novo,
                            'data_alteracao': historico.data_alteracao.strftime('%d/%m/%Y %H:%M:%S'),
                            'usuario': historico.usuario
                        })
                except Exception as e:
                    print(f"DEBUG - Erro ao buscar histórico de descontos: {str(e)}")
                
                material_info = {
                    'solicitacao_id': cotacao.solicitacao_id,
                    'valor_unitario': float(cotacao.valor_unitario) if cotacao.valor_unitario else 0.0,
                    'valor_total': float(cotacao.valor_total) if cotacao.valor_total else 0.0,
                    'descricao': solicitacao_match.material.DescricaoMaterial if solicitacao_match.material else '',
                    'quantidade': float(solicitacao_match.quantidade) if solicitacao_match.quantidade else 0.0,
                    'unidade': solicitacao_match.unidade_medida or 'un',
                    'historico_descontos': historico_descontos
                }
                cotacao_estruturada['materiais'].append(material_info)
        
        cotacoes_estruturadas.append(cotacao_estruturada)
    
    # Se não houver cotações salvas, criar uma estrutura vazia para o template
    if not cotacoes_estruturadas:
        cotacoes_estruturadas = [{
            'id': None,
            'fornecedor_id': None,
            'fornecedor_nome_fantasia': '',
            'fornecedor_cnpj': '',
            'fornecedor_telefone': '',
            'fornecedor_endereco': '',
            'valor_frete': None,
            'prazo_entrega': '',
            'condicao_pagamento': '',
            'observacoes': '',
            'desconto_valor': 0,
            'pdf_path': None,
            'materiais': []
        }]

    # Criar uma lista de TODOS os materiais únicos para o template
    todos_materiais_unicos = []
    materiais_vistos = set()
    
    # Primeiro, adicionar todos os materiais das solicitações do grupo
    for sol in solicitacoes_grupo:
        if sol.id not in materiais_vistos:
            materiais_vistos.add(sol.id)
            material_completo = {
                'id': sol.id,
                'descricao': sol.material.DescricaoMaterial if sol.material else '',
                'quantidade': float(sol.quantidade) if sol.quantidade else 0.0,
                'unidade': sol.unidade_medida or 'un',
                'especificacao': sol.especificacao or '' if hasattr(sol, 'especificacao') else ''
            }
            todos_materiais_unicos.append(material_completo)
    
    # Depois, adicionar materiais das cotações que não estão no grupo
    for cotacao in cotacoes_estruturadas:
        for material in cotacao.get('materiais', []):
            material_id = material.get('solicitacao_id')
            if material_id and material_id not in materiais_vistos:
                materiais_vistos.add(material_id)
                
                # Buscar informações completas do material
                solicitacao_material = solicitacoes_completas.get(material_id)
                if solicitacao_material:
                    material_completo = {
                        'id': material_id,
                        'descricao': solicitacao_material.material.DescricaoMaterial if solicitacao_material.material else '',
                        'quantidade': float(solicitacao_material.quantidade) if solicitacao_material.quantidade else 0.0,
                        'unidade': solicitacao_material.unidade_medida or 'un',
                        'especificacao': solicitacao_material.especificacao or '' if hasattr(solicitacao_material, 'especificacao') else ''
                    }
                    todos_materiais_unicos.append(material_completo)
    
    print(f"DEBUG - Total de materiais únicos encontrados: {len(todos_materiais_unicos)}")
    print(f"DEBUG - Total de cotações: {len(cotacoes_estruturadas)}")
    print(f"DEBUG - Modo grupo: {modo_grupo}")

    return render_template(
        'cotacao_imprimir.html',
        solicitacao=solicitacao,
        solicitacoes_grupo=solicitacoes_grupo,
        todos_materiais=todos_materiais_unicos,
        cotacoes_salvas=cotacoes_estruturadas,
        modo_grupo=modo_grupo,
        grupo_ids=','.join(str(gid) for gid in grupo_ids)
    )
#------------------------------------------------------------------

@routes_bp.route('/listar_solicitacoes_preenchidas', methods=['GET'])
def listar_solicitacoes_preenchidas():
    """Lista solicitações preenchidas com tratamento robusto de erros"""
    
    # Verificação de autenticação
    if 'usuario' not in session:
        flash('🔒 Acesso não autorizado. Faça login para continuar.', 'warning')
        return redirect(url_for('routes_bp.login'))
    
    # Inicializar variáveis com valores padrão
    preenchimentos_por_material = {}
    empresas = []
    usuarios = []
    filtros = {
        'empresa': '',
        'usuario': '',
        'data_inicio': '',
        'data_fim': ''
    }
    
    try:
        # 1. Coletar e validar parâmetros de filtro
        empresa = request.args.get('empresa', '').strip()
        usuario = request.args.get('usuario', '').strip()
        data_inicio_str = request.args.get('data_inicio', '').strip()
        data_fim_str = request.args.get('data_fim', '').strip()
        
        # Armazenar filtros para o template
        filtros['empresa'] = empresa
        filtros['usuario'] = usuario
        filtros['data_inicio'] = data_inicio_str
        filtros['data_fim'] = data_fim_str
        
        # 2. Validar e converter datas
        data_inicio = None
        data_fim = None
        
        if data_inicio_str:
            try:
                data_inicio = datetime.strptime(data_inicio_str, '%Y-%m-%d')
            except ValueError:
                flash('⚠️ Data inicial inválida. Use o formato AAAA-MM-DD.', 'warning')
                data_inicio_str = ''
                filtros['data_inicio'] = ''
        
        if data_fim_str:
            try:
                data_fim = datetime.strptime(data_fim_str, '%Y-%m-%d')
                # Adicionar 1 dia para incluir todo o dia final
                data_fim_ajustada = data_fim + timedelta(days=1)
            except ValueError:
                flash('⚠️ Data final inválida. Use o formato AAAA-MM-DD.', 'warning')
                data_fim_str = ''
                filtros['data_fim'] = ''
        
        # 3. Construir query com tratamento seguro
        query = SolicitacoesPreenchidas.query
        
        # Join com SolicitacoesCompra para filtros de empresa
        query = query.join(SolicitacoesCompra)
        
        # Aplicar filtros com validação
        if empresa:
            # Verificar se a empresa existe na tabela
            empresa_existe = db.session.query(
                SolicitacoesCompra.empresa
            ).filter(SolicitacoesCompra.empresa == empresa).first()
            
            if empresa_existe:
                query = query.filter(SolicitacoesCompra.empresa == empresa)
            else:
                flash(f'⚠️ Empresa "{empresa}" não encontrada nos registros.', 'warning')
        
        if usuario:
            # Verificar se o usuário existe
            usuario_existe = db.session.query(
                SolicitacoesPreenchidas.usuario
            ).filter(SolicitacoesPreenchidas.usuario == usuario).first()
            
            if usuario_existe:
                query = query.filter(SolicitacoesPreenchidas.usuario == usuario)
            else:
                flash(f'⚠️ Usuário "{usuario}" não encontrado nos registros.', 'warning')
        
        if data_inicio:
            query = query.filter(SolicitacoesPreenchidas.data_preenchimento >= data_inicio)
        
        if data_fim:
            query = query.filter(SolicitacoesPreenchidas.data_preenchimento <= data_fim_ajustada)
        
        # 4. Executar query com timeout
        try:
            preenchimentos = query.order_by(
                SolicitacoesPreenchidas.data_preenchimento.desc()
            ).all()
            
            app.logger.info(f'✅ Query executada: {len(preenchimentos)} preenchimentos encontrados')
            
        except SQLAlchemyError as db_error:
            app.logger.error(f'❌ Erro de banco de dados: {str(db_error)}', exc_info=True)
            flash('⛔ Erro ao acessar o banco de dados. Tente novamente.', 'danger')
            raise
        
        # 5. Processar resultados com tratamento de erros individuais
        preenchimentos_por_material = {}
        contador_processados = 0
        contador_erros = 0
        
        for p in preenchimentos:
            try:
                # Obter nome do material com fallback
                material_nome = 'N/A'
                if p.solicitacao and p.solicitacao.material:
                    material_nome = p.solicitacao.material.DescricaoMaterial or 'Material sem descrição'
                elif p.solicitacao:
                    material_nome = f'Solicitação #{p.solicitacao.id} (material não encontrado)'
                else:
                    material_nome = 'Solicitação não encontrada'
                    app.logger.warning(f'⚠️ Preenchimento {p.id} sem solicitação associada')
                
                # Obter nome do fornecedor com tratamento de erro
                fornecedor_nome = 'Fornecedor não encontrado'
                try:
                    fornecedor_nome = get_fornecedor_nome(p.fornecedor_id)
                except Exception as fornecedor_error:
                    app.logger.warning(f'⚠️ Erro ao buscar fornecedor {p.fornecedor_id}: {str(fornecedor_error)}')
                
                # Garantir que valores numéricos não sejam None
                valor_unitario = p.valor_unitario if p.valor_unitario is not None else 0.0
                valor_frete = p.valor_frete if p.valor_frete is not None else 0.0
                valor_total = p.valor_total if p.valor_total is not None else 0.0
                
                # Processar histórico de descontos com tratamento seguro
                historico_descontos = []
                try:
                    if hasattr(p, 'historico_descontos'):
                        historico_descontos = [h.to_dict() for h in p.historico_descontos]
                except Exception as historico_error:
                    app.logger.warning(f'⚠️ Erro ao processar histórico de descontos para {p.id}: {str(historico_error)}')
                
                # Criar estrutura do preenchimento
                preenchimento_info = {
                    'id': p.id,
                    'fornecedor_nome': fornecedor_nome,
                    'solicitacao': p.solicitacao,
                    'valor_unitario': valor_unitario,
                    'valor_frete': valor_frete,
                    'valor_total': valor_total,
                    'prazo_entrega': p.prazo_entrega or 'Não informado',
                    'condicao_pagamento': p.condicao_pagamento or 'Não informada',
                    'status': p.status or 'Desconhecido',
                    'usuario': p.usuario or 'Não informado',
                    'pdf_path': p.pdf_path,
                    'historico_descontos': historico_descontos,
                    'observacoes': p.observacoes or ''
                }
                
                # Agrupar por material
                if material_nome not in preenchimentos_por_material:
                    preenchimentos_por_material[material_nome] = []
                
                preenchimentos_por_material[material_nome].append(preenchimento_info)
                contador_processados += 1
                
            except Exception as process_error:
                contador_erros += 1
                app.logger.error(f'❌ Erro ao processar preenchimento {p.id}: {str(process_error)}', exc_info=True)
                continue
        
        # Log de processamento
        if contador_erros > 0:
            app.logger.warning(f'⚠️ {contador_erros} erro(s) durante o processamento de preenchimentos')
        
        # 6. Obter listas para filtros
        try:
            empresas_query = db.session.query(
                SolicitacoesCompra.empresa
            ).distinct().order_by(
                SolicitacoesCompra.empresa
            ).all()
            
            empresas = [e[0] for e in empresas_query if e[0]]
            
            usuarios_query = db.session.query(
                SolicitacoesPreenchidas.usuario
            ).distinct().order_by(
                SolicitacoesPreenchidas.usuario
            ).all()
            
            usuarios = [u[0] for u in usuarios_query if u[0]]
            
        except SQLAlchemyError as filter_error:
            app.logger.error(f'❌ Erro ao buscar filtros: {str(filter_error)}')
            # Continuar com listas vazias
        
        # 7. Mensagem informativa
        if not preenchimentos_por_material:
            flash('📭 Nenhuma solicitação preenchida encontrada com os filtros aplicados.', 'info')
        else:
            total_materiais = len(preenchimentos_por_material)
            total_cotacoes = sum(len(cotacoes) for cotacoes in preenchimentos_por_material.values())
            flash(f'📊 {total_cotacoes} cotações encontradas em {total_materiais} materiais diferentes.', 'success')
        
        # 8. Renderizar template
        return render_template(
            'listar_solicitacoes_preenchidas.html',
            preenchimentos_por_material=preenchimentos_por_material,
            empresas=empresas,
            usuarios=usuarios,
            filtros=filtros
        )
    
    except SQLAlchemyError as db_error:
        # Erro específico do banco de dados
        app.logger.error(f'❌ Erro de banco de dados em listar_solicitacoes_preenchidas: {str(db_error)}', exc_info=True)
        flash('⛔ Erro crítico no banco de dados. Entre em contato com o administrador.', 'danger')
        
        return render_template(
            'listar_solicitacoes_preenchidas.html',
            preenchimentos_por_material={},
            empresas=[],
            usuarios=[],
            filtros=filtros
        )
    
    except ValueError as val_error:
        # Erro de validação
        app.logger.error(f'❌ Erro de validação: {str(val_error)}', exc_info=True)
        flash(f'⚠️ Erro de validação: {str(val_error)}', 'warning')
        
        return render_template(
            'listar_solicitacoes_preenchidas.html',
            preenchimentos_por_material={},
            empresas=[],
            usuarios=[],
            filtros=filtros
        )
    
    except Exception as e:
        # Erro geral não tratado
        app.logger.error(f'❌ Erro inesperado em listar_solicitacoes_preenchidas: {str(e)}', exc_info=True)
        flash('⛔ Ocorreu um erro inesperado. Tente novamente ou entre em contato com o suporte.', 'danger')
        
        return render_template(
            'listar_solicitacoes_preenchidas.html',
            preenchimentos_por_material={},
            empresas=[],
            usuarios=[],
            filtros=filtros
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
        # ==========================
        # POST
        # ==========================
        if request.method == 'POST':

            preenchimento_ids = request.form.getlist('preenchimento_ids')
            forma_pagamento = request.form.get('forma_pagamento', '').strip()
            condicao_pagamento = request.form.get('condicao_pagamento', '').strip()
            observacoes = request.form.get('observacoes', '').strip()

            if not preenchimento_ids:
                flash('Nenhum preenchimento selecionado.', 'error')
                return redirect(url_for('routes_bp.gerar_pedido_compra'))

            preenchimentos = SolicitacoesPreenchidas.query.filter(
                SolicitacoesPreenchidas.id.in_(preenchimento_ids)
            ).all()

            for preenchimento in preenchimentos:
                if preenchimento.status != 'Aprovado':
                    flash(f'O preenchimento ID {preenchimento.id} não está aprovado.', 'error')
                    return redirect(url_for('routes_bp.gerar_pedido_compra'))

            # Número sequencial
            ultimo_pedido = PedidosCompra.query.order_by(PedidosCompra.id.desc()).first()
            proximo_numero = (ultimo_pedido.id + 1) if ultimo_pedido else 1
            numero_pedido = f"PC{datetime.now().year}{proximo_numero:04d}"

            # Totais
            valor_total = sum(p.valor_total for p in preenchimentos)
            valor_frete_total = sum(p.valor_frete if p.valor_frete else 0 for p in preenchimentos)
            valor_liquido = valor_total - valor_frete_total

            # 🔥 LÓGICA FLEXÍVEL (campo não obrigatório)
            forma_condicao_final = None

            if forma_pagamento and condicao_pagamento:
                forma_condicao_final = f"{forma_pagamento} - {condicao_pagamento}"
            elif forma_pagamento:
                forma_condicao_final = forma_pagamento
            elif condicao_pagamento:
                forma_condicao_final = condicao_pagamento

            # Criar pedido
            pedido = PedidosCompra(
                numero_pedido=numero_pedido,
                usuario=session['usuario'],
                status='Gerado',
                valor_total=valor_total,
                valor_frete=valor_frete_total if valor_frete_total != 0 else None,
                valor_liquido=valor_liquido,
                forma_pagamento=forma_condicao_final,
                observacoes=observacoes if observacoes else None,
                data_criacao=datetime.now()
            )

            pedido.preenchimentos = preenchimentos
            db.session.add(pedido)

            for preenchimento in preenchimentos:
                preenchimento.status = 'Em Processamento'

            db.session.commit()

            # ==========================
            # PDF
            # ==========================

            materiais_por_fornecedor = {}
            fornecedor_ids = {p.fornecedor_id for p in preenchimentos}

            fornecedores_info = {}
            if fornecedor_ids:
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

            upload_dir = app.config['UPLOAD_FOLDER']
            if not os.path.exists(upload_dir):
                os.makedirs(upload_dir)

            pdf_filename = f"pedido_{numero_pedido}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            pdf_path = os.path.join(upload_dir, pdf_filename)

            html_content = render_template(
                'pedido_compra_pdf.html',
                pedido=pedido,
                materiais_por_fornecedor=materiais_por_fornecedor,
                data_criacao=datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                usuario=session['usuario'],
                total_itens=len(preenchimento_ids),
                fornecedores_count=len(materiais_por_fornecedor),
                observacoes=observacoes
            )

            wkhtmltopdf_path = shutil.which('wkhtmltopdf')
            if not wkhtmltopdf_path:
                raise FileNotFoundError("wkhtmltopdf não encontrado no sistema.")

            config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

            options = {
                'encoding': 'UTF-8',
                'enable-local-file-access': '',
                'margin-top': '10mm',
                'margin-right': '10mm',
                'margin-bottom': '10mm',
                'margin-left': '10mm'
            }

            pdfkit.from_string(html_content, pdf_path, configuration=config, options=options)

            pedido.pdf_path = pdf_path
            db.session.commit()

            flash(f'Pedido {numero_pedido} gerado com sucesso!', 'success')
            return redirect(url_for('routes_bp.auditoria_solicitacoes'))

        # ==========================
        # GET
        # ==========================

        preenchimentos = SolicitacoesPreenchidas.query.filter_by(status='Aprovado')\
            .join(SolicitacoesCompra)\
            .join(Materiais)\
            .order_by(Materiais.DescricaoMaterial)\
            .all()

        # 🔥 PUXA AUTOMÁTICO PARA O INPUT (editável)
        condicao_pagamento_padrao = ''
        if preenchimentos:
            condicao_pagamento_padrao = preenchimentos[0].condicao_pagamento or ''

        preenchimentos_por_material = {}
        for preenchimento in preenchimentos:
            material_nome = preenchimento.solicitacao.material.DescricaoMaterial
            if material_nome not in preenchimentos_por_material:
                preenchimentos_por_material[material_nome] = []
            preenchimentos_por_material[material_nome].append(preenchimento)

        return render_template(
            'gerar_pedido_compra.html',
            preenchimentos_por_material=preenchimentos_por_material,
            formas_pagamento=['À Vista', 'A Prazo', 'Boleto', 'Cartão de Crédito', 'Transferência Bancária'],
            condicao_pagamento_padrao=condicao_pagamento_padrao
        )

    except Exception as e:
        db.session.rollback()
        logging.error(f"Erro em gerar_pedido_compra: {str(e)}", exc_info=True)
        flash(f'Erro ao gerar pedido: {str(e)}', 'error')
        return redirect(url_for('routes_bp.listar_solicitacoes_preenchidas'))


#testeaqui
@routes_bp.route('/pedido/<int:pedido_id>/view')
def view_pedido_pdf(pedido_id):
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Buscar o pedido no banco de dados
        pedido = PedidosCompra.query.get_or_404(pedido_id)
        
        # Obter informações dos fornecedores e materiais (mesma lógica usada na geração do PDF)
        fornecedor_ids = {p.fornecedor_id for p in pedido.preenchimentos}
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

        # Organizar materiais por fornecedor
        materiais_por_fornecedor = {}
        for preenchimento in pedido.preenchimentos:
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

        # Renderizar o template com os dados
        return render_template(
            'pedido_compra_pdf.html',
            pedido=pedido,
            materiais_por_fornecedor=materiais_por_fornecedor,
            data_criacao=pedido.data_criacao.strftime('%d/%m/%Y %H:%M:%S'),
            usuario=session['usuario'],
            total_itens=len(pedido.preenchimentos),
            fornecedores_count=len(materiais_por_fornecedor),
            observacoes=pedido.observacoes
        )
        
    except Exception as e:
        flash(f'Erro ao carregar pedido: {str(e)}', 'error')
        logging.error(f"Erro em view_pedido_pdf: {str(e)}")
        return redirect(url_for('routes_bp.listar_pedidos_compra'))
def get_fornecedor_details_full(fornecedor_id):
    """Busca todos os detalhes do fornecedor pelo ID no banco fornecedores.db"""
    try:
        conn = sqlite3.connect(DB_PATH_FORNECEDORES)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, nome_fantasia, cnpj, telefone, email, endereco, bairro, cidade, estado, contato, materiais
            FROM fornecedores 
            WHERE id = ?
        ''', (fornecedor_id,))
        result = cursor.fetchone()
        conn.close()
        if result:
            return {
                'id': result[0],
                'nome_fantasia': result[1],
                'cnpj': result[2],
                'telefone': result[3],
                'email': result[4],
                'endereco': result[5],
                'bairro': result[6],
                'cidade': result[7],
                'estado': result[8],
                'contato': result[9],
                'materiais': result[10]
            }
        return None
    except sqlite3.Error as e:
        logging.error(f"Erro ao buscar fornecedor {fornecedor_id}: {str(e)}")
        return None

# Adicione esta nova rota no app.py (próximo às outras rotas do blueprint routes_bp)
@routes_bp.route('/visualizar_pedido/<int:pedido_id>', methods=['GET'])
def visualizar_pedido(pedido_id):
    if 'usuario' not in session:
        flash('Acesso não autorizado', 'error')
        return redirect(url_for('routes_bp.login'))
    
    try:
        pedido = PedidosCompra.query.get_or_404(pedido_id)
        
        # Agrupar materiais por fornecedor (lógica similar à geração de PDF, adaptada)
        materiais_por_fornecedor = {}
        for preench in pedido.preenchimentos:
            forn_id = str(preench.fornecedor_id)  # Usar string como chave para evitar issues
            if forn_id not in materiais_por_fornecedor:
                info = get_fornecedor_details_full(preench.fornecedor_id)
                materiais_por_fornecedor[forn_id] = {'info': info, 'itens': []}
            
            # Detalhes do item
            item = {
                'material': preench.solicitacao.material.DescricaoMaterial if preench.solicitacao.material else 'Não informado',
                'marca': preench.solicitacao.marca or 'Sem marca',
                'especificacao': preench.solicitacao.especificacao or 'N/A',
                'quantidade': preench.solicitacao.quantidade,
                'unidade': preench.solicitacao.unidade_medida or 'Un',
                'valor_unitario': preench.valor_unitario,
                'valor_total': preench.valor_total,
                'prazo_entrega': preench.prazo_entrega or 'N/A'
            }
            materiais_por_fornecedor[forn_id]['itens'].append(item)
        
        # Calcular total de itens
        total_itens = sum(len(dados['itens']) for dados in materiais_por_fornecedor.values())
        
        # Renderizar o template como HTML (sem conversão para PDF)
        return render_template(
            'pedido_compra_pdf.html',
            pedido=pedido,
            usuario=pedido.usuario or session['usuario'],  # Ajuste conforme necessário (pode ser pedido.usuario)
            total_itens=total_itens,
            materiais_por_fornecedor=materiais_por_fornecedor
        )
    
    except Exception as e:
        flash(f'Erro ao visualizar pedido: {str(e)}', 'error')
        logging.error(f"Erro em visualizar_pedido para ID {pedido_id}: {str(e)}", exc_info=True)
        return redirect(url_for('routes_bp.listar_pedidos_compras'))  # Ajuste para a rota de listagem correta
    
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
        # ────────────────────────────────────────────────
        # 1. Captura de parâmetros
        # ────────────────────────────────────────────────
        pagina       = request.args.get('pagina', 1, type=int)
        por_pagina   = request.args.get('por_pagina', 50, type=int)
        status       = request.args.get('status')
        empresa      = request.args.get('empresa')
        comprador_atribuido = request.args.get('comprador_atribuido')
        data_inicio  = request.args.get('data_inicio')
        data_fim     = request.args.get('data_fim')
        numero_pedido = request.args.get('numero_pedido')

        # Validação básica
        if por_pagina not in [50, 100]:
            por_pagina = 50
        if pagina < 1:
            pagina = 1

        logging.info(f"→ Requisição GET /listar_pedidos_compra | página={pagina} | por_pagina={por_pagina}")
        logging.info(f"   Filtros recebidos: status={status}, empresa={empresa}, comprador_atribuido={comprador_atribuido}, data_inicio={data_inicio}, data_fim={data_fim}, numero_pedido={numero_pedido}")

        # ────────────────────────────────────────────────
        # 2. Carregar empresas e compradores do senhas.txt
        # ────────────────────────────────────────────────
        compradores_empresas = {}
        empresas_unicas = set()
        compradores_unicos = set()

        try:
            with open('senhas.txt', 'r', encoding='utf-8') as f:
                for line in f:
                    partes = line.strip().split('%')
                    if len(partes) >= 4:
                        usuario = partes[0].strip()
                        empresa_usuario = partes[3].strip()
                        pagina_usuario = partes[2].strip().lower() if len(partes) > 2 else ''
                        
                        empresas_unicas.add(empresa_usuario)
                        
                        if 'comprador' in pagina_usuario:
                            compradores_unicos.add(usuario)
                            compradores_empresas[usuario] = empresa_usuario
                            logging.debug(f"Comprador encontrado: {usuario} - Empresa: {empresa_usuario}")
                        
        except Exception as e:
            logging.error(f"Erro lendo senhas.txt: {str(e)}")

        logging.info(f"Compradores carregados: {sorted(compradores_unicos)}")
        logging.info(f"Empresas carregadas: {sorted(empresas_unicas)}")

        # ────────────────────────────────────────────────
        # 3. Query base
        # ────────────────────────────────────────────────
        query = db.session.query(PedidosCompra).join(
            pedido_preenchimento_associacao, PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
        ).join(
            SolicitacoesPreenchidas, pedido_preenchimento_associacao.c.preenchimento_id == SolicitacoesPreenchidas.id
        ).join(
            SolicitacoesCompra, SolicitacoesPreenchidas.solicitacao_id == SolicitacoesCompra.id
        )

        # Filtros
        if status:
            query = query.filter(PedidosCompra.status == status)
            
        if numero_pedido:
            query = query.filter(PedidosCompra.numero_pedido.ilike(f'%{numero_pedido}%'))
            
        if empresa:
            query = query.filter(
                or_(
                    SolicitacoesCompra.empresa == empresa,
                    PedidosCompra.usuario.in_([u for u, e in compradores_empresas.items() if e == empresa])
                )
            )
            
        if comprador_atribuido:
            query = query.filter(SolicitacoesCompra.comprador_atribuido == comprador_atribuido)
            
        if data_inicio:
            query = query.filter(PedidosCompra.data_criacao >= data_inicio)
            
        if data_fim:
            try:
                dt_fim = datetime.strptime(data_fim, '%Y-%m-%d') + timedelta(days=1)
                query = query.filter(PedidosCompra.data_criacao < dt_fim)
            except:
                logging.warning("Formato inválido em data_fim")

        # ────────────────────────────────────────────────
        # 4. Contagem e paginação
        # ────────────────────────────────────────────────
        ids_distintos = query.with_entities(PedidosCompra.id)\
            .distinct()\
            .order_by(
                PedidosCompra.data_criacao.desc(),
                PedidosCompra.id.desc()
            ).all()

        ids_distintos = [row[0] for row in ids_distintos]
        total_itens = len(ids_distintos)
        total_paginas = max(1, (total_itens + por_pagina - 1) // por_pagina)

        if pagina > total_paginas:
            pagina = total_paginas

        inicio = (pagina - 1) * por_pagina
        fim = inicio + por_pagina
        ids_page = ids_distintos[inicio:fim]

        pedidos = db.session.query(PedidosCompra)\
            .join(
                pedido_preenchimento_associacao,
                PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
            ).join(
                SolicitacoesPreenchidas,
                pedido_preenchimento_associacao.c.preenchimento_id == SolicitacoesPreenchidas.id
            ).join(
                SolicitacoesCompra,
                SolicitacoesPreenchidas.solicitacao_id == SolicitacoesCompra.id
            ).filter(
                PedidosCompra.id.in_(ids_page)
            ).distinct().order_by(
                PedidosCompra.data_criacao.desc(),
                PedidosCompra.id.desc()
            ).all()

        # ────────────────────────────────────────────────
        # 5. Buscar informações de fornecedores
        # ────────────────────────────────────────────────
        fornecedor_ids = set()
        for pedido in pedidos:
            for preenchimento in pedido.preenchimentos:
                if preenchimento.fornecedor_id:
                    fornecedor_ids.add(preenchimento.fornecedor_id)

        fornecedores = {}
        if fornecedor_ids:
            conn = get_db_connection(DB_PATH_FORNECEDORES)
            if conn:
                try:
                    cursor = conn.cursor()
                    placeholders = ','.join(['?' for _ in fornecedor_ids])
                    cursor.execute(f'SELECT id, nome_fantasia, cnpj FROM fornecedores WHERE id IN ({placeholders})', 
                                 list(fornecedor_ids))
                    for row in cursor.fetchall():
                        fornecedores[row[0]] = {
                            'nome_fantasia': row[1],
                            'cnpj': format_cnpj(row[2]) if row[2] else 'N/A'
                        }
                except Exception as e:
                    logging.error(f"Erro ao buscar fornecedores: {str(e)}")
                finally:
                    conn.close()

        # ────────────────────────────────────────────────
        # 6. Estruturar dados para o template - ADICIONAR comprador_atribuido AQUI
        # ────────────────────────────────────────────────
        pedidos_completos = []
        for pedido in pedidos:
            preenchimentos_info = []
            for preenchimento in pedido.preenchimentos:
                fornecedor_info = fornecedores.get(preenchimento.fornecedor_id, {
                    'nome_fantasia': 'Fornecedor não encontrado',
                    'cnpj': 'N/A'
                })
                
                empresa_usuario = compradores_empresas.get(pedido.usuario, '')
                if not empresa_usuario and preenchimento.solicitacao:
                    empresa_usuario = preenchimento.solicitacao.empresa
                
                marca = preenchimento.solicitacao.marca if preenchimento.solicitacao and preenchimento.solicitacao.marca else 'Não informado'
                
                # 🔥 CORREÇÃO: Adicionar comprador_atribuido ao preenchimentos_info
                preenchimentos_info.append({
                    'id': preenchimento.id,
                    'marca': marca,
                    'fornecedor_nome': fornecedor_info['nome_fantasia'],
                    'fornecedor_cnpj': fornecedor_info.get('cnpj', 'N/A'),
                    'fornecedor_id': preenchimento.fornecedor_id,
                    'material': preenchimento.solicitacao.material.DescricaoMaterial if preenchimento.solicitacao and preenchimento.solicitacao.material else 'N/A',
                    'empresa': empresa_usuario,
                    'comprador_atribuido': preenchimento.solicitacao.comprador_atribuido if preenchimento.solicitacao else None  # NOVO CAMPO
                })
            pedidos_completos.append({
                'pedido': pedido,
                'preenchimentos': preenchimentos_info,
                'observacoes': pedido.observacoes
            })

        # ────────────────────────────────────────────────
        # 7. Renderizar template
        # ────────────────────────────────────────────────
        return render_template(
            'listar_pedidos_compra.html', 
            pedidos_completos=pedidos_completos,
            empresas=sorted(empresas_unicas),
            compradores=sorted(compradores_unicos),
            filtros={
                'empresa': empresa,
                'comprador_atribuido': comprador_atribuido,
                'data_inicio': data_inicio,
                'data_fim': data_fim,
                'status': status,
                'numero_pedido': numero_pedido
            },
            pagina_atual=pagina,
            total_paginas=total_paginas,
            total_itens=total_itens,
            por_pagina=por_pagina,
            request=request
        )
        
    except Exception as e:
        logging.error(f"Erro ao listar pedidos de compra: {str(e)}", exc_info=True)
        flash(f'Erro ao listar pedidos de compra: {str(e)}', 'error')
        return render_template(
            'listar_pedidos_compra.html', 
            pedidos_completos=[],
            empresas=[],
            compradores=[],
            filtros={},
            pagina_atual=1,
            total_paginas=1,
            total_itens=0,
            por_pagina=50,
            request=request
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
                
                # CORREÇÃO: A marca está na solicitação, não no preenchimento
                marca = preenchimento.solicitacao.marca if preenchimento.solicitacao.marca else 'Não informado'
                
                preenchimentos_info.append({
                    'id': preenchimento.id,
                    'marca': marca,  # Usando a marca da solicitação
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
            SELECT id, nome_fantasia, cnpj, telefone, email, endereco, 
                   bairro, cidade, estado, contato, materiais
            FROM fornecedores
            WHERE nome_fantasia LIKE ? OR cnpj LIKE ?
            ORDER BY nome_fantasia
        ''', (f'%{query}%', f'%{query}%'))
        
        fornecedores = []
        for row in cursor.fetchall():
            fornecedor = {
                'id': row['id'],
                'nome_fantasia': row['nome_fantasia'],
                'cnpj': row['cnpj'],
                'telefone': row['telefone'],
                'email': row['email'],
                'endereco': row['endereco'],
                'bairro': row['bairro'],
                'cidade': row['cidade'],
                'estado': row['estado'],
                'contato': row['contato'],
                'materiais': row['materiais']
            }
            fornecedores.append(fornecedor)
        
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
        # Parâmetros de filtro
        empresa = request.args.get('empresa')
        usuario = request.args.get('usuario')
        ativo = request.args.get('ativo')
        nome_ativo = request.args.get('nome_ativo')
        status = request.args.get('status')
        prioridade_filtro = request.args.get('prioridade')
        data_inicio = request.args.get('data_inicio')
        data_fim = request.args.get('data_fim')
        
        # OBTER USUÁRIO LOGADO
        usuario_logado = session.get('usuario')
        
        # Se não foi especificado um usuário no filtro, usar o usuário logado por padrão
        if not usuario and usuario_logado:
            usuario = usuario_logado

        # Query base para buscar TODAS as solicitações
        query = db.session.query(SolicitacoesCompra)
        
        # Aplicar filtros básicos
        if empresa and empresa != 'Todas':
            query = query.filter(SolicitacoesCompra.empresa == empresa)
        if usuario and usuario != 'Todos':
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

        # Executar query - buscar todas as solicitações
        solicitacoes = query.order_by(SolicitacoesCompra.data_solicitacao.desc()).all()

        # Preparar dados para o template
        auditoria_data = []
        for solicitacao in solicitacoes:
            material = db.session.get(Materiais, solicitacao.cod_material)
            
            # Buscar TODOS os preenchimentos (excluindo rascunhos)
            preenchimentos = db.session.query(SolicitacoesPreenchidas).filter_by(
                solicitacao_id=solicitacao.id
            ).filter(
                SolicitacoesPreenchidas.status != 'Rascunho'
            ).all()
            
            # CASO 1: Solicitação SEM preenchimentos
            if not preenchimentos:
                # VERIFICAR PRIMEIRO O STATUS DE APROVAÇÃO
                if solicitacao.status_aprovacao:
                    # Se tem status_aprovacao, usar esse status
                    if solicitacao.status_aprovacao == 'Reprovado':
                        status_solicitacao = 'Reprovado'
                        prioridade = 0  # Prioridade baixa para reprovados
                    elif solicitacao.status_aprovacao == 'Aprovado':
                        status_solicitacao = 'Aprovado'
                        prioridade = 1  # Prioridade média para aprovados sem preenchimento
                    else:
                        status_solicitacao = solicitacao.status_aprovacao
                        prioridade = 0
                else:
                    # Se não tem status_aprovacao, verificar outros campos
                    solicitacao_status = getattr(solicitacao, 'status', None)
                    
                    # Determinar status baseado no campo da solicitação
                    if solicitacao_status == 'Pendente':
                        status_solicitacao = 'Pendente'
                    else:
                        status_solicitacao = 'Aberta'  # Status padrão
                    
                    prioridade = 0  # Prioridade baixa para solicitações sem preenchimento
                
                # Aplicar filtro de status se especificado
                if status and status != 'Todos':
                    if status != status_solicitacao:
                        continue  # Pular se status não corresponde
                
                # Aplicar filtro de prioridade se especificado
                if prioridade_filtro and prioridade_filtro != '':
                    if int(prioridade_filtro) != prioridade:
                        continue
                
                auditoria_data.append({
                    'solicitacao': solicitacao,
                    'material': material,
                    'preenchimento': None,  # Sem preenchimento
                    'pedido': None,
                    'estoque': None,
                    'requisicoes': [],
                    'fornecedor': {},
                    'prioridade': prioridade,
                    'status_determinado': status_solicitacao  # Adicionar status determinado
                })
            
            # CASO 2: Solicitação COM preenchimentos
            else:
                # Para cada preenchimento, criar um registro na auditoria
                for preenchimento in preenchimentos:
                    # Buscar pedido associado (se existir)
                    pedido = None
                    pedido = db.session.query(PedidosCompra).join(
                        pedido_preenchimento_associacao,
                        PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
                    ).filter(
                        pedido_preenchimento_associacao.c.preenchimento_id == preenchimento.id
                    ).first()
                    
                    # Buscar estoque (se existir)
                    estoque = None
                    estoque = db.session.query(Estoque).filter_by(
                        preenchimento_id=preenchimento.id
                    ).first()
                    
                    # Buscar requisições (se existirem)
                    requisicoes = []
                    requisicoes = db.session.query(Requisicoes).filter_by(
                        preenchimento_id=preenchimento.id
                    ).all()
                    
                    # Aplicar filtro de status se especificado
                    if status and status != 'Todos':
                        # Para solicitações com preenchimentos, verificar status do preenchimento
                        # Mas também considerar o status_aprovacao da solicitação
                        status_match = False
                        
                        # Primeiro verificar status do preenchimento
                        if preenchimento.status == status:
                            status_match = True
                        # Verificar se a solicitação foi reprovada
                        elif solicitacao.status_aprovacao == 'Reprovado' and status == 'Reprovado':
                            status_match = True
                        # Verificar se a solicitação foi aprovada
                        elif solicitacao.status_aprovacao == 'Aprovado' and status == 'Aprovado':
                            status_match = True
                        
                        if not status_match:
                            continue  # Pular se status não corresponde
                    
                    # Calcular prioridade para ordenação - CORREÇÃO IMPORTANTE
                    prioridade = 0
                    
                    # Primeiro verificar se a solicitação foi reprovada
                    if solicitacao.status_aprovacao == 'Reprovado':
                        prioridade = 0  # Prioridade mais baixa para reprovados
                    # Depois verificar status do preenchimento
                    elif preenchimento.status == 'Entregue':
                        prioridade = 3  # Máxima prioridade
                    elif preenchimento.status == 'Aprovado' and pedido:
                        prioridade = 2  # Alta prioridade
                    elif preenchimento.status == 'Aprovado':
                        prioridade = 1  # Média prioridade
                    # Status "Pendente" no preenchimento (se existir)
                    elif preenchimento.status == 'Pendente':
                        prioridade = 0
                    
                    # Aplicar filtro de prioridade se especificado
                    if prioridade_filtro and prioridade_filtro != '':
                        if int(prioridade_filtro) != prioridade:
                            continue
                    
                    # Buscar informações do fornecedor
                    fornecedor_info = {}
                    if preenchimento.fornecedor_id:
                        conn = get_db_connection(DB_PATH_FORNECEDORES)
                        if conn:
                            try:
                                cursor = conn.cursor()
                                cursor.execute('SELECT nome_fantasia, cnpj FROM fornecedores WHERE id = ?', 
                                             (preenchimento.fornecedor_id,))
                                result = cursor.fetchone()
                                if result:
                                    fornecedor_info = {
                                        'nome_fantasia': result[0],
                                        'cnpj': format_cnpj(result[1]) if result[1] else 'N/A'
                                    }
                            finally:
                                conn.close()
                    
                    auditoria_data.append({
                        'solicitacao': solicitacao,
                        'material': material,
                        'preenchimento': preenchimento,
                        'pedido': pedido,
                        'estoque': estoque,
                        'requisicoes': requisicoes,
                        'fornecedor': fornecedor_info,
                        'prioridade': prioridade,
                        'status_determinado': solicitacao.status_aprovacao if solicitacao.status_aprovacao else preenchimento.status
                    })

        # Ordenar por prioridade (mais alta primeiro) e depois por data
        auditoria_data.sort(key=lambda x: (-x['prioridade'], x['solicitacao'].data_solicitacao), reverse=True)

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

        total_registros = len(auditoria_data)

        # Passar o usuário logado para o template
        usuario_logado = session.get('usuario')

        return render_template(
            'auditoria_solicitacoes.html',
            auditoria=auditoria_data,
            empresas=[e[0] for e in empresas if e[0]],
            usuarios=[u[0] for u in usuarios if u[0]],
            nomes_ativos=[n[0] for n in nomes_ativos if n[0]],
            total_registros=total_registros,
            usuario_logado=usuario_logado,
            filtros={
                'empresa': empresa,
                'usuario': usuario,
                'ativo': ativo,
                'nome_ativo': nome_ativo,
                'status': status,
                'prioridade': prioridade_filtro,
                'data_inicio': data_inicio,
                'data_fim': data_fim
            }
        )
    except Exception as e:
        flash(f'Erro ao carregar página de auditoria: {str(e)}', 'error')
        app.logger.error(f'Erro em auditoria_solicitacoes: {str(e)}')
        return render_template(
            'auditoria_solicitacoes.html',
            auditoria=[],
            empresas=[],
            usuarios=[],
            nomes_ativos=[],
            total_registros=0,
            usuario_logado=session.get('usuario'),
            filtros={}
        )

@routes_bp.route('/exportar_auditoria_xlsx', methods=['GET'])
def exportar_auditoria_xlsx():
    if 'usuario' not in session:
        return redirect(url_for('routes_bp.login'))
    
    try:
        # ────────────────────────────────────────────────
        # Parâmetros vindos da URL (mesmos da tela de auditoria)
        # ────────────────────────────────────────────────
        empresa         = request.args.get('empresa')
        usuario         = request.args.get('usuario')
        ativo           = request.args.get('ativo')
        nome_ativo      = request.args.get('nome_ativo')
        status          = request.args.get('status')
        prioridade_filtro = request.args.get('prioridade')
        data_inicio     = request.args.get('data_inicio')
        data_fim        = request.args.get('data_fim')

        # Log para debug
        app.logger.info(f"=== INÍCIO EXPORTAÇÃO ===")
        app.logger.info(f"Parâmetro usuario recebido: '{usuario}'")
        app.logger.info(f"Tipo do parâmetro: {type(usuario)}")

        # CORREÇÃO: Verificar corretamente se é para exportar TODOS os usuários
        # Se usuario for vazio, None, string vazia ou 'Todos' (case insensitive)
        exportar_todos = (usuario is None or 
                         usuario == '' or 
                         usuario.lower() == 'todos' or 
                         usuario.lower() == 'todos os usuários')
        
        app.logger.info(f"exportar_todos: {exportar_todos}")

        # ────────────────────────────────────────────────
        # Query base
        # ────────────────────────────────────────────────
        query = db.session.query(SolicitacoesCompra).outerjoin(
            SolicitacoesPreenchidas,
            SolicitacoesCompra.id == SolicitacoesPreenchidas.solicitacao_id
        )

        # Filtros
        if empresa and empresa != '' and empresa != 'Todas':
            query = query.filter(SolicitacoesCompra.empresa == empresa)
            app.logger.info(f"Aplicando filtro empresa: {empresa}")

        # CORREÇÃO CRÍTICA: Aplicar filtro de usuário APENAS se NÃO for exportar todos
        if not exportar_todos:
            # Se não for exportar todos, filtra pelo usuário específico
            if usuario and usuario.strip():
                query = query.filter(SolicitacoesCompra.usuario == usuario)
                app.logger.info(f"Aplicando filtro para usuário específico: '{usuario}'")
            else:
                app.logger.warning(f"exportar_todos é False mas usuario está vazio: '{usuario}'")
        else:
            app.logger.info("Exportando TODOS os usuários (filtro de usuário ignorado)")

        if ativo and ativo != '':
            query = query.filter(SolicitacoesCompra.ativo == ativo)
            app.logger.info(f"Aplicando filtro ativo: {ativo}")

        if nome_ativo and nome_ativo.strip() and ativo == 'Sim':
            query = query.filter(SolicitacoesCompra.nome_ativo.ilike(f'%{nome_ativo}%'))
            app.logger.info(f"Aplicando filtro nome_ativo: {nome_ativo}")

        if data_inicio and data_inicio.strip():
            query = query.filter(SolicitacoesCompra.data_solicitacao >= data_inicio)
            app.logger.info(f"Aplicando filtro data_inicio: {data_inicio}")

        if data_fim and data_fim.strip():
            try:
                dt_fim = datetime.strptime(data_fim, '%Y-%m-%d')
                data_fim_ajustada = dt_fim + timedelta(days=1)
                query = query.filter(SolicitacoesCompra.data_solicitacao < data_fim_ajustada)
                app.logger.info(f"Aplicando filtro data_fim: {data_fim} (ajustada: {data_fim_ajustada})")
            except Exception as e:
                app.logger.error(f"Erro ao processar data_fim: {e}")

        # Busca os registros
        solicitacoes = query.order_by(SolicitacoesCompra.data_solicitacao.desc()).distinct().all()
        
        app.logger.info(f"Total de solicitações encontradas: {len(solicitacoes)}")
        
        # Se não encontrou nada, log mais detalhado
        if len(solicitacoes) == 0:
            app.logger.warning("NENHUMA SOLICITAÇÃO ENCONTRADA COM OS FILTROS:")
            app.logger.warning(f"  - empresa: {empresa}")
            app.logger.warning(f"  - usuario: {usuario} (exportar_todos: {exportar_todos})")
            app.logger.warning(f"  - ativo: {ativo}")
            app.logger.warning(f"  - nome_ativo: {nome_ativo}")
            app.logger.warning(f"  - status: {status}")
            app.logger.warning(f"  - prioridade_filtro: {prioridade_filtro}")
            app.logger.warning(f"  - data_inicio: {data_inicio}")
            app.logger.warning(f"  - data_fim: {data_fim}")

        # ────────────────────────────────────────────────
        # Preparar dados para o Excel
        # ────────────────────────────────────────────────
        dados = []
        contador_solicitacoes = 0
        contador_preenchimentos = 0

        for solicitacao in solicitacoes:
            contador_solicitacoes += 1
            material = db.session.get(Materiais, solicitacao.cod_material)

            # Buscar preenchimentos (excluindo rascunhos)
            preenchimentos_query = db.session.query(SolicitacoesPreenchidas).filter_by(
                solicitacao_id=solicitacao.id
            ).filter(
                SolicitacoesPreenchidas.status != 'Rascunho'
            )
            
            preenchimentos = preenchimentos_query.all()
            
            app.logger.debug(f"Solicitação {solicitacao.id} - {len(preenchimentos)} preenchimentos encontrados")

            # Caso não tenha preenchimento → trata como 'Aberta'
            if not preenchimentos:
                if status and status.strip() and status != 'Todos' and status != 'Aberta':
                    continue
                
                prioridade = 0
                nivel_prioridade = 'Baixa (Outras)'
                status_formatado = 'Aberta'
                fornecedor_nome = ''
                fornecedor_cnpj = ''
                pedido_numero = ''
                pedido_status = ''
                valor_unitario = 0
                valor_total = 0
                valor_frete = 0
                prazo_entrega = ''
                condicao_pagamento = ''
                observacoes = ''
                
                if prioridade_filtro and prioridade_filtro.strip():
                    if int(prioridade_filtro) != prioridade:
                        continue
                
                dados.append({
                    'ID_Solicitacao': solicitacao.id,
                    'ID_Cotacao': '',
                    'Material': material.DescricaoMaterial if material else 'Material não encontrado',
                    'Cod_Material': material.CodMaterial if material else '',
                    'Especificacao': solicitacao.especificacao,
                    'Marca': solicitacao.marca or '',
                    'Solicitante': solicitacao.usuario,
                    'Comprador_Atribuido': solicitacao.comprador_atribuido or '',
                    'Empresa': solicitacao.empresa,
                    'Ativo': 'Sim' if solicitacao.ativo == 'Sim' else 'Não',
                    'Nome_Ativo': solicitacao.nome_ativo if solicitacao.ativo == 'Sim' else '',
                    'Quantidade': solicitacao.quantidade,
                    'Unidade_Medida': solicitacao.unidade_medida,
                    'Prioridade_Original': solicitacao.prioridade,
                    'Nivel_Prioridade': nivel_prioridade,
                    'Valor_Prioridade': prioridade,
                    'Data_Solicitacao': solicitacao.data_solicitacao.strftime('%d/%m/%Y %H:%M') if solicitacao.data_solicitacao else '',
                    'Status': status_formatado,
                    'Fornecedor': fornecedor_nome,
                    'CNPJ_Fornecedor': fornecedor_cnpj,
                    'Valor_Unitario': valor_unitario,
                    'Valor_Total': valor_total,
                    'Valor_Frete': valor_frete,
                    'Prazo_Entrega': prazo_entrega,
                    'Condicao_Pagamento': condicao_pagamento,
                    'Pedido_Numero': pedido_numero,
                    'Pedido_Status': pedido_status,
                    'Observacoes': observacoes,
                    'Aprovacao': solicitacao.status_aprovacao or 'Pendente',
                    'Aplicacao': solicitacao.aplicacao or '',
                    'Aplicacao_Geral': solicitacao.aplicacao_geral or ''
                })
                continue
            
            # Tem preenchimentos → processa cada um
            for preenchimento in preenchimentos:
                contador_preenchimentos += 1
                
                if status and status.strip() and status != 'Todos':
                    if preenchimento.status != status:
                        continue
                
                # Lógica de prioridade (igual ao original)
                prioridade = 0
                if preenchimento.status == 'Entregue':
                    prioridade = 3
                elif preenchimento.status == 'Aprovado':
                    pedido = db.session.query(PedidosCompra).join(
                        pedido_preenchimento_associacao,
                        PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
                    ).filter(
                        pedido_preenchimento_associacao.c.preenchimento_id == preenchimento.id
                    ).first()
                    prioridade = 2 if pedido else 1
                
                if prioridade_filtro and prioridade_filtro.strip():
                    if int(prioridade_filtro) != prioridade:
                        continue
                
                # Busca fornecedor
                fornecedor_nome = ''
                fornecedor_cnpj = ''
                if preenchimento.fornecedor_id:
                    conn = get_db_connection(DB_PATH_FORNECEDORES)
                    if conn:
                        cursor = conn.cursor()
                        cursor.execute('SELECT nome_fantasia, cnpj FROM fornecedores WHERE id = ?', (preenchimento.fornecedor_id,))
                        result = cursor.fetchone()
                        if result:
                            fornecedor_nome = result[0] or ''
                            fornecedor_cnpj = format_cnpj(result[1]) if result[1] else ''
                        conn.close()
                
                # Busca pedido associado
                pedido = None
                pedido_numero = ''
                pedido_status = ''
                if preenchimento:
                    pedido = db.session.query(PedidosCompra).join(
                        pedido_preenchimento_associacao,
                        PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
                    ).filter(
                        pedido_preenchimento_associacao.c.preenchimento_id == preenchimento.id
                    ).first()
                    if pedido:
                        pedido_numero = pedido.numero_pedido
                        pedido_status = pedido.status or ''
                
                # Formata status
                status_formatado = preenchimento.status or 'Aberta'
                if status_formatado == 'Rascunho':
                    status_formatado = 'Rascunho'
                elif status_formatado == 'Aguardando Aprovação':
                    status_formatado = 'Aguardando Aprovação'
                elif status_formatado == 'Aprovado':
                    status_formatado = 'Aprovado'
                elif status_formatado == 'Reprovado':
                    status_formatado = 'Reprovado'
                elif status_formatado == 'Em Processamento':
                    status_formatado = 'Em Processamento'
                elif status_formatado == 'Entregue':
                    status_formatado = 'Entregue'
                
                nivel_prioridade = {
                    3: 'Máxima (Entregues)',
                    2: 'Alta (Aprovadas + Pedido)',
                    1: 'Média (Aprovadas)',
                    0: 'Baixa (Outras)'
                }.get(prioridade, 'Baixa (Outras)')
                
                dados.append({
                    'ID_Solicitacao': solicitacao.id,
                    'ID_Cotacao': preenchimento.id,
                    'Material': material.DescricaoMaterial if material else 'Material não encontrado',
                    'Cod_Material': material.CodMaterial if material else '',
                    'Especificacao': solicitacao.especificacao,
                    'Marca': solicitacao.marca or '',
                    'Solicitante': solicitacao.usuario,
                    'Comprador_Atribuido': solicitacao.comprador_atribuido or '',
                    'Empresa': solicitacao.empresa,
                    'Ativo': 'Sim' if solicitacao.ativo == 'Sim' else 'Não',
                    'Nome_Ativo': solicitacao.nome_ativo if solicitacao.ativo == 'Sim' else '',
                    'Quantidade': solicitacao.quantidade,
                    'Unidade_Medida': solicitacao.unidade_medida,
                    'Prioridade_Original': solicitacao.prioridade,
                    'Nivel_Prioridade': nivel_prioridade,
                    'Valor_Prioridade': prioridade,
                    'Data_Solicitacao': solicitacao.data_solicitacao.strftime('%d/%m/%Y %H:%M') if solicitacao.data_solicitacao else '',
                    'Status': status_formatado,
                    'Fornecedor': fornecedor_nome,
                    'CNPJ_Fornecedor': fornecedor_cnpj,
                    'Valor_Unitario': preenchimento.valor_unitario or 0,
                    'Valor_Total': preenchimento.valor_total or 0,
                    'Valor_Frete': preenchimento.valor_frete or 0,
                    'Prazo_Entrega': preenchimento.prazo_entrega or '',
                    'Condicao_Pagamento': preenchimento.condicao_pagamento or '',
                    'Pedido_Numero': pedido_numero,
                    'Pedido_Status': pedido_status,
                    'Observacoes': preenchimento.observacoes or '',
                    'Aprovacao': solicitacao.status_aprovacao or 'Pendente',
                    'Aplicacao': solicitacao.aplicacao or '',
                    'Aplicacao_Geral': solicitacao.aplicacao_geral or ''
                })

        app.logger.info(f"Total de solicitações processadas: {contador_solicitacoes}")
        app.logger.info(f"Total de preenchimentos processados: {contador_preenchimentos}")
        app.logger.info(f"Total de linhas geradas para Excel: {len(dados)}")

        if not dados:
            app.logger.warning("Nenhum dado encontrado para exportar!")
            flash('Nenhum dado encontrado para exportar com os filtros aplicados.', 'warning')
            return redirect(url_for('routes_bp.auditoria_solicitacoes'))

        # ────────────────────────────────────────────────
        # Geração do Excel
        # ────────────────────────────────────────────────
        wb = Workbook()
        ws = wb.active
        ws.title = "Auditoria Solicitações"

        # Estilos
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        headers = [
            'ID Solicitação', 'ID Cotação', 'Material', 'Cód Material', 'Especificação', 'Marca',
            'Solicitante', 'Comprador Atribuído', 'Empresa', 'Ativo', 'Nome Ativo', 'Quantidade',
            'Unidade', 'Prioridade Original', 'Nível Prioridade', 'Valor Prioridade',
            'Data Solicitação', 'Status', 'Fornecedor', 'CNPJ Fornecedor', 'Valor Unitário',
            'Valor Total', 'Valor Frete', 'Prazo Entrega', 'Condição Pagamento',
            'Número Pedido', 'Status Pedido', 'Observações', 'Aprovação',
            'Aplicação', 'Aplicação Geral'
        ]

        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        # Dados
        for row_num, item in enumerate(dados, 2):
            row = [
                item['ID_Solicitacao'],
                item['ID_Cotacao'],
                item['Material'],
                item['Cod_Material'],
                item['Especificacao'],
                item['Marca'],
                item['Solicitante'],
                item['Comprador_Atribuido'],
                item['Empresa'],
                item['Ativo'],
                item['Nome_Ativo'],
                item['Quantidade'],
                item['Unidade_Medida'],
                item['Prioridade_Original'],
                item['Nivel_Prioridade'],
                item['Valor_Prioridade'],
                item['Data_Solicitacao'],
                item['Status'],
                item['Fornecedor'],
                item['CNPJ_Fornecedor'],
                item['Valor_Unitario'],
                item['Valor_Total'],
                item['Valor_Frete'],
                item['Prazo_Entrega'],
                item['Condicao_Pagamento'],
                item['Pedido_Numero'],
                item['Pedido_Status'],
                item['Observacoes'],
                item['Aprovacao'],
                item['Aplicacao'],
                item['Aplicacao_Geral']
            ]
            for col_num, value in enumerate(row, 1):
                ws.cell(row=row_num, column=col_num, value=value)

        # Ajuste automático de largura
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 3, 40)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Salvar em memória
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        # Nome do arquivo indicando se é completo ou filtrado
        if exportar_todos:
            filename = f"auditoria_solicitacoes_TODOS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        else:
            # Sanitizar nome do usuário para evitar problemas com caracteres especiais
            usuario_sanitizado = usuario.replace(' ', '_').replace('\\', '_').replace('/', '_')
            filename = f"auditoria_solicitacoes_{usuario_sanitizado}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        response = make_response(output.getvalue())
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        response.headers['Content-Disposition'] = f'attachment; filename={filename}'

        app.logger.info(f"Arquivo gerado com sucesso: {filename}")
        return response

    except Exception as e:
        flash(f'Erro ao exportar relatório: {str(e)}', 'error')
        app.logger.error(f'Erro em exportar_auditoria_xlsx: {str(e)}', exc_info=True)
        return redirect(url_for('routes_bp.auditoria_solicitacoes'))
    
def get_fornecedor_cnpj(fornecedor_id):
    conn = get_db_connection(DB_PATH_FORNECEDORES)
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute('SELECT cnpj FROM fornecedores WHERE id = ?', (fornecedor_id,))
            result = cursor.fetchone()
            return format_cnpj(result['cnpj']) if result and result['cnpj'] else 'N/A'
        finally:
            conn.close()
    return 'N/A'

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
        return True
    except sqlite3.Error as e:
        print(f"Erro ao criar a tabela Auditoria: {str(e)}")
        return False

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
        # Obter parâmetros de filtro
        empresa = request.args.get('empresa', 'Todas')
        periodo = request.args.get('periodo', 'Últimos 30 dias')
        status_filtro = request.args.get('status', 'Todos')
        comprador_filtro = request.args.get('comprador', 'Todos')
        
        # Calcular datas com base no período
        data_atual = get_local_time()
        data_inicio = None
        
        if periodo == 'Últimos 30 dias':
            data_inicio = data_atual - timedelta(days=30)
        elif periodo == 'Últimos 6 meses':
            data_inicio = data_atual - timedelta(days=180)
        elif periodo == '2025':
            data_inicio = datetime(2025, 1, 1)
        elif periodo == '2024':
            data_inicio = datetime(2024, 1, 1)
        
        # CONSULTA BASE OTIMIZADA
        query = SolicitacoesCompra.query
        
        # Aplicar filtros
        if data_inicio:
            query = query.filter(SolicitacoesCompra.data_solicitacao >= data_inicio)
        
        if empresa and empresa != 'Todas':
            query = query.filter(SolicitacoesCompra.empresa == empresa)
        
        if status_filtro and status_filtro != 'Todos':
            if status_filtro == 'Pendente':
                query = query.filter(SolicitacoesCompra.status_aprovacao.in_([None, '', 'Pendente']))
            elif status_filtro == 'Aprovado':
                query = query.filter(SolicitacoesCompra.status_aprovacao == 'Aprovado')
            elif status_filtro == 'Reprovado':
                query = query.filter(SolicitacoesCompra.status_aprovacao == 'Reprovado')
        
        if comprador_filtro and comprador_filtro != 'Todos':
            query = query.filter(SolicitacoesCompra.comprador_atribuido == comprador_filtro)
        
        # Executar consulta
        solicitacoes = query.all()
        total_solicitacoes = len(solicitacoes)
        
        # DADOS BÁSICOS
        aprovadas = sum(1 for s in solicitacoes if s.status_aprovacao == 'Aprovado')
        pendentes = sum(1 for s in solicitacoes if s.status_aprovacao in [None, '', 'Pendente'])
        rejeitadas = sum(1 for s in solicitacoes if s.status_aprovacao == 'Reprovado')
        
        # 1. PRODUTIVIDADE DOS COMPRADORES
        compradores_metrics = {}
        
        for solicitacao in solicitacoes:
            comprador = solicitacao.comprador_atribuido or 'Não Atribuído'
            
            if comprador not in compradores_metrics:
                compradores_metrics[comprador] = {
                    'pendentes': 0,
                    'aprovadas': 0,
                    'rejeitadas': 0,
                    'valor_total': 0,
                    'tempo_medio': 0,
                    'contador_tempo': 0
                }
            
            status = solicitacao.status_aprovacao
            if status == 'Aprovado':
                compradores_metrics[comprador]['aprovadas'] += 1
            elif status in [None, '', 'Pendente']:
                compradores_metrics[comprador]['pendentes'] += 1
            elif status == 'Reprovado':
                compradores_metrics[comprador]['rejeitadas'] += 1
            
            # Buscar preenchimentos para calcular valor
            preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                solicitacao_id=solicitacao.id
            ).filter(
                SolicitacoesPreenchidas.status != 'Rascunho'
            ).all()
            
            valor_solicitacao = sum(p.valor_total for p in preenchimentos if p and p.valor_total)
            compradores_metrics[comprador]['valor_total'] += valor_solicitacao
            
            # Calcular tempo médio para solicitações aprovadas
            if status == 'Aprovado' and solicitacao.data_solicitacao:
                dias = (data_atual - solicitacao.data_solicitacao).days
                compradores_metrics[comprador]['tempo_medio'] += dias
                compradores_metrics[comprador]['contador_tempo'] += 1
        
        # Converter para lista
        compradores_lista = []
        for nome, dados in compradores_metrics.items():
            tempo_medio = dados['tempo_medio'] / dados['contador_tempo'] if dados['contador_tempo'] > 0 else 0
            total = dados['pendentes'] + dados['aprovadas'] + dados['rejeitadas']
            taxa_aprovacao = (dados['aprovadas'] / total * 100) if total > 0 else 0
            
            compradores_lista.append({
                'nome': nome,
                'pendentes': dados['pendentes'],
                'aprovadas': dados['aprovadas'],
                'rejeitadas': dados['rejeitadas'],
                'valor_total': dados['valor_total'],
                'tempo_medio_dias': round(tempo_medio, 1),
                'taxa_aprovacao': round(taxa_aprovacao, 1),
                'total': total
            })
        
        compradores_lista.sort(key=lambda x: x['total'], reverse=True)
        
        # 2. PRODUTOS PENDENTES
        produtos_pendentes = {}
        
        for solicitacao in solicitacoes:
            if solicitacao.status_aprovacao in [None, '', 'Pendente']:
                material = db.session.get(Materiais, solicitacao.cod_material)
                if material:
                    # Truncar nome do produto se muito longo
                    produto_nome = material.DescricaoMaterial
                    if len(produto_nome) > 40:
                        produto_nome = produto_nome[:40] + '...'
                    
                    chave = f"{material.CodMaterial}|{solicitacao.empresa}"
                    
                    if chave not in produtos_pendentes:
                        produtos_pendentes[chave] = {
                            'produto': produto_nome,
                            'quantidade': 0,
                            'valor_total': 0,
                            'empresa': solicitacao.empresa,
                            'dias_pendente': (data_atual - solicitacao.data_solicitacao).days if solicitacao.data_solicitacao else 0
                        }
                    
                    produtos_pendentes[chave]['quantidade'] += solicitacao.quantidade
                    
                    # Calcular valor estimado
                    preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                        solicitacao_id=solicitacao.id
                    ).filter(
                        SolicitacoesPreenchidas.status != 'Rascunho'
                    ).all()
                    
                    if preenchimentos:
                        valor_medio = sum(p.valor_unitario for p in preenchimentos if p and p.valor_unitario) / len(preenchimentos)
                        produtos_pendentes[chave]['valor_total'] += valor_medio * solicitacao.quantidade
        
        produtos_pendentes_lista = sorted(
            produtos_pendentes.values(),
            key=lambda x: x['dias_pendente'],
            reverse=True
        )[:8]  # Top 8
        
        # 3. VOLUME DE COMPRAS POR EMPRESA
        volume_por_empresa = {}
        for solicitacao in solicitacoes:
            empresa_nome = solicitacao.empresa
            if empresa_nome not in volume_por_empresa:
                volume_por_empresa[empresa_nome] = {
                    'total': 0,
                    'valor_total': 0
                }
            
            volume_por_empresa[empresa_nome]['total'] += 1
            
            # Calcular valor
            preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                solicitacao_id=solicitacao.id
            ).filter(
                SolicitacoesPreenchidas.status != 'Rascunho'
            ).all()
            
            valor = sum(p.valor_total for p in preenchimentos if p and p.valor_total)
            volume_por_empresa[empresa_nome]['valor_total'] += valor
        
        # Calcular porcentagens e crescimento simplificado
        volume_total = sum(dados['total'] for dados in volume_por_empresa.values())
        volume_empresa_data = []
        
        for empresa_nome, dados in volume_por_empresa.items():
            porcentagem = round((dados['total'] / volume_total) * 100, 1) if volume_total > 0 else 0
            volume_empresa_data.append({
                'empresa': empresa_nome,
                'volume': dados['total'],
                'valor_total': dados['valor_total'],
                'crescimento': 0,  # Simplificado para demo
                'porcentagem': porcentagem
            })
        
        volume_empresa_data.sort(key=lambda x: x['volume'], reverse=True)
        
        # 4. PENDÊNCIAS POR EMPRESA
        pendencias_por_empresa = {}
        for solicitacao in solicitacoes:
            if solicitacao.status_aprovacao in [None, '', 'Pendente']:
                empresa_nome = solicitacao.empresa
                
                if empresa_nome not in pendencias_por_empresa:
                    pendencias_por_empresa[empresa_nome] = {
                        'quantidade': 0,
                        'dias_total': 0,
                        'valor_pendente': 0
                    }
                
                pendencias_por_empresa[empresa_nome]['quantidade'] += 1
                
                # Calcular dias pendente
                dias = (data_atual - solicitacao.data_solicitacao).days if solicitacao.data_solicitacao else 0
                pendencias_por_empresa[empresa_nome]['dias_total'] += dias
                
                # Calcular valor pendente
                preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                    solicitacao_id=solicitacao.id
                ).filter(
                    SolicitacoesPreenchidas.status != 'Rascunho'
                ).all()
                
                valor = sum(p.valor_total for p in preenchimentos if p and p.valor_total)
                pendencias_por_empresa[empresa_nome]['valor_pendente'] += valor
        
        pendencias_lista = []
        for empresa_nome, dados in pendencias_por_empresa.items():
            tempo_medio = dados['dias_total'] / dados['quantidade'] if dados['quantidade'] > 0 else 0
            pendencias_lista.append({
                'empresa': empresa_nome,
                'pendencias': dados['quantidade'],
                'tempo_medio_dias': round(tempo_medio, 1),
                'valor_pendente': dados['valor_pendente']
            })
        
        pendencias_lista.sort(key=lambda x: x['pendencias'], reverse=True)
        
        # 5. MAIOR VOLUME DE GASTOS POR VEÍCULO
        gastos_por_veiculo = {}
        for solicitacao in solicitacoes:
            if solicitacao.ativo == 'Sim' and solicitacao.nome_ativo:
                veiculo = solicitacao.nome_ativo
                
                if veiculo not in gastos_por_veiculo:
                    gastos_por_veiculo[veiculo] = {
                        'valor_total': 0,
                        'solicitacoes': 0,
                        'itens_total': 0
                    }
                
                # Buscar preenchimentos
                preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                    solicitacao_id=solicitacao.id
                ).filter(
                    SolicitacoesPreenchidas.status != 'Rascunho'
                ).all()
                
                valor = sum(p.valor_total for p in preenchimentos if p and p.valor_total)
                gastos_por_veiculo[veiculo]['valor_total'] += valor
                gastos_por_veiculo[veiculo]['solicitacoes'] += 1
                gastos_por_veiculo[veiculo]['itens_total'] += solicitacao.quantidade
        
        gastos_veiculo_lista = [
            {
                'veiculo': veiculo,
                'valor_total': dados['valor_total'],
                'solicitacoes': dados['solicitacoes'],
                'itens_total': dados['itens_total'],
                'valor_medio_por_item': dados['valor_total'] / dados['itens_total'] if dados['itens_total'] > 0 else 0
            }
            for veiculo, dados in gastos_por_veiculo.items()
        ]
        gastos_veiculo_lista.sort(key=lambda x: x['valor_total'], reverse=True)
        
        # 6. MAIOR VOLUME DE COMPRAS POR VEÍCULO (simplificado)
        compras_por_veiculo = {}
        for solicitacao in solicitacoes:
            if solicitacao.ativo == 'Sim' and solicitacao.nome_ativo:
                veiculo = solicitacao.nome_ativo
                if veiculo not in compras_por_veiculo:
                    compras_por_veiculo[veiculo] = 0
                compras_por_veiculo[veiculo] += solicitacao.quantidade
        
        compras_veiculo_lista = [
            {
                'veiculo': veiculo,
                'quantidade_total': quantidade,
                'solicitacoes': sum(1 for s in solicitacoes if s.nome_ativo == veiculo)
            }
            for veiculo, quantidade in compras_por_veiculo.items()
        ]
        compras_veiculo_lista.sort(key=lambda x: x['quantidade_total'], reverse=True)
        
        # 7. SOLICITAÇÕES COM MAIOR TEMPO EM ABERTO
        tempo_aberto_lista = []
        for solicitacao in solicitacoes:
            if solicitacao.status_aprovacao in [None, '', 'Pendente']:
                dias_aberto = (data_atual - solicitacao.data_solicitacao).days if solicitacao.data_solicitacao else 0
                
                if dias_aberto > 0:
                    material = db.session.get(Materiais, solicitacao.cod_material)
                    material_nome = material.DescricaoMaterial[:30] + '...' if material and len(material.DescricaoMaterial) > 30 else material.DescricaoMaterial if material else ''
                    
                    # Calcular prioridade simplificada
                    preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                        solicitacao_id=solicitacao.id
                    ).filter(
                        SolicitacoesPreenchidas.status != 'Rascunho'
                    ).all()
                    
                    valor_pendente = sum(p.valor_total for p in preenchimentos if p and p.valor_total)
                    prioridade = dias_aberto * (1 + (valor_pendente / 10000))
                    
                    tempo_aberto_lista.append({
                        'id': solicitacao.id,
                        'dias_aberto': dias_aberto,
                        'empresa': solicitacao.empresa,
                        'material': material_nome,
                        'valor_pendente': valor_pendente,
                        'prioridade': round(prioridade, 2),
                        'comprador': solicitacao.comprador_atribuido or 'Não Atribuído'
                    })
        
        tempo_aberto_lista.sort(key=lambda x: x['prioridade'], reverse=True)
        tempo_aberto_lista = tempo_aberto_lista[:10]
        
        # 8. COMPARAÇÃO DE GASTOS POR EMPRESA (simplificado)
        gastos_por_empresa = {}
        for solicitacao in solicitacoes:
            empresa_nome = solicitacao.empresa
            if empresa_nome not in gastos_por_empresa:
                gastos_por_empresa[empresa_nome] = 0
            
            # Buscar preenchimentos
            preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                solicitacao_id=solicitacao.id
            ).filter(
                SolicitacoesPreenchidas.status != 'Rascunho'
            ).all()
            
            valor = sum(p.valor_total for p in preenchimentos if p and p.valor_total)
            gastos_por_empresa[empresa_nome] += valor
        
        gastos_empresa_lista = []
        for empresa_nome, gastos in gastos_por_empresa.items():
            # Determinar tendência simplificada
            if gastos > 100000:
                tendencia = 'alta'
                simbolo = '↗'
            elif gastos > 50000:
                tendencia = 'estável'
                simbolo = '→'
            else:
                tendencia = 'baixa'
                simbolo = '↘'
            
            gastos_empresa_lista.append({
                'empresa': empresa_nome,
                'gastos_total': gastos,
                'variacao': 0,  # Simplificado para demo
                'tendencia': tendencia,
                'simbolo': simbolo
            })
        
        gastos_empresa_lista.sort(key=lambda x: x['gastos_total'], reverse=True)
        
        # 9. EVOLUÇÃO DE COMPPRAS NOS ÚLTIMOS MESES (dados de exemplo)
        meses_nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
        
        # Gerar dados de exemplo baseados nas empresas existentes
        empresas_existentes = list(set(s.empresa for s in solicitacoes))
        evolucao_compras = {}
        
        for empresa in empresas_existentes[:3]:  # Máximo 3 empresas
            # Gerar dados aleatórios para demonstração
            import random
            dados = [random.randint(50, 200) for _ in range(6)]  # 6 meses
            evolucao_compras[empresa] = dados
        
        # Nomes dos últimos 6 meses
        mes_atual = data_atual.month
        meses_labels = []
        for i in range(6):
            mes_num = mes_atual - (5 - i)
            ano = data_atual.year
            if mes_num <= 0:
                mes_num += 12
                ano -= 1
            meses_labels.append(f"{meses_nomes[mes_num - 1]}/{str(ano)[-2:]}")
        
        # 10. PRODUTIVIDADE GERAL POR EMPRESA
        produtividade_por_empresa = {}
        
        for solicitacao in solicitacoes:
            empresa_nome = solicitacao.empresa
            
            if empresa_nome not in produtividade_por_empresa:
                produtividade_por_empresa[empresa_nome] = {
                    'compradores': set(),
                    'solicitacoes': 0,
                    'aprovadas': 0,
                    'valor_total': 0
                }
            
            produtividade_por_empresa[empresa_nome]['solicitacoes'] += 1
            
            if solicitacao.comprador_atribuido:
                produtividade_por_empresa[empresa_nome]['compradores'].add(solicitacao.comprador_atribuido)
            
            if solicitacao.status_aprovacao == 'Aprovado':
                produtividade_por_empresa[empresa_nome]['aprovadas'] += 1
            
            # Calcular valor
            preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                solicitacao_id=solicitacao.id
            ).filter(
                SolicitacoesPreenchidas.status != 'Rascunho'
            ).all()
            
            valor = sum(p.valor_total for p in preenchimentos if p and p.valor_total)
            produtividade_por_empresa[empresa_nome]['valor_total'] += valor
        
        produtividade_lista = []
        for empresa_nome, dados in produtividade_por_empresa.items():
            taxa_aprovacao = (dados['aprovadas'] / dados['solicitacoes'] * 100) if dados['solicitacoes'] > 0 else 0
            
            produtividade_lista.append({
                'empresa': empresa_nome,
                'compradores': len(dados['compradores']),
                'solicitacoes': dados['solicitacoes'],
                'aprovadas': dados['aprovadas'],
                'taxa_aprovacao': round(taxa_aprovacao, 1),
                'valor_total': dados['valor_total']
            })
        
        produtividade_lista.sort(key=lambda x: x['taxa_aprovacao'], reverse=True)
        
        # OBTER LISTAS PARA FILTROS
        empresas_disponiveis = db.session.query(
            SolicitacoesCompra.empresa
        ).distinct().order_by(
            SolicitacoesCompra.empresa
        ).all()
        
        compradores_disponiveis = db.session.query(
            SolicitacoesCompra.comprador_atribuido
        ).filter(
            SolicitacoesCompra.comprador_atribuido.isnot(None),
            SolicitacoesCompra.comprador_atribuido != ''
        ).distinct().order_by(
            SolicitacoesCompra.comprador_atribuido
        ).all()
        
        empresas_filtro = ['Todas'] + [e[0] for e in empresas_disponiveis if e[0]]
        compradores_filtro = ['Todos'] + [c[0] for c in compradores_disponiveis if c[0]]
        
        # FORMATAR VALORES
        def formatar_valor(valor):
            try:
                valor = float(valor)
                if valor >= 1000000:
                    return f"R$ {valor/1000000:.1f}M"
                elif valor >= 1000:
                    return f"R$ {valor/1000:.1f}K"
                else:
                    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except:
                return "R$ 0,00"
        
        # Calcular tempo médio de aprovação
        tempo_medio_aprovacao = 0
        contador_tempo = 0
        for solicitacao in solicitacoes:
            if solicitacao.status_aprovacao == 'Aprovado' and solicitacao.data_solicitacao:
                dias = (data_atual - solicitacao.data_solicitacao).days
                tempo_medio_aprovacao += dias
                contador_tempo += 1
        
        tempo_medio_aprovacao = tempo_medio_aprovacao / contador_tempo if contador_tempo > 0 else 0
        
        # Calcular valor total e médio
        valor_total = 0
        for solicitacao in solicitacoes:
            preenchimentos = SolicitacoesPreenchidas.query.filter_by(
                solicitacao_id=solicitacao.id
            ).filter(
                SolicitacoesPreenchidas.status != 'Rascunho'
            ).all()
            
            for preenchimento in preenchimentos:
                if preenchimento and preenchimento.valor_total:
                    valor_total += preenchimento.valor_total
        
        valor_medio_por_solicitacao = valor_total / total_solicitacoes if total_solicitacoes > 0 else 0
        
        # Adicionar a função formatar_valor ao contexto do template
        import sys
        sys.path.append('.')
        
        return render_template(
            'dashboard.html',
            # Dados básicos
            total_solicitacoes=total_solicitacoes,
            aprovadas=aprovadas,
            pendentes=pendentes,
            rejeitadas=rejeitadas,
            valor_total=formatar_valor(valor_total),
            tempo_medio_aprovacao=round(tempo_medio_aprovacao, 1),
            valor_medio_por_solicitacao=formatar_valor(valor_medio_por_solicitacao),
            
            # Dados para gráficos e tabelas
            compradores_produtividade=compradores_lista[:8],
            produtos_pendentes=produtos_pendentes_lista[:8],
            volume_empresa=volume_empresa_data[:6],
            pendencias_empresa=pendencias_lista[:6],
            gastos_veiculo=gastos_veiculo_lista[:5],
            compras_veiculo=compras_veiculo_lista[:5],
            tempo_aberto=tempo_aberto_lista,
            gastos_empresa=gastos_empresa_lista[:6],
            evolucao_compras=evolucao_compras,
            evolucao_meses_labels=meses_labels,
            produtividade_geral=produtividade_lista[:6],
            
            # Filtros
            empresas_filtro=empresas_filtro,
            compradores_filtro=compradores_filtro,
            empresa_selecionada=empresa,
            periodo_selecionado=periodo,
            status_selecionado=status_filtro,
            comprador_selecionado=comprador_filtro,
            
            # Função auxiliar
            formatar_valor=formatar_valor
        )
        
    except Exception as e:
        flash(f'Erro ao carregar dashboard: {str(e)}', 'error')
        app.logger.error(f'Erro no dashboard: {str(e)}', exc_info=True)
        
        # Retornar dados vazios em caso de erro
        return render_template(
            'dashboard.html',
            total_solicitacoes=0,
            aprovadas=0,
            pendentes=0,
            rejeitadas=0,
            valor_total="R$ 0,00",
            tempo_medio_aprovacao=0,
            valor_medio_por_solicitacao="R$ 0,00",
            compradores_produtividade=[],
            produtos_pendentes=[],
            volume_empresa=[],
            pendencias_empresa=[],
            gastos_veiculo=[],
            compras_veiculo=[],
            tempo_aberto=[],
            gastos_empresa=[],
            evolucao_compras={},
            evolucao_meses_labels=[],
            produtividade_geral=[],
            empresas_filtro=['Todas'],
            compradores_filtro=['Todos'],
            empresa_selecionada='Todas',
            periodo_selecionado='Últimos 30 dias',
            status_selecionado='Todos',
            comprador_selecionado='Todos'
        )

#------------------------------------------------------------------------------------
    
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
                
                # CORREÇÃO: A marca está na solicitação, não no preenchimento
                marca = preenchimento.solicitacao.marca if preenchimento.solicitacao.marca else 'Não informado'
                
                preenchimentos_info.append({
                    'id': preenchimento.id,
                    'marca': marca,  # Usando a marca da solicitação
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
        valor_unitario_str = data.get('valor_unitario')
        valor_frete_str = data.get('valor_frete', '0')
        valor_unitario_original_str = data.get('valor_unitario_original')
        observacoes = data.get('observacoes')
        
        # Validações
        if not preenchimento_id:
            return jsonify({'success': False, 'message': 'ID do preenchimento é obrigatório'}), 400
        
        if not valor_unitario_str:
            return jsonify({'success': False, 'message': 'Valor unitário é obrigatório'}), 400
        
        # Função robusta para conversão
        def safe_currency_convert(value_str):
            if not value_str:
                return 0.0
            try:
                # Remove caracteres não numéricos exceto vírgula e ponto
                cleaned = re.sub(r'[^\d,.]', '', str(value_str))
                # Remove pontos de milhar (preserva apenas o último ponto como decimal)
                if ',' in cleaned and '.' in cleaned:
                    # Formato: 1.234,56 -> remove pontos de milhar
                    cleaned = cleaned.replace('.', '')
                    cleaned = cleaned.replace(',', '.')
                elif ',' in cleaned:
                    # Formato: 1234,56
                    cleaned = cleaned.replace(',', '.')
                # Se só tem ponto, assume que é decimal
                return float(cleaned) if cleaned else 0.0
            except (ValueError, TypeError):
                return 0.0
        
        valor_unitario = safe_currency_convert(valor_unitario_str)
        valor_frete = safe_currency_convert(valor_frete_str)
        valor_unitario_original = safe_currency_convert(valor_unitario_original_str)
        
        # Valida se os valores são números válidos
        if valor_unitario <= 0:
            return jsonify({'success': False, 'message': 'Valor unitário deve ser maior que zero'}), 400
        
        # Buscar o preenchimento
        preenchimento = SolicitacoesPreenchidas.query.get_or_404(preenchimento_id)
        
        # Salvar valores anteriores
        valor_unitario_anterior = preenchimento.valor_unitario
        valor_frete_anterior = preenchimento.valor_frete
        
        # Calcular novo valor total
        quantidade = preenchimento.solicitacao.quantidade
        novo_valor_total = (valor_unitario * quantidade) + valor_frete
        
        # Atualizar os valores
        preenchimento.valor_unitario = valor_unitario
        preenchimento.valor_frete = valor_frete if valor_frete > 0 else None
        preenchimento.valor_total = novo_valor_total
        preenchimento.observacoes = observacoes
        
        db.session.commit()
        
        return jsonify({
            'success': True, 
            'message': 'Valores atualizados com sucesso',
            'novo_valor_total': novo_valor_total,
            'valor_unitario': valor_unitario,
            'valor_frete': valor_frete
        })
        
    except Exception as e:
        db.session.rollback()
        return jsonify({
            'success': False, 
            'message': f'Erro ao atualizar valores: {str(e)}'
        }), 500
    
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
        # Usar SQLAlchemy em vez de sqlite3 direto
        from sqlalchemy import inspect
        
        inspector = inspect(db.engine)
        columns = [col['name'] for col in inspector.get_columns('SolicitacoesCompra')]
        
        if 'status_aprovacao' not in columns:
            # Usar SQLAlchemy para executar o ALTER TABLE
            db.engine.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN status_aprovacao TEXT")
            logging.info("Coluna status_aprovacao adicionada à tabela SolicitacoesCompra")
            print("✓ Coluna status_aprovacao adicionada com sucesso!")
        else:
            print("✓ Coluna status_aprovacao já existe na tabela")
        
        return True
    except Exception as e:
        logging.error(f"Erro na migração status_aprovacao: {str(e)}")
        print(f"✗ Erro na migração: {str(e)}")
        return False
    
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
    
def migrate_solicitacoes_preenchidas_status():
    """Migra o status padrão para suportar 'Rascunho'"""
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Verificar se existem registros com status vazio ou nulo
        cursor.execute("UPDATE SolicitacoesPreenchidas SET status = 'Rascunho' WHERE status IS NULL OR status = ''")
        conn.commit()
        
        # Verificar se a atualização foi bem-sucedida
        cursor.execute("SELECT COUNT(*) FROM SolicitacoesPreenchidas WHERE status = 'Rascunho'")
        count = cursor.fetchone()[0]
        
        logging.info(f"Migração de status concluída: {count} registros atualizados para 'Rascunho'")
        print(f"✓ Migração de status concluída: {count} registros atualizados")
        
        conn.close()
        return True
    except sqlite3.Error as e:
        logging.error(f"Erro na migração de status: {str(e)}")
        print(f"✗ Erro na migração: {str(e)}")
        return False
    

@routes_bp.route('/teste_data_hora')
def teste_data_hora():
    from datetime import datetime
    return f"""
    <h3>Teste de Data/Hora</h3>
    <p>datetime.utcnow(): {datetime.utcnow()}</p>
    <p>datetime.now(): {datetime.now()}</p>
    <p>get_local_time(): {get_local_time()}</p>
    <p>Ano atual: {get_local_time().year}</p>
    """

@app.template_filter('format_brasil_time')
def format_brasil_time(dt):
    if dt is None:
        return ""
    
    # Converte para fuso horário de Brasília se não tiver informação de timezone
    if dt.tzinfo is None:
        from datetime import timezone, timedelta
        brasil_tz = timezone(timedelta(hours=-3))
        dt = dt.replace(tzinfo=timezone.utc).astimezone(brasil_tz)
    
    return dt.strftime('%d/%m/%Y %H:%M')

def add_aplicacao_geral_column():
    """Adiciona a coluna aplicacao_geral na tabela SolicitacoesCompra"""
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        
        # Verificar se a coluna já existe
        cursor.execute("PRAGMA table_info(SolicitacoesCompra)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'aplicacao_geral' not in columns:
            cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN aplicacao_geral TEXT")
            conn.commit()
            logging.info("Coluna aplicacao_geral adicionada à tabela SolicitacoesCompra")
            print("✓ Coluna aplicacao_geral adicionada com sucesso!")
        else:
            print("✓ Coluna aplicacao_geral já existe na tabela")
        
        conn.close()
        return True
    except sqlite3.Error as e:
        logging.error(f"Erro ao adicionar coluna aplicacao_geral: {str(e)}")
        print(f"✗ Erro ao adicionar coluna aplicacao_geral: {str(e)}")
        return False
    
#Tela comprador
# No app.py, substitua a função get_compradores por esta versão
# Definir o caminho absoluto para senhas.txt
SENHAS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'senhas.txt')

# Verificar se o arquivo existe no início
print(f"Verificando caminho de senhas.txt: {SENHAS_FILE}")
print(f"Arquivo existe: {os.path.exists(SENHAS_FILE)}")
    
def get_next_comprador(aplicacao):
    """
    Retorna o comprador a ser atribuído com base na aplicação.
    Se já houver um comprador atribuído a essa aplicação, reutiliza o mesmo.
    Caso contrário, distribui de forma balanceada entre os compradores disponíveis.
    """
    try:
        compradores = get_compradores()
        if not compradores:
            return None

        # Normaliza nome da aplicação
        aplicacao = aplicacao.strip().lower()

        # Verifica se já há comprador atribuído para essa aplicação
        solicit_existente = SolicitacoesCompra.query.filter_by(aplicacao=aplicacao).filter(
            SolicitacoesCompra.comprador_atribuido.isnot(None)
        ).first()

        if solicit_existente and solicit_existente.comprador_atribuido:
            # Reutiliza o comprador já existente para essa aplicação
            return solicit_existente.comprador_atribuido

        # Caso seja uma nova aplicação (sem histórico), distribui automaticamente
        todas_solicitacoes = SolicitacoesCompra.query.filter(
            SolicitacoesCompra.comprador_atribuido.isnot(None)
        ).all()

        # Conta quantas solicitações cada comprador já tem (para balanceamento)
        contagem = {c: 0 for c in compradores}
        for sol in todas_solicitacoes:
            if sol.comprador_atribuido in contagem:
                contagem[sol.comprador_atribuido] += 1

        # Seleciona o comprador com menor carga de solicitações
        comprador_escolhido = min(contagem, key=contagem.get)
        return comprador_escolhido

    except Exception as e:
        logging.error(f"Erro em get_next_comprador: {str(e)}")
        return None


# Adicione esta migração no init_db ou em uma função separada (ex: no bloco de migrações no final do app.py)
def add_comprador_atribuido_column():
    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(SolicitacoesCompra)")
        columns = [col[1] for col in cursor.fetchall()]
        if 'comprador_atribuido' not in columns:
            cursor.execute("ALTER TABLE SolicitacoesCompra ADD COLUMN comprador_atribuido TEXT")
            conn.commit()
            logging.info("Coluna comprador_atribuido adicionada à tabela SolicitacoesCompra")
            print("✓ Coluna comprador_atribuido adicionada com sucesso!")
        else:
            print("✓ Coluna comprador_atribuido já existe na tabela")
        conn.close()
        return True
    except sqlite3.Error as e:
        logging.error(f"Erro ao adicionar coluna comprador_atribuido: {str(e)}")
        print(f"✗ Erro ao adicionar coluna: {str(e)}")
        return False

# No bloco if __name__ == '__main__': adicione esta chamada à migração
add_comprador_atribuido_column()

# Adicione esta rota no app.py, no bloco de rotas (após outras rotas semelhantes)
@routes_bp.route('/api/solicitacoes/<int:solicitacao_id>/atribuir_comprador', methods=['POST'])
def atribuir_comprador(solicitacao_id):
    if 'usuario' not in session:
        return jsonify({'success': False, 'error': 'Usuário não autenticado'}), 401
    try:
        data = request.get_json()
        comprador = data.get('comprador')
        if not comprador:
            return jsonify({'success': False, 'error': 'Nenhum comprador selecionado'}), 400
        solicitacao = SolicitacoesCompra.query.get_or_404(solicitacao_id)
        solicitacao.comprador_atribuido = comprador
        db.session.commit()
        return jsonify({
            'success': True,
            'message': f'Solicitação atribuída a {comprador} com sucesso!'
        })
    except Exception as e:
        db.session.rollback()
        logging.error(f"Erro ao atribuir comprador: {str(e)}")
        return jsonify({'success': False, 'error': f'Erro ao atribuir comprador: {str(e)}'}), 500
    
@app.template_filter('basename')
def basename_filter(path):
    """
    Filtro Jinja2 para extrair apenas o nome do arquivo de um caminho completo.
    Ex: '/Uploads/cotacoes/123.pdf' → '123.pdf'
    """
    if not path:
        return ''
    
    return basename(path)
# Registre o filtro (ADICIONE ESTA LINHA)
app.jinja_env.filters['format_brasil_time'] = format_brasil_time
app.jinja_env.filters['basename'] = basename_filter

# No app.py, substitua ou adicione a rota para listar_solicitacoes_finalizadas


@routes_bp.route('/listar_solicitacoes_finalizadas', methods=['GET'])
def listar_solicitacoes_finalizadas():
    if 'usuario' not in session:
        flash('Acesso não autorizado', 'error')
        return redirect(url_for('routes_bp.login'))
    
    try:
        solicitacoes = SolicitacoesCompra.query.join(Materiais).all()
        empresas = db.session.query(SolicitacoesCompra.empresa).distinct().order_by(SolicitacoesCompra.empresa).all()
        empresas = [e[0] for e in empresas if e[0]]
        usuarios = db.session.query(SolicitacoesCompra.usuario).distinct().order_by(SolicitacoesCompra.usuario).all()
        usuarios = [u[0] for u in usuarios if u[0]]
        aplicacoes = db.session.query(SolicitacoesCompra.aplicacao).distinct().order_by(SolicitacoesCompra.aplicacao).all()
        aplicacoes = [a[0] for a in aplicacoes if a[0]]
        
        compradores = get_compradores()
        print(f"Compradores enviados ao template: {compradores}")  # Depuração
        logging.info(f"Compradores enviados ao template: {compradores}")
        
        if not compradores:
            flash('Nenhum comprador encontrado no arquivo senhas.txt. Verifique o arquivo ou a função get_compradores.', 'error')
        
        return render_template(
            'listar_solicitacoes.html',
            solicitacoes=solicitacoes,
            empresas=empresas,
            usuarios=usuarios,
            aplicacoes=aplicacoes,
            compradores=compradores  # Garante que compradores seja sempre passado
        )
    
    except Exception as e:
        flash(f'Erro ao carregar solicitações: {str(e)}', 'error')
        logging.error(f"Erro em listar_solicitacoes_finalizadas: {str(e)}")
        print(f"ERRO: Erro em listar_solicitacoes_finalizadas: {str(e)}")
        return render_template(
            'listar_solicitacoes.html',
            solicitacoes=[],
            empresas=[],
            usuarios=[],
            aplicacoes=[],
            compradores=[]  # Passa lista vazia em caso de erro
        )
    
def get_fornecedor_details(fornecedor_id):
    """Busca nome_fantasia e cnpj do fornecedor pelo ID no banco fornecedores.db"""
    try:
        conn = sqlite3.connect(DB_PATH_FORNECEDORES)
        cursor = conn.cursor()
        cursor.execute('SELECT nome_fantasia, cnpj FROM fornecedores WHERE id = ?', (fornecedor_id,))
        result = cursor.fetchone()
        conn.close()
        if result:
            return result[0], result[1]  # nome_fantasia, cnpj
        return None, None
    except sqlite3.Error as e:
        logging.error(f"Erro ao buscar fornecedor {fornecedor_id}: {str(e)}")
        return None, None

@routes_bp.route('/api/fornecedores', methods=['GET'])
def api_fornecedores():
    """API para fornecedores com paginação no servidor"""
    if 'usuario' not in session:
        return jsonify({'error': 'Não autenticado'}), 401
    
    try:
        # Parâmetros de paginação
        page = request.args.get('page', 1, type=int)
        per_page = request.args.get('per_page', 25, type=int)
        search_name = request.args.get('search_name', '').strip()
        search_material = request.args.get('search_material', '').strip()
        
        # Conexão com o banco de fornecedores
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if not conn:
            return jsonify({'error': 'Erro ao conectar ao banco'}), 500
        
        try:
            cursor = conn.cursor()
            
            # Query base com filtros
            query = '''
                SELECT id, nome_fantasia, cnpj, telefone, email, endereco, 
                       bairro, cidade, estado, contato, materiais
                FROM fornecedores 
                WHERE 1=1
            '''
            params = []
            
            if search_name:
                query += ' AND (nome_fantasia LIKE ? OR cnpj LIKE ?)'
                params.extend([f'%{search_name}%', f'%{search_name}%'])
            
            if search_material:
                query += ' AND materiais LIKE ?'
                params.append(f'%{search_material}%')
            
            # Ordenação
            query += ' ORDER BY nome_fantasia'
            
            # Contagem total (para paginação)
            count_query = 'SELECT COUNT(*) FROM fornecedores WHERE 1=1'
            count_params = []
            
            if search_name:
                count_query += ' AND (nome_fantasia LIKE ? OR cnpj LIKE ?)'
                count_params.extend([f'%{search_name}%', f'%{search_name}%'])
            
            if search_material:
                count_query += ' AND materiais LIKE ?'
                count_params.append(f'%{search_material}%')
            
            cursor.execute(count_query, count_params)
            total_count = cursor.fetchone()[0]
            
            # Query com paginação
            query += ' LIMIT ? OFFSET ?'
            offset = (page - 1) * per_page
            params.extend([per_page, offset])
            
            cursor.execute(query, params)
            fornecedores = []
            
            for row in cursor.fetchall():
                fornecedor = {
                    'id': row[0],
                    'nome_fantasia': row[1],
                    'cnpj': row[2],
                    'telefone': row[3],
                    'email': row[4],
                    'endereco': row[5],
                    'bairro': row[6],
                    'cidade': row[7],
                    'estado': row[8],
                    'contato': row[9],
                    'materiais': row[10],
                    'cnpj_formatado': format_cnpj(row[2]) if row[2] else ''
                }
                fornecedores.append(fornecedor)
            
            return jsonify({
                'success': True,
                'fornecedores': fornecedores,
                'pagination': {
                    'page': page,
                    'per_page': per_page,
                    'total_count': total_count,
                    'total_pages': (total_count + per_page - 1) // per_page
                }
            })
            
        finally:
            conn.close()
            
    except Exception as e:
        logging.error(f"Erro na API de fornecedores: {str(e)}")
        return jsonify({'error': str(e)}), 500
    
@routes_bp.route('/api/excluir_grupo_aplicacao', methods=['POST'])
def excluir_grupo_aplicacao():
    """Exclui todas as solicitações de um grupo (aplicação) e seus relacionamentos"""
    if 'usuario' not in session:
        return jsonify({'success': False, 'error': 'Usuário não autenticado'}), 401
    
    try:
        data = request.get_json()
        aplicacao = data.get('aplicacao')
        
        if not aplicacao:
            return jsonify({'success': False, 'error': 'Aplicação não especificada'}), 400
        
        usuario = session['usuario']
        print(f"🚨 Usuário {usuario} solicitando exclusão do grupo: {aplicacao}")
        
        # 1. Buscar todas as solicitações APROVADAS desta aplicação
        solicitacoes = SolicitacoesCompra.query.filter_by(
            aplicacao=aplicacao,
            status_aprovacao='Aprovado'
        ).all()
        
        if not solicitacoes:
            return jsonify({'success': False, 'error': f'Nenhuma solicitação APROVADA encontrada para "{aplicacao}"'}), 404
        
        ids_solicitacoes = [s.id for s in solicitacoes]
        count_solicitacoes = len(ids_solicitacoes)
        
        print(f"📋 Encontradas {count_solicitacoes} solicitações para exclusão")
        
        # 2. Buscar todos os preenchimentos relacionados
        preenchimentos = SolicitacoesPreenchidas.query.filter(
            SolicitacoesPreenchidas.solicitacao_id.in_(ids_solicitacoes)
        ).all()
        
        ids_preenchimentos = [p.id for p in preenchimentos]
        count_preenchimentos = len(ids_preenchimentos)
        
        print(f"📋 Encontrados {count_preenchimentos} preenchimentos para exclusão")
        
        # 3. Buscar e excluir histórico de descontos
        historicos = HistoricoDescontos.query.filter(
            HistoricoDescontos.preenchimento_id.in_(ids_preenchimentos)
        ).all()
        
        for historico in historicos:
            db.session.delete(historico)
        
        count_historicos = len(historicos)
        print(f"📋 Encontrados {count_historicos} históricos de desconto para exclusão")
        
        # 4. Verificar se há pedidos de compra associados
        pedidos_associados = []
        if ids_preenchimentos:
            # Buscar pedidos que têm esses preenchimentos
            pedidos = PedidosCompra.query.join(
                pedido_preenchimento_associacao,
                PedidosCompra.id == pedido_preenchimento_associacao.c.pedido_id
            ).filter(
                pedido_preenchimento_associacao.c.preenchimento_id.in_(ids_preenchimentos)
            ).distinct().all()
            
            pedidos_associados = [p.numero_pedido for p in pedidos]
        
        if pedidos_associados:
            return jsonify({
                'success': False,
                'error': f'Não é possível excluir: existem pedidos de compra associados ({len(pedidos_associados)} pedidos)',
                'pedidos': pedidos_associados
            }), 400
        
        # 5. Verificar se há registros em outras tabelas
        # Estoque
        estoques = Estoque.query.filter(
            Estoque.preenchimento_id.in_(ids_preenchimentos)
        ).all()
        
        if estoques:
            for estoque in estoques:
                db.session.delete(estoque)
            print(f"📋 Excluídos {len(estoques)} registros de estoque")
        
        # Requisições
        requisicoes = Requisicoes.query.filter(
            Requisicoes.preenchimento_id.in_(ids_preenchimentos)
        ).all()
        
        if requisicoes:
            for requisicao in requisicoes:
                db.session.delete(requisicao)
            print(f"📋 Excluídos {len(requisicoes)} registros de requisições")
        
        # Auditoria
        auditorias = Auditoria.query.filter(
            Auditoria.solicitacao_id.in_(ids_solicitacoes)
        ).all()
        
        if auditorias:
            for auditoria in auditorias:
                db.session.delete(auditoria)
            print(f"📋 Excluídos {len(auditorias)} registros de auditoria")
        
        # 6. Excluir preenchimentos
        for preenchimento in preenchimentos:
            db.session.delete(preenchimento)
        
        # 7. Excluir solicitações
        for solicitacao in solicitacoes:
            db.session.delete(solicitacao)
        
        # 8. Commit da transação
        db.session.commit()
        
        # 9. Log detalhado
        logging.info(f"""
        ✅ GRUPO EXCLUÍDO COM SUCESSO
        Aplicação: {aplicacao}
        Usuário: {usuario}
        Data/Hora: {get_local_time()}
        Estatísticas:
          - Solicitações: {count_solicitacoes}
          - Preenchimentos: {count_preenchimentos}
          - Históricos: {count_historicos}
          - Estoques: {len(estoques) if estoques else 0}
          - Requisições: {len(requisicoes) if requisicoes else 0}
          - Auditorias: {len(auditorias) if auditorias else 0}
        """)
        
        return jsonify({
            'success': True,
            'message': f'✅ Grupo "{aplicacao}" excluído com sucesso!',
            'details': {
                'solicitacoes_excluidas': count_solicitacoes,
                'preenchimentos_excluidas': count_preenchimentos,
                'historicos_excluidos': count_historicos,
                'estoques_excluidos': len(estoques) if estoques else 0,
                'requisicoes_excluidas': len(requisicoes) if requisicoes else 0,
                'auditorias_excluidas': len(auditorias) if auditorias else 0
            }
        })
        
    except Exception as e:
        db.session.rollback()
        logging.error(f"❌ ERRO AO EXCLUIR GRUPO {aplicacao}: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': f'Erro ao excluir grupo: {str(e)}'
        }), 500

@routes_bp.route('/editar_fornecedor/<int:fornecedor_id>', methods=['GET', 'POST'])
def editar_fornecedor(fornecedor_id):
    """Rota para editar fornecedor (GET para carregar, POST para salvar)"""
    if 'usuario' not in session:
        flash('Você precisa estar logado para editar fornecedores.', 'warning')
        return redirect(url_for('routes_bp.login'))
    
    conn = get_db_connection(DB_PATH_FORNECEDORES)
    if not conn:
        flash('Erro ao conectar ao banco de dados.', 'error')
        return redirect(url_for('routes_bp.lista_fornecedores'))
    
    try:
        cursor = conn.cursor()
        
        if request.method == 'GET':
            # Buscar dados do fornecedor para edição
            cursor.execute('''
                SELECT id, nome_fantasia, cnpj, telefone, email, endereco, 
                       bairro, cidade, estado, contato, materiais
                FROM fornecedores WHERE id = ?
            ''', (fornecedor_id,))
            
            fornecedor_raw = cursor.fetchone()
            
            if not fornecedor_raw:
                flash('Fornecedor não encontrado.', 'error')
                return redirect(url_for('routes_bp.lista_fornecedores'))
            
            # Converter para dicionário
            fornecedor = {
                'id': fornecedor_raw[0],
                'nome_fantasia': fornecedor_raw[1],
                'cnpj': fornecedor_raw[2],
                'telefone': fornecedor_raw[3],
                'email': fornecedor_raw[4],
                'endereco': fornecedor_raw[5],
                'bairro': fornecedor_raw[6],
                'cidade': fornecedor_raw[7],
                'estado': fornecedor_raw[8],
                'contato': fornecedor_raw[9],
                'materiais': fornecedor_raw[10]
            }
            
            return render_template('editar_fornecedor.html', fornecedor=fornecedor)
        
        elif request.method == 'POST':
            # Obter dados do formulário de edição
            nome_fantasia = request.form.get('nome_fantasia', '').strip()
            telefone = request.form.get('telefone', '').strip()
            email = request.form.get('email', '').strip()
            endereco = request.form.get('endereco', '').strip()
            bairro = request.form.get('bairro', '').strip()
            cidade = request.form.get('cidade', '').strip()
            estado = request.form.get('estado', '').strip()
            contato = request.form.get('contato', '').strip()
            materiais = request.form.get('materiais', '').strip()
            
            # Validações (exceto CNPJ)
            if not all([nome_fantasia, telefone, email, endereco, bairro, cidade, estado, contato, materiais]):
                flash('Todos os campos são obrigatórios.', 'error')
                return redirect(url_for('routes_bp.editar_fornecedor', fornecedor_id=fornecedor_id))
            
            if not validate_email(email):
                flash('Email inválido. Insira um email válido.', 'error')
                return redirect(url_for('routes_bp.editar_fornecedor', fornecedor_id=fornecedor_id))
            
            if len(telefone) < 10:
                flash('Telefone inválido. Deve conter pelo menos 10 dígitos.', 'error')
                return redirect(url_for('routes_bp.editar_fornecedor', fornecedor_id=fornecedor_id))
            
            # Atualizar fornecedor (exceto CNPJ)
            cursor.execute('''
                UPDATE fornecedores SET 
                    nome_fantasia = ?,
                    telefone = ?,
                    email = ?,
                    endereco = ?,
                    bairro = ?,
                    cidade = ?,
                    estado = ?,
                    contato = ?,
                    materiais = ?
                WHERE id = ?
            ''', (nome_fantasia, telefone, email, endereco, bairro, cidade, estado, contato, materiais, fornecedor_id))
            
            conn.commit()
            flash('Fornecedor atualizado com sucesso!', 'success')
            return redirect(url_for('routes_bp.lista_fornecedores'))
            
    except Exception as e:
        if conn:
            conn.rollback()
        flash(f'Erro ao editar fornecedor: {str(e)}', 'error')
        return redirect(url_for('routes_bp.lista_fornecedores'))
    finally:
        if conn:
            conn.close()

@routes_bp.route('/excluir_fornecedor/<int:fornecedor_id>', methods=['POST'])
def excluir_fornecedor(fornecedor_id):
    """Exclui um fornecedor - versão simplificada e testada"""
    if 'usuario' not in session:
        flash('Você precisa estar logado.', 'warning')
        return redirect(url_for('routes_bp.login'))
    
    try:
        # Conectar ao banco de fornecedores
        conn = sqlite3.connect(DB_PATH_FORNECEDORES)
        cursor = conn.cursor()
        
        # 1. Buscar nome do fornecedor para mensagem
        cursor.execute('SELECT nome_fantasia FROM fornecedores WHERE id = ?', (fornecedor_id,))
        resultado = cursor.fetchone()
        
        if not resultado:
            flash('Fornecedor não encontrado.', 'error')
            return redirect(url_for('routes_bp.lista_fornecedores'))
        
        nome_fornecedor = resultado[0]
        
        # 2. Verificar se está em uso (opcional - comente se quiser forçar exclusão)
        try:
            conn_main = sqlite3.connect(DATABASE)
            cursor_main = conn_main.cursor()
            cursor_main.execute('SELECT COUNT(*) FROM SolicitacoesPreenchidas WHERE fornecedor_id = ?', (fornecedor_id,))
            uso_count = cursor_main.fetchone()[0]
            conn_main.close()
            
            if uso_count > 0:
                flash(f'❌ Fornecedor "{nome_fornecedor}" está em {uso_count} cotações. Não pode ser excluído.', 'error')
                return redirect(url_for('routes_bp.lista_fornecedores'))
        except Exception as e:
            print(f"AVISO: Não foi possível verificar uso - {e}")
            # Continua mesmo sem verificar
        
        # 3. Excluir fornecedor
        cursor.execute('DELETE FROM fornecedores WHERE id = ?', (fornecedor_id,))
        conn.commit()
        conn.close()
        
        flash(f'✅ Fornecedor "{nome_fornecedor}" excluído com sucesso!', 'success')
        
    except Exception as e:
        print(f"ERRO CRÍTICO ao excluir: {str(e)}")
        flash(f'❌ Erro ao excluir fornecedor: {str(e)}', 'error')
    
    return redirect(url_for('routes_bp.lista_fornecedores'))

    
@routes_bp.route('/api/fornecedor/<int:fornecedor_id>', methods=['GET'])
def api_get_fornecedor(fornecedor_id):
    """API para obter dados de um fornecedor específico"""
    if 'usuario' not in session:
        return jsonify({'success': False, 'error': 'Não autenticado'}), 401
    
    try:
        conn = get_db_connection(DB_PATH_FORNECEDORES)
        if not conn:
            return jsonify({'success': False, 'error': 'Erro ao conectar ao banco'}), 500
        
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, nome_fantasia, cnpj, telefone, email, endereco, 
                   bairro, cidade, estado, contato, materiais
            FROM fornecedores WHERE id = ?
        ''', (fornecedor_id,))
        
        fornecedor_raw = cursor.fetchone()
        conn.close()
        
        if not fornecedor_raw:
            return jsonify({'success': False, 'error': 'Fornecedor não encontrado'}), 404
        
        fornecedor = {
            'id': fornecedor_raw[0],
            'nome_fantasia': fornecedor_raw[1],
            'cnpj': fornecedor_raw[2],
            'telefone': fornecedor_raw[3],
            'email': fornecedor_raw[4],
            'endereco': fornecedor_raw[5],
            'bairro': fornecedor_raw[6],
            'cidade': fornecedor_raw[7],
            'estado': fornecedor_raw[8],
            'contato': fornecedor_raw[9],
            'materiais': fornecedor_raw[10],
            'cnpj_formatado': format_cnpj(fornecedor_raw[2]) if fornecedor_raw[2] else ''
        }
        
        return jsonify({
            'success': True,
            'fornecedor': fornecedor
        })
        
    except Exception as e:
        logging.error(f"Erro na API de fornecedor específico: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@routes_bp.route('/duplicar_grupo_aplicacao', methods=['POST'])
def duplicar_grupo_aplicacao():
    """Duplica todas as solicitações de um grupo (aplicação)"""
    if 'usuario' not in session:
        return jsonify({'success': False, 'error': 'Usuário não autenticado'}), 401
    
    try:
        data = request.get_json()
        aplicacao_original = data.get('aplicacao')
        
        if not aplicacao_original:
            return jsonify({'success': False, 'error': 'Aplicação não especificada'}), 400
        
        # Buscar todas as solicitações da aplicação original
        solicitacoes_originais = SolicitacoesCompra.query.filter(
            SolicitacoesCompra.aplicacao == aplicacao_original,
            SolicitacoesCompra.status_aprovacao == "Aprovado"
        ).all()
        
        if not solicitacoes_originais:
            return jsonify({'success': False, 'error': 'Nenhuma solicitação encontrada para esta aplicação'}), 404
        
        # Criar nova aplicação com sufixo _copia
        aplicacao_nova = f"{aplicacao_original}_copia"
        usuario = session['usuario']
        
        # Contador de duplicações bem-sucedidas
        duplicacoes_criadas = 0
        
        # Duplicar cada solicitação
        for solicitacao_original in solicitacoes_originais:
            try:
                # Criar nova solicitação baseada na original
                nova_solicitacao = SolicitacoesCompra(
                    cod_material=solicitacao_original.cod_material,
                    especificacao=solicitacao_original.especificacao,
                    quantidade=solicitacao_original.quantidade,
                    unidade_medida=solicitacao_original.unidade_medida,
                    aplicacao=aplicacao_nova,
                    aplicacao_geral=aplicacao_nova,  # Aplicação geral também atualizada
                    empresa=solicitacao_original.empresa,
                    usuario=usuario,
                    foto_path=solicitacao_original.foto_path,
                    marca=solicitacao_original.marca,
                    ativo=solicitacao_original.ativo,
                    nome_ativo=solicitacao_original.nome_ativo,
                    prioridade=solicitacao_original.prioridade,
                    status_aprovacao=None,  # Nova solicitação começa sem aprovação
                    comprador_atribuido=None  # Resetar comprador atribuído
                )
                
                db.session.add(nova_solicitacao)
                duplicacoes_criadas += 1
                
            except Exception as e:
                db.session.rollback()
                logging.error(f"Erro ao duplicar solicitação {solicitacao_original.id}: {str(e)}")
                return jsonify({
                    'success': False, 
                    'error': f'Erro ao duplicar solicitação {solicitacao_original.id}: {str(e)}'
                }), 500
        
        # Commit de todas as novas solicitações
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': f'Grupo duplicado com sucesso! {duplicacoes_criadas} solicitações criadas.',
            'aplicacao_nova': aplicacao_nova,
            'count': duplicacoes_criadas
        })
        
    except Exception as e:
        db.session.rollback()
        logging.error(f"Erro geral ao duplicar grupo: {str(e)}")
        return jsonify({
            'success': False, 
            'error': f'Erro ao duplicar grupo: {str(e)}'
        }), 500   
    
@routes_bp.route('/reprovar_solicitacao_individual/<int:id>', methods=['POST'])
def reprovar_solicitacao_individual(id):
    if 'usuario' not in session:
        return jsonify({'success': False, 'message': 'Usuário não autenticado'}), 401
    
    try:
        # Buscar a solicitação
        solicitacao = SolicitacoesCompra.query.get_or_404(id)
        usuario = session['usuario']
        
        # Verificar se já está reprovada
        if solicitacao.status_aprovacao == 'Reprovado':
            return jsonify({
                'success': False, 
                'message': 'Esta solicitação já está reprovada.'
            }), 400
        
        # REMOVIDA A VERIFICAÇÃO DE COTAÇÕES
        # A reprovação agora é permitida independentemente de ter cotações
        
        # Obter motivo do corpo da requisição
        data = request.get_json()
        motivo = data.get('motivo', '').strip() if data else ''
        
        # Atualizar status para reprovado
        solicitacao.status_aprovacao = 'Reprovado'
        
        # Adicionar observação se houver motivo
        if motivo:
            if solicitacao.observacoes_col:
                solicitacao.observacoes_col += f"\nREPROVADA INDIVIDUALMENTE - Motivo: {motivo} - Usuário: {usuario}"
            else:
                solicitacao.observacoes_col = f"REPROVADA INDIVIDUALMENTE - Motivo: {motivo} - Usuário: {usuario}"
        
        # Registrar log
        ip = request.remote_addr
        registrar_log(usuario, 'reprovar_solicitacao_individual', 
                     f'Solicitação ID {id} reprovada individualmente. Motivo: {motivo}', ip)
        
        db.session.commit()
        
        return jsonify({
            'success': True,
            'message': 'Solicitação reprovada com sucesso!'
        })
        
    except Exception as e:
        db.session.rollback()
        logging.error(f"Erro ao reprovar solicitação {id}: {str(e)}")
        return jsonify({
            'success': False, 
            'message': f'Erro ao reprovar solicitação: {str(e)}'
        }), 500
    
    
  # Registro do Blueprint (apenas uma vez)
app.register_blueprint(routes_bp)  

# Inicialização do banco de dados e execução do app
if __name__ == '__main__':
    create_database()
    create_auditoria_table()
    create_fornecedores_db()
    create_solicitacoes_preenchidas_table()
    create_historico_descontos_table()
    
    with app.app_context():
        db.create_all()
        
        if not atualizar_estrutura_requisicoes():
            print("Falha na migração de requisicoes. Verifique os logs.")
        
        migrate_solicitacoes_compra()
        migrate_solicitacoes_compra_status_aprovacao()
        migrate_observacoes_col()
        add_observacoes_column_to_solicitacoes_preenchidas()
        migrate_solicitacoes_preenchidas_status()
        verificar_coluna_observacoes()
        add_comprovante_pagamento_column()
        add_aplicacao_geral_column()  # NOVA MIGRAÇÃO AQUI
        add_comprador_atribuido_column()
        
        print("✓ Todas as migrações concluídas!")
    
    logging.basicConfig(
        filename='app_errors.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        encoding='utf-8'
    )
    
    app.run(debug=True, host='0.0.0.0', port=80, threaded=False, use_reloader=False)
