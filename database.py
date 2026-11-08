import os
import pandas as pd
from flask import session, has_request_context
from sqlalchemy import create_engine, text
from werkzeug.security import generate_password_hash

# Lê a URL do Supabase configurada nas variáveis de ambiente do Render
DATABASE_URL = os.getenv('DATABASE_URL')

if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Garante que o sslmode=require está presente para conexões seguras
if DATABASE_URL and "sslmode=" not in DATABASE_URL:
    separator = "&" if "?" in DATABASE_URL else "?"
    DATABASE_URL = f"{DATABASE_URL}{separator}sslmode=require"

if DATABASE_URL:
    engine = create_engine(DATABASE_URL)
else:
    os.makedirs('database_alunos', exist_ok=True)
    engine = create_engine('sqlite:///database_alunos/local_fallback.db')

EXCEL_FILES = [
    'clientes', 'compras', 'custos_fixos', 'despesas',
    'encomendas', 'financeira_dados', 'ingredientes',
    * 'lista_compras', 'lista_geral', 'mao_de_obra',
    'maodeobra', 'materiais', 'orçamentos', 'receitas'
]

def get_safe_user_email():
    if has_request_context():
        user_email = session.get('user_email', 'visitante')
    else:
        user_email = 'visitante'
    return "".join(c if c.isalnum() else "_" for c in user_email)

def get_table_name(filename):
    base_name = filename.replace('.xlsx', '').replace('-', '_')
    user_suffix = get_safe_user_email()
    return f"{base_name}_{user_suffix}"

def get_excel_path(filename):
    table_name = get_table_name(filename)
    try:
        query = f"SELECT * FROM {table_name}"
        df = pd.read_sql(query, con=engine)
    except Exception:
        df = pd.DataFrame()
        try:
            df.to_sql(table_name, con=engine, if_exists='replace', index=False)
        except Exception as e:
            print(f"Erro ao criar tabela {table_name}: {e}")
    return df

def inicializar_bd_aluno():
    """Cria as tabelas globais e garante o utilizador administrador no arranque com hash seguro."""
    try:
        with engine.begin() as conn:
            # 1. Cria a tabela de utilizadores
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS utilizadores (
                    id SERIAL PRIMARY KEY,
                    nome VARCHAR(150) NOT NULL,
                    email VARCHAR(150) UNIQUE NOT NULL,
                    password VARCHAR(255) NOT NULL,
                    tipo VARCHAR(50) NOT NULL,
                    telemovel VARCHAR(50)
                );
            """))

            # 2. Cria a tabela de recuperação
            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS pedidos_recuperacao (
                    id SERIAL PRIMARY KEY,
                    email VARCHAR(150) NOT NULL,
                    pass_temporaria VARCHAR(50) NOT NULL,
                    data_pedido TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    estado VARCHAR(50) DEFAULT 'Pendente'
                );
            """))
            
            # Gera o hash dinamicamente usando a biblioteca oficial do Werkzeug
            senha_hash = generate_password_hash("1234")
            
            # Insere ou atualiza o administrador principal com o hash fresco
            conn.execute(text("""
                INSERT INTO utilizadores (nome, email, password, tipo, telemovel) 
                VALUES (:nome, :email, :password, :tipo, :telemovel)
                ON CONFLICT (email) 
                DO UPDATE SET password = EXCLUDED.password;
            """), {
                "nome": "Cátia Cardoso",
                "email": "cadoces023@gmail.com",
                "password": senha_hash,
                "tipo": "admin",
                "telemovel": "351910313162"
            })
            print("Base de dados inicializada e administrador configurado com sucesso!")
                
    except Exception as e:
        print(f"Erro ao inicializar base de dados: {e}")