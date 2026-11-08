import os
import pandas as pd
from flask import session, has_request_context
from sqlalchemy import create_engine, text
from werkzeug.security import generate_password_hash

# Lê a URL do Supabase configurada nas variáveis de ambiente do Render (com suporte a sslmode)
DATABASE_URL = os.getenv('DATABASE_URL')

if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

if DATABASE_URL:
    engine = create_engine(DATABASE_URL)
else:
    os.makedirs('database_alunos', exist_ok=True)
    engine = create_engine('sqlite:///database_alunos/local_fallback.db')

EXCEL_FILES = [
    'clientes', 'compras', 'custos_fixos', 'despesas',
    'encomendas', 'financeira_dados', 'ingredientes',
    'lista_compras', 'lista_geral', 'mao_de_obra',
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
    """Cria apenas as tabelas globais essenciais e o admin no arranque."""
    try:
        with engine.begin() as conn:
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

            conn.execute(text("""
                CREATE TABLE IF NOT EXISTS pedidos_recuperacao (
                    id SERIAL PRIMARY KEY,
                    email VARCHAR(150) NOT NULL,
                    pass_temporaria VARCHAR(50) NOT NULL,
                    data_pedido TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    estado VARCHAR(50) DEFAULT 'Pendente'
                );
            """))
            
            result = conn.execute(text("SELECT COUNT(*) FROM utilizadores WHERE email = :email"), {"email": "cadoces023@gmail.com"}).scalar()
            senha_hash = generate_password_hash("1234")
            
            if result == 0:
                conn.execute(text("""
                    INSERT INTO utilizadores (nome, email, password, tipo, telemovel) 
                    VALUES (:nome, :email, :password, :tipo, :telemovel)
                """), {
                    "nome": "Cátia Cardoso",
                    "email": "cadoces023@gmail.com",
                    "password": senha_hash,
                    "tipo": "admin",
                    "telemovel": "351910313162"
                })
                print("Administrador criado com sucesso!")
            else:
                conn.execute(text("""
                    UPDATE utilizadores SET password = :password WHERE email = :email
                """), {"password": senha_hash, "email": "cadoces023@gmail.com"})
                
    except Exception as e:
        print(f"Erro ao inicializar base de dados: {e}")

    # NOTA: Removemos daqui o ciclo dos EXCEL_FILES no arranque global.
    # Os ficheiros Excel por utilizador serão carregados dinamicamente apenas após o login.