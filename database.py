import os
import pandas as pd
from flask import session
from sqlalchemy import create_engine, text

# Lê a URL do Supabase configurada nas variáveis de ambiente do Render
DATABASE_URL = os.getenv('DATABASE_URL')

# Ajuste automático caso a URL comece com postgres:// (compatibilidade com SQLAlchemy)
if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Cria o motor de ligação ao PostgreSQL (Supabase) se a variável existir, senão usa fallback local para testes
if DATABASE_URL:
    engine = create_engine(DATABASE_URL)
else:
    # Fallback local SQLite para desenvolvimento no teu PC se não houver DATABASE_URL
    os.makedirs('database_alunos', exist_ok=True)
    engine = create_engine('sqlite:///database_alunos/local_fallback.db')

# Lista de todos os módulos/tabelas que a app usa
EXCEL_FILES = [
    'clientes', 'compras', 'custos_fixos', 'despesas',
    'encomendas', 'financeira_dados', 'ingredientes',
    'lista_compras', 'lista_geral', 'mao_de_obra',
    'maodeobra', 'materiais', 'orçamentos', 'receitas'
]

def get_safe_user_email():
    """Retorna o email limpo do aluno autenticado ou 'visitante'."""
    user_email = session.get('user_email', 'visitante')
    return "".join(c if c.isalnum() else "_" for c in user_email)

def get_table_name(filename):
    """Converte o nome do ficheiro Excel (ex: 'clientes.xlsx') num nome de tabela seguro (ex: 'clientes')."""
    base_name = filename.replace('.xlsx', '').replace('-', '_')
    user_suffix = get_safe_user_email()
    # Exemplo de nome de tabela na BD: clientes_catia_gmail_com
    return f"{base_name}_{user_suffix}"

def get_excel_path(filename):
    """
    Mantém compatibilidade com funções antigas: 
    Lê os dados da tabela do utilizador no Supabase e devolve um DataFrame do Pandas.
    Se a tabela não existir na base de dados, cria-a vazia.
    """
    table_name = get_table_name(filename)
    
    try:
        # Tenta ler a tabela do utilizador do Supabase
        query = f"SELECT * FROM {table_name}"
        df = pd.read_sql(query, con=engine)
    except Exception:
        # Se a tabela não existir, cria um DataFrame vazio e grava na base de dados
        df = pd.DataFrame()
        try:
            df.to_sql(table_name, con=engine, if_exists='replace', index=False)
        except Exception as e:
            print(f"Erro ao criar tabela {table_name}: {e}")
            
    return df

def inicializar_bd_aluno():
    """Inicializa todas as tabelas necessárias para o aluno autenticado quando faz login."""
    for filename in EXCEL_FILES:
        get_excel_path(filename)