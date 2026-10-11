import os
import pandas as pd
from flask import session

DB_DIR = 'database_alunos'

# Lista de todos os ficheiros Excel que a app usa
EXCEL_FILES = [
    'clientes.xlsx', 'compras.xlsx', 'custos_fixos.xlsx', 'despesas.xlsx',
    'encomendas.xlsx', 'financeira_dados.xlsx', 'ingredientes.xlsx',
    'lista_compras.xlsx', 'lista-geral.xlsx', 'mao_de-obra.xlsx',
    'maodeobra.xlsx', 'materiais.xlsx', 'orçamentos.xlsx', 'receitas.xlsx'
]

def get_user_folder():
    """Retorna a pasta específica do aluno autenticado dentro de database_alunos."""
    if not os.path.exists(DB_DIR):
        os.makedirs(DB_DIR)
        
    user_email = session.get('user_email', 'visitante')
    safe_name = "".join(c if c.isalnum() else "_" for c in user_email)
    
    user_path = os.path.join(DB_DIR, safe_name)
    if not os.path.exists(user_path):
        os.makedirs(user_path)
        
    return user_path

def get_excel_path(filename):
    """Retorna o caminho do Excel na pasta do aluno. Se não existir, cria um Excel vazio/limpo."""
    user_folder = get_user_folder()
    user_file = os.path.join(user_folder, filename)
    
    if not os.path.exists(user_file):
        df = pd.DataFrame()
        df.to_excel(user_file, index=False)
        
    return user_file

def inicializar_bd_aluno():
    """Cria/inicializa todos os ficheiros Excel limpos na pasta do aluno quando ele faz login."""
    for filename in EXCEL_FILES:
        get_excel_path(filename)