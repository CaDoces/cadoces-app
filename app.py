import os
import datetime
import random
import string
import re
import pandas as pd
import sqlite3
from flask import Flask, render_template_string, session, redirect, url_for, request
from werkzeug.security import generate_password_hash, check_password_hash

# Importação dos módulos da CáDoces
from routes_receitas import receitas_bp
from routes_ingredientes import ingredientes_bp
from routes_clientes import clientes_bp
from routes_encomendas import encomendas_bp, EXCEL_ORCAMENTOS, EXCEL_ENCOMENDAS, carregar_dados
from routes_compras import compras_bp
from routes_despesas import despesas_bp
from routes_maodeobra import maodeobra_bp
from routes_suporte import suporte_bp
from routes_financeira import financeira_bp
from routes_portfolio import portfolio_bp
from database import inicializar_bd_aluno

app = Flask(__name__)
app.secret_key = 'cadoces_segredo_super_importante'

EXCEL_UTILIZADORES = 'utilizadores.xlsx'

def gerar_password_aleatoria():
    minusculas = string.ascii_lowercase
    maiusculas = string.ascii_uppercase
    digitos = string.digits
    especiais = "!@#$%^&*"
    
    pwd = [
        random.choice(minusculas),
        random.choice(maiusculas),
        random.choice(digitos),
        random.choice(especiais)
    ]
    
    todos = minusculas + maiusculas + digitos + especiais
    for _ in range(6):
        pwd.append(random.choice(todos))
        
    random.shuffle(pwd)
    return "".join(pwd)

def validar_password_forte(password):
    if len(password) < 8:
        return "A palavra-passe deve ter pelo menos 8 caracteres."
    if not re.search(r"[A-Z]", password):
        return "A palavra-passe deve conter pelo menos uma letra maiúscula."
    if not re.search(r"[a-z]", password):
        return "A palavra-passe deve conter pelo menos uma letra minúscula."
    if not re.search(r"[0-9]", password):
        return "A palavra-passe deve conter pelo menos um número."
    if not re.search(r"[!@#$%^&*(),.?\":{}|<>]", password):
        return "A palavra-passe deve conter pelo menos um caractere especial (ex: !@#$%^&*)."
    return None

def carregar_utilizadores():
    if not os.path.exists(EXCEL_UTILIZADORES):
        df = pd.DataFrame([
            {
                'Nome': 'Cátia Cardoso', 
                'Email': 'cadoces023@gmail.com', 
                'Password': generate_password_hash('1234'), 
                'Tipo': 'admin',
                'Telemovel': '351910313162'
            }
        ])
        df.to_excel(EXCEL_UTILIZADORES, index=False)
    else:
        df = pd.read_excel(EXCEL_UTILIZADORES)
        if 'catia@cadoces.com' in df['Email'].values:
            df.loc[df['Email'] == 'catia@cadoces.com', 'Email'] = 'cadoces023@gmail.com'
            df.to_excel(EXCEL_UTILIZADORES, index=False)
        if 'Telemovel' not in df.columns:
            df['Telemovel'] = ''
            df.to_excel(EXCEL_UTILIZADORES, index=False)
    
    df = pd.read_excel(EXCEL_UTILIZADORES)
    df['Nome'] = df['Nome'].astype(str)
    df['Email'] = df['Email'].astype(str)
    
    df['Telemovel'] = df['Telemovel'].fillna('').astype(str)
    df['Telemovel'] = df['Telemovel'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
    df['Telemovel'] = df['Telemovel'].replace('nan', '')
    
    return df

# Registo de todos os Blueprints
app.register_blueprint(receitas_bp)
app.register_blueprint(ingredientes_bp)
app.register_blueprint(clientes_bp)
app.register_blueprint(encomendas_bp)
app.register_blueprint(compras_bp)
app.register_blueprint(despesas_bp)
app.register_blueprint(maodeobra_bp)
app.register_blueprint(suporte_bp)
app.register_blueprint(financeira_bp)
app.register_blueprint(portfolio_bp)

@app.before_request
def verificar_autenticacao():
    rotas_livres = ['login', 'recuperar_password', 'static']
    if request.endpoint not in rotas_livres and 'user_email' not in session:
        return redirect(url_for('login'))

@app.route("/login", methods=["GET", "POST"])
def login():
    erro = None
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")
        
        df = carregar_utilizadores()
        utilizador = df[df['Email'] == email]
        
        if not utilizador.empty:
            hashed_pw = utilizador.iloc[0]['Password']
            if check_password_hash(hashed_pw, password):
                session['user_email'] = email
                session['user_name'] = utilizador.iloc[0]['Nome']
                session['user_type'] = utilizador.iloc[0]['Tipo']
                session['user_telemovel'] = str(utilizador.iloc[0].get('Telemovel', ''))
                
                inicializar_bd_aluno()
                
                return redirect(url_for('index'))
        
        erro = "Email ou palavra-passe incorretos."

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Login</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #fdfbf7; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; color: #333; padding: 15px; box-sizing: border-box; }
            .login-card { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 30px; width: 100%; max-width: 380px; box-shadow: 0 4px 10px rgba(0,0,0,0.03); box-sizing: border-box; }
            h2 { margin-top: 0; color: #111; text-align: center; margin-bottom: 25px; }
            .form-group { margin-bottom: 20px; }
            label { display: block; font-size: 13px; font-weight: bold; color: #333; margin-bottom: 6px; }
            .input-container { position: relative; display: flex; align-items: center; }
            input { width: 100%; padding: 12px; padding-right: 45px; border: 2px solid #f3e5d0; border-radius: 8px; box-sizing: border-box; font-size: 14px; outline: none; background: #fff; }
            .toggle-btn { position: absolute; right: 12px; background: none; border: none; cursor: pointer; font-size: 16px; color: #888; }
            .btn { width: 100%; background: #dfa673; color: #fff; border: none; padding: 12px; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 14px; }
            .btn:hover { background: #c8915f; }
            .error { color: #c62828; font-size: 13px; text-align: center; margin-bottom: 15px; background: #ffebee; padding: 8px; border-radius: 6px; }
            .forgot-link { display: block; text-align: right; font-size: 12px; color: #b85d19; text-decoration: none; margin-top: 6px; font-weight: bold; }
            .forgot-link:hover { text-decoration: underline; }
        </style>
    </head>
    <body>
        <div class="login-card">
            <h2>CáDoces 🧁</h2>
            {% if erro %}
                <div class="error">{{ erro }}</div>
            {% endif %}
            <form method="POST">
                <div class="form-group">
                    <label>Email</label>
                    <input type="email" name="email" required placeholder="cadoces023@gmail.com">
                </div>
                <div class="form-group">
                    <label>Palavra-passe</label>
                    <div class="input-container">
                        <input type="password" name="password" id="loginPassword" required placeholder="••••••••">
                        <button type="button" class="toggle-btn" onclick="togglePassword('loginPassword', this)">👁️</button>
                    </div>
                    <a href="/recuperar_password" class="forgot-link">Esqueceu-se da palavra-passe?</a>
                </div>
                <button type="submit" class="btn">Entrar</button>
            </form>
        </div>

        <script>
            function togglePassword(fieldId, btn) {
                var input = document.getElementById(fieldId);
                if (input.type === "password") {
                    input.type = "text";
                    btn.textContent = "👁️‍🗨️";
                } else {
                    input.type = "password";
                    btn.textContent = "👁️";
                }
            }
        </script>
    </body>
    </html>
    """, erro=erro)

@app.route("/recuperar_password", methods=["GET", "POST"])
def recuperar_password():
    mensagem = None
    erro = None
    
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        df = carregar_utilizadores()
        
        user_indices = df[df['Email'] == email].index
        if not user_indices.empty:
            idx = user_indices[0]
            nova_pass_temp = gerar_password_aleatoria()
            df.loc[idx, 'Password'] = generate_password_hash(nova_pass_temp)
            df.to_excel(EXCEL_UTILIZADORES, index=False)
            
            mensagem = f"A sua palavra-passe temporária é: <strong>{nova_pass_temp}</strong><br><br>Copie-a e utilize-a para entrar no sistema. Depois de iniciar sessão, altere-a para uma nova palavra-passe à sua escolha."
        else:
            erro = "Email não encontrado na base de dados de utilizadores."

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Reposição de Palavra-passe</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #fdfbf7; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; color: #333; padding: 15px; box-sizing: border-box; }
            .card-box { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 30px; width: 100%; max-width: 420px; box-shadow: 0 4px 10px rgba(0,0,0,0.03); box-sizing: border-box; }
            h2 { margin-top: 0; color: #111; text-align: center; margin-bottom: 15px; font-size: 20px; }
            p.sub { font-size: 13px; color: #666; text-align: center; margin-bottom: 20px; line-height: 1.4; }
            .form-group { margin-bottom: 20px; }
            label { display: block; font-size: 13px; font-weight: bold; color: #333; margin-bottom: 6px; }
            input { width: 100%; padding: 12px; border: 2px solid #f3e5d0; border-radius: 8px; box-sizing: border-box; font-size: 14px; outline: none; background: #fff; }
            .btn { width: 100%; background: #dfa673; color: #fff; border: none; padding: 12px; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 14px; }
            .btn:hover { background: #c8915f; }
            .msg { background: #e8f5e9; color: #2e7d32; padding: 12px; border-radius: 8px; margin-bottom: 20px; font-size: 13px; line-height: 1.5; border: 1px solid #c8e6c9; word-break: break-all; }
            .error { color: #c62828; font-size: 13px; text-align: center; margin-bottom: 15px; background: #ffebee; padding: 8px; border-radius: 6px; }
            .back-link { display: block; text-align: center; margin-top: 18px; color: #b85d19; text-decoration: none; font-weight: bold; font-size: 13px; }
            .back-link:hover { text-decoration: underline; }
        </style>
    </head>
    <body>
        <div class="card-box">
            <h2>Repor Palavra-passe 🔑</h2>
            <p class="sub">Introduza o email associado à sua conta para gerar uma palavra-passe temporária.</p>
            
            {% if mensagem %}
                <div class="msg">{{ mensagem | safe }}</div>
            {% endif %}
            
            {% if erro %}
                <div class="error">{{ erro }}</div>
            {% endif %}

            <form method="POST">
                <div class="form-group">
                    <label>Email de Acesso</label>
                    <input type="email" name="email" required placeholder="cadoces023@gmail.com">
                </div>
                <button type="submit" class="btn">Gerar Palavra-passe Temporária</button>
            </form>
            
            <a href="/login" class="back-link">← Voltar ao Login</a>
        </div>
    </body>
    </html>
    """, mensagem=mensagem, erro=erro)

@app.route("/alterar_password", methods=["GET", "POST"])
def alterar_password():
    if 'user_email' not in session:
        return redirect(url_for('login'))
        
    mensagem = None
    erro = None
    
    if request.method == "POST":
        pass_atual = request.form.get("password_atual")
        nova_pass = request.form.get("nova_password")
        conf_pass = request.form.get("confirmar_password")
        
        df = carregar_utilizadores()
        email_atual = session['user_email']
        user_indices = df[df['Email'] == email_atual].index
        
        if not user_indices.empty:
            idx = user_indices[0]
            hashed_pw = df.loc[idx, 'Password']
            
            erro_forte = validar_password_forte(nova_pass)
            
            if not check_password_hash(hashed_pw, pass_atual):
                erro = "A palavra-passe atual está incorreta."
            elif erro_forte:
                erro = erro_forte
            elif nova_pass != conf_pass:
                erro = "A nova palavra-passe e a confirmação não coincidem."
            else:
                df.loc[idx, 'Password'] = generate_password_hash(nova_pass)
                df.to_excel(EXCEL_UTILIZADORES, index=False)
                mensagem = "Palavra-passe alterada com sucesso!"
        else:
            erro = "Utilizador não encontrado."

    msg_html = f'<div class="msg" style="background:#e8f5e9; color:#2e7d32; padding:12px; border-radius:8px; margin-bottom:20px; font-weight:bold;">{mensagem}</div>' if mensagem else ''
    err_html = f'<div class="error" style="color:#c62828; font-size:13px; text-align:center; margin-bottom:15px; background:#ffebee; padding:8px; border-radius:6px;">{erro}</div>' if erro else ''

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Alterar Palavra-passe</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 20px; color: #333; box-sizing: border-box; }
            .container { max-width: 500px; margin: auto; }
            .card-box { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); box-sizing: border-box; }
            .form-group { margin-bottom: 18px; }
            label { display: block; font-size: 13px; font-weight: bold; color: #333; margin-bottom: 6px; }
            .input-container { position: relative; display: flex; align-items: center; }
            .form-control { width: 100%; padding: 10px; padding-right: 40px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; outline: none; box-sizing: border-box; }
            .toggle-btn { position: absolute; right: 10px; background: none; border: none; cursor: pointer; font-size: 15px; color: #888; }
            .btn-submit { width: 100%; background: #dfa673; color: #fff; border: none; padding: 12px; border-radius: 8px; cursor: pointer; font-weight: bold; font-size: 14px; }
            .btn-submit:hover { background: #c8915f; }
            .back-link { display: inline-block; margin-bottom: 15px; color: #b85d19; text-decoration: none; font-weight: bold; font-size: 14px; }
            .rules { background: #f9f0e6; border: 1px solid #f3e5d0; padding: 12px; border-radius: 8px; font-size: 12px; color: #555; margin-bottom: 15px; line-height: 1.5; }

            @media screen and (max-width: 768px) {
                body { padding: 15px; }
                .card-box { padding: 20px; }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <a href="/" class="back-link">← Voltar ao Painel Principal</a>
            <div class="card-box">
                <h2 style="margin-top:0; color:#111; font-size:20px; margin-bottom:15px;">Alterar Palavra-passe</h2>
                <div class="rules">
                    <strong>Regras de Segurança:</strong>
                    <ul style="margin: 5px 0 0 -15px; padding: 0;">
                        <li>Mínimo de 8 caracteres</li>
                        <li>Pelo menos uma letra maiúscula e uma minúscula</li>
                        <li>Pelo menos um número</li>
                        <li>Pelo menos um caractere especial (!@#$%^&*)</li>
                    </ul>
                </div>
                {{ msg_html | safe }}
                {{ err_html | safe }}
                <form method="POST">
                    <div class="form-group">
                        <label>Palavra-passe Atual</label>
                        <div class="input-container">
                            <input type="password" name="password_atual" id="passAtual" class="form-control" required placeholder="••••••••">
                            <button type="button" class="toggle-btn" onclick="togglePassword('passAtual', this)">👁️</button>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Nova Palavra-passe</label>
                        <div class="input-container">
                            <input type="password" name="nova_password" id="novaPass" class="form-control" required placeholder="••••••••">
                            <button type="button" class="toggle-btn" onclick="togglePassword('novaPass', this)">👁️</button>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Confirmar Nova Palavra-passe</label>
                        <div class="input-container">
                            <input type="password" name="confirmar_password" id="confPass" class="form-control" required placeholder="••••••••">
                            <button type="button" class="toggle-btn" onclick="togglePassword('confPass', this)">👁️</button>
                        </div>
                    </div>
                    <button type="submit" class="btn-submit">Atualizar Palavra-passe</button>
                </form>
            </div>
        </div>

        <script>
            function togglePassword(fieldId, btn) {
                var input = document.getElementById(fieldId);
                if (input.type === "password") {
                    input.type = "text";
                    btn.textContent = "👁️‍🗨️";
                } else {
                    input.type = "password";
                    btn.textContent = "👁️";
                }
            }
        </script>
    </body>
    </html>
    """, msg_html=msg_html, err_html=err_html)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route("/admin/alunos", methods=["GET", "POST"])
def gerir_alunos():
    if session.get('user_type') != 'admin':
        return "Acesso não autorizado.", 403
        
    df = carregar_utilizadores()
    mensagem = None
    
    if request.method == "POST":
        nome = request.form.get("nome")
        email = request.form.get("email")
        telemovel = request.form.get("telemovel", "").strip()
        
        if not df[df['Email'] == email].empty:
            mensagem = "⚠️ Já existe um utilizador com este email!"
        else:
            password_temporaria = gerar_password_aleatoria()
            
            novo_aluno = {
                'Nome': nome,
                'Email': email,
                'Password': generate_password_hash(password_temporaria),
                'Tipo': 'aluno',
                'Telemovel': telemovel
            }
            df = pd.concat([df, pd.DataFrame([novo_aluno])], ignore_index=True)
            df.to_excel(EXCEL_UTILIZADORES, index=False)
            
            mensagem = f"✅ Aluno adicionado com sucesso! Palavra-passe temporária gerada: <strong>{password_temporaria}</strong> (Guarde esta palavra-passe para partilhar com o aluno)."

    linhas_alunos = ""
    for idx, row in df.iterrows():
        nome_val = str(row.get('Nome', '')).replace("'", "\\'")
        email_val = str(row.get('Email', '')).replace("'", "\\'")
        tel_val = str(row.get('Telemovel', ''))
        if tel_val == 'nan' or not tel_val:
            tel_val = ''
            
        tipo_perfil = str(row.get('Tipo', 'aluno'))
        
        btn_editar = f'<button type="button" onclick="abrirModalEditarAluno({idx}, \'{nome_val}\', \'{email_val}\', \'{tel_val}\')" style="background: none; border: none; cursor: pointer; font-size: 16px;" title="Editar">✏️</button>' if tipo_perfil != 'admin' else ''
        btn_apagar = f'<a href="/admin/alunos/apagar/{idx}" onclick="return confirm(\'Remover o acesso deste aluno?\')" style="color: #c62828; text-decoration: none; font-size: 16px;" title="Apagar Acesso">🗑️</a>' if tipo_perfil != 'admin' else '<i>Protegido</i>'

        linhas_alunos += f"""
        <tr>
            <td data-label="Nome">{row.get('Nome')}</td>
            <td data-label="Email">{row.get('Email')}</td>
            <td data-label="Telemóvel">{tel_val if tel_val else '<span style="color: #999; font-style: italic;">Não definido</span>'}</td>
            <td data-label="Perfil"><span style="background: {'#e8f5e9' if tipo_perfil=='admin' else '#fff3e0'}; color: {'#2e7d32' if tipo_perfil=='admin' else '#b85d19'}; padding: 4px 8px; border-radius: 6px; font-weight: bold; font-size: 12px;">{tipo_perfil.upper()}</span></td>
            <td data-label="Ações">
                <div style="display: flex; align-items: center; gap: 12px; justify-content: flex-end;">
                    {btn_editar}
                    {btn_apagar}
                </div>
            </td>
        </tr>
        """

    msg_html = f'<div style="background: #e8f5e9; color: #2e7d32; padding: 15px; border-radius: 8px; margin-bottom: 20px; font-size: 14px; border: 1px solid #c8e6c9;">{mensagem}</div>' if mensagem else ''

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Gestão de Alunos</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; color: #333; box-sizing: border-box; }
            .container { max-width: 950px; margin: auto; }
            .card-box { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 30px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); box-sizing: border-box; }
            .card-title { font-size: 18px; font-weight: bold; color: #111; margin-bottom: 20px; }
            .form-grid { display: grid; grid-template-columns: 1fr 1fr 1fr auto; gap: 15px; align-items: flex-end; }
            .form-group label { display: block; font-size: 13px; font-weight: bold; color: #333; margin-bottom: 6px; }
            .form-control { width: 100%; padding: 10px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; outline: none; box-sizing: border-box; }
            .btn-submit { background: #dfa673; color: #fff; border: none; padding: 11px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; font-size: 14px; }
            .btn-submit:hover { background: #c8915f; }
            table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            th { background: transparent; color: #888; font-size: 12px; font-weight: bold; padding: 12px 10px; text-align: left; border-bottom: 2px solid #f3e5d0; text-transform: uppercase; }
            td { padding: 12px 10px; border-bottom: 1px solid #f3e5d0; color: #444; font-size: 14px; }
            .back-link { display: inline-block; margin-bottom: 15px; color: #b85d19; text-decoration: none; font-weight: bold; font-size: 14px; }
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4); padding-top: 60px; overflow-y: auto; box-sizing: border-box; }
            .modal-content { background-color: #fffdf9; margin: auto; padding: 30px; border: 2px solid #f3e5d0; width: 450px; max-width: 90%; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); position: relative; box-sizing: border-box; }
            .close-btn { position: absolute; right: 20px; top: 20px; font-size: 18px; cursor: pointer; color: #888; font-weight: bold; }

            @media screen and (max-width: 768px) {
                body { padding: 15px; }
                .form-grid { grid-template-columns: 1fr; }
                table, thead, tbody, th, td, tr { display: block; }
                thead tr { position: absolute; top: -9999px; left: -9999px; }
                tr { background: #fff; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 15px; padding: 15px; }
                td { border: none !important; position: relative; padding-left: 45% !important; text-align: right; padding-top: 8px; padding-bottom: 8px; }
                td::before { position: absolute; left: 15px; width: 40%; padding-right: 10px; white-space: nowrap; text-align: left; font-weight: bold; color: #888; content: attr(data-label); }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <a href="/" class="back-link">← Voltar ao Painel Principal</a>
            
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 10px;">
                <h1>Gestão de Alunos e Acessos</h1>
                <div>
                    <a href="/alterar_password" style="color: #b85d19; text-decoration: none; font-weight: bold; font-size: 14px; margin-right: 15px;">🔑 Alterar Palavra-passe</a>
                    <a href="/logout" style="color: #c62828; text-decoration: none; font-weight: bold; font-size: 14px;">Terminar Sessão 🚪</a>
                </div>
            </div>

            {{ msg_html | safe }}

            <div class="card-box">
                <div class="card-title">Adicionar Novo Aluno</div>
                <p style="font-size: 13px; color: #666; margin-top: -10px; margin-bottom: 15px;">Ao registar um novo aluno, será gerada automaticamente uma <strong>palavra-passe temporária única</strong> e associado o telemóvel para os testes de WhatsApp.</p>
                <form method="POST">
                    <div class="form-grid">
                        <div class="form-group">
                            <label>Nome do Aluno</label>
                            <input type="text" name="nome" class="form-control" required placeholder="Ex: Maria Silva">
                        </div>
                        <div class="form-group">
                            <label>Email de Acesso</label>
                            <input type="email" name="email" class="form-control" required placeholder="aluno@email.com">
                        </div>
                        <div class="form-group">
                            <label>Telemóvel (ex: 351912345678)</label>
                            <input type="text" name="telemovel" class="form-control" placeholder="3519...">
                        </div>
                        <div>
                            <button type="submit" class="btn-submit" style="width: 100%;">Gerir Acesso</button>
                        </div>
                    </div>
                </form>
            </div>

            <div class="card-box">
                <div class="card-title">Utilizadores Registados</div>
                <table>
                    <thead>
                        <tr>
                            <th>Nome</th>
                            <th>Email</th>
                            <th>Telemóvel</th>
                            <th>Tipo de Perfil</th>
                            <th style="text-align: right;">Ações</th>
                        </tr>
                    </thead>
                    <tbody>
                        {{ linhas_alunos | safe }}
                    </tbody>
                </table>
            </div>
        </div>

        <!-- MODAL EDITAR ALUNO -->
        <div id="modalEditarAluno" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModalEditarAluno()">&times;</span>
                <h3 style="margin-top:0; color:#111; font-size:20px; margin-bottom:20px;">Editar Dados do Aluno</h3>
                <form id="formEditarAluno" method="POST">
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label>Nome do Aluno</label>
                        <input type="text" name="nome" id="edit_nome" class="form-control" required>
                    </div>
                    <div class="form-group" style="margin-bottom: 15px;">
                        <label>Email de Acesso</label>
                        <input type="email" name="email" id="edit_email" class="form-control" required>
                    </div>
                    <div class="form-group" style="margin-bottom: 20px;">
                        <label>Telemóvel (ex: 351912345678)</label>
                        <input type="text" name="telemovel" id="edit_telemovel" class="form-control" placeholder="3519...">
                    </div>
                    <div style="display: flex; justify-content: flex-end; gap: 10px; flex-wrap: wrap;">
                        <button type="button" class="btn-submit" style="background: #ccc;" onclick="fecharModalEditarAluno()">Cancelar</button>
                        <button type="submit" class="btn-submit">Guardar Alterações</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function abrirModalEditarAluno(idx, nome, email, telemovel) {
                document.getElementById('formEditarAluno').action = "/admin/alunos/editar/" + idx;
                document.getElementById('edit_nome').value = nome;
                document.getElementById('edit_email').value = email;
                document.getElementById('edit_telemovel').value = telemovel === 'nan' ? '' : telemovel;
                document.getElementById('modalEditarAluno').style.display = 'block';
            }

            function fecharModalEditarAluno() {
                document.getElementById('modalEditarAluno').style.display = 'none';
            }
        </script>
    </body>
    </html>
    """, msg_html=msg_html, linhas_alunos=linhas_alunos)

@app.route("/admin/alunos/editar/<int:id_idx>", methods=["POST"])
def editar_aluno(id_idx):
    if session.get('user_type') != 'admin':
        return "Acesso não autorizado.", 403
        
    df = carregar_utilizadores()
    if id_idx in df.index and str(df.loc[id_idx, 'Tipo']) != 'admin':
        df.loc[id_idx, 'Nome'] = str(request.form.get("nome"))
        df.loc[id_idx, 'Email'] = str(request.form.get("email"))
        df.loc[id_idx, 'Telemovel'] = str(request.form.get("telemovel", "")).strip()
        df.to_excel(EXCEL_UTILIZADORES, index=False)
        
    return redirect(url_for('gerir_alunos'))

@app.route("/admin/alunos/apagar/<int:id_idx>")
def apagar_aluno(id_idx):
    if session.get('user_type') != 'admin':
        return "Acesso não autorizado.", 403
        
    df = carregar_utilizadores()
    if id_idx in df.index and str(df.loc[id_idx, 'Tipo']) != 'admin':
        df = df.drop(id_idx)
        df.to_excel(EXCEL_UTILIZADORES, index=False)
        
    return redirect(url_for('gerir_alunos'))

# --- PAINEL PRINCIPAL ---

@app.route("/")
def index():
    user_name = session.get('user_name', 'Cátia Cardoso')
    user_type = session.get('user_type', 'aluno')
    user_email = session.get('user_email', 'admin')
    
    hoje = datetime.date.today()
    limite = hoje + datetime.timedelta(days=3)
    
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'DataEntrega', 'Estado', 'Total'])
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome'])
    
    alertas_urgentes = []
    
    for _, r in df_orc.iterrows():
        d_str = str(r.get('DataEntrega', ''))
        if d_str and d_str != 'nan' and d_str != 'None':
            try:
                data_entrega = datetime.datetime.strptime(d_str[:10], '%Y-%m-%d').date()
                if hoje <= data_entrega <= limite:
                    dias_diff = (data_entrega - hoje).days
                    tempo_txt = "Hoje!" if dias_diff == 0 else ("Amanhã!" if dias_diff == 1 else f"Daqui a {dias_diff} dias")
                    alertas_urgentes.append({
                        'data': data_entrega.strftime('%d/%m/%Y'),
                        'tempo': tempo_txt,
                        'cliente': r.get('Cliente', 'Cliente'),
                        'nome': r.get('Nome', 'Orçamento'),
                        'tipo': 'Orçamento / Pendente'
                    })
            except Exception:
                pass

    for _, r in df_enc.iterrows():
        d_str = str(r.get('DataEntrega', ''))
        if d_str and d_str != 'nan' and d_str != 'None':
            try:
                data_entrega = datetime.datetime.strptime(d_str[:10], '%Y-%m-%d').date()
                if hoje <= data_entrega <= limite:
                    dias_diff = (data_entrega - hoje).days
                    tempo_txt = "Hoje!" if dias_diff == 0 else ("Amanhã!" if dias_diff == 1 else f"Daqui a {dias_diff} dias")
                    alertas_urgentes.append({
                        'data': data_entrega.strftime('%d/%m/%Y'),
                        'tempo': tempo_txt,
                        'cliente': r.get('Cliente', 'Cliente'),
                        'nome': r.get('Nome', 'Encomenda'),
                        'tipo': 'Encomenda Confirmada'
                    })
            except Exception:
                pass

    html_alertas = ""
    if alertas_urgentes:
        itens_alerta_html = ""
        for alt in alertas_urgentes:
            itens_alerta_html += f"""
            <div style="background: #fff; border-left: 4px solid #D99B6A; padding: 12px 16px; border-radius: 8px; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 4px rgba(74,59,50,0.02); flex-wrap: wrap; gap: 8px;">
                <div>
                    <span style="background: #F5EBE0; color: #4A3B32; padding: 3px 8px; border-radius: 6px; font-size: 11px; font-weight: bold; margin-right: 8px;">{alt['tempo']}</span>
                    <strong>{alt['nome']}</strong> para <em>{alt['cliente']}</em>
                </div>
                <div style="font-size: 13px; color: #8C7A6B; font-weight: 600;">
                    📅 {alt['data']} ({alt['tipo']})
                </div>
            </div>
            """
        
        html_alertas = f"""
        <div style="background: #FFFDF9; border: 1.5px solid #EEDCCE; border-radius: 16px; padding: 22px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(74,59,50,0.03);">
            <h3 style="margin: 0 0 12px 0; color: #4A3B32; font-size: 17px; display: flex; align-items: center; gap: 8px;">
                🔔 Alertas de Encomendas Próximas (Próximos 3 dias)
            </h3>
            {itens_alerta_html}
        </div>
        """

    stock_critico = 0
    try:
        conn = sqlite3.connect('cadoces.db')
        cursor = conn.execute('SELECT COUNT(*) FROM ingredientes_db WHERE user_email = ? AND (stock <= 0 OR stock_unidades <= 0)', (user_email,))
        row = cursor.fetchone()
        if row:
            stock_critico = row[0]
        conn.close()
    except Exception:
        pass

    alerta_stock_html = f"""
    <div style="background: {'#ffebee' if stock_critico > 0 else '#e8f5e9'}; border: 1px solid {'#ffcdd2' if stock_critico > 0 else '#c8e6c9'}; border-radius: 12px; padding: 16px 20px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; flex-wrap: wrap; gap: 12px;">
        <div style="display: flex; align-items: center; gap: 12px;">
            <span style="font-size: 22px;">{'⚠️' if stock_critico > 0 else '✅'}</span>
            <div>
                <strong style="color: {'#c62828' if stock_critico > 0 else '#2e7d32'}; font-size: 14px;">
                    {'Atenção ao Stock!' if stock_critico > 0 else 'Stock Atualizado!'}
                </strong>
                <div style="font-size: 13px; color: #555;">
                    {'Tens ' + str(stock_critico) + ' ingrediente(s) esgotado(s) ou com stock crítico.' if stock_critico > 0 else 'Todos os teus ingredientes têm stock suficiente.'}
                </div>
            </div>
        </div>
        {'<a href="/compras" style="background: #c62828; color: #fff; text-decoration: none; padding: 8px 14px; border-radius: 8px; font-size: 12px; font-weight: bold;">Ver Lista de Compras</a>' if stock_critico > 0 else ''}
    </div>
    """

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Painel Principal</title>
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
            :root {
                --cor-principal: #4A3B32;
                --cor-caramelo: #D99B6A;
                --cor-bege-claro: #F5EBE0;
                --borda-bege: #EEDCCE;
                --cor-fundo: #FAF8F5;
                --cor-card: #FFFFFF;
            }

            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: var(--cor-fundo);
                margin: 0;
                padding: 30px;
                color: var(--cor-principal);
                box-sizing: border-box;
            }
            .container {
                max-width: 1100px;
                margin: auto;
            }
            .header-bar {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 25px;
                background: var(--cor-card);
                border: 1px solid var(--borda-bege);
                padding: 20px 30px;
                border-radius: 16px;
                box-shadow: 0 4px 12px rgba(74,59,50,0.04);
                flex-wrap: wrap;
                gap: 15px;
            }
            .brand-area {
                display: flex;
                align-items: center;
                gap: 22px;
            }
            .brand-logo {
                width: 105px;
                height: 105px;
                border-radius: 50%;
                object-fit: cover;
                border: 3px solid var(--cor-caramelo);
                box-shadow: 0 4px 12px rgba(217,155,106,0.25);
                background-color: #FFFFFF;
            }
            .user-welcome {
                font-size: 15px;
                color: #8C7A6B;
            }
            .user-welcome strong {
                color: var(--cor-principal);
                font-size: 16px;
            }
            .top-links {
                display: flex;
                gap: 12px;
                align-items: center;
                flex-wrap: wrap;
            }
            .top-links a {
                text-decoration: none;
                font-weight: 600;
                font-size: 14px;
                padding: 10px 16px;
                border-radius: 10px;
                transition: all 0.2s;
            }
            .btn-alunos {
                background: var(--cor-bege-claro);
                color: var(--cor-principal);
                border: 1px solid var(--borda-bege);
            }
            .btn-alunos:hover {
                background: var(--borda-bege);
            }
            .btn-logout {
                color: #C0392B;
                background: #FDF3F2;
                border: 1px solid #FADBD8;
            }
            .btn-logout:hover {
                background: #FADBD8;
            }
            h1 {
                color: #2C1E1A;
                text-align: center;
                margin-bottom: 25px;
                font-size: 26px;
                font-weight: 700;
                letter-spacing: -0.5px;
            }
            h1 span {
                color: var(--cor-caramelo);
            }
            .grid {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 22px;
            }
            .card {
                background: var(--cor-card);
                border: 1px solid var(--borda-bege);
                border-radius: 16px;
                padding: 28px;
                text-decoration: none;
                color: var(--cor-principal);
                box-shadow: 0 2px 6px rgba(74,59,50,0.02);
                transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
                display: flex;
                flex-direction: column;
                justify-content: space-between;
                min-height: 140px;
                text-align: left;
                box-sizing: border-box;
            }
            .card:hover {
                transform: translateY(-4px);
                box-shadow: 0 8px 20px rgba(74,59,50,0.07);
                border-color: var(--cor-caramelo);
            }
            .card-top {
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                margin-bottom: 14px;
            }
            .card-icon {
                font-size: 26px;
                color: var(--cor-caramelo);
            }
            .card h3 {
                margin: 0 0 6px 0;
                color: #2C1E1A;
                font-size: 18px;
                font-weight: 700;
                letter-spacing: -0.3px;
            }
            .card p {
                margin: 0;
                color: #8C7A6B;
                font-size: 13.5px;
                font-weight: 400;
            }

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {
                body { padding: 15px; }
                .grid { grid-template-columns: 1fr; gap: 15px; }
                .header-bar { flex-direction: column; align-items: stretch; text-align: center; }
                .brand-area { flex-direction: column; gap: 12px; align-items: center; }
                .top-links { justify-content: center; width: 100%; }
                .top-links a { flex: 1; text-align: center; justify-content: center; }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-bar">
                <div class="brand-area">
                    <img src="{{ url_for('static', filename='logo.jpg') }}" alt="CáDoces Logo" class="brand-logo" onerror="this.style.display='none'; document.getElementById('erro-logo').style.display='inline';">
                    <span id="erro-logo" style="display:none; color: red; font-size: 12px; font-weight: bold;">⚠️ Logo não encontrado</span>
                    <div class="user-welcome">
                        Sessão de: <strong>{{ user_name }}</strong>
                    </div>
                </div>
                <div class="top-links">
                    {% if user_type == 'admin' %}
                        <a href="/admin/alunos" class="btn-alunos"><i class="fa-solid fa-users-rectangle"></i> Gerir Alunos</a>
                    {% endif %}
                    <a href="/alterar_password" class="btn-alunos"><i class="fa-solid fa-key"></i> Alterar Palavra-passe</a>
                    <a href="/logout" class="btn-logout"><i class="fa-solid fa-arrow-right-from-bracket"></i> Terminar Sessão</a>
                </div>
            </div>

            <h1>CáDoces - <span>Gestão de Cake Design</span></h1>
            
            {{ alerta_stock_html | safe }}
            {{ html_alertas | safe }}

            <div class="grid">
                <a href="/ingredientes" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Lista Geral</h3>
                            <p>Gerir ingredientes e materiais</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-boxes-stacked"></i></div>
                    </div>
                </a>
                <a href="/receitas" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Receitas</h3>
                            <p>Criar e gerir receitas</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-utensils"></i></div>
                    </div>
                </a>
                <a href="/despesas" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Despesas</h3>
                            <p>Despesas operacionais</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-receipt"></i></div>
                    </div>
                </a>
                <a href="/encomendas" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Orçamentos / Encomendas</h3>
                            <p>Orçamentos e encomendas</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-file-invoice-dollar"></i></div>
                    </div>
                </a>
                <a href="/clientes" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Clientes</h3>
                            <p>Gestão de clientes</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-address-book"></i></div>
                    </div>
                </a>
                <a href="/compras" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Lista de Compras</h3>
                            <p>Gerir lista de compras</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-cart-shopping"></i></div>
                    </div>
                </a>
                <a href="/maodeobra" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Mão de Obra</h3>
                            <p>Tarefas e custos de trabalho</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-clock-rotate-left"></i></div>
                    </div>
                </a>
                <a href="/financeira" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Análise Financeira</h3>
                            <p>Análise de custos e lucros</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-chart-pie"></i></div>
                    </div>
                </a>
                <a href="/portfolio" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Portefólio</h3>
                            <p>Galeria de trabalhos e bolos</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-images"></i></div>
                    </div>
                </a>
                <a href="/suporte" class="card">
                    <div class="card-top">
                        <div>
                            <h3>Suporte</h3>
                            <p>Reportar um problema</p>
                        </div>
                        <div class="card-icon"><i class="fa-solid fa-headset"></i></div>
                    </div>
                </a>
            </div>
        </div>
    </body>
    </html>
    """, user_name=user_name, user_type=user_type, html_alertas=html_alertas, alerta_stock_html=alerta_stock_html)

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)