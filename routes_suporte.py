from flask import Blueprint, render_template_string, request, redirect, url_for, session
import pandas as pd
import os
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

suporte_bp = Blueprint('suporte_bp', __name__)

EXCEL_SUPORTE_GLOBAL = 'suporte_global.xlsx'

MEU_EMAIL = "o_tue_email@gmail.com"        
PASSWORD_EMAIL = "abcd efgh ijkl mnop"    

def carregar_suporte_global():
    if not os.path.exists(EXCEL_SUPORTE_GLOBAL):
        df = pd.DataFrame(columns=['ID', 'Data', 'UserEmail', 'UserName', 'Area', 'Descricao', 'Estado'])
        df.to_excel(EXCEL_SUPORTE_GLOBAL, index=False)
    return pd.read_excel(EXCEL_SUPORTE_GLOBAL)

def guardar_suporte_global(df):
    df.to_excel(EXCEL_SUPORTE_GLOBAL, index=False)

def enviar_email_notificacao(area, descricao, data_str, autor):
    if "o_tue_email@gmail.com" in MEU_EMAIL:
        return
    try:
        msg = MIMEMultipart()
        msg['From'] = MEU_EMAIL
        msg['To'] = MEU_EMAIL
        msg['Subject'] = f"⚙️ Novo Pedido de Suporte - CáDoces ({area})"
        
        corpo = f"""
        Novo pedido de suporte submetido na aplicação:
        
        - Aluno/Utilizador: {autor}
        - Data/Hora: {data_str}
        - Área: {area}
        - Descrição:
        {descricao}
        
        ---
        App CáDoces - Sistema Local
        """
        msg.attach(MIMEText(corpo, 'plain', 'utf-8'))
        
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(MEU_EMAIL, PASSWORD_EMAIL)
        servidor.sendmail(MEU_EMAIL, MEU_EMAIL, msg.as_string())
        servidor.quit()
    except Exception as e:
        print(f"Erro ao enviar email: {e}")

@suporte_bp.route("/suporte", methods=["GET", "POST"])
def gerir_suporte():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    user_type = session.get('user_type', 'aluno')
    user_name = session.get('user_name', 'Utilizador')
    is_admin = (user_email == 'cadoces023@gmail.com' or user_type == 'admin')

    df = carregar_suporte_global()

    df_utilizador = df if is_admin else (df[df['UserEmail'] == user_email] if not df.empty else pd.DataFrame())

    linhas_pedidos = ""
    if not df_utilizador.empty:
        for idx, row in df_utilizador.iterrows():
            id_pedido = row.get('ID', idx)
            data_str = str(row.get('Data', ''))
            autor = f"👤 {row.get('UserName', 'Utilizador')} ({row.get('UserEmail', '')})" if is_admin else ""
            area = str(row.get('Area', ''))
            descricao = str(row.get('Descricao', ''))
            estado = str(row.get('Estado', 'Pendente'))
            
            if estado == 'Resolvido':
                cor_estado_bg = '#e8f5e9'
                cor_estado_txt = '#2e7d32'
            elif estado == 'Em Análise':
                cor_estado_bg = '#fff3e0'
                cor_estado_txt = '#e65100'
            else:
                cor_estado_txt = '#c62828'
                cor_estado_bg = '#ffebee'

            acoes_admin = ""
            if is_admin:
                acoes_admin = f"""
                <div style="display: flex; gap: 6px; align-items: center; margin-top: 10px; flex-wrap: wrap;">
                    <span style="font-size: 12px; color: #666; font-weight: bold;">Alterar Estado:</span>
                    <a href="/suporte/estado/{id_pedido}/Pendente" style="text-decoration: none; font-size: 11px; padding: 4px 10px; border-radius: 4px; background: #ffebee; color: #c62828; font-weight: bold;">Pendente</a>
                    <a href="/suporte/estado/{id_pedido}/Em Análise" style="text-decoration: none; font-size: 11px; padding: 4px 10px; border-radius: 4px; background: #fff3e0; color: #e65100; font-weight: bold;">Em Análise</a>
                    <a href="/suporte/estado/{id_pedido}/Resolvido" style="text-decoration: none; font-size: 11px; padding: 4px 10px; border-radius: 4px; background: #e8f5e9; color: #2e7d32; font-weight: bold;">Resolvido</a>
                </div>
                """

            linhas_pedidos += f"""
            <div class="support-ticket">
                <div style="flex: 1; word-break: break-word;">
                    {f'<div style="font-size: 13px; font-weight: bold; color: #b85d19; margin-bottom: 4px;">{autor}</div>' if autor else ''}
                    <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 6px; flex-wrap: wrap;">
                        <span style="font-size: 13px; color: #888;">{data_str}</span>
                        <strong style="color: #444; font-size: 13px;">[{area}]</strong>
                        <span style="background: {cor_estado_bg}; color: {cor_estado_txt}; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: bold;">{estado}</span>
                    </div>
                    <div style="font-size: 14px; color: #333; line-height: 1.4;">{descricao}</div>
                    {acoes_admin}
                </div>
                <div style="margin-left: 15px; align-self: flex-start;">
                    <a href="/suporte/apagar/{id_pedido}" onclick="return confirm('Tens a certeza que queres remover este pedido?')" style="text-decoration: none; font-size: 16px;" title="Apagar">🗑️</a>
                </div>
            </div>
            """
            
    if not linhas_pedidos:
        linhas_pedidos = '<div style="color: #888; font-size: 14px; padding: 5px 0;">Ainda não existem pedidos de suporte registados.</div>'

    topo_html = '<a href="/" class="back-link">← Voltar ao Painel Principal</a>' if is_admin else f'''
    <div class="session-bar">
        <span style="font-size: 14px; color: #555;">Sessão de Aluno: <strong>{user_name}</strong></span>
        <a href="/logout" style="color: #c62828; text-decoration: none; font-weight: bold; font-size: 14px;">Terminar Sessão 🚪</a>
    </div>
    '''

    titulo_seccao = "📥 Central de Suporte (Painel de Administrador)" if is_admin else "Os meus pedidos de suporte"
    subtitulo_topo = "Gestão de incidentes reportados pelos alunos" if is_admin else "Reportar um problema na aplicação"

    formulario_aluno_html = '' if is_admin else '''
    <div class="card-box">
        <div class="card-title"><span>⚙️</span> Novo pedido de suporte</div>
        <form action="/suporte/enviar" method="POST">
            <div class="form-group">
                <label>Área onde ocorreu o problema</label>
                <select name="area" class="form-control">
                    <option value="Receitas">Receitas</option>
                    <option value="Ingredientes">Ingredientes e Materiais</option>
                    <option value="Clientes">Clientes</option>
                    <option value="Encomendas">Orçamentos / Encomendas</option>
                    <option value="Compras">Lista de Compras</option>
                    <option value="Despesas">Despesas</option>
                    <option value="Mão de Obra">Mão de Obra</option>
                    <option value="Análise Financeira">Análise Financeira</option>
                    <option value="Outro">Outro</option>
                </select>
            </div>
            <div class="form-group">
                <label>Descrição do problema *</label>
                <textarea name="descricao" class="form-control form-textarea" placeholder="Descreve com o máximo de detalhe possível o que aconteceu..." required></textarea>
            </div>
            <button type="submit" class="btn-submit">✈️ Enviar pedido</button>
        </form>
    </div>
    '''

    return render_template_string(f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Suporte</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; color: #333; }}
            .container {{ max-width: 900px; margin: auto; }}
            .top-header {{ margin-bottom: 25px; }}
            .top-header h1 {{ margin: 0; color: #111; font-size: 24px; }}
            .top-header p {{ margin: 5px 0 0 0; color: #666; font-size: 14px; }}
            .card-box {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 30px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
            .card-title {{ font-size: 18px; font-weight: bold; color: #111; margin-bottom: 20px; display: flex; align-items: center; gap: 8px; }}
            .form-group {{ margin-bottom: 20px; }}
            .form-group label {{ display: block; font-size: 13px; font-weight: bold; color: #222; margin-bottom: 8px; }}
            .form-control {{ width: 100%; padding: 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; outline: none; box-sizing: border-box; }}
            .form-textarea {{ height: 130px; resize: vertical; }}
            .btn-submit {{ background: #dfa673; color: #fff; border: none; padding: 12px 22px; border-radius: 8px; cursor: pointer; font-weight: bold; font-size: 14px; display: flex; align-items: center; gap: 8px; }}
            .btn-submit:hover {{ background: #c8915f; }}
            .back-link {{ display: inline-block; margin-bottom: 15px; color: #b85d19; text-decoration: none; font-weight: bold; font-size: 14px; }}
            .back-link:hover {{ text-decoration: underline; }}
            .session-bar {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; background: #fffdf9; border: 2px solid #f3e5d0; padding: 12px 20px; border-radius: 10px; flex-wrap: wrap; gap: 10px; }}
            .support-ticket {{ background: #fff; border: 1px solid #f3e5d0; border-radius: 10px; padding: 18px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: flex-start; }}

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {{
                body {{ padding: 15px; }}
                .card-box {{ padding: 20px; }}
                .support-ticket {{ flex-direction: column; gap: 12px; }}
                .support-ticket > div:last-child {{ align-self: flex-end; margin-left: 0; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            {topo_html}
            
            <div class="top-header">
                <h1>Suporte</h1>
                <p>{subtitulo_topo}</p>
            </div>

            {formulario_aluno_html}

            <div class="card-box">
                <div class="card-title" style="margin-bottom: 15px;">{titulo_seccao}</div>
                <div>
                    {linhas_pedidos}
                </div>
            </div>
        </div>
    </body>
    </html>
    """)

@suporte_bp.route("/suporte/enviar", methods=["POST"])
def enviar_suporte():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    area = request.form.get("area")
    descricao = request.form.get("descricao")
    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")
    user_name = session.get('user_name', 'Utilizador')
    
    df = carregar_suporte_global()
    novo_id = int(df['ID'].max()) + 1 if not df.empty and 'ID' in df.columns and not df['ID'].dropna().empty else 1
    
    nova_linha = {
        'ID': novo_id,
        'Data': data_atual,
        'UserEmail': user_email,
        'UserName': user_name,
        'Area': area,
        'Descricao': descricao,
        'Estado': 'Pendente'
    }
    
    df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
    guardar_suporte_global(df)
    
    enviar_email_notificacao(area, descricao, data_atual, f"{user_name} ({user_email})")
    
    return redirect(url_for('suporte_bp.gerir_suporte'))

@suporte_bp.route("/suporte/estado/<int:id_pedido>/<novo_estado>")
def alterar_estado_suporte(id_pedido, novo_estado):
    user_email = session.get('user_email')
    if not user_email or (user_email != 'cadoces023@gmail.com' and session.get('user_type') != 'admin'):
        return redirect('/')

    df = carregar_suporte_global()
    if not df.empty and 'ID' in df.columns:
        idx = df[df['ID'] == id_pedido].index
        if not idx.empty:
            df.loc[idx, 'Estado'] = novo_estado
            guardar_suporte_global(df)
            
    return redirect(url_for('suporte_bp.gerir_suporte'))

@suporte_bp.route("/suporte/apagar/<int:id_pedido>")
def apagar_suporte(id_pedido):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    is_admin = (user_email == 'cadoces023@gmail.com' or session.get('user_type') == 'admin')
    
    df = carregar_suporte_global()
    if not df.empty and 'ID' in df.columns:
        if is_admin:
            df = df[df['ID'] != id_pedido]
        else:
            df = df[(df['ID'] != id_pedido) | (df['UserEmail'] != user_email)]
        guardar_suporte_global(df)
        
    return redirect(url_for('suporte_bp.gerir_suporte'))