from flask import Blueprint, render_template_string, request, redirect, url_for, session
import pandas as pd
import os
from werkzeug.utils import secure_filename

portfolio_bp = Blueprint('portfolio_bp', __name__)

EXCEL_PORTFOLIO = 'portfolio.xlsx'
UPLOAD_FOLDER_BASE = 'static/uploads/portfolio'

os.makedirs(UPLOAD_FOLDER_BASE, exist_ok=True)

def obter_utilizador_info():
    email = session.get('user_email', 'admin')
    if email == 'cadoces023@gmail.com':
        return 'admin', EXCEL_PORTFOLIO, UPLOAD_FOLDER_BASE
    else:
        email_limpo = email.replace('@', '_at_').replace('.', '_')
        nome_excel = f"portfolio_{email_limpo}.xlsx"
        pasta_upload = os.path.join(UPLOAD_FOLDER_BASE, email_limpo)
        os.makedirs(pasta_upload, exist_ok=True)
        return email_limpo, nome_excel, pasta_upload

def carregar_portfolio():
    _, excel_path, _ = obter_utilizador_info()
    if not os.path.exists(excel_path):
        df = pd.DataFrame(columns=['ID', 'Titulo', 'Cliente', 'Categoria', 'Data', 'Imagem'])
        df.to_excel(excel_path, index=False)
    return pd.read_excel(excel_path)

@portfolio_bp.route("/portfolio", methods=['GET', 'POST'])
def gerir_portfolio():
    if not session.get('user_email'):
        return redirect('/')

    email_limpo, excel_path, pasta_upload = obter_utilizador_info()

    if request.method == 'POST':
        titulo = request.form.get('titulo')
        cliente = request.form.get('cliente')
        categoria = request.form.get('categoria')
        data = request.form.get('data')
        
        file = request.files.get('imagem')
        filename = ''
        
        if file and file.filename != '':
            filename = secure_filename(file.filename)
            file.save(os.path.join(pasta_upload, filename))
            
        df = carregar_portfolio()
        novo_id = int(df['ID'].max() + 1) if not df.empty and pd.notna(df['ID'].max()) else 1
        
        novo_registo = pd.DataFrame([{
            'ID': novo_id,
            'Titulo': titulo,
            'Cliente': cliente,
            'Categoria': categoria,
            'Data': data,
            'Imagem': filename
        }])
        
        df = pd.concat([df, novo_registo], ignore_index=True)
        df.to_excel(excel_path, index=False)
        return redirect(url_for('portfolio_bp.gerir_portfolio'))

    df = carregar_portfolio()
    categoria_selecionada = request.args.get('cat', 'Todos')
    
    if categoria_selecionada != 'Todos' and not df.empty:
        df_filtrado = df[df['Categoria'].astype(str) == categoria_selecionada]
    else:
        df_filtrado = df
    
    cards_html = ""
    if not df_filtrado.empty:
        for _, row in df_filtrado.iterrows():
            img_name = row.get('Imagem', '')
            if img_name:
                if email_limpo == 'admin':
                    img_url = url_for('static', filename=f'uploads/portfolio/{img_name}')
                else:
                    img_url = url_for('static', filename=f'uploads/portfolio/{email_limpo}/{img_name}')
            else:
                img_url = 'https://via.placeholder.com/300?text=Sem+Foto'
                
            cards_html += f"""
            <div style="background: #FFFFFF; border: 1px solid #EEDCCE; border-radius: 16px; overflow: hidden; box-shadow: 0 2px 6px rgba(74,59,50,0.02); display: flex; flex-direction: column;">
                <img src="{img_url}" alt="{row.get('Titulo', '')}" style="width: 100%; height: 200px; object-fit: cover; border-bottom: 1px solid #EEDCCE;">
                <div style="padding: 20px; flex: 1; display: flex; flex-direction: column; justify-content: space-between;">
                    <div>
                        <span style="background: #F5EBE0; color: #4A3B32; padding: 4px 10px; border-radius: 8px; font-size: 11.5px; font-weight: 600;">{row.get('Categoria', 'Geral')}</span>
                        <h3 style="margin: 12px 0 6px 0; color: #2C1E1A; font-size: 16px; font-weight: 700;">{row.get('Titulo', 'Sem Título')}</h3>
                        <p style="margin: 0; color: #8C7A6B; font-size: 13.5px;">Cliente: <strong>{row.get('Cliente', 'N/D')}</strong></p>
                        <p style="margin: 4px 0 0 0; color: #A09080; font-size: 12.5px;">Data: {row.get('Data', '-')}</p>
                    </div>
                </div>
            </div>
            """
    else:
        cards_html = '<div style="grid-column: 1 / -1; text-align: center; color: #8C7A6B; padding: 40px; background: #FFFFFF; border: 1px solid #EEDCCE; border-radius: 16px;">Ainda não existem trabalhos registados nesta categoria.</div>'

    categorias = ['Todos', 'Aniversário', 'Casamento', 'Batizado', 'Massa Levedada', 'Outro']
    botoes_filtro_html = ""
    for cat in categorias:
        ativo = 'background: var(--cor-caramelo); color: white; border-color: var(--cor-caramelo);' if categoria_selecionada == cat else 'background: #FFFFFF; color: var(--cor-principal); border-color: var(--borda-bege);'
        botoes_filtro_html += f"""
        <a href="/portfolio?cat={cat}" style="text-decoration: none; padding: 8px 16px; border-radius: 20px; font-weight: 600; font-size: 13.5px; border: 1px solid; {ativo} transition: all 0.2s;">{cat}</a>
        """

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Portefólio</title>
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
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: var(--cor-fundo); margin: 0; padding: 30px; color: var(--cor-principal); }
            .container { max-width: 1100px; margin: auto; }
            .header-bar { display: flex; justify-content: space-between; align-items: center; background: var(--cor-card); border: 1px solid var(--borda-bege); padding: 20px 30px; border-radius: 16px; margin-bottom: 25px; box-shadow: 0 4px 12px rgba(74,59,50,0.04); flex-wrap: wrap; gap: 15px; }
            .header-left { display: flex; align-items: center; gap: 20px; }
            .back-btn { text-decoration: none; color: var(--cor-caramelo); font-size: 20px; font-weight: bold; background: var(--cor-bege-claro); padding: 10px 16px; border-radius: 10px; border: 1px solid var(--borda-bege); }
            h1 { margin: 0; color: #2C1E1A; font-size: 24px; font-weight: 700; }
            p.sub { margin: 5px 0 0 0; color: #8C7A6B; font-size: 14px; }
            .btn-primary { background: var(--cor-caramelo); color: #fff; border: none; padding: 12px 20px; border-radius: 10px; cursor: pointer; font-weight: 600; text-decoration: none; display: flex; align-items: center; gap: 8px; font-size: 14px; transition: background 0.2s; }
            .btn-primary:hover { background: #c58857; }
            
            .filtros-bar { display: flex; gap: 10px; margin-bottom: 30px; flex-wrap: wrap; align-items: center; }
            
            .grid-galeria { display: grid; grid-template-columns: repeat(3, 1fr); gap: 22px; }
            
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(74,59,50,0.4); overflow-y: auto; padding: 40px 0; }
            .modal-content { background-color: var(--cor-card); margin: auto; padding: 30px; border: 1px solid var(--borda-bege); width: 450px; max-width: 90%; border-radius: 16px; box-shadow: 0 8px 24px rgba(74,59,50,0.1); position: relative; box-sizing: border-box; }
            .close-btn { position: absolute; right: 20px; top: 20px; font-size: 16px; cursor: pointer; color: #8C7A6B; font-weight: bold; background: var(--cor-bege-claro); width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 1px solid var(--borda-bege); }
            .form-group { margin-bottom: 18px; }
            .form-group label { display: block; font-weight: 600; color: #2C1E1A; margin-bottom: 6px; font-size: 13.5px; }
            .form-control { width: 100%; padding: 10px 14px; border: 1px solid var(--borda-bege); border-radius: 10px; background: #FAF8F5; font-size: 14px; box-sizing: border-box; outline: none; }
            .form-control:focus { border-color: var(--cor-caramelo); background: #FFFFFF; }
            .modal-footer { display: flex; justify-content: flex-end; gap: 12px; border-top: 1px solid var(--borda-bege); padding-top: 20px; margin-top: 20px; flex-wrap: wrap; }
            .btn-secondary { background: #FFFFFF; border: 1px solid var(--borda-bege); color: var(--cor-principal); padding: 10px 20px; border-radius: 10px; cursor: pointer; font-weight: 600; }

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {
                body { padding: 15px; }
                .grid-galeria { grid-template-columns: 1fr; gap: 15px; }
                .header-bar { flex-direction: column; align-items: stretch; text-align: left; }
                .header-left { flex-direction: column; align-items: flex-start; gap: 12px; }
                .btn-primary { width: 100%; justify-content: center; }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-bar">
                <div class="header-left">
                    <a href="/" class="back-btn"><i class="fa-solid fa-arrow-left"></i></a>
                    <div>
                        <h1>Portefólio de Trabalhos</h1>
                        <p class="sub">Galeria fotográfica dos bolos e criações do CáDoces</p>
                    </div>
                </div>
                <button class="btn-primary" onclick="abrirModal()"><i class="fa-solid fa-plus"></i> Adicionar Trabalho</button>
            </div>

            <div class="filtros-bar">
                <span style="font-size: 13.5px; font-weight: 600; color: #8C7A6B; margin-right: 5px;">Filtrar por:</span>
                {{ botoes_filtro_html | safe }}
            </div>

            <div class="grid-galeria">
                {{ cards_html | safe }}
            </div>
        </div>

        <div id="modalPortfolio" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModal()">&times;</span>
                <h3 style="margin-top:0; color:#2C1E1A; font-size:20px; font-weight:700;">Adicionar Novo Trabalho</h3>
                <form action="/portfolio" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Título / Tema do Bolo *</label>
                        <input type="text" name="titulo" class="form-control" placeholder="Ex: Bolo Aniversário - Dinossauros" required>
                    </div>
                    <div class="form-group">
                        <label>Cliente</label>
                        <input type="text" name="cliente" class="form-control" placeholder="Nome do cliente...">
                    </div>
                    <div class="form-group">
                        <label>Categoria</label>
                        <select name="categoria" class="form-control">
                            <option value="Aniversário">Aniversário</option>
                            <option value="Casamento">Casamento</option>
                            <option value="Batizado">Batizado</option>
                            <option value="Massa Levedada">Massa Levedada</option>
                            <option value="Outro">Outro</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Data de Realização</label>
                        <input type="date" name="data" class="form-control">
                    </div>
                    <div class="form-group">
                        <label>Fotografia *</label>
                        <input type="file" name="imagem" class="form-control" accept="image/*" required style="padding: 7px; background: #FFFFFF;">
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModal()">Cancelar</button>
                        <button type="submit" class="btn-primary">Guardar Trabalho</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function abrirModal() { document.getElementById('modalPortfolio').style.display = 'block'; }
            function fecharModal() { document.getElementById('modalPortfolio').style.display = 'none'; }
        </script>
    </body>
    </html>
    """, cards_html=cards_html, botoes_filtro_html=botoes_filtro_html)