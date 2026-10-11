from flask import Blueprint, render_template_string, request, redirect, url_for, send_from_directory
import pandas as pd
import os
from werkzeug.utils import secure_filename

portfolio_bp = Blueprint('portfolio_bp', __name__)

EXCEL_PORTFOLIO = 'portfolio.xlsx'
UPLOAD_FOLDER = 'static/uploads/portfolio'

# Garantir que a pasta de uploads existe
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def carregar_portfolio():
    if not os.path.exists(EXCEL_PORTFOLIO):
        df = pd.DataFrame(columns=['ID', 'Titulo', 'Cliente', 'Categoria', 'Data', 'Imagem'])
        df.to_excel(EXCEL_PORTFOLIO, index=False)
    return pd.read_excel(EXCEL_PORTFOLIO)

@portfolio_bp.route("/portfolio", methods=['GET', 'POST'])
def gerir_portfolio():
    if request.method == 'POST':
        titulo = request.form.get('titulo')
        cliente = request.form.get('cliente')
        categoria = request.form.get('categoria')
        data = request.form.get('data')
        
        file = request.files.get('imagem')
        filename = ''
        
        if file and file.filename != '':
            filename = secure_filename(file.filename)
            file.save(os.path.join(UPLOAD_FOLDER, filename))
            
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
        df.to_excel(EXCEL_PORTFOLIO, index=False)
        return redirect(url_for('portfolio_bp.gerir_portfolio'))

    df = carregar_portfolio()
    
    # Gerar os cartões HTML da galeria
    cards_html = ""
    if not df.empty:
        for _, row in df.iterrows():
            img_name = row.get('Imagem', '')
            img_url = url_for('static', filename=f'uploads/portfolio/{img_name}') if img_name else 'https://via.placeholder.com/300?text=Sem+Foto'
            cards_html += f"""
            <div style="background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.02); display: flex; flex-direction: column;">
                <img src="{img_url}" alt="{row.get('Titulo', '')}" style="width: 100%; height: 200px; object-fit: cover; border-bottom: 2px solid #f3e5d0;">
                <div style="padding: 15px; flex: 1; display: flex; flex-direction: column; justify-content: space-between;">
                    <div>
                        <span style="background: #f9f0e6; color: #8c4410; padding: 3px 8px; border-radius: 6px; font-size: 11px; font-weight: bold;">{row.get('Categoria', 'Geral')}</span>
                        <h3 style="margin: 10px 0 5px 0; color: #111; font-size: 16px;">{row.get('Titulo', 'Sem Título')}</h3>
                        <p style="margin: 0; color: #666; font-size: 13px;">Cliente: <strong>{row.get('Cliente', 'N/D')}</strong></p>
                        <p style="margin: 3px 0 0 0; color: #888; font-size: 12px;">Data: {row.get('Data', '-')}</p>
                    </div>
                </div>
            </div>
            """
    else:
        cards_html = '<div style="grid-column: span 3; text-align: center; color: #888; padding: 40px;">Ainda não existem trabalhos registados no portefólio.</div>'

    return render_template_string(f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <title>CáDoces - Portefólio</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; }}
            .container {{ max-width: 1100px; margin: auto; }}
            .header-bar {{ display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; }}
            .header-left {{ display: flex; align-items: center; gap: 20px; }}
            .back-btn {{ text-decoration: none; color: #b85d19; font-size: 22px; font-weight: bold; background: #f9f0e6; padding: 8px 14px; border-radius: 8px; }}
            h1 {{ margin: 0; color: #111; font-size: 24px; }}
            p.sub {{ margin: 5px 0 0 0; color: #666; font-size: 14px; }}
            .btn-primary {{ background: #dfa673; color: #fff; border: none; padding: 12px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; }}
            .btn-primary:hover {{ background: #c8915f; }}
            .grid-galeria {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }}
            
            .modal {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4); overflow-y: auto; padding: 40px 0; }}
            .modal-content {{ background-color: #fffdf9; margin: auto; padding: 30px; border: 2px solid #f3e5d0; width: 500px; border-radius: 14px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); position: relative; }}
            .close-btn {{ position: absolute; right: 20px; top: 20px; font-size: 18px; cursor: pointer; color: #888; font-weight: bold; background: #f9f0e6; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 1px solid #f3e5d0; }}
            .form-group {{ margin-bottom: 18px; }}
            .form-group label {{ display: block; font-weight: bold; color: #222; margin-bottom: 6px; font-size: 13px; }}
            .form-control {{ width: 100%; padding: 10px 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; box-sizing: border-box; outline: none; }}
            .modal-footer {{ display: flex; justify-content: flex-end; gap: 12px; border-top: 2px solid #f3e5d0; padding-top: 20px; margin-top: 20px; }}
            .btn-secondary {{ background: #fff; border: 2px solid #f3e5d0; color: #444; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-bar">
                <div class="header-left">
                    <a href="/" class="back-btn">&#8592;</a>
                    <div>
                        <h1>Portefólio de Trabalhos</h1>
                        <p class="sub">Galeria fotográfica dos bolos e criações do CáDoces</p>
                    </div>
                </div>
                <button class="btn-primary" onclick="abrirModal()">+ Adicionar Trabalho</button>
            </div>

            <div class="grid-galeria">
                {cards_html}
            </div>
        </div>

        <!-- MODAL ADICIONAR -->
        <div id="modalPortfolio" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModal()">&times;</span>
                <h3 style="margin-top:0; color:#8c4410; font-size:20px;">Adicionar Novo Trabalho</h3>
                <form action="/portfolio" method="POST" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>Título / Tema do Bolo *</label>
                        <input type="text" name="titulo" class="form-control" placeholder="Ex: Bolo Aniversário 3 Anos - Dinossauros" required>
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
                        <input type="file" name="imagem" class="form-control" accept="image/*" required style="padding: 6px;">
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModal()">Cancelar</button>
                        <button type="submit" class="btn-primary">Guardar Trabalho</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function abrirModal() {{ document.getElementById('modalPortfolio').style.display = 'block'; }}
            function fecharModal() {{ document.getElementById('modalPortfolio').style.display = 'none'; }}
        </script>
    </body>
    </html>
    """)