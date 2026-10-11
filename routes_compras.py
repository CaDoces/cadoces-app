from flask import Blueprint, render_template_string, request, redirect, url_for, send_file, session
import pandas as pd
import os
import sqlite3
from io import BytesIO

compras_bp = Blueprint('compras_bp', __name__)

EXCEL_COMPRAS = 'lista_compras.xlsx'
DB_PATH = 'cadoces.db'

def obter_ficheiro_utilizador(nome_base):
    email = session.get('user_email', 'admin')
    if email == 'cadoces023@gmail.com':
        return nome_base
    else:
        email_limpo = email.replace('@', '_at_').replace('.', '_')
        nome, ext = os.path.splitext(nome_base)
        return f"{nome}_{email_limpo}{ext}"

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def carregar_dados(ficheiro_base, colunas_padrao):
    ficheiro = obter_ficheiro_utilizador(ficheiro_base)
    if not os.path.exists(ficheiro):
        df = pd.DataFrame(columns=colunas_padrao)
        df.to_excel(ficheiro, index=False)
    return pd.read_excel(ficheiro)

def guardar_dados(ficheiro_base, df):
    ficheiro = obter_ficheiro_utilizador(ficheiro_base)
    df.to_excel(ficheiro, index=False)

def obter_materiais_bd():
    """Busca os ingredientes/materiais disponíveis na base de dados"""
    materiais = []
    try:
        conn = get_db_connection()
        for tabela in ['ingredientes_db', 'ingredientes', 'materiais']:
            try:
                cursor = conn.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{tabela}'")
                if cursor.fetchone():
                    rows = conn.execute(f"SELECT nome FROM {tabela} ORDER BY nome").fetchall()
                    materiais = [row['nome'] for row in rows]
                    break
            except Exception:
                pass
        conn.close()
    except Exception:
        pass
    return materiais

def atualizar_material_e_stock_bd(nome_item, novo_preco, qtd_adicionada):
    """Atualiza o preço e soma a quantidade comprada ao stock na base de dados"""
    if not nome_item:
        return
    try:
        conn = get_db_connection()
        for tabela in ['ingredientes_db', 'ingredientes', 'materiais']:
            try:
                cursor = conn.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{tabela}'")
                if cursor.fetchone():
                    colunas = [col['name'] for col in conn.execute(f"PRAGMA table_info({tabela})").fetchall()]
                    
                    if novo_preco > 0:
                        for col_preco in ['valor', 'preco']:
                            if col_preco in colunas:
                                conn.execute(f"UPDATE {tabela} SET {col_preco} = ? WHERE LOWER(nome) = LOWER(?)", (float(novo_preco), nome_item.strip()))
                    
                    if qtd_adicionada > 0:
                        for col_stock in ['stock', 'quantidade', 'existencias']:
                            if col_stock in colunas:
                                conn.execute(f"UPDATE {tabela} SET {col_stock} = COALESCE({col_stock}, 0) + ? WHERE LOWER(nome) = LOWER(?)", (float(qtd_adicionada), nome_item.strip()))
                                break
                                
                    conn.commit()
            except Exception as ex:
                print(f"Erro na tabela {tabela}:", ex)
        conn.close()
    except Exception as e:
        print("Aviso ao sincronizar stock e materiais:", e)

@compras_bp.route("/compras")
def gerir_compras():
    df = carregar_dados(EXCEL_COMPRAS, ['ID', 'Item', 'Quantidade', 'Preco', 'Comprado'])
    materiais_disponiveis = obter_materiais_bd()
    
    total_itens = len(df)
    pendentes_count = 0
    comprados_count = 0
    valor_pendente = 0.0
    valor_comprado = 0.0
    valor_total = 0.0
    
    linhas_tabela = ""
    
    if not df.empty:
        for idx, row in df.iterrows():
            id_item = row.get('ID', idx)
            nome = str(row.get('Item', ''))
            qtd = float(row.get('Quantidade', 1) or 1)
            preco = float(row.get('Preco', 0) or 0)
            subtotal = qtd * preco
            comprado = bool(row.get('Comprado', False))
            
            valor_total += subtotal
            if comprado:
                comprados_count += 1
                valor_comprado += subtotal
            else:
                pendentes_count += 1
                valor_pendente += subtotal
                
            estado_checked = "checked" if comprado else ""
            
            linhas_tabela += f"""
            <tr class="linha-item" data-texto="{nome.lower()}">
                <td data-label="Estado">
                    <input type="checkbox" {estado_checked} onchange="window.location.href='/compras/toggle/{id_item}'" style="transform: scale(1.2); cursor: pointer;" title="Marcar como comprado e atualizar stock">
                </td>
                <td data-label="Item" style="font-weight: 600; color: #111; text-decoration: {'line-through' if comprado else 'none'};">{nome}</td>
                <td data-label="Qtd a Comprar">{qtd}</td>
                <td data-label="Preço Unit. (€)">
                    <form action="/compras/atualizar_preco/{id_item}" method="POST" style="display:inline; display: flex; gap: 6px; align-items: center; justify-content: flex-end;">
                        <input type="number" step="0.01" name="preco" value="{preco:.2f}" style="width: 80px; padding: 4px 8px; border: 1px solid #f3e5d0; border-radius: 6px; font-size: 13px;" />
                        <button type="submit" style="background: #dfa673; color: white; border: none; padding: 4px 8px; border-radius: 6px; cursor: pointer; font-size: 11px; font-weight: bold;" title="Guardar Preço e Atualizar BD">💾</button>
                    </form>
                </td>
                <td data-label="Subtotal" style="font-weight: bold;">€{subtotal:.2f}</td>
                <td data-label="Ações" style="text-align: right;">
                    <a href="/compras/apagar/{id_item}" style="text-decoration: none; font-size: 16px;" title="Apagar">🗑️</a>
                </td>
            </tr>
            """

    options_materiais = "".join([f'<option value="{m}">{m}</option>' for m in materiais_disponiveis])

    conteudo_principal = ""
    if df.empty:
        conteudo_principal = f"""
        <div class="empty-state">
            <div class="empty-icon">🛒</div>
            <h2>Lista de compras vazia</h2>
            <p>Adicione itens da sua lista de materiais ou crie novos para controlar os seus gastos e stock.</p>
            <button class="btn-primary" onclick="abrirModal()" style="margin: 20px auto 0 auto;">+ Adicionar Item</button>
        </div>
        """
    else:
        conteudo_principal = f"""
        <div class="content-box">
            <table class="table-custom">
                <thead>
                    <tr>
                        <th style="width: 40px;"></th>
                        <th>Item</th>
                        <th>Qtd a Comprar</th>
                        <th style="text-align: right;">Preço Unit. (€)</th>
                        <th>Subtotal</th>
                        <th style="text-align: right;">Ações</th>
                    </tr>
                </thead>
                <tbody id="tabelaComprasCorpo">
                    {linhas_tabela}
                </tbody>
            </table>
        </div>
        """

    return render_template_string(f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Lista de Compras</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; }}
            .container {{ max-width: 1200px; margin: auto; }}
            .header-bar {{ display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; gap: 15px; flex-wrap: wrap; }}
            .header-left h1 {{ margin: 0; color: #111; font-size: 24px; }}
            .header-left p {{ margin: 5px 0 0 0; color: #666; font-size: 14px; }}
            .header-actions {{ display: flex; gap: 12px; flex-wrap: wrap; }}
            .btn-export {{ background: #fff; border: 2px solid #f3e5d0; color: #333; padding: 10px 18px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; }}
            .btn-export:hover {{ background: #f9f0e6; }}
            .btn-primary {{ background: #dfa673; color: #fff; border: none; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; width: fit-content; }}
            .btn-primary:hover {{ background: #c8915f; }}
            .search-box {{ margin-bottom: 25px; }}
            .search-input {{ width: 100%; padding: 14px 20px; border: 2px solid #f3e5d0; border-radius: 10px; background: #fffdf9; font-size: 15px; box-sizing: border-box; outline: none; }}
            .search-input:focus {{ border-color: #dfa673; }}
            .cards-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px; }}
            .card-metric {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 20px 25px; }}
            .card-header {{ display: flex; justify-content: space-between; align-items: center; font-size: 14px; font-weight: bold; color: #333; margin-bottom: 12px; }}
            .card-value {{ font-size: 28px; font-weight: bold; color: #111; margin-bottom: 4px; }}
            .card-sub {{ font-size: 13px; color: #666; }}
            .empty-state {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 60px 20px; text-align: center; margin-top: 20px; }}
            .empty-icon {{ font-size: 40px; background: #fdf3e3; width: 70px; height: 70px; display: flex; align-items: center; justify-content: center; border-radius: 16px; margin: 0 auto 20px auto; }}
            .content-box {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 10px 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
            table.table-custom {{ width: 100%; border-collapse: collapse; }}
            table.table-custom th {{ background-color: transparent; color: #888; font-size: 13px; font-weight: bold; padding: 16px 12px; text-align: left; border-bottom: 2px solid #f3e5d0; text-transform: uppercase; }}
            table.table-custom td {{ padding: 16px 12px; border-bottom: 1px solid #f3e5d0; color: #444; font-size: 14px; }}
            .modal {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4); padding: 40px 0; overflow-y: auto; }}
            .modal-content {{ background-color: #fffdf9; margin: auto; padding: 30px; border: 2px solid #f3e5d0; width: 90%; max-width: 450px; border-radius: 14px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); position: relative; box-sizing: border-box; }}
            .close-btn {{ position: absolute; right: 20px; top: 20px; font-size: 22px; cursor: pointer; color: #888; font-weight: bold; }}
            .form-group {{ margin-bottom: 18px; }}
            .form-group label {{ display: block; font-weight: bold; color: #222; margin-bottom: 6px; font-size: 13px; }}
            .form-control {{ width: 100%; padding: 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; box-sizing: border-box; outline: none; }}
            .modal-footer {{ display: flex; justify-content: flex-end; gap: 12px; border-top: 2px solid #f3e5d0; padding-top: 20px; margin-top: 20px; }}
            .btn-secondary {{ background: #fff; border: 2px solid #f3e5d0; color: #444; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; }}

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {{
                body {{ padding: 15px; }}
                .header-bar {{ flex-direction: column; text-align: center; }}
                .header-actions {{ justify-content: center; }}
                .cards-grid {{ grid-template-columns: 1fr; gap: 15px; }}
                table.table-custom, table.table-custom thead, table.table-custom tbody, table.table-custom th, table.table-custom td, table.table-custom tr {{
                    display: block;
                }}
                table.table-custom thead tr {{
                    position: absolute;
                    top: -9999px;
                    left: -9999px;
                }}
                table.table-custom tr {{
                    background: #fff;
                    border: 2px solid #f3e5d0;
                    border-radius: 12px;
                    margin-bottom: 15px;
                    padding: 15px;
                }}
                table.table-custom td {{
                    border: none !important;
                    position: relative;
                    padding-left: 55% !important;
                    text-align: right;
                    padding-top: 8px;
                    padding-bottom: 8px;
                }}
                table.table-custom td::before {{
                    position: absolute;
                    left: 15px;
                    width: 45%;
                    padding-right: 10px;
                    white-space: nowrap;
                    text-align: left;
                    font-weight: bold;
                    color: #888;
                    content: attr(data-label);
                }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-bar">
                <div class="header-left">
                    <h1>Lista de Compras</h1>
                    <p>{total_itens} itens • Total: {valor_total:.2f} €</p>
                </div>
                <div class="header-actions">
                    <a href="/compras/exportar" class="btn-export">📥 Exportar</a>
                    <button class="btn-primary" onclick="abrirModal()">+ Adicionar Item</button>
                </div>
            </div>

            <div class="search-box">
                <input type="text" id="inputPesquisa" class="search-input" placeholder="🔍 Pesquisar itens..." onkeyup="filtrarItens()">
            </div>

            <div class="cards-grid">
                <div class="card-metric">
                    <div class="card-header"><span>Por Comprar</span> 🛒</div>
                    <div class="card-value">{valor_pendente:.2f} €</div>
                    <div class="card-sub">{pendentes_count} itens pendentes</div>
                </div>
                <div class="card-metric">
                    <div class="card-header"><span>Já Comprado</span> ✓</div>
                    <div class="card-value">{valor_comprado:.2f} €</div>
                    <div class="card-sub">{comprados_count} itens comprados</div>
                </div>
                <div class="card-metric">
                    <div class="card-header"><span>Total Geral</span> €</div>
                    <div class="card-value">{valor_total:.2f} €</div>
                    <div class="card-sub">{total_itens} itens na lista</div>
                </div>
            </div>

            {conteudo_principal}
        </div>

        <div id="modalItem" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModal()">&times;</span>
                <h2 style="margin-top:0; color:#111; font-size:20px;">Adicionar à Lista de Compras</h2>
                
                <form action="/compras/adicionar" method="POST">
                    <div class="form-group">
                        <label>Selecionar da Base de Dados de Materiais</label>
                        <select name="item_select" class="form-control" onchange="document.getElementById('inputItemNome').value = this.value;">
                            <option value="">-- Escolher dos materiais existentes --</option>
                            {options_materiais}
                        </select>
                    </div>

                    <div class="form-group">
                        <label>Ou Nome do Item *</label>
                        <input type="text" id="inputItemNome" name="item" class="form-control" placeholder="Ex: Farinha, Manteiga..." required>
                    </div>

                    <div class="form-group">
                        <label>Quantidade a Comprar</label>
                        <input type="number" step="0.01" name="quantidade" class="form-control" value="1">
                    </div>

                    <div class="form-group">
                        <label>Preço Unitário (€)</label>
                        <input type="number" step="0.01" name="preco" class="form-control" value="0.00">
                    </div>

                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModal()">Cancelar</button>
                        <button type="submit" class="btn-primary">Adicionar</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function abrirModal() {{ document.getElementById('modalItem').style.display = 'block'; }}
            function fecharModal() {{ document.getElementById('modalItem').style.display = 'none'; }}

            function filtrarItens() {{
                let termo = document.getElementById('inputPesquisa').value.toLowerCase();
                let linhas = document.querySelectorAll('.linha-item');
                linhas.forEach(linha => {{
                    let texto = linha.getAttribute('data-texto');
                    linha.style.display = texto.includes(termo) ? '' : 'none';
                }});
            }}
        </script>
    </body>
    </html>
    """)

@compras_bp.route("/compras/adicionar", methods=["POST"])
def adicionar_compra():
    item = request.form.get("item") or request.form.get("item_select")
    qtd = float(request.form.get("quantidade", 1) or 1)
    preco = float(request.form.get("preco", 0) or 0)
    
    if not item:
        return redirect(url_for('compras_bp.gerir_compras'))

    df = carregar_dados(EXCEL_COMPRAS, ['ID', 'Item', 'Quantidade', 'Preco', 'Comprado'])
    novo_id = int(df['ID'].max()) + 1 if not df.empty and 'ID' in df.columns and not df['ID'].dropna().empty else 1
    
    nova_linha = {
        'ID': novo_id,
        'Item': item.strip(),
        'Quantidade': qtd,
        'Preco': preco,
        'Comprado': False
    }
    df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
    guardar_dados(EXCEL_COMPRAS, df)
    
    if preco > 0:
        atualizar_material_e_stock_bd(item, preco, 0)
        
    return redirect(url_for('compras_bp.gerir_compras'))

@compras_bp.route("/compras/atualizar_preco/<int:id_item>", methods=["POST"])
def atualizar_preco_compra(id_item):
    novo_preco = float(request.form.get("preco", 0) or 0)
    df = carregar_dados(EXCEL_COMPRAS, ['ID', 'Item', 'Quantidade', 'Preco', 'Comprado'])
    
    if not df.empty:
        idx = df[df['ID'] == id_item].index
        if not idx.empty:
            nome_item = str(df.loc[idx, 'Item'].values[0])
            df.loc[idx, 'Preco'] = novo_preco
            guardar_dados(EXCEL_COMPRAS, df)
            
            atualizar_material_e_stock_bd(nome_item, novo_preco, 0)
            
    return redirect(url_for('compras_bp.gerir_compras'))

@compras_bp.route("/compras/toggle/<int:id_item>")
def toggle_compra(id_item):
    df = carregar_dados(EXCEL_COMPRAS, ['ID', 'Item', 'Quantidade', 'Preco', 'Comprado'])
    if not df.empty:
        idx = df[df['ID'] == id_item].index
        if not idx.empty:
            estado_atual = bool(df.loc[idx, 'Comprado'].values[0])
            novo_estado = not estado_atual
            df.loc[idx, 'Comprado'] = novo_estado
            guardar_dados(EXCEL_COMPRAS, df)
            
            if novo_estado:
                nome_item = str(df.loc[idx, 'Item'].values[0])
                qtd = float(df.loc[idx, 'Quantidade'].values[0] or 0)
                preco = float(df.loc[idx, 'Preco'].values[0] or 0)
                atualizar_material_e_stock_bd(nome_item, preco, qtd)
                
    return redirect(url_for('compras_bp.gerir_compras'))

@compras_bp.route("/compras/apagar/<int:id_item>")
def apagar_compra(id_item):
    df = carregar_dados(EXCEL_COMPRAS, ['ID', 'Item', 'Quantidade', 'Preco', 'Comprado'])
    if not df.empty:
        df = df[df['ID'] != id_item]
        guardar_dados(EXCEL_COMPRAS, df)
    return redirect(url_for('compras_bp.gerir_compras'))

@compras_bp.route("/compras/exportar")
def exportar_compras():
    df = carregar_dados(EXCEL_COMPRAS, ['ID', 'Item', 'Quantidade', 'Preco', 'Comprado'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Lista de Compras')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='lista_compras.xlsx')