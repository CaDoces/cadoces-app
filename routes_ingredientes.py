import os
import pandas as pd
import sqlite3
from flask import Blueprint, redirect, render_template_string, request, send_file, url_for, session

try:
    from database import get_excel_path
except ImportError:
    def get_excel_path(filename):
        return filename

SHEET_NAME = "Ingredientes"

ingredientes_bp = Blueprint('ingredientes_bp', __name__)

def get_db_connection():
    conn = sqlite3.connect('cadoces.db')
    conn.row_factory = sqlite3.Row
    return conn

def parse_val(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace('€', '').strip()
    if '.' in val_str and ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def carregar_dados_db(user_email):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ingredientes_db (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_email TEXT,
            nome TEXT,
            unidade TEXT,
            quantidade REAL,
            valor REAL,
            stock REAL,
            stock_unidades REAL DEFAULT 0,
            alergenios TEXT DEFAULT ''
        )
    ''')
    try:
        cursor.execute("ALTER TABLE ingredientes_db ADD COLUMN stock_unidades REAL DEFAULT 0")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    try:
        cursor.execute("ALTER TABLE ingredientes_db ADD COLUMN alergenios TEXT DEFAULT ''")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    cursor.execute("SELECT id as original_index, nome, unidade, quantidade, valor, stock, stock_unidades, alergenios FROM ingredientes_db WHERE user_email = ?", (user_email,))
    rows = cursor.fetchall()
    conn.close()
    
    return [dict(row) for row in rows]

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CáDoces - Gestão de Ingredientes</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAF8F5; color: #4A3B32; margin: 0; padding: 30px; }
        .container { max-width: 1200px; margin: auto; }
        .header-bar { display: flex; justify-content: space-between; align-items: center; background: #FFFFFF; padding: 20px 25px; border: 1px solid #EEDCCE; border-radius: 14px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); flex-wrap: wrap; gap: 15px; }
        .header-left { display: flex; align-items: center; gap: 20px; }
        .back-btn { text-decoration: none; color: #4A3B32; font-size: 18px; font-weight: bold; background: #F5EBE0; padding: 10px 16px; border-radius: 10px; border: 1px solid #EEDCCE; transition: background 0.2s; display: inline-flex; align-items: center; justify-content: center; }
        .back-btn:hover { background: #EEDCCE; }
        h1 { margin: 0; color: #2C1E1A; font-size: 24px; font-weight: 700; }
        p.sub { margin: 4px 0 0 0; color: #8C7A6B; font-size: 14px; }
        .top-actions { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; position: relative; }
        .btn-outline { background: #F5EBE0; border: 1px solid #EEDCCE; color: #4A3B32; padding: 10px 16px; border-radius: 10px; cursor: pointer; font-weight: 600; display: inline-flex; align-items: center; gap: 6px; text-decoration: none; font-size: 14px; transition: all 0.2s; }
        .btn-outline:hover { background: #EEDCCE; }
        .btn-primary { background: #D99B6A; color: #FFFFFF; border: none; padding: 11px 20px; border-radius: 10px; cursor: pointer; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; font-size: 14px; transition: background 0.2s; box-shadow: 0 2px 4px rgba(217,155,106,0.2); }
        .btn-primary:hover { background: #C48757; }
        .dropdown { position: relative; display: inline-block; }
        .dropdown-content { display: none; position: absolute; right: 0; background-color: #FFFFFF; min-width: 160px; box-shadow: 0px 8px 16px rgba(74,59,50,0.1); border-radius: 10px; border: 1px solid #EEDCCE; z-index: 10; overflow: hidden; }
        .dropdown-content a { color: #4A3B32; padding: 10px 14px; text-decoration: none; display: block; font-size: 13px; font-weight: 500; }
        .dropdown-content a:hover { background-color: #FAF8F5; color: #D99B6A; }
        .dropdown.active .dropdown-content { display: block; }
        .search-bar { width: 100%; padding: 12px 16px; border: 1px solid #EEDCCE; border-radius: 10px; font-size: 14px; margin-top: 25px; margin-bottom: 25px; box-sizing: border-box; background: #FFFFFF; color: #4A3B32; box-shadow: 0 2px 6px rgba(74,59,50,0.02); }
        .search-bar:focus { outline: none; border-color: #D99B6A; }
        .content-box { background: #FFFFFF; border: 1px solid #EEDCCE; border-radius: 14px; padding: 25px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); margin-top: 25px; }
        table.table-custom { width: 100%; border-collapse: collapse; margin-top: 5px; }
        table.table-custom th { background-color: #FBF7F2; color: #734D2E; padding: 14px 16px; text-align: left; border-bottom: 2px solid #EEDCCE; font-size: 14px; font-weight: 600; }
        table.table-custom td { padding: 14px 16px; border-bottom: 1px solid #F5EBE0; color: #4A3B32; font-size: 14px; vertical-align: middle; }
        table.table-custom tr:hover td { background-color: #FAF8F5; }
        .cost-val { color: #D99B6A; font-weight: 600; }
        .stock-verde { color: #27AE60; font-weight: bold; }
        .stock-vermelho { color: #C0392B; font-weight: bold; }
        .action-icons { display: flex; gap: 12px; align-items: center; }
        .action-btn { background: none; border: none; cursor: pointer; font-size: 16px; color: #8C7A6B; padding: 0; text-decoration: none; transition: color 0.2s; }
        .action-btn:hover { color: #4A3B32; }
        .action-btn.delete:hover { color: #C0392B; }
        .form-grid { display: grid; grid-template-columns: 2fr 1fr 1fr 1fr 1fr 1.5fr auto; gap: 15px; background: #FAF8F5; padding: 20px; border-radius: 12px; border: 1px solid #EEDCCE; margin-bottom: 25px; align-items: end; }
        .form-group label { display: block; font-size: 12px; margin-bottom: 6px; color: #734D2E; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .form-group input { width: 100%; padding: 10px 12px; border: 1px solid #EEDCCE; border-radius: 8px; box-sizing: border-box; background: #FFFFFF; color: #4A3B32; font-size: 14px; }
        .form-group input:focus { outline: none; border-color: #D99B6A; }
        .import-box { background: #F5EBE0; padding: 20px; border-radius: 12px; border: 1px solid #EEDCCE; margin-bottom: 25px; display: none; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 15px; }
        .select-col { display: none; width: 30px; text-align: center; }
        .unit-stock-control { display: inline-flex; align-items: center; gap: 6px; background: #FBF7F2; border: 1px solid #EEDCCE; padding: 4px 8px; border-radius: 8px; }
        .unit-btn { background: #FFFFFF; border: 1px solid #EEDCCE; color: #4A3B32; width: 22px; height: 22px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 12px; display: inline-flex; align-items: center; justify-content: center; transition: background 0.2s; }
        .unit-btn:hover { background: #EEDCCE; color: #2C1E1A; }

        /* Estilos Mobile-First / Responsivo para Telemóveis */
        @media screen and (max-width: 768px) {
            body { padding: 15px; }
            .form-grid { grid-template-columns: 1fr; gap: 15px; }
            table.table-custom, table.table-custom thead, table.table-custom tbody, table.table-custom th, table.table-custom td, table.table-custom tr {
                display: block;
            }
            table.table-custom thead tr {
                position: absolute;
                top: -9999px;
                left: -9999px;
            }
            table.table-custom tr {
                background: #FFFFFF;
                border: 1px solid #EEDCCE;
                border-radius: 12px;
                margin-bottom: 15px;
                padding: 15px;
                box-shadow: 0 2px 4px rgba(74,59,50,0.02);
            }
            table.table-custom td {
                border: none !important;
                position: relative;
                padding-left: 45% !important;
                text-align: right;
                padding-top: 8px;
                padding-bottom: 8px;
            }
            table.table-custom td::before {
                position: absolute;
                left: 15px;
                width: 40%;
                padding-right: 10px;
                white-space: nowrap;
                text-align: left;
                font-weight: 600;
                color: #734D2E;
                content: attr(data-label);
            }
            .action-icons {
                justify-content: flex-end;
            }
        }
    </style>
    <script>
        let selectionMode = false;
        function toggleSelectionMode() {
            selectionMode = !selectionMode;
            let cols = document.querySelectorAll('.select-col');
            cols.forEach(el => el.style.display = selectionMode ? 'table-cell' : 'none');
            let btn = document.getElementById('btn-selecionar');
            btn.style.background = selectionMode ? '#EEDCCE' : '#F5EBE0';
            btn.style.color = selectionMode ? '#2C1E1A' : '#4A3B32';
        }
        function toggleDropdown(id) {
            let drop = document.getElementById(id);
            drop.classList.toggle('active');
        }
        window.onclick = function(event) {
            if (!event.target.matches('.btn-outline') && !event.target.closest('.dropdown')) {
                document.querySelectorAll('.dropdown').forEach(d => d.classList.remove('active'));
            }
        }
        function filtrarTabela() {
            let input = document.getElementById('searchInput').value.toLowerCase();
            let rows = document.querySelectorAll('#tabelaIngredientes tbody tr');
            rows.forEach(row => {
                let nome = row.cells[2] ? row.cells[2].textContent.toLowerCase() : '';
                if (nome.includes(input) || input === '') {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }
    </script>
</head>
<body>
<div class="container">
    <div class="header-bar">
        <div class="header-left">
            <a href="/" class="back-btn">&#8592;</a>
            <div>
                <h1>Ingredientes e Materiais</h1>
                <p class="sub">Gestão integrada de stock por unidades e custos</p>
            </div>
        </div>
        <div class="top-actions">
            <a href="{{ url_for('ingredientes_bp.lista_compras') }}" class="btn-outline" style="background: #FDF3E3; border-color: #F3E5D0; color: #B85D19;">🛒 Lista de Compras</a>
            <button id="btn-selecionar" onclick="toggleSelectionMode()" class="btn-outline">🗂 Selecionar</button>
            <div class="dropdown" id="dropdownOrdenar">
                <button onclick="toggleDropdown('dropdownOrdenar')" class="btn-outline">↕ Ordenar</button>
                <div class="dropdown-content">
                    <a href="{{ url_for('ingredientes_bp.ingredientes', ordenar='az') }}">Nome A-Z</a>
                    <a href="{{ url_for('ingredientes_bp.ingredientes', ordenar='za') }}">Nome Z-A</a>
                    <a href="{{ url_for('ingredientes_bp.ingredientes', ordenar='qtd_asc') }}">Quantidade ↑</a>
                    <a href="{{ url_for('ingredientes_bp.ingredientes', ordenar='qtd_desc') }}">Quantidade ↓</a>
                </div>
            </div>
            <a href="#" onclick="let box=document.getElementById('import-box'); box.style.display = box.style.display === 'flex' ? 'none' : 'flex'; return false;" class="btn-outline">📤 Importar/Exportar</a>
            <a href="#" onclick="window.print(); return false;" class="btn-outline">🖨 Imprimir</a>
            <a href="/" class="btn-outline">🏠 Menu</a>
            <a href="#form-adicionar" onclick="document.getElementById('form-adicionar').style.display='grid';" class="btn-primary">+ Adicionar</a>
        </div>
    </div>

    {% if pagina == 'lista_compras' %}
        <div class="content-box">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 15px;">
                <div>
                    <h2 style="margin: 0; color: #2C1E1A; font-size: 20px;">Lista de Compras Recomendada</h2>
                    <p class="sub">Ingredientes com stock esgotado ou inferior a 1 unidade</p>
                </div>
                <a href="{{ url_for('ingredientes_bp.ingredientes') }}" class="btn-outline">← Voltar ao Stock</a>
            </div>
            <table class="table-custom">
                <thead>
                    <tr>
                        <th>Ingrediente / Material</th>
                        <th>Stock Atual</th>
                        <th>Estado</th>
                        <th>Preço Estimado</th>
                    </tr>
                </thead>
                <tbody>
                    {% for item in lista_compras %}
                    <tr>
                        <td data-label="Ingrediente"><strong>{{ item.nome }}</strong></td>
                        <td data-label="Stock Atual">{{ item.stock_unidades|int if item.stock_unidades == item.stock_unidades|int else item.stock_unidades }} {{ item.unidade }} ({{ item.stock|int if item.stock == item.stock|int else item.stock }}{{ item.unidade }})</td>
                        <td data-label="Estado"><span style="background: #FDEDEC; color: #C0392B; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 12px;">Esgotado / Baixo</span></td>
                        <td data-label="Preço Estimado" class="cost-val">€{{ "%.2f"|format(item.valor) }}</td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="4" style="text-align: center; color: #27AE60; padding: 30px; font-weight: bold;">🎉 Está tudo em ordem! Nenhum ingrediente em falta no stock.</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    {% elif pagina == 'editar' %}
        <div class="content-box">
            <h2>Editar Item</h2>
            <p class="sub" style="margin-bottom: 20px;">Atualize os dados do ingrediente ou material</p>
            <form action="{{ url_for('ingredientes_bp.atualizar', index=index) }}" method="POST" class="form-grid">
                <div class="form-group">
                    <label>Nome</label>
                    <input type="text" name="nome" value="{{ item.nome }}" required>
                </div>
                <div class="form-group">
                    <label>Unidade</label>
                    <input type="text" name="unidade" value="{{ item.unidade }}" required>
                </div>
                <div class="form-group">
                    <label>Qtd. Compra</label>
                    <input type="text" name="quantidade" value="{{ item.quantidade }}" required>
                </div>
                <div class="form-group">
                    <label>Valor (€)</label>
                    <input type="text" name="valor" value="{{ item.valor }}" required>
                </div>
                <div class="form-group">
                    <label>Stock Unid.</label>
                    <input type="text" name="stock_unidades" value="{{ item.stock_unidades }}" required>
                </div>
                <div class="form-group">
                    <label>Alergénios</label>
                    <input type="text" name="alergenios" value="{{ item.alergenios }}" placeholder="Ex: Glúten, Ovos">
                </div>
                <div class="form-group">
                    <button type="submit" class="btn-primary" style="width: 100%; justify-content: center;">Atualizar</button>
                </div>
            </form>
        </div>
    {% elif pagina == 'ingredientes' %}
        <input type="text" id="searchInput" onkeyup="filtrarTabela()" placeholder="🔍 Pesquisar ingrediente ou material..." class="search-bar">

        <div id="import-box" class="import-box">
            <form action="{{ url_for('ingredientes_bp.importar_excel') }}" method="POST" enctype="multipart/form-data" style="display: flex; align-items: center; gap: 15px; margin: 0; flex-wrap: wrap;">
                <div>
                    <label style="font-size: 12px; font-weight: 600; color: #734D2E; display: block; margin-bottom: 4px; text-transform: uppercase;">Importar Ficheiro Excel (.xlsx ou .csv)</label>
                    <input type="file" name="ficheiro" accept=".xlsx, .xls, .csv" required style="font-size: 13px; color: #4A3B32;">
                </div>
                <button type="submit" class="btn-primary" style="padding: 10px 16px;">📤 Enviar e Importar</button>
            </form>
            <div>
                <a href="{{ url_for('ingredientes_bp.exportar_excel') }}" class="btn-outline" style="background: #FFFFFF;">📥 Exportar Excel</a>
            </div>
        </div>

        <form id="form-adicionar" action="{{ url_for('ingredientes_bp.adicionar') }}" method="POST" class="form-grid" style="display: none;">
            <div class="form-group">
                <label>Nome</label>
                <input type="text" name="nome" placeholder="Ex: Farinha T55" required>
            </div>
            <div class="form-group">
                <label>Unidade</label>
                <input type="text" name="unidade" placeholder="Ex: g, kg, ml" required>
            </div>
            <div class="form-group">
                <label>Qtd. Compra</label>
                <input type="text" name="quantidade" placeholder="1000" required>
            </div>
            <div class="form-group">
                <label>Valor (€)</label>
                <input type="text" name="valor" placeholder="1,50" required>
            </div>
            <div class="form-group">
                <label>Stock Unid.</label>
                <input type="text" name="stock_unidades" placeholder="Ex: 2" required>
            </div>
            <div class="form-group">
                <label>Alergénios</label>
                <input type="text" name="alergenios" placeholder="Ex: Glúten, Lactose">
            </div>
            <div class="form-group">
                <button type="submit" class="btn-primary" style="width: 100%; justify-content: center;">Guardar</button>
            </div>
        </form>

        <div class="content-box">
            <table id="tabelaIngredientes" class="table-custom">
                <thead>
                    <tr>
                        <th class="select-col"></th>
                        <th>Unidade</th>
                        <th>Nome</th>
                        <th>Compra</th>
                        <th>Stock Qtd. (Total)</th>
                        <th>Stock Unidades</th>
                        <th>Alergénios</th>
                        <th>Custo/Un.</th>
                        <th>Ações</th>
                    </tr>
                </thead>
                <tbody>
                    {% for item in lista %}
                    <tr id="item-{{ item.original_index }}">
                        <td class="select-col"><input type="checkbox" name="selecionado" value="{{ item.original_index }}"></td>
                        <td data-label="Unidade">{{ item.unidade }}</td>
                        <td data-label="Nome"><strong>{{ item.nome }}</strong></td>
                        <td data-label="Compra">€{{ "%.2f"|format(item.valor) }} / {{ item.quantidade|int if item.quantidade == item.quantidade|int else item.quantidade }}{{ item.unidade }}</td>
                        <td data-label="Stock Qtd. (Total)">
                            <strong class="{% if item.stock >= (3 * item.quantidade) %}stock-verde{% elif item.stock <= item.quantidade %}stock-vermelho{% endif %}">
                                {{ item.stock|int if item.stock == item.stock|int else item.stock }}{{ item.unidade }}
                            </strong>
                        </td>
                        <td data-label="Stock Unidades">
                            <div class="unit-stock-control">
                                <form action="{{ url_for('ingredientes_bp.decrementar_stock', index=item.original_index) }}" method="POST" style="display:inline;">
                                    <button type="submit" class="unit-btn" title="Menos uma unidade">-</button>
                                </form>
                                <form action="{{ url_for('ingredientes_bp.atualizar_stock_unidades', index=item.original_index) }}" method="POST" style="display:inline-flex; align-items: center; margin: 0;">
                                    <input type="text" name="stock_unidades" value="{{ item.stock_unidades|int if item.stock_unidades == item.stock_unidades|int else item.stock_unidades }}" onchange="this.form.submit()" style="width: 50px; padding: 2px 4px; border: 1px solid #EEDCCE; border-radius: 4px; text-align: center; font-size: 13px; font-weight: 600; background: #FFFFFF; color: #4A3B32;" title="Escreve e prime Enter ou clica fora" />
                                </form>
                                <form action="{{ url_for('ingredientes_bp.incrementar_stock', index=item.original_index) }}" method="POST" style="display:inline;">
                                    <button type="submit" class="unit-btn" title="Mais uma unidade">+</button>
                                </form>
                            </div>
                        </td>
                        <td data-label="Alergénios"><span style="font-size: 12px; color: #b85d19; font-weight: 600;">{{ item.alergenios if item.alergenios else 'Nenhum' }}</span></td>
                        <td data-label="Custo/Un." class="cost-val">€{{ "%.4f"|format(item.custo_unitario) }}/{{ item.unidade }}</td>
                        <td data-label="Ações">
                            <div class="action-icons">
                                <a href="{{ url_for('ingredientes_bp.editar', index=item.original_index) }}" class="action-btn" title="Editar">✏</a>
                                <form action="{{ url_for('ingredientes_bp.apagar', index=item.original_index) }}" method="POST" style="margin:0;" onsubmit="return confirm('Tem certeza que deseja apagar?');">
                                    <button type="submit" class="action-btn delete" title="Apagar">🗑️</button>
                                </form>
                            </div>
                        </td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="9" style="text-align: center; color: #8C7A6B; padding: 30px;">Ainda não existem ingredientes registados.</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    {% endif %}
</div>
</body>
</html>
"""

@ingredientes_bp.route("/ingredientes")
def ingredientes():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    dados_brutos = carregar_dados_db(user_email)
    lista = []
    for row in dados_brutos:
        nome = str(row.get('nome', ''))
        unidade = str(row.get('unidade', ''))
        quantidade = parse_val(row.get('quantidade', 1))
        valor = parse_val(row.get('valor', 0))
        stock_unidades = parse_val(row.get('stock_unidades', 0))
        alergenios = str(row.get('alergenios', ''))
        
        stock = stock_unidades * quantidade
        custo_unitario = (valor / quantidade) if quantidade > 0 else 0
        lista.append({
            "original_index": row['original_index'],
            "nome": nome,
            "unidade": unidade,
            "quantidade": quantidade,
            "valor": valor,
            "stock": stock,
            "stock_unidades": stock_unidades,
            "alergenios": alergenios,
            "custo_unitario": custo_unitario
        })
    
    ordenar = request.args.get('ordenar')
    if ordenar == 'az':
        lista = sorted(lista, key=lambda x: x['nome'].lower())
    elif ordenar == 'za':
        lista = sorted(lista, key=lambda x: x['nome'].lower(), reverse=True)
    elif ordenar == 'qtd_asc':
        lista = sorted(lista, key=lambda x: x['quantidade'])
    elif ordenar == 'qtd_desc':
        lista = sorted(lista, key=lambda x: x['quantidade'], reverse=True)

    return render_template_string(HTML_TEMPLATE, pagina="ingredientes", lista=lista)

@ingredientes_bp.route("/ingredientes/lista_compras")
def lista_compras():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    dados_brutos = carregar_dados_db(user_email)
    lista_compras = []
    for row in dados_brutos:
        stock_unidades = parse_val(row.get('stock_unidades', 0))
        quantidade = parse_val(row.get('quantidade', 1))
        valor = parse_val(row.get('valor', 0))
        stock = stock_unidades * quantidade
        
        if stock_unidades <= 0:
            lista_compras.append({
                "nome": str(row.get('nome', '')),
                "unidade": str(row.get('unidade', '')),
                "stock_unidades": stock_unidades,
                "stock": stock,
                "valor": valor
            })

    return render_template_string(HTML_TEMPLATE, pagina="lista_compras", lista_compras=lista_compras)

@ingredientes_bp.route("/adicionar", methods=["POST"])
def adicionar():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    nome = request.form.get("nome")
    unidade = request.form.get("unidade")
    quantidade = parse_val(request.form.get("quantidade", 1))
    valor = parse_val(request.form.get("valor", 0))
    stock_unidades = parse_val(request.form.get("stock_unidades", 0))
    alergenios = request.form.get("alergenios", "")
    stock = stock_unidades * quantidade

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO ingredientes_db (user_email, nome, unidade, quantidade, valor, stock, stock_unidades, alergenios)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (user_email, nome, unidade, quantidade, valor, stock, stock_unidades, alergenios))
    conn.commit()
    conn.close()

    return redirect(url_for("ingredientes_bp.ingredientes"))

@ingredientes_bp.route("/editar/<int:index>")
def editar(index):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id as original_index, nome, unidade, quantidade, valor, stock, stock_unidades, alergenios FROM ingredientes_db WHERE id = ? AND user_email = ?", (index, user_email))
    row = cursor.fetchone()
    conn.close()

    if row:
        return render_template_string(HTML_TEMPLATE, pagina="editar", item=dict(row), index=index)
    return redirect(url_for("ingredientes_bp.ingredientes"))

@ingredientes_bp.route("/atualizar/<int:index>", methods=["POST"])
def atualizar(index):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    nome = request.form.get("nome")
    unidade = request.form.get("unidade")
    quantidade = parse_val(request.form.get("quantidade", 1))
    valor = parse_val(request.form.get("valor", 0))
    stock_unidades = parse_val(request.form.get("stock_unidades", 0))
    alergenios = request.form.get("alergenios", "")
    stock = stock_unidades * quantidade

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE ingredientes_db
        SET nome = ?, unidade = ?, quantidade = ?, valor = ?, stock = ?, stock_unidades = ?, alergenios = ?
        WHERE id = ? AND user_email = ?
    ''', (nome, unidade, quantidade, valor, stock, stock_unidades, alergenios, index, user_email))
    conn.commit()
    conn.close()

    return redirect(url_for("ingredientes_bp.ingredientes"))

@ingredientes_bp.route("/incrementar_stock/<int:index>", methods=["POST"])
def incrementar_stock(index):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT quantidade, stock_unidades FROM ingredientes_db WHERE id = ? AND user_email = ?", (index, user_email))
    row = cursor.fetchone()
    if row:
        quantidade = parse_val(row['quantidade'])
        novo_stock_unidades = parse_val(row['stock_unidades']) + 1
        novo_stock = novo_stock_unidades * quantidade
        
        cursor.execute('''
            UPDATE ingredientes_db
            SET stock_unidades = ?, stock = ?
            WHERE id = ? AND user_email = ?
        ''', (novo_stock_unidades, novo_stock, index, user_email))
        conn.commit()
    conn.close()

    return redirect(url_for("ingredientes_bp.ingredientes") + f"#item-{index}")

@ingredientes_bp.route("/decrementar_stock/<int:index>", methods=["POST"])
def decrementar_stock(index):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT quantidade, stock_unidades FROM ingredientes_db WHERE id = ? AND user_email = ?", (index, user_email))
    row = cursor.fetchone()
    if row:
        quantidade = parse_val(row['quantidade'])
        novo_stock_unidades = max(0.0, parse_val(row['stock_unidades']) - 1)
        novo_stock = novo_stock_unidades * quantidade
        
        cursor.execute('''
            UPDATE ingredientes_db
            SET stock_unidades = ?, stock = ?
            WHERE id = ? AND user_email = ?
        ''', (novo_stock_unidades, novo_stock, index, user_email))
        conn.commit()
    conn.close()

    return redirect(url_for("ingredientes_bp.ingredientes") + f"#item-{index}")

@ingredientes_bp.route("/atualizar_stock_unidades/<int:index>", methods=["POST"])
def atualizar_stock_unidades(index):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    stock_unidades_raw = request.form.get("stock_unidades", "0")
    stock_unidades = parse_val(stock_unidades_raw)

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT quantidade FROM ingredientes_db WHERE id = ? AND user_email = ?", (index, user_email))
    row = cursor.fetchone()
    if row:
        quantidade = parse_val(row['quantidade'])
        novo_stock = stock_unidades * quantidade
        
        cursor.execute('''
            UPDATE ingredientes_db
            SET stock_unidades = ?, stock = ?
            WHERE id = ? AND user_email = ?
        ''', (stock_unidades, novo_stock, index, user_email))
        conn.commit()
    conn.close()

    return redirect(url_for("ingredientes_bp.ingredientes") + f"#item-{index}")

@ingredientes_bp.route("/apagar/<int:index>", methods=["POST"])
def apagar(index):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ingredientes_db WHERE id = ? AND user_email = ?", (index, user_email))
    conn.commit()
    conn.close()

    return redirect(url_for("ingredientes_bp.ingredientes"))

@ingredientes_bp.route("/exportar")
def exportar_excel():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    dados_brutos = carregar_dados_db(user_email)
    if not dados_brutos:
        df = pd.DataFrame(columns=["nome", "unidade", "quantidade", "valor", "stock", "stock_unidades", "alergenios"])
    else:
        df = pd.DataFrame(dados_brutos)
        if 'original_index' in df.columns:
            df = df.drop(columns=['original_index'])
            
    output_path = get_excel_path(f"ingredientes_{user_email.replace('@', '_').replace('.', '_')}.xlsx")
    df.to_excel(output_path, index=False, sheet_name=SHEET_NAME)
    return send_file(output_path, as_attachment=True)

@ingredientes_bp.route("/importar", methods=["POST"])
def importar_excel():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    ficheiro = request.files.get("ficheiro")
    if ficheiro and ficheiro.filename:
        ext = ficheiro.filename.split(".")[-1].lower()
        try:
            if ext == "csv":
                df = pd.read_csv(ficheiro)
            elif ext in ["xls", "xlsx"]:
                df = pd.read_excel(ficheiro)
            else:
                return redirect(url_for("ingredientes_bp.ingredientes"))

            conn = get_db_connection()
            cursor = conn.cursor()

            for _, row in df.iterrows():
                row_dict = {str(c).strip().lower(): v for c, v in row.items()}
                nome = str(row_dict.get("nome", row_dict.get("ingrediente", "Sem Nome")))
                unidade = str(row_dict.get("unidade", "g"))
                quantidade = parse_val(row_dict.get("quantidade", row_dict.get("qtd", 1)))
                valor = parse_val(row_dict.get("valor", row_dict.get("preço", row_dict.get("preco", 0))))
                stock_unidades = parse_val(row_dict.get("stock_unidades", row_dict.get("stock unidades", 0)))
                alergenios = str(row_dict.get("alergenios", row_dict.get("alergénios", "")))
                stock = stock_unidades * quantidade

                cursor.execute('''
                    INSERT INTO ingredientes_db (user_email, nome, unidade, quantidade, valor, stock, stock_unidades, alergenios)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (user_email, nome, unidade, quantidade, valor, stock, stock_unidades, alergenios))

            conn.commit()
            conn.close()
        except Exception as e:
            print("Erro ao importar:", e)

    return redirect(url_for("ingredientes_bp.ingredientes"))