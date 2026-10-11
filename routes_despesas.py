from flask import Blueprint, render_template_string, request, redirect, url_for, session
import sqlite3

despesas_bp = Blueprint('despesas_bp', __name__)

DB_PATH = 'cadoces.db'

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def inicializar_tabelas():
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS custos_fixos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_email TEXT,
            nome TEXT,
            valor REAL,
            frequencia TEXT,
            percentagem REAL
        )
    ''')
    conn.commit()
    conn.close()

inicializar_tabelas()

def carregar_custos(user_email):
    conn = get_db_connection()
    custos = conn.execute(
        'SELECT * FROM custos_fixos WHERE user_email = ?', 
        (user_email,)
    ).fetchall()
    conn.close()
    return custos

@despesas_bp.route("/despesas")
def listar_custos():
    user_email = session.get('user_email')
    if not user_email:
        return redirect(url_for('login'))

    lista_custos = carregar_custos(user_email)
    
    total_mensal = 0.0
    gastos_producao = 0.0
    num_custos = len(lista_custos)
    
    linhas_tabela = ""
    
    if lista_custos:
        for row in lista_custos:
            nome = row['nome'] or 'Sem Nome'
            valor = float(row['valor'] or 0)
            freq = row['frequencia'] or 'Mensal'
            perc = float(row['percentagem'] or 100)
            
            if freq == 'Anual':
                valor_base_mensal = valor / 12
            elif freq == 'Trimestral':
                valor_base_mensal = valor / 3
            else:
                valor_base_mensal = valor
                
            total_mensal += valor_base_mensal
            
            custo_prod = valor_base_mensal * (perc / 100)
            gastos_producao += custo_prod
            
            linhas_tabela += f"""
            <tr>
                <td data-label="Nome"><strong>{nome}</strong></td>
                <td data-label="Frequência">{freq}</td>
                <td data-label="Valor">€{valor:.2f}</td>
                <td data-label="% Produção">{perc:.1f}%</td>
                <td data-label="Mensal">€{custo_prod:.2f}</td>
                <td data-label="Ações" style="text-align: right;"><a href="/despesas/apagar/{row['id']}" style="color: #d9534f; text-decoration: none; font-weight: bold;">🗑️ Apagar</a></td>
            </tr>
            """
    else:
        linhas_tabela = '<tr><td colspan="6" style="text-align:center; color:#888; padding: 20px;">Ainda não tens custos fixos registados.</td></tr>'

    return render_template_string(f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Custos Fixos</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; }}
            .container {{ max-width: 1100px; margin: auto; }}
            .header-bar {{ display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; gap: 15px; flex-wrap: wrap; }}
            .header-left {{ display: flex; align-items: center; gap: 20px; }}
            .back-btn {{ text-decoration: none; color: #b85d19; font-size: 22px; font-weight: bold; background: #f9f0e6; padding: 8px 14px; border-radius: 8px; }}
            h1 {{ margin: 0; color: #111; font-size: 24px; }}
            p.sub {{ margin: 5px 0 0 0; color: #666; font-size: 14px; }}
            .btn-primary {{ background: #dfa673; color: #fff; border: none; padding: 12px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; display: flex; align-items: center; gap: 6px; width: fit-content; }}
            .btn-primary:hover {{ background: #c8915f; }}

            .cards-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px; }}
            .card {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 14px; padding: 25px; position: relative; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
            .card.highlight {{ background: #fcf3e3; border-color: #e6c594; }}
            .card-title {{ font-size: 14px; color: #666; margin-bottom: 10px; font-weight: bold; }}
            .card-value {{ font-size: 32px; font-weight: bold; color: #111; }}
            .card-icon {{ position: absolute; right: 25px; top: 25px; background: #f2e2cb; padding: 10px; border-radius: 8px; font-size: 18px; }}

            .calc-section {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 25px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 6px rgba(0,0,0,0.02); gap: 20px; flex-wrap: wrap; }}
            .calc-left {{ flex: 1; min-width: 280px; }}
            .calc-right {{ background: #fcf3e3; border: 2px solid #e6c594; border-radius: 14px; padding: 20px 30px; display: flex; align-items: center; justify-content: space-between; min-width: 280px; position: relative; box-sizing: border-box; }}
            .form-control {{ width: 100%; padding: 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; box-sizing: border-box; margin-top: 8px; }}

            .content-box {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
            table.table-custom {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            table.table-custom th {{ background-color: #f9f0e6; color: #8c4410; padding: 12px; text-align: left; border-bottom: 2px solid #f3e5d0; }}
            table.table-custom td {{ padding: 12px; border-bottom: 1px solid #f3e5d0; color: #444; }}

            .modal {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4); overflow-y: auto; padding: 30px 0; }}
            .modal-content {{ background-color: #fffdf9; margin: auto; padding: 30px; border: 2px solid #f3e5d0; width: 90%; max-width: 500px; border-radius: 14px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); position: relative; box-sizing: border-box; }}
            .close-btn {{ position: absolute; right: 20px; top: 20px; font-size: 22px; cursor: pointer; color: #888; font-weight: bold; }}
            .form-group {{ margin-bottom: 20px; }}
            .form-group label {{ display: block; font-weight: bold; color: #222; margin-bottom: 8px; }}
            .modal-footer {{ display: flex; justify-content: flex-end; gap: 12px; }}
            .btn-secondary {{ background: #fff; border: 2px solid #f3e5d0; color: #444; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; }}

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {{
                body {{ padding: 15px; }}
                .header-bar {{ flex-direction: column; align-items: stretch; text-align: center; }}
                .header-left {{ flex-direction: column; gap: 10px; }}
                .cards-grid {{ grid-template-columns: 1fr; gap: 15px; }}
                .calc-section {{ flex-direction: column; align-items: stretch; }}
                .calc-right {{ width: 100%; }}
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
                    <a href="/" class="back-btn">&#8592;</a>
                    <div>
                        <h1>Custos Fixos</h1>
                        <p class="sub">Gerencie seus custos fixos e variáveis</p>
                    </div>
                </div>
                <button class="btn-primary" onclick="abrirModal()">+ Novo Custo Fixo</button>
            </div>

            <div class="cards-grid">
                <div class="card highlight">
                    <div class="card-title">Total Mensal</div>
                    <div class="card-value">€{total_mensal:.2f}</div>
                    <div class="card-icon">📄</div>
                </div>
                <div class="card">
                    <div class="card-title">Gastos com Produção</div>
                    <div class="card-value">€{gastos_producao:.2f}</div>
                    <div class="card-icon">📄</div>
                </div>
                <div class="card">
                    <div class="card-title">Nº de Custos Fixos</div>
                    <div class="card-value">{num_custos}</div>
                    <div class="card-icon">📄</div>
                </div>
            </div>

            <div class="calc-section">
                <div class="calc-left">
                    <h3 style="margin-top:0; color:#111; font-size: 20px;">Cálculo por Encomenda</h3>
                    <div style="max-width: 350px; margin-top: 15px;">
                        <label style="font-weight: bold; color: #333; font-size: 14px;">Média de encomendas por mês</label>
                        <input type="number" id="inputEncomendas" class="form-control" placeholder="Ex: 15" min="1" value="1" oninput="calcularEncomenda()">
                    </div>
                </div>
                <div class="calc-right">
                    <div>
                        <div style="font-size: 13px; color: #666; font-weight: bold; margin-bottom: 5px;">Custo fixo por encomenda</div>
                        <div id="custoPorEncomenda" style="font-size: 28px; font-weight: bold; color: #111;">€{gastos_producao:.2f}</div>
                    </div>
                    <div style="background: #f2e2cb; padding: 12px; border-radius: 10px; font-size: 20px; font-weight: bold;">÷</div>
                </div>
            </div>

            <div class="content-box">
                <h3 style="margin-top:0; color:#111;">Lista de Custos Fixos</h3>
                <table class="table-custom">
                    <thead>
                        <tr>
                            <th>Nome</th>
                            <th>Frequência</th>
                            <th>Valor</th>
                            <th>% Produção</th>
                            <th>Mensal</th>
                            <th style="text-align: right;">Ações</th>
                        </tr>
                    </thead>
                    <tbody>
                        {linhas_tabela}
                    </tbody>
                </table>
            </div>
        </div>

        <div id="modalCusto" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModal()">&times;</span>
                <h2 style="margin-top:0; color:#111; font-size:22px;">Adicionar Custo Fixo</h2>
                
                <form action="/despesas/adicionar" method="POST">
                    <div class="form-group">
                        <label>Nome do Custo Fixo</label>
                        <input type="text" name="nome" class="form-control" placeholder="Ex: Conta de Luz" required>
                    </div>

                    <div style="display: flex; gap: 15px; flex-wrap: wrap;">
                        <div class="form-group" style="flex: 1; min-width: 150px;">
                            <label>Valor (€)</label>
                            <input type="number" step="0.01" name="valor" class="form-control" value="0" min="0" required>
                        </div>
                        <div class="form-group" style="flex: 1; min-width: 150px;">
                            <label>Frequência</label>
                            <select name="frequencia" class="form-control">
                                <option value="Mensal">Mensal</option>
                                <option value="Trimestral">Trimestral</option>
                                <option value="Anual">Anual</option>
                            </select>
                        </div>
                    </div>

                    <div class="form-group">
                        <label>% Direcionada à Produção</label>
                        <input type="number" name="percentagem" class="form-control" value="100" min="0" max="100">
                    </div>

                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModal()">Cancelar</button>
                        <button type="submit" class="btn-primary">Adicionar</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            const gastosProducaoTotal = {gastos_producao};

            function abrirModal() {{
                document.getElementById('modalCusto').style.display = 'block';
            }}
            function fecharModal() {{
                document.getElementById('modalCusto').style.display = 'none';
            }}
            function calcularEncomenda() {{
                const qtd = parseFloat(document.getElementById('inputEncomendas').value) || 0;
                const elem = document.getElementById('custoPorEncomenda');
                if (qtd > 0) {{
                    const resultado = gastosProducaoTotal / qtd;
                    elem.innerText = '€' + resultado.toFixed(2);
                }} else {{
                    elem.innerText = '€0.00';
                }}
            }}
            calcularEncomenda();
        </script>
    </body>
    </html>
    """)

@despesas_bp.route("/despesas/adicionar", methods=["POST"])
def adicionar_custo():
    user_email = session.get('user_email')
    if not user_email:
        return redirect(url_for('login'))

    nome = request.form.get("nome")
    valor = float(request.form.get("valor", 0) or 0)
    frequencia = request.form.get("frequencia")
    percentagem = float(request.form.get("percentagem", 100) or 100)
    
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO custos_fixos (user_email, nome, valor, frequencia, percentagem)
        VALUES (?, ?, ?, ?, ?)
    ''', (user_email, nome, valor, frequencia, percentagem))
    conn.commit()
    conn.close()
    
    return redirect(url_for('despesas_bp.listar_custos'))

@despesas_bp.route("/despesas/apagar/<int:id_custo>")
def apagar_custo(id_custo):
    user_email = session.get('user_email')
    if not user_email:
        return redirect(url_for('login'))

    conn = get_db_connection()
    conn.execute(
        'DELETE FROM custos_fixos WHERE id = ? AND user_email = ?', 
        (id_custo, user_email)
    )
    conn.commit()
    conn.close()
    
    return redirect(url_for('despesas_bp.listar_custos'))