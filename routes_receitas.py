from flask import Blueprint, render_template_string, request, redirect, url_for, session
import sqlite3
import json

receitas_bp = Blueprint('receitas_bp', __name__)

DB_PATH = 'cadoces.db'

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def inicializar_tabelas():
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS receitas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_email TEXT,
            nome TEXT,
            categoria TEXT,
            descricao TEXT,
            preparacao TEXT,
            notas TEXT,
            tempo INTEGER,
            ingredientes_usados TEXT,
            custo_total REAL
        )
    ''')
    conn.commit()
    conn.close()

inicializar_tabelas()

def carregar_receitas(user_email):
    conn = get_db_connection()
    receitas = conn.execute(
        'SELECT * FROM receitas WHERE user_email = ?', 
        (user_email,)
    ).fetchall()
    conn.close()
    return receitas

def carregar_ingredientes_detalhes(user_email):
    conn = get_db_connection()
    ingredientes_dict = {}
    try:
        cursor = conn.execute(
            'SELECT nome, unidade, quantidade, valor FROM ingredientes_db WHERE user_email = ?', 
            (user_email,)
        )
        for row in cursor.fetchall():
            nome = row['nome']
            if nome:
                quantidade = float(row['quantidade'] or 1)
                valor = float(row['valor'] or 0)
                unidade = row['unidade'] or ''
                custo_unitario = (valor / quantidade) if quantidade > 0 else 0.0
                ingredientes_dict[nome] = {"custo_unitario": custo_unitario, "unidade": unidade}
    except Exception as e:
        print("Aviso ao carregar ingredientes_db, a tentar tabela alternativa...", e)
        try:
            cursor = conn.execute(
                'SELECT nome, unidade, quantidade, valor FROM ingredientes WHERE user_email = ?', 
                (user_email,)
            )
            for row in cursor.fetchall():
                nome = row['nome']
                if nome:
                    quantidade = float(row['quantidade'] or 1)
                    valor = float(row['valor'] or 0)
                    unidade = row['unidade'] or ''
                    custo_unitario = (valor / quantidade) if quantidade > 0 else 0.0
                    ingredientes_dict[nome] = {"custo_unitario": custo_unitario, "unidade": unidade}
        except Exception as e2:
            print("Erro total ao carregar ingredientes:", e2)
    finally:
        conn.close()
    return ingredientes_dict

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CáDoces - Gestão de Receitas</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAF8F5; color: #4A3B32; margin: 0; padding: 30px; }
        .container { max-width: 1200px; margin: auto; }
        .header-bar { display: flex; justify-content: space-between; align-items: center; background: #FFFFFF; padding: 20px 25px; border: 1px solid #EEDCCE; border-radius: 14px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); flex-wrap: wrap; gap: 15px; }
        .header-left { display: flex; align-items: center; gap: 20px; }
        .back-btn { text-decoration: none; color: #4A3B32; font-size: 18px; font-weight: bold; background: #F5EBE0; padding: 10px 16px; border-radius: 10px; border: 1px solid #EEDCCE; transition: background 0.2s; display: inline-flex; align-items: center; justify-content: center; }
        .back-btn:hover { background: #EEDCCE; }
        h1 { margin: 0; color: #2C1E1A; font-size: 24px; font-weight: 700; }
        p.sub { margin: 4px 0 0 0; color: #8C7A6B; font-size: 14px; }
        .top-actions { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
        .btn-outline { background: #F5EBE0; border: 1px solid #EEDCCE; color: #4A3B32; padding: 10px 16px; border-radius: 10px; cursor: pointer; font-weight: 600; display: inline-flex; align-items: center; gap: 6px; text-decoration: none; font-size: 14px; transition: all 0.2s; }
        .btn-outline:hover { background: #EEDCCE; }
        .btn-primary { background: #D99B6A; color: #FFFFFF; border: none; padding: 11px 20px; border-radius: 10px; cursor: pointer; font-weight: 600; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; font-size: 14px; transition: background 0.2s; box-shadow: 0 2px 4px rgba(217,155,106,0.2); }
        .btn-primary:hover { background: #C48757; }
        .content-box { background: #FFFFFF; border: 1px solid #EEDCCE; border-radius: 14px; padding: 25px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); margin-top: 25px; }
        table.table-custom { width: 100%; border-collapse: collapse; margin-top: 5px; }
        table.table-custom th { background-color: #FBF7F2; color: #734D2E; padding: 14px 16px; text-align: left; border-bottom: 2px solid #EEDCCE; font-size: 14px; font-weight: 600; }
        table.table-custom td { padding: 14px 16px; border-bottom: 1px solid #F5EBE0; color: #4A3B32; font-size: 14px; vertical-align: middle; }
        table.table-custom tr:hover td { background-color: #FAF8F5; }
        .cost-val { color: #D99B6A; font-weight: 600; }
        .action-icons { display: flex; gap: 12px; align-items: center; }
        .action-btn { background: none; border: none; cursor: pointer; font-size: 16px; color: #8C7A6B; padding: 0; text-decoration: none; transition: color 0.2s; }
        .action-btn:hover { color: #4A3B32; }
        .action-btn.delete:hover { color: #C0392B; }
        
        /* Modal */
        .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(74,59,50,0.4); overflow-y: auto; padding: 40px 0; }
        .modal-content { background-color: #FFFFFF; margin: auto; padding: 30px; border: 1px solid #EEDCCE; width: 550px; max-width: 90%; border-radius: 14px; box-shadow: 0 8px 24px rgba(74,59,50,0.1); position: relative; box-sizing: border-box; }
        .close-btn { position: absolute; right: 20px; top: 20px; font-size: 20px; cursor: pointer; color: #8C7A6B; font-weight: bold; background: #F5EBE0; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 1px solid #EEDCCE; }
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; font-size: 12px; margin-bottom: 6px; color: #734D2E; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .form-control { width: 100%; padding: 10px 12px; border: 1px solid #EEDCCE; border-radius: 8px; box-sizing: border-box; background: #FAF8F5; color: #4A3B32; font-size: 14px; outline: none; }
        .form-control:focus { border-color: #D99B6A; background: #FFFFFF; }
        .modal-footer { display: flex; justify-content: flex-end; gap: 12px; border-top: 1px solid #EEDCCE; padding-top: 20px; margin-top: 20px; flex-wrap: wrap; }

        @media screen and (max-width: 768px) {
            body { padding: 15px; }
            .header-bar { flex-direction: column; align-items: stretch; }
            .header-left { flex-direction: column; align-items: flex-start; gap: 12px; }
            .top-actions { width: 100%; justify-content: stretch; }
            .top-actions a, .top-actions button { flex: 1; justify-content: center; }
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
            .action-icons { justify-content: flex-end; }
        }
    </style>
    <script>
        const dadosIngredientes = {{ custos_js | safe }};

        function abrirModalReceita() {
            document.getElementById('modalReceita').style.display = 'block';
        }

        function fecharModalReceita() {
            document.getElementById('modalReceita').style.display = 'none';
        }

        function atualizarUnidade(selectElement) {
            const nomeIngrediente = selectElement.value;
            const containerLinha = selectElement.parentElement;
            const spanUnidade = containerLinha.querySelector('.unidade-label');
            
            if (dadosIngredientes[nomeIngrediente]) {
                spanUnidade.textContent = dadosIngredientes[nomeIngrediente].unidade;
            } else {
                spanUnidade.textContent = '';
            }
        }

        function adicionarLinhaIngrediente(containerId = 'ingredientesContainer', ingredienteVal = '', qtdVal = '') {
            const container = document.getElementById(containerId);
            const row = document.createElement('div');
            row.style.cssText = "display: flex; gap: 10px; margin-bottom: 10px; align-items: center;";
            
            let optionsHtml = '<option value="" disabled ' + (!ingredienteVal ? 'selected' : '') + '>Selecionar ingrediente...</option>';
            for (const [ing, info] of Object.entries(dadosIngredientes)) {
                let selected = (ing === ingredienteVal) ? 'selected' : '';
                optionsHtml += `<option value="${ing}" ${selected}>${ing}</option>`;
            }

            let unidadeInicial = '';
            if (ingredienteVal && dadosIngredientes[ingredienteVal]) {
                unidadeInicial = dadosIngredientes[ingredienteVal].unidade;
            }

            row.innerHTML = `
                <select name="ingrediente[]" class="form-control" style="flex: 2;" onchange="atualizarUnidade(this)" required>
                    ${optionsHtml}
                </select>
                <input type="number" step="any" name="quantidade[]" class="form-control" placeholder="Qtd" value="${qtdVal}" style="flex: 1;" required>
                <span class="unidade-label" style="font-size: 13px; font-weight: 600; color: #734D2E; min-width: 35px;">${unidadeInicial}</span>
                <button type="button" onclick="this.parentElement.remove()" style="background: none; border: none; color: #C0392B; font-size: 18px; cursor: pointer;" title="Remover">🗑️</button>
            `;
            container.appendChild(row);
        }
    </script>
</head>
<body>
<div class="container">
    <div class="header-bar">
        <div class="header-left">
            <a href="/" class="back-btn">&#8592;</a>
            <div>
                <h1>Receitas e Custos</h1>
                <p class="sub">Gestão de fichas técnicas e cálculo automático de custos</p>
            </div>
        </div>
        <div class="top-actions">
            <a href="/" class="btn-outline">🏠 Menu</a>
            <button onclick="abrirModalReceita()" class="btn-primary">+ Adicionar Receita</button>
        </div>
    </div>

    <div class="content-box">
        <table class="table-custom">
            <thead>
                <tr>
                    <th>Nome da Receita</th>
                    <th>Categoria</th>
                    <th>Descrição</th>
                    <th>Tempo</th>
                    <th>Custo Total</th>
                    <th>Ações</th>
                </tr>
            </thead>
            <tbody>
                {{ linhas_tabela | safe }}
            </tbody>
        </table>
    </div>
</div>

<!-- Modal Adicionar Receita -->
<div id="modalReceita" class="modal">
    <div class="modal-content">
        <span class="close-btn" onclick="fecharModalReceita()">&times;</span>
        <h2 style="margin-top:0; color:#2C1E1A; font-size:20px;">Adicionar Nova Receita</h2>
        
        <form action="{{ url_for('receitas_bp.adicionar_receita') }}" method="POST">
            <div class="form-group">
                <label>Nome da Receita *</label>
                <input type="text" name="nome" class="form-control" placeholder="Ex: Bolo de Aniversário Baunilha" required>
            </div>
            <div class="form-group">
                <label>Categoria</label>
                <select name="categoria" class="form-control">
                    <option value="Bolos de Aniversário">Bolos de Aniversário</option>
                    <option value="Bolachas e Biscoitos">Bolachas e Biscoitos</option>
                    <option value="Massas Levedadas">Massas Levedadas</option>
                    <option value="Outro">Outro</option>
                </select>
            </div>
            <div class="form-group">
                <label>Descrição / Resumo</label>
                <input type="text" name="descricao" class="form-control" placeholder="Breve descrição...">
            </div>
            <div class="form-group">
                <label>Tempo de Preparação (min)</label>
                <input type="number" name="tempo" class="form-control" placeholder="Ex: 45">
            </div>
            
            <div class="form-group">
                <label style="display: flex; justify-content: space-between; align-items: center;">
                    <span>Ingredientes Necessários</span>
                    <button type="button" onclick="adicionarLinhaIngrediente('ingredientesContainer')" class="btn-outline" style="padding: 4px 10px; font-size: 12px;">+ Adicionar</button>
                </label>
                <div id="ingredientesContainer" style="margin-top: 8px;">
                    <!-- Linha inicial gerada por script ou HTML -->
                </div>
            </div>

            <div class="form-group">
                <label>Modo de Preparação</label>
                <textarea name="preparacao" class="form-control" rows="3" placeholder="Passo a passo..."></textarea>
            </div>
            <div class="form-group">
                <label>Notas</label>
                <textarea name="notas" class="form-control" rows="2" placeholder="Notas adicionais..."></textarea>
            </div>

            <div class="modal-footer">
                <button type="button" class="btn-outline" onclick="fecharModalReceita()">Cancelar</button>
                <button type="submit" class="btn-primary">Guardar Receita</button>
            </div>
        </form>
    </div>
</div>

<script>
    // Inicializar uma linha em branco no modal de adicionar
    document.addEventListener("DOMContentLoaded", function() {
        adicionarLinhaIngrediente('ingredientesContainer');
    });
</script>
</body>
</html>
"""

EDITAR_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CáDoces - Editar Receita</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAF8F5; color: #4A3B32; margin: 0; padding: 30px; }
        .container { max-width: 900px; margin: auto; }
        .header-bar { display: flex; justify-content: space-between; align-items: center; background: #FFFFFF; padding: 20px 25px; border: 1px solid #EEDCCE; border-radius: 14px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); margin-bottom: 25px; }
        .back-btn { text-decoration: none; color: #4A3B32; font-size: 18px; font-weight: bold; background: #F5EBE0; padding: 10px 16px; border-radius: 10px; border: 1px solid #EEDCCE; display: inline-flex; align-items: center; justify-content: center; }
        .back-btn:hover { background: #EEDCCE; }
        h1 { margin: 0; color: #2C1E1A; font-size: 22px; font-weight: 700; }
        .content-box { background: #FFFFFF; border: 1px solid #EEDCCE; border-radius: 14px; padding: 30px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); }
        .form-group { margin-bottom: 18px; }
        .form-group label { display: block; font-size: 12px; margin-bottom: 6px; color: #734D2E; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
        .form-control { width: 100%; padding: 10px 12px; border: 1px solid #EEDCCE; border-radius: 8px; box-sizing: border-box; background: #FAF8F5; color: #4A3B32; font-size: 14px; outline: none; }
        .form-control:focus { border-color: #D99B6A; background: #FFFFFF; }
        .btn-outline { background: #F5EBE0; border: 1px solid #EEDCCE; color: #4A3B32; padding: 8px 14px; border-radius: 8px; cursor: pointer; font-weight: 600; text-decoration: none; font-size: 13px; display: inline-flex; align-items: center; gap: 6px; }
        .btn-outline:hover { background: #EEDCCE; }
        .btn-primary { background: #D99B6A; color: #FFFFFF; border: none; padding: 12px 20px; border-radius: 10px; cursor: pointer; font-weight: 600; text-decoration: none; font-size: 14px; box-shadow: 0 2px 4px rgba(217,155,106,0.2); }
        .btn-primary:hover { background: #C48757; }
        .footer-actions { display: flex; justify-content: flex-end; gap: 12px; border-top: 1px solid #EEDCCE; padding-top: 20px; margin-top: 20px; }
    </style>
    <script>
        const dadosIngredientes = {{ custos_js | safe }};

        function atualizarUnidade(selectElement) {
            const nomeIngrediente = selectElement.value;
            const containerLinha = selectElement.parentElement;
            const spanUnidade = containerLinha.querySelector('.unidade-label');
            
            if (dadosIngredientes[nomeIngrediente]) {
                spanUnidade.textContent = dadosIngredientes[nomeIngrediente].unidade;
            } else {
                spanUnidade.textContent = '';
            }
        }

        function adicionarLinhaIngrediente(containerId = 'ingredientesContainer', ingredienteVal = '', qtdVal = '') {
            const container = document.getElementById(containerId);
            const row = document.createElement('div');
            row.style.cssText = "display: flex; gap: 10px; margin-bottom: 10px; align-items: center;";
            
            let optionsHtml = '<option value="" disabled ' + (!ingredienteVal ? 'selected' : '') + '>Selecionar ingrediente...</option>';
            for (const [ing, info] of Object.entries(dadosIngredientes)) {
                let selected = (ing === ingredienteVal) ? 'selected' : '';
                optionsHtml += `<option value="${ing}" ${selected}>${ing}</option>`;
            }

            let unidadeInicial = '';
            if (ingredienteVal && dadosIngredientes[ingredienteVal]) {
                unidadeInicial = dadosIngredientes[ingredienteVal].unidade;
            }

            row.innerHTML = `
                <select name="ingrediente[]" class="form-control" style="flex: 2;" onchange="atualizarUnidade(this)" required>
                    ${optionsHtml}
                </select>
                <input type="number" step="any" name="quantidade[]" class="form-control" placeholder="Qtd" value="${qtdVal}" style="flex: 1;" required>
                <span class="unidade-label" style="font-size: 13px; font-weight: 600; color: #734D2E; min-width: 35px;">${unidadeInicial}</span>
                <button type="button" onclick="this.parentElement.remove()" style="background: none; border: none; color: #C0392B; font-size: 18px; cursor: pointer;" title="Remover">🗑️</button>
            `;
            container.appendChild(row);
        }
    </script>
</head>
<body>
<div class="container">
    <div class="header-bar">
        <div style="display: flex; align-items: center; gap: 15px;">
            <a href="{{ url_for('receitas_bp.listar_receitas') }}" class="back-btn">&#8592;</a>
            <h1>Editar Receita: {{ receita.nome }}</h1>
        </div>
    </div>

    <div class="content-box">
        <form action="{{ url_for('receitas_bp.atualizar_receita', id=receita.id) }}" method="POST">
            <div class="form-group">
                <label>Nome da Receita *</label>
                <input type="text" name="nome" class="form-control" value="{{ receita.nome }}" required>
            </div>
            <div class="form-group">
                <label>Categoria</label>
                <select name="categoria" class="form-control">
                    <option value="Bolos de Aniversário" {% if receita.categoria == 'Bolos de Aniversário' %}selected{% endif %}>Bolos de Aniversário</option>
                    <option value="Bolachas e Biscoitos" {% if receita.categoria == 'Bolachas e Biscoitos' %}selected{% endif %}>Bolachas e Biscoitos</option>
                    <option value="Massas Levedadas" {% if receita.categoria == 'Massas Levedadas' %}selected{% endif %}>Massas Levedadas</option>
                    <option value="Outro" {% if receita.categoria == 'Outro' %}selected{% endif %}>Outro</option>
                </select>
            </div>
            <div class="form-group">
                <label>Descrição / Resumo</label>
                <input type="text" name="descricao" class="form-control" value="{{ receita.descricao or '' }}">
            </div>
            <div class="form-group">
                <label>Tempo de Preparação (min)</label>
                <input type="number" name="tempo" class="form-control" value="{{ receita.tempo or 0 }}">
            </div>
            
            <div class="form-group">
                <label style="display: flex; justify-content: space-between; align-items: center;">
                    <span>Ingredientes Necessários</span>
                    <button type="button" onclick="adicionarLinhaIngrediente('ingredientesContainer')" class="btn-outline">+ Adicionar</button>
                </label>
                <div id="ingredientesContainer" style="margin-top: 8px;">
                    <!-- Preenchido via script com os ingredientes existentes -->
                </div>
            </div>

            <div class="form-group">
                <label>Modo de Preparação</label>
                <textarea name="preparacao" class="form-control" rows="4">{{ receita.preparacao or '' }}</textarea>
            </div>
            <div class="form-group">
                <label>Notas</label>
                <textarea name="notas" class="form-control" rows="2">{{ receita.notas or '' }}</textarea>
            </div>

            <div class="footer-actions">
                <a href="{{ url_for('receitas_bp.listar_receitas') }}" class="btn-outline">Cancelar</a>
                <button type="submit" class="btn-primary">Guardar Alterações</button>
            </div>
        </form>
    </div>
</div>

<script>
    document.addEventListener("DOMContentLoaded", function() {
        const ingredientesAnteriores = {{ ingredientes_atuais_json | safe }};
        if (ingredientesAnteriores && ingredientesAnteriores.length > 0) {
            ingredientesAnteriores.forEach(item => {
                adicionarLinhaIngrediente('ingredientesContainer', item.nome, item.quantidade);
            });
        } else {
            adicionarLinhaIngrediente('ingredientesContainer');
        }
    });
</script>
</body>
</html>
"""

@receitas_bp.route("/receitas")
def listar_receitas():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    lista_receitas = carregar_receitas(user_email)
    ingredientes_dict = carregar_ingredientes_detalhes(user_email)
    
    linhas_tabela = ""
    if lista_receitas:
        for row in lista_receitas:
            categoria_txt = row['categoria'] if 'categoria' in row.keys() and row['categoria'] else 'Geral'
            custo_total = float(row['custo_total'] or 0)
            linhas_tabela += f"""
            <tr>
                <td data-label="Nome da Receita"><strong>{row['nome'] or 'Sem Nome'}</strong></td>
                <td data-label="Categoria"><span style="background: #F5EBE0; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 600; color: #734D2E;">{categoria_txt}</span></td>
                <td data-label="Descrição">{row['descricao'] or '-'}</td>
                <td data-label="Tempo">{row['tempo'] or 0} min</td>
                <td data-label="Custo Total" class="cost-val">€{custo_total:.2f}</td>
                <td data-label="Ações">
                    <div class="action-icons">
                        <a href="{url_for('receitas_bp.editar_receita', id=row['id'])}" class="action-btn" title="Editar">✏</a>
                        <form action="{url_for('receitas_bp.apagar_receita', id=row['id'])}" method="POST" style="margin:0;" onsubmit="return confirm('Tem certeza que deseja apagar esta receita?');">
                            <button type="submit" class="action-btn delete" title="Apagar">🗑️</button>
                        </form>
                    </div>
                </td>
            </tr>
            """
    else:
        linhas_tabela = '<tr><td colspan="6" style="text-align:center; color:#8C7A6B; padding: 30px;">Nenhuma receita registada.</td></tr>'

    custos_js = json.dumps(ingredientes_dict)

    return render_template_string(
        HTML_TEMPLATE,
        linhas_tabela=linhas_tabela,
        custos_js=custos_js
    )

@receitas_bp.route("/receitas/adicionar", methods=["POST"])
def adicionar_receita():
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    nome = request.form.get("nome")
    categoria = request.form.get("categoria")
    descricao = request.form.get("descricao")
    preparacao = request.form.get("preparacao")
    notas = request.form.get("notas")
    tempo = int(request.form.get("tempo", 0) or 0)
    
    ingredientes_selecionados = request.form.getlist("ingrediente[]")
    quantidades_inseridas = request.form.getlist("quantidade[]")
    
    ingredientes_info = carregar_ingredientes_detalhes(user_email)
    
    lista_final = []
    custo_total_calculado = 0.0
    
    for ing, qtd in zip(ingredientes_selecionados, quantidades_inseridas):
        if ing:
            qtd_num = float(qtd) if qtd else 0.0
            unidade_ing = ingredientes_info[ing]["unidade"] if ing in ingredientes_info else ""
            lista_final.append(f"{qtd_num}{unidade_ing} {ing}")
            
            if ing in ingredientes_info:
                custo_total_calculado += ingredientes_info[ing]["custo_unitario"] * qtd_num
            
    ingredientes_str = " | ".join(lista_final)
    
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO receitas (user_email, nome, categoria, descricao, preparacao, notas, tempo, ingredientes_usados, custo_total)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (user_email, nome, categoria, descricao, preparacao, notas, tempo, ingredientes_str, custo_total_calculado))
    conn.commit()
    conn.close()
    
    return redirect(url_for('receitas_bp.listar_receitas'))

@receitas_bp.route("/receitas/editar/<int:id>", methods=["GET"])
def editar_receita(id):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    conn = get_db_connection()
    receita = conn.execute('SELECT * FROM receitas WHERE id = ? AND user_email = ?', (id, user_email)).fetchone()
    conn.close()

    if not receita:
        return redirect(url_for('receitas_bp.listar_receitas'))

    ingredientes_info = carregar_ingredientes_detalhes(user_email)
    custos_js = json.dumps(ingredientes_info)

    # Converter string de ingredientes guardados (ex: "500g Farinha | 2x Ovos") de volta para lista estruturada para o formulário
    ingredientes_atuais = []
    ingredientes_str = receita['ingredientes_usados'] or ""
    if ingredientes_str:
        partes = ingredientes_str.split(" | ")
        for p in partes:
            # Tentar extrair quantidade e nome do formato gerado
            p = p.strip()
            for ing_nome in ingredientes_info.keys():
                if p.endswith(ing_nome):
                    qtd_str = p[:-len(ing_nome)].strip()
                    # Remover unidade do final se existir
                    unidade = ingredientes_info[ing_nome]["unidade"]
                    if unidade and qtd_str.endswith(unidade):
                        qtd_str = qtd_str[:-len(unidade)].strip()
                    try:
                        qtd_num = float(qtd_str)
                    except:
                        qtd_num = 1.0
                    ingredientes_atuais.append({"nome": ing_nome, "quantidade": qtd_num})
                    break

    ingredientes_atuais_json = json.dumps(ingredientes_atuais)

    return render_template_string(
        EDITAR_TEMPLATE,
        receita=receita,
        custos_js=custos_js,
        ingredientes_atuais_json=ingredientes_atuais_json
    )

@receitas_bp.route("/receitas/atualizar/<int:id>", methods=["POST"])
def atualizar_receita(id):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    nome = request.form.get("nome")
    categoria = request.form.get("categoria")
    descricao = request.form.get("descricao")
    preparacao = request.form.get("preparacao")
    notas = request.form.get("notas")
    tempo = int(request.form.get("tempo", 0) or 0)
    
    ingredientes_selecionados = request.form.getlist("ingrediente[]")
    quantidades_inseridas = request.form.getlist("quantidade[]")
    
    ingredientes_info = carregar_ingredientes_detalhes(user_email)
    
    lista_final = []
    custo_total_calculado = 0.0
    
    for ing, qtd in zip(ingredientes_selecionados, quantidades_inseridas):
        if ing:
            qtd_num = float(qtd) if qtd else 0.0
            unidade_ing = ingredientes_info[ing]["unidade"] if ing in ingredientes_info else ""
            lista_final.append(f"{qtd_num}{unidade_ing} {ing}")
            
            if ing in ingredientes_info:
                custo_total_calculado += ingredientes_info[ing]["custo_unitario"] * qtd_num
            
    ingredientes_str = " | ".join(lista_final)

    conn = get_db_connection()
    conn.execute('''
        UPDATE receitas
        SET nome = ?, categoria = ?, descricao = ?, preparacao = ?, notas = ?, tempo = ?, ingredientes_usados = ?, custo_total = ?
        WHERE id = ? AND user_email = ?
    ''', (nome, categoria, descricao, preparacao, notas, tempo, ingredientes_str, custo_total_calculado, id, user_email))
    conn.commit()
    conn.close()

    return redirect(url_for('receitas_bp.listar_receitas'))

@receitas_bp.route("/receitas/apagar/<int:id>", methods=["POST"])
def apagar_receita(id):
    user_email = session.get('user_email')
    if not user_email:
        return redirect('/')

    conn = get_db_connection()
    conn.execute('DELETE FROM receitas WHERE id = ? AND user_email = ?', (id, user_email))
    conn.commit()
    conn.close()

    return redirect(url_for('receitas_bp.listar_receitas'))