from flask import Blueprint, render_template_string, request, redirect, url_for, send_file, session
import pandas as pd
import os
from io import BytesIO
import urllib.parse

clientes_bp = Blueprint('clientes_bp', __name__)

def obter_ficheiro_utilizador(nome_base):
    email = session.get('user_email', 'admin')
    if email == 'cadoces023@gmail.com':
        return nome_base
    else:
        email_limpo = email.replace('@', '_at_').replace('.', '_')
        nome, ext = os.path.splitext(nome_base)
        return f"{nome}_{email_limpo}{ext}"

def carregar_dados(ficheiro_base, colunas_padrao):
    ficheiro = obter_ficheiro_utilizador(ficheiro_base)
    try:
        if not os.path.exists(ficheiro):
            df = pd.DataFrame(columns=colunas_padrao)
            df.to_excel(ficheiro, index=False)
            return df
        return pd.read_excel(ficheiro)
    except Exception:
        df = pd.DataFrame(columns=colunas_padrao)
        df.to_excel(ficheiro, index=False)
        return df

def guardar_dados(ficheiro_base, df):
    ficheiro = obter_ficheiro_utilizador(ficheiro_base)
    df.to_excel(ficheiro, index=False)

@clientes_bp.route("/clientes")
def gerir_clientes():
    df_cli = carregar_dados('clientes.xlsx', ['ID', 'Nome', 'Telefone', 'Email', 'Notas'])
    df_orc = carregar_dados('orcamentos.xlsx', ['ID', 'Cliente', 'Total'])
    df_enc = carregar_dados('encomendas.xlsx', ['ID', 'Cliente', 'Valor'])
    
    linhas_tabela = ""
    
    if not df_cli.empty:
        for idx, row in df_cli.iterrows():
            id_cli = row.get('ID', idx)
            nome = str(row.get('Nome', '')) if pd.notna(row.get('Nome')) else ''
            nome_limpo = nome.strip().lower()
            telefone = str(row.get('Telefone', '-')) if pd.notna(row.get('Telefone')) else '-'
            email = str(row.get('Email', '-')) if pd.notna(row.get('Email')) else '-'
            notas = str(row.get('Notas', '-')) if pd.notna(row.get('Notas')) else '-'
            
            pedidos_orc = 0
            total_gasto_orc = 0.0
            if not df_orc.empty and 'Cliente' in df_orc.columns:
                match_orc = df_orc[df_orc['Cliente'].astype(str).str.strip().str.lower() == nome_limpo]
                pedidos_orc = len(match_orc)
                if 'Total' in match_orc.columns:
                    total_gasto_orc = match_orc['Total'].fillna(0).astype(float).sum()
                
            pedidos_enc = 0
            total_gasto_enc = 0.0
            if not df_enc.empty and 'Cliente' in df_enc.columns:
                match_enc = df_enc[df_enc['Cliente'].astype(str).str.strip().str.lower() == nome_limpo]
                pedidos_enc = len(match_enc)
                if 'Valor' in match_enc.columns:
                    total_gasto_enc = match_enc['Valor'].fillna(0).astype(float).sum()
                
            total_pedidos = pedidos_orc + pedidos_enc
            total_gasto = total_gasto_orc + total_gasto_enc
            
            telefone_str = telefone if telefone and telefone != 'nan' else '-'
            email_str = email if email and email != 'nan' else '-'
            notas_str = notas if notas and notas != 'nan' else '-'
            
            linhas_tabela += f"""
            <tr class="linha-cliente" data-texto="{nome.lower()} {telefone_str.lower()} {email_str.lower()} {notas_str.lower()}">
                <td data-label="Nome">
                    <a href="/clientes/historico/{id_cli}" style="font-weight: 600; color: #b85d19; text-decoration: none;" title="Ver Histórico e Fidelização">
                        {nome} 📋
                    </a>
                </td>
                <td data-label="Telefone">{telefone_str}</td>
                <td data-label="Email">{email_str}</td>
                <td data-label="Notas">{notas_str}</td>
                <td data-label="Pedidos">📄 {total_pedidos}</td>
                <td data-label="Total Gasto"><strong style="color: #b85d19;">€{total_gasto:.2f}</strong></td>
                <td data-label="Ações" style="text-align: right;">
                    <button onclick="abrirModalEditar('{id_cli}', '{nome}', '{telefone_str}', '{email_str}', '{notas_str}')" style="background:none; border:none; cursor:pointer; font-size: 16px; margin-right: 8px;" title="Editar">✏️</button>
                    <a href="/clientes/apagar/{id_cli}" onclick="return confirm('Tens a certeza que queres apagar este cliente?')" style="text-decoration: none; font-size: 16px;" title="Apagar">🗑️</a>
                </td>
            </tr>
            """
            
    if not linhas_tabela:
        linhas_tabela = '<tr id="sem-resultados"><td colspan="7" style="text-align:center; color:#888; padding: 30px;">Nenhum cliente encontrado.</td></tr>'

    template = """
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Gestão de Clientes</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; color: #333; }
            .container { max-width: 1200px; margin: auto; }
            .header-bar { display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; gap: 15px; }
            .header-left { display: flex; align-items: center; gap: 20px; }
            .back-btn { text-decoration: none; color: #b85d19; font-size: 22px; font-weight: bold; background: #f9f0e6; padding: 8px 14px; border-radius: 8px; }
            h1 { margin: 0; color: #111; font-size: 24px; }
            p.sub { margin: 5px 0 0 0; color: #666; font-size: 14px; }
            .header-actions { display: flex; gap: 12px; flex-wrap: wrap; }
            .btn-export { background: #fff; border: 2px solid #f3e5d0; color: #333; padding: 10px 18px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; }
            .btn-export:hover { background: #f9f0e6; }
            .btn-primary { background: #dfa673; color: #fff; border: none; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; }
            .btn-primary:hover { background: #c8915f; }
            .search-box { margin-bottom: 25px; }
            .search-input { width: 100%; padding: 14px 20px; border: 2px solid #f3e5d0; border-radius: 10px; background: #fffdf9; font-size: 15px; box-sizing: border-box; outline: none; }
            .search-input:focus { border-color: #dfa673; }
            .content-box { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 10px 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
            table.table-custom { width: 100%; border-collapse: collapse; }
            table.table-custom th { background-color: transparent; color: #888; font-size: 13px; font-weight: bold; padding: 16px 12px; text-align: left; border-bottom: 2px solid #f3e5d0; text-transform: uppercase; letter-spacing: 0.5px; }
            table.table-custom td { padding: 16px 12px; border-bottom: 1px solid #f3e5d0; color: #444; font-size: 14px; }
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4); overflow-y: auto; padding: 40px 0; }
            .modal-content { background-color: #fffdf9; margin: auto; padding: 30px; border: 2px solid #f3e5d0; width: 90%; max-width: 500px; border-radius: 14px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); position: relative; box-sizing: border-box; }
            .close-btn { position: absolute; right: 20px; top: 20px; font-size: 22px; cursor: pointer; color: #888; font-weight: bold; }
            .form-group { margin-bottom: 18px; }
            .form-group label { display: block; font-weight: bold; color: #222; margin-bottom: 6px; font-size: 13px; }
            .form-control { width: 100%; padding: 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; box-sizing: border-box; outline: none; }
            .form-control:focus { border-color: #dfa673; }
            .modal-footer { display: flex; justify-content: flex-end; gap: 12px; border-top: 2px solid #f3e5d0; padding-top: 20px; margin-top: 20px; }
            .btn-secondary { background: #fff; border: 2px solid #f3e5d0; color: #444; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; }

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {
                body { padding: 15px; }
                .header-bar { flex-direction: column; align-items: stretch; text-align: center; }
                .header-left { flex-direction: column; gap: 10px; }
                .header-actions { justify-content: center; }
                table.table-custom, table.table-custom thead, table.table-custom tbody, table.table-custom th, table.table-custom td, table.table-custom tr {
                    display: block;
                }
                table.table-custom thead tr {
                    position: absolute;
                    top: -9999px;
                    left: -9999px;
                }
                table.table-custom tr {
                    background: #fff;
                    border: 2px solid #f3e5d0;
                    border-radius: 12px;
                    margin-bottom: 15px;
                    padding: 15px;
                    box-shadow: 0 2px 6px rgba(0,0,0,0.02);
                }
                table.table-custom td {
                    border: none !important;
                    position: relative;
                    padding-left: 55% !important;
                    text-align: right;
                    padding-top: 10px;
                    padding-bottom: 10px;
                    font-size: 14px;
                }
                table.table-custom td::before {
                    position: absolute;
                    left: 15px;
                    width: 45%;
                    padding-right: 10px;
                    white-space: nowrap;
                    text-align: left;
                    font-weight: bold;
                    color: #888;
                    content: attr(data-label);
                }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-bar">
                <div class="header-left">
                    <a href="/" class="back-btn">&#8592;</a>
                    <div>
                        <h1>Clientes</h1>
                        <p class="sub">Gestão de clientes e fidelização</p>
                    </div>
                </div>
                <div class="header-actions">
                    <a href="/clientes/exportar" class="btn-export">📥 Exportar</a>
                    <button class="btn-primary" onclick="abrirModalNovo()">+ Novo Cliente</button>
                </div>
            </div>

            <div class="search-box">
                <input type="text" id="inputPesquisa" class="search-input" placeholder="🔍 Pesquisar clientes..." onkeyup="filtrarClientesTempoReal()">
            </div>

            <div class="content-box">
                <table class="table-custom">
                    <thead>
                        <tr>
                            <th>Nome</th>
                            <th>Telefone</th>
                            <th>Email</th>
                            <th>Notas</th>
                            <th>Pedidos</th>
                            <th>Total Gasto</th>
                            <th style="text-align: right;">Ações</th>
                        </tr>
                    </thead>
                    <tbody id="tabelaClientesCorpo">
                        {{ linhas_tabela | safe }}
                    </tbody>
                </table>
            </div>
        </div>

        <div id="modalCliente" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModal()">&times;</span>
                <h2 id="tituloModal" style="margin-top:0; color:#111; font-size:20px;">Novo Cliente</h2>
                <form id="formCliente" action="/clientes/salvar" method="POST">
                    <input type="hidden" name="id" id="clienteId">
                    <div class="form-group">
                        <label>Nome *</label>
                        <input type="text" name="nome" id="clienteNome" class="form-control" placeholder="Nome do cliente" required>
                    </div>
                    <div class="form-group">
                        <label>Telefone</label>
                        <input type="text" name="telefone" id="clienteTelefone" class="form-control" placeholder="Ex: 912345678">
                    </div>
                    <div class="form-group">
                        <label>Email</label>
                        <input type="email" name="email" id="clienteEmail" class="form-control" placeholder="exemplo@email.com">
                    </div>
                    <div class="form-group">
                        <label>Notas</label>
                        <textarea name="notas" id="clienteNotas" class="form-control" rows="3" placeholder="Observações sobre o cliente..."></textarea>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModal()">Cancelar</button>
                        <button type="submit" class="btn-primary">Guardar Cliente</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function abrirModalNovo() {
                document.getElementById('tituloModal').innerText = 'Novo Cliente';
                document.getElementById('clienteId').value = '';
                document.getElementById('clienteNome').value = '';
                document.getElementById('clienteTelefone').value = '';
                document.getElementById('clienteEmail').value = '';
                document.getElementById('clienteNotas').value = '';
                document.getElementById('modalCliente').style.display = 'block';
            }

            function abrirModalEditar(id, nome, telefone, email, notas) {
                document.getElementById('tituloModal').innerText = 'Editar Cliente';
                document.getElementById('clienteId').value = id;
                document.getElementById('clienteNome').value = nome;
                document.getElementById('clienteTelefone').value = telefone !== '-' ? telefone : '';
                document.getElementById('clienteEmail').value = email !== '-' ? email : '';
                document.getElementById('clienteNotas').value = notas !== '-' ? notas : '';
                document.getElementById('modalCliente').style.display = 'block';
            }

            function fecharModal() {
                document.getElementById('modalCliente').style.display = 'none';
            }

            function filtrarClientesTempoReal() {
                let termo = document.getElementById('inputPesquisa').value.toLowerCase();
                let linhas = document.querySelectorAll('.linha-cliente');
                linhas.forEach(linha => {
                    let textoLinha = linha.getAttribute('data-texto');
                    if (textoLinha.includes(termo)) {
                        linha.style.display = '';
                    } else {
                        linha.style.display = 'none';
                    }
                });
            }
        </script>
    </body>
    </html>
    """
    return render_template_string(template, linhas_tabela=linhas_tabela)

@clientes_bp.route("/clientes/historico/<int:id_cli>")
def historico_cliente(id_cli):
    df_cli = carregar_dados('clientes.xlsx', ['ID', 'Nome', 'Telefone', 'Email', 'Notas'])
    cliente_res = df_cli[df_cli['ID'].astype(float) == float(id_cli)]
    
    if cliente_res.empty:
        return "Cliente não encontrado.", 404
        
    row_cli = cliente_res.iloc[0]
    nome_cliente = str(row_cli.get('Nome', 'Cliente')).strip()
    nome_limpo = nome_cliente.lower()
    
    telefone = row_cli.get('Telefone', '-')
    whatsapp_html = ""
    
    if pd.notna(telefone) and str(telefone).strip() != '' and str(telefone).strip() != '-':
        tel_str = str(telefone).strip()
        tel_limpo = ''.join(filter(str.isdigit, tel_str))
        if tel_limpo:
            if not tel_limpo.startswith('351') and len(tel_limpo) == 9:
                tel_limpo = '351' + tel_limpo
            msg_wpp = urllib.parse.quote(f"Olá {nome_cliente}, falamos a respeito da sua encomenda na CáDoces...")
            whatsapp_html = f' <a href="https://wa.me/{tel_limpo}?text={msg_wpp}" target="_blank" style="background: #25d366; color: white; padding: 4px 10px; border-radius: 6px; text-decoration: none; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 4px; margin-left: 10px;" title="Enviar mensagem via WhatsApp">💬 WhatsApp</a>'
            
    email = row_cli.get('Email', '-')
    if pd.isna(email) or str(email).lower() == 'nan' or not str(email).strip():
        email = '-'
        
    notas = row_cli.get('Notas', '-')
    if pd.isna(notas) or str(notas).lower() == 'nan' or not str(notas).strip():
        notas = 'Sem notas'
    
    df_orc = carregar_dados('orcamentos.xlsx', ['ID', 'Nome', 'Cliente', 'DataEntrega', 'Total', 'Estado'])
    df_enc = carregar_dados('encomendas.xlsx', ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome'])
    
    orcamentos_cli = df_orc[df_orc['Cliente'].astype(str).str.strip().str.lower() == nome_limpo] if not df_orc.empty and 'Cliente' in df_orc.columns else pd.DataFrame()
    encomendas_cli = df_enc[df_enc['Cliente'].astype(str).str.strip().str.lower() == nome_limpo] if not df_enc.empty and 'Cliente' in df_enc.columns else pd.DataFrame()
    
    pontos = 0
    if not orcamentos_cli.empty:
        for _, r in orcamentos_cli.iterrows():
            total_val = float(r.get('Total', 0) or 0)
            if total_val >= 20:
                pontos += 1
                
    if not encomendas_cli.empty:
        for _, r in encomendas_cli.iterrows():
            val_val = float(r.get('Valor', 0) or 0)
            if val_val >= 20:
                pontos += 1
                
    pontos_carimbo = min(10, pontos)
    premio_atingido = pontos >= 10
    
    carimbos_html = ""
    for i in range(1, 11):
        if i <= pontos_carimbo:
            carimbos_html += '<span style="font-size: 22px; margin-right: 6px;" title="Carimbo obtido">🟢</span>'
        else:
            carimbos_html += '<span style="font-size: 22px; margin-right: 6px; opacity: 0.3;" title="Em falta">⚪</span>'
    
    linhas_orc = ""
    for _, r in orcamentos_cli.iterrows():
        linhas_orc += f"""
        <tr>
            <td data-label="Nome do Orçamento"><strong>{r.get('Nome', 'Orçamento')}</strong></td>
            <td data-label="Data de Entrega">{r.get('DataEntrega', '-')}</td>
            <td data-label="Estado">{r.get('Estado', 'Pendente')}</td>
            <td data-label="Total"><strong style="color: #b85d19;">€{float(r.get('Total', 0)):.2f}</strong></td>
        </tr>
        """
    if not linhas_orc:
        linhas_orc = '<tr><td colspan="4" style="text-align:center; color:#888; padding: 20px;">Sem orçamentos registados para este cliente.</td></tr>'
        
    linhas_enc = ""
    for _, r in encomendas_cli.iterrows():
        linhas_enc += f"""
        <tr>
            <td data-label="Referência"><strong>Encomenda #{r.get('ID', '')} {r.get('Nome', 'Encomenda')}</strong></td>
            <td data-label="Data de Entrega">{r.get('DataEntrega', '-')}</td>
            <td data-label="Estado">{r.get('Estado', 'Aceite')}</td>
            <td data-label="Valor"><strong style="color: #b85d19;">€{float(r.get('Valor', 0)):.2f}</strong></td>
        </tr>
        """
    if not linhas_enc:
        linhas_enc = '<tr><td colspan="4" style="text-align:center; color:#888; padding: 20px;">Sem encomendas registadas para este cliente.</td></tr>'

    template_historico = """
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Histórico de Clientes - CáDoces</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; color: #333; }
            .container { max-width: 1000px; margin: auto; }
            .header-bar { display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; gap: 15px; }
            .back-btn { text-decoration: none; color: #b85d19; font-size: 16px; font-weight: bold; background: #f9f0e6; padding: 8px 16px; border-radius: 8px; border: 1px solid #f3e5d0; }
            .card { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 25px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
            h2 { margin-top: 0; color: #8c4410; font-size: 20px; }
            table.table-custom { width: 100%; border-collapse: collapse; margin-top: 10px; }
            table.table-custom th { background-color: #f9f0e6; color: #8c4410; padding: 12px; text-align: left; border-bottom: 2px solid #f3e5d0; font-size: 13px; }
            table.table-custom td { padding: 12px; border-bottom: 1px solid #f3e5d0; color: #444; font-size: 14px; }
            .info-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-top: 15px; font-size: 14px; align-items: center; }
            .fidelidade-box { background: {{ '#e8f5e9' if premio_atingido else '#fff9c4' }}; border: 2px solid {{ '#c8e6c9' if premio_atingido else '#ffee58' }}; border-radius: 12px; padding: 20px; margin-top: 20px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {
                body { padding: 15px; }
                .header-bar { flex-direction: column; text-align: center; }
                .info-grid { grid-template-columns: 1fr; }
                table.table-custom, table.table-custom thead, table.table-custom tbody, table.table-custom th, table.table-custom td, table.table-custom tr {
                    display: block;
                }
                table.table-custom thead tr {
                    position: absolute;
                    top: -9999px;
                    left: -9999px;
                }
                table.table-custom tr {
                    background: #fff;
                    border: 2px solid #f3e5d0;
                    border-radius: 12px;
                    margin-bottom: 15px;
                    padding: 15px;
                }
                table.table-custom td {
                    border: none !important;
                    position: relative;
                    padding-left: 55% !important;
                    text-align: right;
                    padding-top: 8px;
                    padding-bottom: 8px;
                }
                table.table-custom td::before {
                    position: absolute;
                    left: 15px;
                    width: 45%;
                    padding-right: 10px;
                    white-space: nowrap;
                    text-align: left;
                    font-weight: bold;
                    color: #888;
                    content: attr(data-label);
                }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-bar">
                <a href="/clientes" class="back-btn">&#8592; Voltar aos Clientes</a>
                <span style="font-size: 18px; font-weight: bold; color: #111;">Ficha de Cliente e Fidelização</span>
            </div>

            <div class="card">
                <h2>👤 {{ nome_cliente }}</h2>
                <div class="info-grid">
                    <div><strong>📞 Telefone:</strong> {{ telefone }} {{ whatsapp_html | safe }}</div>
                    <div><strong>📧 Email:</strong> {{ email }}</div>
                    <div><strong>📝 Notas:</strong> {{ notas }}</div>
                </div>

                <div class="fidelidade-box">
                    <div>
                        <strong>🎁 Cartão de Fidelização (Compras ≥ 20€ = 1 Carimbo)</strong>
                        <div style="margin-top: 8px;">
                            {{ carimbos_html | safe }} <span style="font-weight: bold; margin-left: 10px;">({{ pontos_carimbo }}/10)</span>
                        </div>
                    </div>
                    <div>
                        {% if premio_atingido %}
                            <span style="background: #2e7d32; color: #fff; padding: 8px 14px; border-radius: 8px; font-weight: bold; font-size: 13px;">🎉 PRÉMIO DESBLOQUEADO! Escolher: Bolo Caseirinho / 1kg Bolachas / Trança</span>
                        {% else %}
                            <span style="color: #666; font-size: 13px;">Faltam {{ 10 - pontos_carimbo }} compra(s) ≥ 20€ para o prémio.</span>
                        {% endif %}
                    </div>
                </div>
            </div>

            <div class="card">
                <h2>📄 Orçamentos do Cliente</h2>
                <table class="table-custom">
                    <thead>
                        <tr>
                            <th>Nome do Orçamento</th>
                            <th>Data de Entrega</th>
                            <th>Estado</th>
                            <th>Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        {{ linhas_orc | safe }}
                    </tbody>
                </table>
            </div>

            <div class="card">
                <h2>🛍️ Encomendas do Cliente</h2>
                <table class="table-custom">
                    <thead>
                        <tr>
                            <th>Referência</th>
                            <th>Data de Entrega</th>
                            <th>Estado</th>
                            <th>Valor</th>
                        </tr>
                    </thead>
                    <tbody>
                        {{ linhas_enc | safe }}
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(template_historico, nome_cliente=nome_cliente, telefone=telefone, whatsapp_html=whatsapp_html, email=email, notas=notas, carimbos_html=carimbos_html, pontos_carimbo=pontos_carimbo, premio_atingido=premio_atingido, linhas_orc=linhas_orc, linhas_enc=linhas_enc)

@clientes_bp.route("/clientes/salvar", methods=["POST"])
def salvar_cliente():
    id_cliente = request.form.get("id")
    nome = request.form.get("nome")
    telefone = request.form.get("telefone")
    email = request.form.get("email")
    notas = request.form.get("notas")
    
    df = carregar_dados('clientes.xlsx', ['ID', 'Nome', 'Telefone', 'Email', 'Notas'])
    
    if id_cliente and str(id_cliente).strip() != "":
        try:
            id_num = int(float(id_cliente))
            idx = df[df['ID'].astype(float) == id_num].index
            if not idx.empty:
                df.loc[idx, ['Nome', 'Telefone', 'Email', 'Notas']] = [nome, telefone, email, notas]
            else:
                novo_id = int(df['ID'].max()) + 1 if not df.empty and 'ID' in df.columns and not df['ID'].dropna().empty else 1
                nova_linha = {'ID': novo_id, 'Nome': nome, 'Telefone': telefone, 'Email': email, 'Notas': notas}
                df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        except Exception:
            pass
    else:
        novo_id = int(df['ID'].max()) + 1 if not df.empty and 'ID' in df.columns and not df['ID'].dropna().empty else 1
        nova_linha = {
            'ID': novo_id,
            'Nome': nome,
            'Telefone': telefone,
            'Email': email,
            'Notas': notas
        }
        df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        
    guardar_dados('clientes.xlsx', df)
    return redirect(url_for('clientes_bp.gerir_clientes'))

@clientes_bp.route("/clientes/apagar/<int:id_cli>")
def apagar_cliente(id_cli):
    df = carregar_dados('clientes.xlsx', ['ID', 'Nome', 'Telefone', 'Email', 'Notas'])
    if not df.empty:
        df = df[df['ID'] != id_cli]
        guardar_dados('clientes.xlsx', df)
    return redirect(url_for('clientes_bp.gerir_clientes'))

@clientes_bp.route("/clientes/exportar")
def exportar_clientes():
    df = carregar_dados('clientes.xlsx', ['ID', 'Nome', 'Telefone', 'Email', 'Notas'])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Clientes')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='clientes.xlsx')