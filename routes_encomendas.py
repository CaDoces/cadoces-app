from flask import Blueprint, render_template_string, request, redirect, url_for, jsonify, Response, send_file, session
import pandas as pd
import os
import datetime
import calendar
import json
import urllib.parse
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from io import BytesIO
from PIL import Image as PILImage

encomendas_bp = Blueprint('encomendas_bp', __name__)

EXCEL_ORCAMENTOS = 'orcamentos.xlsx'
EXCEL_ENCOMENDAS = 'encomendas.xlsx'
EXCEL_RECEITAS = 'receitas.xlsx'
EXCEL_INGREDIENTES = 'ingredientes.xlsx' 
EXCEL_MAO_OBRA = 'mao_de_obra.xlsx'
EXCEL_DESPESAS = 'despesas.xlsx'
EXCEL_CLIENTES = 'clientes.xlsx'

# --- AS TUAS FRASES PERSONALIZADAS PARA O WHATSAPP ---
FRASE_WHATSAPP_ADMIN = "Olá {cliente}! Segue em anexo a proposta de orçamento para '{nome}' . Agradeço desde já a sua preferência!"
FRASE_WHATSAPP_ALUNO = "Olá {cliente}! (Simulação Pedagógica) Proposta de orçamento para '{nome}'. Agradeço a sua preferência"

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
    if not os.path.exists(ficheiro):
        df = pd.DataFrame(columns=colunas_padrao)
        df.to_excel(ficheiro, index=False)
    df = pd.read_excel(ficheiro)
    for col in ['DataEntrega', 'Notas', 'Estado', 'Cliente', 'Nome', 'Receita', 'Material', 'MaoObra', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas']:
        if col in df.columns:
            df[col] = df[col].astype(object)
    return df

def guardar_dados(ficheiro_base, df):
    ficheiro = obter_ficheiro_utilizador(ficheiro_base)
    df.to_excel(ficheiro, index=False)

def obter_primeiro_valor_valido(row, colunas_possiveis):
    for col in colunas_possiveis:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != '':
            return str(row[col])
    for val in row.values:
        if pd.notna(val) and str(val).strip() != '':
            return str(val)
    return "Sem Nome"

@encomendas_bp.route("/api/item/detalhes")
def api_item_detalhes():
    tipo = request.args.get('tipo', '')
    nome = request.args.get('nome', '')
    
    custo = 0.0
    mao_obra_custo = 0.0
    
    if tipo == 'receita':
        df_rec = carregar_dados(EXCEL_RECEITAS, ['ID', 'Nome da Receita', 'Custo Ingredientes', 'Custo Total', 'Tempo'])
        col = next((c for c in ['Nome da Receita', 'Nome', 'Receita', 'Título'] if c in df_rec.columns), df_rec.columns[0])
        res = df_rec[df_rec[col].astype(str) == nome]
        if not res.empty:
            r = res.iloc[0]
            custo = float(r.get('Custo Ingredientes', r.get('Custo Total', 5.00)) or 5.00)
            tempo_min = float(r.get('Tempo', r.get('Tempo (min)', 30)) or 30)
            mao_obra_custo = (tempo_min / 60) * 8.00
            
    elif tipo == 'material':
        df_ing = carregar_dados(EXCEL_INGREDIENTES, ['ID', 'Nome', 'Custo Unitario', 'Preço'])
        col = next((c for c in ['Nome', 'Ingrediente', 'Designação', 'Material'] if c in df_ing.columns), df_ing.columns[0])
        res = df_ing[df_ing[col].astype(str) == nome]
        if not res.empty:
            r = res.iloc[0]
            custo = float(r.get('Custo Unitario', r.get('Preço', 1.00)) or 1.00)
            
    elif tipo == 'mao_obra':
        df_mob = carregar_dados(EXCEL_MAO_OBRA, ['ID', 'Descricao', 'Valor'])
        col = next((c for c in ['Descricao', 'Descrição', 'Funcao', 'Função', 'Nome'] if c in df_mob.columns), df_mob.columns[0])
        res = df_mob[df_mob[col].astype(str) == nome]
        if not res.empty:
            r = res.iloc[0]
            custo = float(r.get('Valor', 8.00) or 8.00)
            
    return jsonify({'sucesso': True, 'custo': custo, 'mao_obra': mao_obra_custo})

@encomendas_bp.route("/encomendas/folha_cozinha")
def folha_cozinha():
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
    df_cli = carregar_dados(EXCEL_CLIENTES, ['ID', 'Nome', 'Telefone', 'Contacto'])
    
    if not df_enc.empty and not df_cli.empty:
        df_enc['Cliente_lower'] = df_enc['Cliente'].astype(str).str.strip().str.lower()
        df_cli['Nome_lower'] = df_cli['Nome'].astype(str).str.strip().str.lower()
        df_enc = pd.merge(df_enc, df_cli[['Nome_lower', 'Telefone']].rename(columns={'Telefone': 'TelefoneCli'}), left_on='Cliente_lower', right_on='Nome_lower', how='left')

    if not df_enc.empty:
        df_enc = df_enc[~df_enc['Estado'].astype(str).str.lower().isin(['cancelado'])]
        df_enc = df_enc.sort_values(by='DataEntrega', na_position='last')

    linhas_folha = ""
    if not df_enc.empty:
        for idx, row in df_enc.iterrows():
            cliente = row.get('Cliente', 'Cliente')
            nome_bolo = row.get('Nome', 'Encomenda')
            data_entrega = str(row.get('DataEntrega', '-'))[:10]
            estado = row.get('Estado', 'Aceite')
            notas = row.get('NotasAlergenios', '')
            num_pessoas = row.get('NumPessoas', '-')
            if pd.isna(num_pessoas) or str(num_pessoas).strip() == '' or str(num_pessoas).strip() == 'nan':
                num_pessoas = 'A definir'
            else:
                num_pessoas = f"{num_pessoas} pessoas"

            telefone = row.get('TelefoneCli', '-')
            if pd.isna(telefone) or str(telefone).strip() == '' or str(telefone).strip() == 'nan':
                telefone = 'Não registado'

            itens_str = str(row.get('Itens', '[]'))
            try:
                itens_lista = json.loads(itens_str) if itens_str and itens_str != 'nan' else []
            except:
                itens_lista = []

            receitas_str = "A definir / Personalizada"
            materiais_str = "Caixa padrão / Base MDF"
            
            lista_rec = [item.get('nome') for item in itens_lista if item.get('tipo') == 'receita']
            lista_mat = [item.get('nome') for item in itens_lista if item.get('tipo') == 'material']
            
            if lista_rec:
                receitas_str = ", ".join(lista_rec)
            if lista_mat:
                materiais_str = ", ".join(lista_mat)

            linhas_folha += f"""
            <tr>
                <td data-label="Data & Hora" style="font-weight: bold; text-align: center;">
                    {data_entrega}
                    <div style="font-size: 11px; color: #b85d19; margin-top: 2px;">🕒 A definir</div>
                </td>
                <td data-label="Cliente & Contacto"><strong>{cliente}</strong><br><span style="font-size: 12px; color: #555;">📞 {telefone}</span></td>
                <td data-label="Detalhes & Receita">
                    <strong>{nome_bolo}</strong><br>
                    <span style="font-size: 12px; color: #444;">
                        🍰 <b>Receita / Massa:</b> {receitas_str}<br>
                        👥 <b>Fatias / Pessoas:</b> {num_pessoas}
                    </span>
                </td>
                <td data-label="Embalagem & Conservação">
                    <span style="font-size: 12px; color: #333; display: block; margin-bottom: 2px;">📦 {materiais_str}</span>
                    <span style="font-size: 12px; color: #c0392b; display: block;">🌡️ Cuidado / Frio</span>
                </td>
                <td data-label="Estado"><span style="padding: 3px 8px; background: #f3e5d0; border-radius: 4px; font-weight: bold; font-size: 12px;">{estado}</span></td>
                <td data-label="Alergénios / Notas" style="font-size: 12px;">⚠️ {notas if pd.notna(notas) and str(notas).strip() != '' else 'Nenhuma'}</td>
                <td data-label="Progresso" style="text-align: center; font-size: 13px; white-space: nowrap;">
                    <span style="display: block; margin-bottom: 3px;">☐ Massa</span>
                    <span style="display: block; margin-bottom: 3px;">☐ Recheio</span>
                    <span style="display: block;">☐ Embalado</span>
                </td>
            </tr>
            """
    else:
        linhas_folha = '<tr><td colspan="7" style="text-align: center; color: #888; padding: 30px;">Sem encomendas ativas para produzir.</td></tr>'

    html_folha = f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Folha de Cozinha - CáDoces</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fff; margin: 0; padding: 25px; color: #333; }}
            .container {{ max-width: 1150px; margin: auto; }}
            .no-print {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; }}
            .btn-voltar {{ text-decoration: none; color: #b85d19; background: #f9f0e6; padding: 8px 16px; border-radius: 8px; font-weight: bold; border: 1px solid #f3e5d0; }}
            .btn-imprimir {{ background: #dfa673; color: white; border: none; padding: 10px 20px; border-radius: 8px; font-weight: bold; cursor: pointer; font-size: 14px; }}
            h1 {{ color: #8c4410; margin-bottom: 5px; }}
            p.sub {{ color: #666; margin-top: 0; margin-bottom: 20px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th, td {{ border: 1px solid #ddd; padding: 10px; font-size: 12.5px; text-align: left; vertical-align: top; }}
            th {{ background-color: #f9f0e6; color: #8c4410; }}
            @media screen and (max-width: 768px) {{
                body {{ padding: 15px; }}
                .no-print {{ flex-direction: column; gap: 10px; align-items: stretch; text-align: center; }}
                table, thead, tbody, th, td, tr {{ display: block; }}
                thead tr {{ position: absolute; top: -9999px; left: -9999px; }}
                tr {{ background: #fff; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 15px; padding: 15px; }}
                td {{ border: none !important; position: relative; padding-left: 50% !important; text-align: right; padding-top: 8px; padding-bottom: 8px; }}
                td::before {{ position: absolute; left: 15px; width: 45%; padding-right: 10px; white-space: nowrap; text-align: left; font-weight: bold; color: #888; content: attr(data-label); }}
            }}
            @media print {{
                .no-print {{ display: none; }}
                body {{ padding: 0; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="no-print">
                <a href="/encomendas?tab=encomendas" class="btn-voltar">&#8592; Voltar às Encomendas</a>
                <button onclick="window.print()" class="btn-imprimir">🖨️ Imprimir Folha de Cozinha</button>
            </div>

            <h1>🧁 CáDoces - Folha de Produção para a Cozinha</h1>
            <p class="sub">Plano detalhado de fabrico, receitas, embalagens do orçamento e progresso.</p>

            <table>
                <thead>
                    <tr>
                        <th>Data & Hora</th>
                        <th>Cliente & Contacto</th>
                        <th>Detalhes & Receita</th>
                        <th>Embalagem & Conservação</th>
                        <th>Estado</th>
                        <th>Alergénios / Notas</th>
                        <th style="text-align: center;">Progresso</th>
                    </tr>
                </thead>
                <tbody>
                    {linhas_folha}
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return render_template_string(html_folha)

@encomendas_bp.route("/encomendas/adicionar_calendario/<int:id_orcamento>")
def adicionar_calendario_ics(id_orcamento):
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'Contacto', 'DataEntrega', 'Total'])
    orcamento = df_orc[df_orc['ID'].astype(str) == str(id_orcamento)]
    
    if orcamento.empty:
        return "Orçamento não encontrado.", 404
        
    row = orcamento.iloc[0]
    cliente = row.get('Cliente', 'Cliente CáDoces')
    nome_bolo = row.get('Nome', 'Encomenda CáDoces')
    valor = row.get('Total', 0)
    
    data_str = str(row.get('DataEntrega', ''))
    if not data_str or data_str == 'nan' or data_str == 'None':
        data_formatada = datetime.datetime.now().strftime('%Y%m%d')
    else:
        data_formatada = data_str.replace('-', '')

    ics_conteudo = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//CáDoces//Gestão de Encomendas//PT
BEGIN:VEVENT
SUMMARY:🎂 Encomenda CáDoces - {cliente} ({nome_bolo})
DESCRIPTION:Encomenda para {cliente}. Valor total: €{float(valor):.2f}
DTSTART;VALUE=DATE:{data_formatada}
DTEND;VALUE=DATE:{data_formatada}
END:VEVENT
END:VCALENDAR"""

    response = Response(ics_conteudo, mimetype="text/calendar")
    response.headers["Content-Disposition"] = f"attachment; filename=encomenda_{id_orcamento}.ics"
    return response

@encomendas_bp.route("/encomendas/aprovar_orcamento/<int:id_orcamento>")
def aprovar_orcamento(id_orcamento):
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'DataEntrega', 'Total', 'Estado', 'Itens', 'NotasAlergenios', 'NumPessoas'])
    orcamento = df_orc[df_orc['ID'].astype(str) == str(id_orcamento)]
    
    if not orcamento.empty:
        row = orcamento.iloc[0]
        df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
        novo_id_enc = int(df_enc['ID'].max() + 1) if not df_enc.empty and 'ID' in df_enc.columns and pd.notna(df_enc['ID'].max()) else 1
        
        nova_encomenda = pd.DataFrame([{
            'ID': novo_id_enc,
            'Cliente': row.get('Cliente', 'Cliente'),
            'DataEntrega': row.get('DataEntrega', ''),
            'Estado': 'Aceite',
            'Valor': row.get('Total', 0),
            'Nome': row.get('Nome', 'Encomenda'),
            'Itens': row.get('Itens', '[]'),
            'SinalPago': 'Pendente',
            'RestantePago': 'Pendente',
            'NotasAlergenios': row.get('NotasAlergenios', ''),
            'NumPessoas': row.get('NumPessoas', '')
        }])
        df_enc = pd.concat([df_enc, nova_encomenda], ignore_index=True)
        guardar_dados(EXCEL_ENCOMENDAS, df_enc)
        
        df_orc = df_orc[df_orc['ID'].astype(str) != str(id_orcamento)]
        guardar_dados(EXCEL_ORCAMENTOS, df_orc)
        
    return redirect(url_for('encomendas_bp.gerir_encomendas', tab='encomendas'))

@encomendas_bp.route("/encomendas/duplicar/<int:id_encomenda>")
def duplicar_encomenda(id_encomenda):
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
    encomenda = df_enc[df_enc['ID'].astype(str) == str(id_encomenda)]
    
    if encomenda.empty:
        return "Encomenda não encontrada.", 404
        
    row = encomenda.iloc[0]
    novo_id_enc = int(df_enc['ID'].max() + 1) if not df_enc.empty and 'ID' in df_enc.columns and pd.notna(df_enc['ID'].max()) else 1
    
    nova_encomenda = pd.DataFrame([{
        'ID': novo_id_enc,
        'Cliente': row.get('Cliente', 'Cliente'),
        'DataEntrega': datetime.date.today().strftime('%Y-%m-%d'),
        'Estado': 'Aceite',
        'Valor': row.get('Valor', 0),
        'Nome': f"Repetição: {row.get('Nome', 'Encomenda')}",
        'Itens': row.get('Itens', '[]'),
        'SinalPago': 'Pendente',
        'RestantePago': 'Pendente',
        'NotasAlergenios': row.get('NotasAlergenios', ''),
        'NumPessoas': row.get('NumPessoas', '')
    }])
    
    df_enc = pd.concat([df_enc, nova_encomenda], ignore_index=True)
    guardar_dados(EXCEL_ENCOMENDAS, df_enc)
    
    return redirect(url_for('encomendas_bp.gerir_encomendas', tab='encomendas'))

@encomendas_bp.route("/encomendas/atualizar_estado/<int:id_encomenda>", methods=['POST'])
def atualizar_estado_encomenda(id_encomenda):
    novo_estado = request.form.get('estado')
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
    mask = df_enc['ID'].astype(str) == str(id_encomenda)
    if any(mask):
        df_enc.loc[mask, 'Estado'] = novo_estado
        guardar_dados(EXCEL_ENCOMENDAS, df_enc)
    return redirect(url_for('encomendas_bp.gerir_encomendas', tab='encomendas'))

@encomendas_bp.route("/encomendas/atualizar_pagamento/<int:id_encomenda>", methods=['POST'])
def atualizar_pagamento_encomenda(id_encomenda):
    sinal = request.form.get('sinal_pago')
    restante = request.form.get('restante_pago')
    
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
    mask = df_enc['ID'].astype(str) == str(id_encomenda)
    if any(mask):
        if sinal: df_enc.loc[mask, 'SinalPago'] = sinal
        if restante: df_enc.loc[mask, 'RestantePago'] = restante
        guardar_dados(EXCEL_ENCOMENDAS, df_enc)
    return redirect(url_for('encomendas_bp.gerir_encomendas', tab='encomendas'))

@encomendas_bp.route("/encomendas/editar_encomenda/<int:id_encomenda>", methods=['POST'])
def editar_encomenda(id_encomenda):
    novo_nome = request.form.get('nome')
    novo_cliente = request.form.get('cliente')
    nova_data = request.form.get('data_entrega')
    novo_valor = request.form.get('total')
    novas_notas = request.form.get('notas_alergenios')
    
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
    mask = df_enc['ID'].astype(str) == str(id_encomenda)
    if any(mask):
        if novo_nome: df_enc.loc[mask, 'Nome'] = novo_nome
        if novo_cliente: df_enc.loc[mask, 'Cliente'] = novo_cliente
        if nova_data: df_enc.loc[mask, 'DataEntrega'] = nova_data
        if novo_valor: 
            val_limpo = str(novo_valor).replace(',', '.')
            df_enc.loc[mask, 'Valor'] = float(val_limpo) if val_limpo else 0.0
        if novas_notas is not None: df_enc.loc[mask, 'NotasAlergenios'] = novas_notas
        guardar_dados(EXCEL_ENCOMENDAS, df_enc)
    return redirect(url_for('encomendas_bp.gerir_encomendas', tab='encomendas'))

@encomendas_bp.route("/encomendas/apagar/<int:id_encomenda>")
def apagar_encomenda(id_encomenda):
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Nome', 'Itens', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
    df_enc = df_enc[df_enc['ID'].astype(str) != str(id_encomenda)]
    guardar_dados(EXCEL_ENCOMENDAS, df_enc)
    return redirect(url_for('encomendas_bp.gerir_encomendas', tab='encomendas'))

@encomendas_bp.route("/encomendas/apagar_orcamento/<int:id_orcamento>")
def apagar_orcamento(id_orcamento):
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'Contacto', 'Notas', 'Estado', 'Total', 'Receita', 'Material', 'MaoObra', 'DespesaExtra', 'DataEntrega', 'Itens', 'NotasAlergenios', 'NumPessoas'])
    df_orc = df_orc[df_orc['ID'].astype(str) != str(id_orcamento)]
    guardar_dados(EXCEL_ORCAMENTOS, df_orc)
    return redirect(url_for('encomendas_bp.gerir_encomendas', tab='orcamentos'))

@encomendas_bp.route("/encomendas/pdf/<int:id_orcamento>")
def exportar_orcamento_pdf(id_orcamento):
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'DataEntrega', 'Total', 'Estado', 'NotasAlergenios', 'NumPessoas'])
    orcamento = df_orc[df_orc['ID'].astype(str) == str(id_orcamento)]
    
    if orcamento.empty:
        return "Orçamento não encontrado.", 404
        
    row = orcamento.iloc[0]
    cliente = row.get('Cliente', 'Cliente CáDoces')
    nome_orcamento = row.get('Nome', 'Proposta CáDoces')
    data_entrega = row.get('DataEntrega', '-')
    total = float(row.get('Total', 0))
    num_pessoas = row.get('NumPessoas', '-')
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elementos = []
    
    styles = getSampleStyleSheet()
    
    estilo_sub = ParagraphStyle(
        'SubDoc',
        parent=styles['Normal'],
        fontSize=13,
        textColor=colors.HexColor('#8c4410'),
        spaceAfter=12,
        fontName='Helvetica-Bold'
    )
    
    estilo_seccao = ParagraphStyle(
        'SecaoDoc',
        parent=styles['Heading2'],
        fontSize=11,
        textColor=colors.HexColor('#8c4410'),
        spaceBefore=10,
        spaceAfter=4,
        fontName='Helvetica-Bold'
    )
    
    estilo_texto = ParagraphStyle(
        'TextoNormal',
        parent=styles['Normal'],
        fontSize=8.5,
        textColor=colors.HexColor('#333333'),
        spaceAfter=3,
        leading=11.5
    )

    estilo_rodape = ParagraphStyle(
        'RodapeDoc',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor('#777777'),
        alignment=1,
        leading=11
    )

    caminho_logo = None
    for ext in ['jpg', 'jpeg', 'png']:
        path_teste = os.path.join('static', f'logo.{ext}')
        if os.path.exists(path_teste):
            caminho_logo = path_teste
            break

    if caminho_logo:
        try:
            pil_img = PILImage.open(caminho_logo)
            img_w, img_h = pil_img.size
            altura_max = 55
            largura_proporcional = (img_w / img_h) * altura_max
            if largura_proporcional > 80:
                largura_proporcional = 80
                altura_max = (img_h / img_w) * largura_proporcional
        except Exception:
            largura_proporcional = 55
            altura_max = 55

        logo_img = Image(caminho_logo, width=largura_proporcional, height=altura_max)
        logo_img.hAlign = 'CENTER'
        
        tabela_moldura_logo = Table([[logo_img]], colWidths=[95], rowHeights=[65])
        tabela_moldura_logo.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#ffffff')),
            ('BOX', (0,0), (-1,-1), 1.5, colors.HexColor('#dfa673')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        
        titulo_doc = Paragraph("Proposta de Orçamento / Encomenda Personalizada", estilo_sub)
        tabela_topo = Table([[tabela_moldura_logo, titulo_doc]], colWidths=[95, 439])
        tabela_topo.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))
        elementos.append(tabela_topo)
    else:
        elementos.append(Paragraph("Proposta de Orçamento / Encomenda Personalizada", estilo_sub))

    elementos.append(Spacer(1, 4))
    
    dados_cliente = [
        [Paragraph("<b>Nome do(a) Cliente:</b>", estilo_texto), Paragraph(str(cliente), estilo_texto)],
        [Paragraph("<b>Data do Evento / Entrega:</b>", estilo_texto), Paragraph(str(data_entrega), estilo_texto)],
        [Paragraph("<b>Tema / Projeto:</b>", estilo_texto), Paragraph(str(nome_orcamento), estilo_texto)],
        [Paragraph("<b>Número de Pessoas / Fatias:</b>", estilo_texto), Paragraph(str(num_pessoas) if pd.notna(num_pessoas) and str(num_pessoas).strip() != '' else 'Não especificado', estilo_texto)],
        [Paragraph("<b>Data do Pedido:</b>", estilo_texto), Paragraph(datetime.date.today().strftime('%Y-%m-%d'), estilo_texto)],
    ]
    
    tabela_cliente = Table(dados_cliente, colWidths=[140, 394])
    tabela_cliente.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#fffdf9')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#f3e5d0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
    ]))
    
    elementos.append(tabela_cliente)
    elementos.append(Spacer(1, 6))
    
    elementos.append(Paragraph("Detalhes da Escolha do Cliente", estilo_seccao))
    
    detalhes_escolha = (
        f"<b>Projeto / Tema:</b> {nome_orcamento}<br/>"
        f"<b>Para quantas pessoas:</b> {num_pessoas if pd.notna(num_pessoas) and str(num_pessoas).strip() != '' else 'A definir'}<br/>"
        "<b>Massa:</b> A definir / Personalizada<br/>"
        "<b>Recheio:</b> A definir / Personalizado<br/>"
        "<b>Cobertura:</b> Buttercream liso / Decoração temática"
    )
    
    notas_alergenios_val = row.get('NotasAlergenios', '')
    if pd.notna(notas_alergenios_val) and str(notas_alergenios_val).strip() != '':
        detalhes_escolha += f"<br/><b>Notas e Alergénios:</b> {str(notas_alergenios_val)}"

    dados_valores = [
        [Paragraph(detalhes_escolha, estilo_texto), Paragraph(f"<b>Total<br/><font size=13 color='#b85d19'>€{total:.2f}</font></b>", ParagraphStyle('Val', parent=estilo_texto, alignment=1))]
    ]
    
    tabela_valores = Table(dados_valores, colWidths=[404, 130])
    tabela_valores.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#fffdf9')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#f3e5d0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    
    elementos.append(tabela_valores)
    elementos.append(Spacer(1, 6))

    elementos.append(Paragraph("Normas, Condições e Cuidados", estilo_seccao))
    
    texto_intro = "Cada bolo é confeccionado artesanalmente, com dedicação e atenção ao detalhe, para tornar cada momento ainda mais especial. Para garantir a melhor experiência, agradeço a leitura das seguintes condições."
    elementos.append(Paragraph(texto_intro, estilo_texto))
    
    texto_normas = (
        "<b>Encomendas e prazos</b><br/>"
        "As encomendas devem ser feitas com pelo menos 7 dias de antecedência, estando sujeitas à disponibilidade da agenda. Os orçamentos têm validade limitada, indicada no momento do envio.<br/><br/>"
        "<b>Pagamentos</b><br/>"
        "A confirmação da encomenda requer o pagamento de um sinal de 50% do valor total. O valor restante é pago no momento do levantamento ou da entrega, através de MB Way ou dinheiro. A data só fica reservada após a confirmação do pagamento do sinal.<br/><br/>"
        "<b>Levantamento e transporte</b><br/>"
        "As encomendas destinam-se, por norma, a levantamento no local, sendo a morada indicada após a confirmação. As entregas podem ser combinadas pontualmente, mediante disponibilidade e aplicação de uma taxa adicional.<br/>"
        "Após o levantamento, o transporte do bolo é da responsabilidade do cliente. Para evitar danos:<br/>"
        "&bull; Transporte o bolo numa superfície plana, estável e nivelada, preferencialmente no chão do veículo ou numa bagageira plana.<br/>"
        "&bull; Mantenha o veículo fresco e evite a exposição solar direta, travagens bruscas e curvas acentuadas.<br/>"
        "&bull; Não transporte o bolo ao colo nem coloque a caixa inclinada.<br/>"
        "&bull; Evite movimentar ou abrir a caixa durante o transporte.<br/><br/>"
        "<b>Conservação e armazenamento</b><br/>"
        "Para preservar a qualidade, o sabor e a decoração do bolo, é importante seguir os cuidados de conservação indicados no momento do levantamento.<br/>"
        "&bull; Quando necessário, conserve o bolo no frigorífico, especialmente se contiver natas, queijo-creme, fruta fresca ou outros recheios perecíveis.<br/>"
        "&bull; Mantenha-o dentro da caixa, protegido de odores fortes e de possíveis contaminações.<br/>"
        "&bull; Evite a exposição ao calor, à luz solar direta e a mudanças bruscas de temperatura.<br/>"
        "&bull; Não deixe o bolo à temperatura ambiente durante períodos prolongados, sobretudo nos dias mais quentes.<br/>"
        "&bull; Siga as indicações fornecidas pela CáDoces quanto à conservação e ao momento ideal para servir.<br/>"
        "<b>Atenção:</b> o calor, a humidade e o manuseamento inadequado podem comprometer a estabilidade e o aspeto da decoração. A CáDoces não se responsabiliza por danos resultantes de transporte, armazenamento ou conservação inadequados após o levantamento.<br/><br/>"
        "<b>Cancelamentos</b><br/>"
        "Em caso de cancelamento com menos de 48 horas de antecedência, o sinal não será devolvido, nos termos aplicáveis. Caso o cancelamento se deva a um imprevisto da responsabilidade da CáDoces, o valor pago será integralmente devolvido.<br/><br/>"
        "<b>Personalização</b><br/>"
        "Todos os bolos são personalizados e confecionados artesanalmente. Podem ocorrer ligeiras diferenças em relação às referências apresentadas, devido à natureza do trabalho manual, à disponibilidade dos materiais ou às condições climatéricas. Cada criação é única e feita especialmente para si.<br/><br/>"
        "<b>Embalagem</b><br/>"
        "Todos os bolos altos/decorados incluem base de MDF reforçada e caixa alta, adequadas ao seu acondicionamento e transporte. Ainda assim, é essencial manuseá-los com cuidado.<br/><br/>"
        "<b>Alergénios</b><br/>"
        "Os produtos podem conter alergénios, nomeadamente glúten, leite, ovos e frutos de casca rija, ou estar sujeitos a contaminação cruzada, dependendo dos ingredientes e das condições de preparação. Informe sempre sobre quaisquer alergias ou intolerâncias antes de efetuar a encomenda.<br/><br/>"
        "Obrigada por escolher a CáDoces e por confiar numa criação feita com dedicação para os seus momentos especiais.<br/><br/>"
        "<b>Doces que contam histórias. </b>"
    )
    elementos.append(Paragraph(texto_normas, estilo_texto))
    
    elementos.append(Spacer(1, 10))
    texto_rodape_final = (
        "CáDoces – Cake Design (Braga) &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Instagram: @cadoces_23<br/>"
        "Email: cadoces023@gmail.com &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Telemóvel: +351 910 313 162"
    )
    
    tabela_rodape = Table([[Paragraph(texto_rodape_final, estilo_rodape)]], colWidths=[534])
    tabela_rodape.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#fffdf9')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#f3e5d0')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ]))
    elementos.append(tabela_rodape)
    
    doc.build(elementos)
    buffer.seek(0)
    
    return send_file(
        buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=f"orcamento_cadoces_{id_orcamento}.pdf"
    )

@encomendas_bp.route("/encomendas/cronograma")
def cronograma_cozinha():
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'DataEntrega', 'Estado', 'Total'])
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor'])
    
    tarefas = []
    
    for _, r in df_orc.iterrows():
        d_str = str(r.get('DataEntrega', ''))
        if d_str and d_str != 'nan' and d_str != 'None':
            try:
                data_entrega = datetime.datetime.strptime(d_str[:10], '%Y-%m-%d').date()
                nome_item = str(r.get('Nome', 'Orçamento')).lower()
                cliente = r.get('Cliente', 'Cliente')
                nome_original = r.get('Nome', 'Orçamento')
                
                if any(k in nome_item for k in ['levedada', 'babka', 'panettone', 'coroa', 'massa']):
                    d_inicio = data_entrega - datetime.timedelta(days=2)
                    tarefas.append({
                        'data': d_inicio.strftime('%Y-%m-%d'),
                        'tipo': 'Fermentação / Massa Levedada',
                        'descricao': f'Iniciar fermentação longa e preparações para: <strong>{nome_original}</strong> ({cliente})'
                    })
                    tarefas.append({
                        'data': data_entrega.strftime('%Y-%m-%d'),
                        'tipo': 'Cozer e Expedir',
                        'descricao': f'Cozer e expedir <strong>{nome_original}</strong> para o cliente <strong>{cliente}</strong>'
                    })
                elif any(k in nome_item for k in ['decorado', 'aniversario', 'casamento', 'batizado', 'infantil', 'andar']):
                    d_3d = data_entrega - datetime.timedelta(days=3)
                    d_1d = data_entrega - datetime.timedelta(days=1)
                    tarefas.append({
                        'data': d_3d.strftime('%Y-%m-%d'),
                        'tipo': 'Preparação & Decoração (3 dias)',
                        'descricao': f'Início de preparações, massas e decorações para bolo decorado: <strong>{nome_original}</strong> ({cliente})'
                    })
                    tarefas.append({
                        'data': d_1d.strftime('%Y-%m-%d'),
                        'tipo': 'Recheio e Montagem',
                        'descricao': f'Recheio, montagem e estrutura do bolo: <strong>{nome_original}</strong> ({cliente})'
                    })
                    tarefas.append({
                        'data': data_entrega.strftime('%Y-%m-%d'),
                        'tipo': 'Entrega Final',
                        'descricao': f'🎂 Entrega final de <strong>{nome_original}</strong> ao cliente <strong>{cliente}</strong>'
                    })
                else:
                    tarefas.append({
                        'data': data_entrega.strftime('%Y-%m-%d'),
                        'tipo': 'Produção e Entrega (No próprio dia)',
                        'descricao': f'Preparar, cozer e expedir (Bolo caseirinho/Muffins): <strong>{nome_original}</strong> para <strong>{cliente}</strong>'
                    })
            except Exception:
                pass

    tarefas = sorted(tarefas, key=lambda x: x['data'])
    
    linhas_cronograma = ""
    for t in tarefas:
        badge_cor = "#b85d19"
        if 'Fermentação' in t['tipo'] or 'Preparação' in t['tipo']: badge_cor = "#2980b9"
        elif 'Montagem' in t['tipo'] or 'Recheio' in t['tipo']: badge_cor = "#d35400"
        elif 'Entrega' in t['tipo'] or 'Cozer' in t['tipo']: badge_cor = "#27ae60"
        
        linhas_cronograma += f"""
        <tr>
            <td data-label="Data" style="font-weight: bold; width: 130px;">{t['data']}</td>
            <td data-label="Fase de Trabalho" style="width: 220px;"><span style="background: {badge_cor}; color: white; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: bold;">{t['tipo']}</span></td>
            <td data-label="Tarefa / Descrição">{t['descricao']}</td>
        </tr>
        """
        
    if not linhas_cronograma:
        linhas_cronograma = '<tr><td colspan="3" style="text-align: center; color: #888; padding: 30px;">Sem tarefas agendadas no cronograma.</td></tr>'

    html_cronograma = f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Cronograma de Cozinha - CáDoces</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; color: #333; }}
            .container {{ max-width: 1000px; margin: auto; }}
            .header-bar {{ display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; gap: 15px; flex-wrap: wrap; }}
            .back-btn {{ text-decoration: none; color: #b85d19; font-size: 16px; font-weight: bold; background: #f9f0e6; padding: 8px 16px; border-radius: 8px; border: 1px solid #f3e5d0; }}
            .content-box {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
            table.table-custom {{ width: 100%; border-collapse: collapse; }}
            table.table-custom th {{ background-color: #f9f0e6; color: #8c4410; padding: 14px; text-align: left; border-bottom: 2px solid #f3e5d0; font-size: 14px; }}
            table.table-custom td {{ padding: 14px; border-bottom: 1px solid #f3e5d0; color: #444; font-size: 14px; }}

            @media screen and (max-width: 768px) {{
                body {{ padding: 15px; }}
                .header-bar {{ flex-direction: column; text-align: center; }}
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
                    padding-left: 50% !important;
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
                <a href="/encomendas?tab=orcamentos" class="back-btn">&#8592; Voltar às Encomendas</a>
                <span style="font-size: 18px; font-weight: bold; color: #111;">📋 Cronograma Inteligente de Cozinha</span>
            </div>

            <div class="content-box">
                <p style="margin-top: 0; color: #666; font-size: 14px; margin-bottom: 20px;">Plano adaptado automaticamente: <strong>3 dias antes</strong> para bolos decorados, <strong>fermentação prévia</strong> para massas levedadas, e produção no <strong>próprio dia</strong> para caseirinhos.</p>
                <table class="table-custom">
                    <thead>
                        <tr>
                            <th>Data</th>
                            <th>Fase de Trabalho</th>
                            <th>Tarefa / Descrição</th>
                        </tr>
                    </thead>
                    <tbody>
                        {linhas_cronograma}
                    </tbody>
                </table>
            </div>
        </div>
    </body>
    </html>
    """
    return render_template_string(html_cronograma)

@encomendas_bp.route("/encomendas/calendario")
def vista_calendario():
    ano = int(request.args.get('ano', datetime.datetime.now().year))
    mes = int(request.args.get('mes', datetime.datetime.now().month))
    
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'DataEntrega', 'Estado', 'Total'])
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor'])
    
    todas_encomendas = []
    for _, r in df_orc.iterrows():
        d = str(r.get('DataEntrega', ''))
        if d and d != 'nan' and d != 'None':
            todas_encomendas.append({'data': d[:10], 'titulo': r.get('Nome', 'Orçamento'), 'cliente': r.get('Cliente', '')})
            
    for _, r in df_enc.iterrows():
        d = str(r.get('DataEntrega', ''))
        if d and d != 'nan' and d != 'None':
            todas_encomendas.append({'data': d[:10], 'titulo': 'Encomenda', 'cliente': r.get('Cliente', '')})
            
    contagem_dias = {}
    for item in todas_encomendas:
        data_key = item['data']
        contagem_dias[data_key] = contagem_dias.get(data_key, 0) + 1

    cal = calendar.Calendar(firstweekday=0)
    dias_mes = cal.monthdayscalendar(ano, mes)
    meses_pt = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    
    mes_anterior = mes - 1 if mes > 1 else 12
    ano_anterior = ano if mes > 1 else ano - 1
    mes_seguinte = mes + 1 if mes < 12 else 1
    ano_seguinte = ano if mes < 12 else ano + 1

    html_cal = f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Calendário de Encomendas - CáDoces</title>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; color: #333; }}
            .container {{ max-width: 1100px; margin: auto; }}
            .header-bar {{ display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; gap: 15px; flex-wrap: wrap; }}
            .back-btn {{ text-decoration: none; color: #b85d19; font-size: 16px; font-weight: bold; background: #f9f0e6; padding: 8px 16px; border-radius: 8px; border: 1px solid #f3e5d0; }}
            .nav-mes {{ display: flex; align-items: center; gap: 20px; font-size: 20px; font-weight: bold; color: #8c4410; flex-wrap: wrap; justify-content: center; }}
            .nav-mes a {{ text-decoration: none; color: #b85d19; background: #f9f0e6; padding: 6px 14px; border-radius: 8px; font-size: 16px; }}
            .legenda {{ display: flex; gap: 25px; background: #fffdf9; padding: 15px 25px; border-radius: 10px; border: 2px solid #f3e5d0; margin-bottom: 20px; font-size: 14px; align-items: center; flex-wrap: wrap; }}
            .legenda-item {{ display: flex; align-items: center; gap: 8px; font-weight: bold; }}
            .bola {{ width: 16px; height: 16px; border-radius: 50%; display: inline-block; }}
            table.calendario-table {{ width: 100%; border-collapse: collapse; background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
            table.calendario-table th {{ background-color: #f9f0e6; color: #8c4410; padding: 15px; text-align: center; font-size: 15px; border-bottom: 2px solid #f3e5d0; }}
            table.calendario-table td {{ width: 14.28%; height: 110px; vertical-align: top; border: 1px solid #f3e5d0; padding: 8px; position: relative; background: #fffdf9; }}
            table.calendario-table td.outro-mes {{ background: #f9f6f0; color: #bbb; }}
            .dia-numero {{ font-weight: bold; font-size: 14px; color: #444; margin-bottom: 6px; display: block; }}
            .badge-encomendas {{ display: block; padding: 6px 8px; border-radius: 8px; font-size: 12px; font-weight: bold; text-align: center; margin-top: 4px; color: white; }}

            @media screen and (max-width: 768px) {{
                body {{ padding: 15px; }}
                .header-bar {{ flex-direction: column; text-align: center; }}
                table.calendario-table td {{ height: 80px; padding: 4px; font-size: 11px; }}
                .badge-encomendas {{ font-size: 9px; padding: 3px 4px; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header-bar">
                <a href="/encomendas?tab=orcamentos" class="back-btn">&#8592; Voltar às Encomendas</a>
                <div class="nav-mes">
                    <a href="/encomendas/calendario?ano={ano_anterior}&mes={mes_anterior}">&#9664; Mês Anterior</a>
                    <span>{meses_pt[mes]} de {ano}</span>
                    <a href="/encomendas/calendario?ano={ano_seguinte}&mes={mes_seguinte}">Mês Seguinte &#9654;</a>
                </div>
            </div>

            <div class="legenda">
                <span><strong>Capacidade diária (Máximo: 5 encomendas):</strong></span>
                <div class="legenda-item"><span class="bola" style="background: #27ae60;"></span> 1 a 3 (Folgado - Verde)</div>
                <div class="legenda-item"><span class="bola" style="background: #e67e22;"></span> 4 (Atenção - Laranja)</div>
                <div class="legenda-item"><span class="bola" style="background: #c0392b;"></span> 5 (Esgotado - Vermelho)</div>
            </div>

            <table class="calendario-table">
                <thead>
                    <tr>
                        <th>Segunda</th><th>Terça</th><th>Quarta</th><th>Quinta</th><th>Sexta</th><th>Sábado</th><th>Domingo</th>
                    </tr>
                </thead>
                <tbody>
    """

    for semana in dias_mes:
        html_cal += "<tr>"
        for dia in semana:
            if dia == 0:
                html_cal += '<td class="outro-mes"></td>'
            else:
                data_str = f"{ano}-{str(mes).zfill(2)}-{str(dia).zfill(2)}"
                qtd = contagem_dias.get(data_str, 0)
                
                if qtd == 0:
                    badge_html = ""
                elif qtd <= 3:
                    badge_html = f'<span class="badge-encomendas" style="background: #27ae60;">{qtd} encomenda{"s" if qtd > 1 else ""}</span>'
                elif qtd == 4:
                    badge_html = f'<span class="badge-encomendas" style="background: #e67e22;">4 encomendas</span>'
                else:
                    badge_html = f'<span class="badge-encomendas" style="background: #c0392b;">{qtd} (ESGOTADO)</span>'

                html_cal += f"""
                <td>
                    <span class="dia-numero">{dia}</span>
                    {badge_html}
                </td>
                """
        html_cal += "</tr>"

    html_cal += """
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """
    return render_template_string(html_cal)

@encomendas_bp.route("/encomendas", methods=['GET', 'POST'])
def gerir_encomendas():
    if request.method == 'POST':
        id_editar = request.form.get('id_editar')
        nome = request.form.get('nome')
        cliente = request.form.get('cliente') or request.form.get('novo_cliente')
        data_entrega = request.form.get('data_entrega')
        total_raw = request.form.get('total_0', request.form.get('total', 0))
        total_limpo = str(total_raw).replace(',', '.')
        total = float(total_limpo) if total_limpo.strip() != '' else 0.0
        
        notas_alergenios = request.form.get('notas_alergenios', '')
        num_pessoas = request.form.get('num_pessoas', '')
        
        df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'Contacto', 'Notas', 'Estado', 'Total', 'Receita', 'Material', 'MaoObra', 'DespesaExtra', 'DataEntrega', 'Itens', 'NotasAlergenios', 'NumPessoas'])
        
        if id_editar and str(id_editar).strip() != '':
            mask = df_orc['ID'].astype(str) == str(id_editar)
            if any(mask):
                df_orc.loc[mask, 'Nome'] = nome
                df_orc.loc[mask, 'Cliente'] = cliente
                df_orc.loc[mask, 'DataEntrega'] = data_entrega
                df_orc.loc[mask, 'Total'] = total
                df_orc.loc[mask, 'NotasAlergenios'] = notas_alergenios
                df_orc.loc[mask, 'NumPessoas'] = num_pessoas
                guardar_dados(EXCEL_ORCAMENTOS, df_orc)
        else:
            novo_id = int(df_orc['ID'].max() + 1) if not df_orc.empty and 'ID' in df_orc.columns and pd.notna(df_orc['ID'].max()) else 1
            
            novo_registo = pd.DataFrame([{
                'ID': novo_id,
                'Nome': nome,
                'Cliente': cliente,
                'DataEntrega': data_entrega,
                'Total': total,
                'Estado': 'Pendente',
                'Itens': '[]',
                'NotasAlergenios': notas_alergenios,
                'NumPessoas': num_pessoas
            }])
            
            df_orc = pd.concat([df_orc, novo_registo], ignore_index=True)
            guardar_dados(EXCEL_ORCAMENTOS, df_orc)
            
        return redirect(url_for('encomendas_bp.gerir_encomendas', tab='orcamentos'))

    tab = request.args.get('tab', 'orcamentos')
    pesquisa = request.args.get('q', '').strip().lower()
    filtro_estado = request.args.get('estado', 'todos')
    
    df_orc = carregar_dados(EXCEL_ORCAMENTOS, ['ID', 'Nome', 'Cliente', 'Contacto', 'Notas', 'Estado', 'Total', 'Receita', 'Material', 'MaoObra', 'DespesaExtra', 'DataEntrega', 'Itens', 'NotasAlergenios', 'NumPessoas'])
    df_enc = carregar_dados(EXCEL_ENCOMENDAS, ['ID', 'Cliente', 'DataEntrega', 'Estado', 'Valor', 'Notas', 'Receita', 'Material', 'MaoObra', 'DespesaExtra', 'Itens', 'Nome', 'SinalPago', 'RestantePago', 'NotasAlergenios', 'NumPessoas'])
    df_rec = carregar_dados(EXCEL_RECEITAS, ['ID', 'Nome da Receita', 'Tempo'])
    df_ing = carregar_dados(EXCEL_INGREDIENTES, ['ID', 'Nome'])
    df_mob = carregar_dados(EXCEL_MAO_OBRA, ['ID', 'Descricao', 'Valor'])
    df_cli = carregar_dados(EXCEL_CLIENTES, ['ID', 'Nome', 'Contacto', 'Telefone'])
    
    num_orc = len(df_orc)
    num_enc = len(df_enc)
    
    df_enc_filtrado = df_enc
    if pesquisa:
        df_enc_filtrado = df_enc_filtrado[df_enc_filtrado['Cliente'].astype(str).str.lower().str.contains(pesquisa, na=False)]
    if filtro_estado != 'todos':
        df_enc_filtrado = df_enc_filtrado[df_enc_filtrado['Estado'].astype(str).str.lower() == filtro_estado.lower()]

    linhas_orc = ""
    if not df_orc.empty:
        for idx, row in df_orc.iterrows():
            id_orc = row.get('ID', idx)
            estado = row.get('Estado', 'Pendente')
            
            if estado == "Aprovado":
                cor_Badge = "#27ae60"
                bg_Badge = "#eafaf1"
            elif estado == "Cancelado":
                cor_Badge = "#c0392b"
                bg_Badge = "#fdedec"
            else:
                cor_Badge = "#b85d19"
                bg_Badge = "#fdf3e3"
            
            r_nome = str(row.get('Nome', '')).replace('"', '&quot;')
            r_cliente = str(row.get('Cliente', '')).replace('"', '&quot;')
            r_data = str(row.get('DataEntrega', ''))[:10]
            r_total = float(row.get('Total', 0) or 0)
            r_notas = str(row.get('NotasAlergenios', '') or '').replace('"', '&quot;')
            r_num_pessoas = str(row.get('NumPessoas', '') or '').replace('"', '&quot;')
            
            telemovel_utilizador = session.get('user_telemovel', '')
            email_atual = session.get('user_email', 'admin')
            if email_atual == 'cadoces023@gmail.com':
                texto_modelo = FRASE_WHATSAPP_ADMIN
            else:
                texto_modelo = FRASE_WHATSAPP_ALUNO

            mensagem_final = texto_modelo.format(cliente=r_cliente, nome=r_nome, valor=r_total)
            msg_whatsapp = urllib.parse.quote(mensagem_final)
            link_whatsapp = f"https://wa.me/{telemovel_utilizador}?text={msg_whatsapp}" if telemovel_utilizador else f"https://wa.me/?text={msg_whatsapp}"
            
            link_email = f"mailto:?subject=Orçamento CáDoces - {r_nome}&body=Olá {r_cliente},%0D%0A%0D%0ASegue em anexo a proposta de orçamento para '{r_nome}' no valor de €{r_total:.2f}.%0D%0A%0D%0MObrigada,%0DCáDoces"

            linhas_orc += f"""
            <tr>
                <td data-label="Nome do Orçamento"><strong>{row.get('Nome', 'Sem Nome')}</strong></td>
                <td data-label="Cliente">{row.get('Cliente', 'Sem Cliente')}</td>
                <td data-label="Nº Pessoas">{r_num_pessoas if r_num_pessoas else '-'}</td>
                <td data-label="Valor Total"><strong style="color: #b85d19; font-size: 15px;">€{r_total:.2f}</strong></td>
                <td data-label="Estado"><span style="background: {bg_Badge}; color: {cor_Badge}; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 12px;">{estado}</span></td>
                <td data-label="Ações" style="text-align: right;">
                    <div style="display: flex; align-items: center; gap: 10px; justify-content: flex-end;">
                        <button type="button" onclick="abrirModalEditar('{id_orc}', '{r_nome}', '{r_cliente}', '{r_data}', '{r_total}', '{r_notas}', '{r_num_pessoas}')" style="background: none; border: none; cursor: pointer; font-size: 16px;" title="Editar">✏️</button>
                        <a href="/encomendas/pdf/{id_orc}" style="color: #b85d19; text-decoration: none; font-size: 16px;" title="Descarregar PDF">📄</a>
                        <a href="{link_whatsapp}" target="_blank" style="color: #27ae60; text-decoration: none; font-size: 16px;" title="Enviar por WhatsApp">💬</a>
                        <a href="{link_email}" style="color: #2980b9; text-decoration: none; font-size: 16px;" title="Enviar por Email">✉️</a>
                        <a href="/encomendas/adicionar_calendario/{id_orc}" style="color: #27ae60; text-decoration: none; font-size: 16px;" title="Calendário .ics">📅</a>
                        <a href="/encomendas/aprovar_orcamento/{id_orc}" style="color: #27ae60; text-decoration: none; font-weight: bold; font-size: 12px; background: #eafaf1; padding: 4px 8px; border-radius: 6px;">✅ Encomenda</a>
                        <a href="/encomendas/apagar_orcamento/{id_orc}" onclick="return confirm('Tem certeza que deseja apagar este orçamento?')" style="color: #c0392b; text-decoration: none; font-size: 16px;" title="Apagar">🗑️</a>
                    </div>
                </td>
            </tr>
            """
    else:
        linhas_orc = '<tr><td colspan="6" style="text-align:center; color:#888; padding: 30px;">Nenhum orçamento registado.</td></tr>'

    linhas_enc = ""
    if not df_enc_filtrado.empty:
        for idx, row in df_enc_filtrado.iterrows():
            id_enc = row.get('ID', idx)
            estado_atual = row.get('Estado', 'Aceite')
            data_entrega = row.get('DataEntrega', '')
            if pd.isna(data_entrega) or str(data_entrega) == 'nan' or str(data_entrega) == 'None':
                data_entrega = ''

            cliente_val = str(row.get('Cliente', 'Cliente desconhecido')).replace("'", "\\'")
            nome_val = str(row.get('Nome', 'Encomenda')).replace("'", "\\'")
            valor_val = float(row.get('Valor', 0) or 0)
            notas_val = str(row.get('NotasAlergenios', '') or '')
            num_pessoas_val = str(row.get('NumPessoas', '') or '')
            
            sinal_pago = str(row.get('SinalPago', 'Pendente'))
            restante_pago = str(row.get('RestantePago', 'Pendente'))
            
            bg_sinal = "#27ae60" if sinal_pago == "Pago" else "#e67e22"
            bg_restante = "#27ae60" if restante_pago == "Pago" else "#e67e22"
            
            bg_select = "#2980b9"
            if estado_atual == "Em Produção": bg_select = "#d35400"
            elif estado_atual == "Concluído": bg_select = "#27ae60"
            elif estado_atual == "Cancelado": bg_select = "#c0392b"

            itens_str = str(row.get('Itens', '[]'))
            try:
                itens_json = json.dumps(json.loads(itens_str) if itens_str and itens_str != 'nan' else [])
            except:
                itens_json = '[]'

            select_options = "".join([f"<option value='{st}' {'selected' if st == estado_atual else ''}>{st}</option>" for st in ["Aceite", "Em Produção", "Concluído", "Cancelado"]])

            badge_notas = f"<br><span style='font-size: 11px; background: #fdf3e3; color: #b85d19; padding: 2px 6px; border-radius: 4px; display: inline-block; margin-top: 3px;' title='{notas_val}'>⚠️ {notas_val[:30]}{'...' if len(notas_val) > 30 else ''}</span>" if notas_val.strip() else ""

            linhas_enc += f"""
            <tr>
                <td data-label="Cliente"><strong>{row.get('Cliente', 'Cliente desconhecido')}</strong><br><span style="font-size: 12px; color: #777;">{row.get('Nome', '')}</span>{badge_notas}</td>
                <td data-label="Data de Entrega">{data_entrega}</td>
                <td data-label="Nº Pessoas">{num_pessoas_val if num_pessoas_val else '-'}</td>
                <td data-label="Estado">
                    <form action="/encomendas/atualizar_estado/{id_enc}" method="POST" style="margin: 0;">
                        <select name="estado" onchange="this.form.submit()" style="background: {bg_select}; color: white; border: none; padding: 6px 12px; border-radius: 8px; font-weight: bold; font-size: 13px; cursor: pointer;">
                            {select_options}
                        </select>
                    </form>
                </td>
                <td data-label="Valor"><strong style="color: #b85d19; font-size: 15px;">€{valor_val:.2f}</strong></td>
                <td data-label="Pagamentos">
                    <form action="/encomendas/atualizar_pagamento/{id_enc}" method="POST" style="display: flex; gap: 6px; margin: 0; justify-content: flex-end;">
                        <select name="sinal_pago" onchange="this.form.submit()" style="background: {bg_sinal}; color: white; border: none; padding: 5px 8px; border-radius: 6px; font-size: 12px; font-weight: bold; cursor: pointer;" title="Sinal 50%">
                            <option value="Pendente" {'selected' if sinal_pago == 'Pendente' else ''}>Sinal: Pendente</option>
                            <option value="Pago" {'selected' if sinal_pago == 'Pago' else ''}>Sinal: Pago</option>
                        </select>
                        <select name="restante_pago" onchange="this.form.submit()" style="background: {bg_restante}; color: white; border: none; padding: 5px 8px; border-radius: 6px; font-size: 12px; font-weight: bold; cursor: pointer;" title="Restante 50%">
                            <option value="Pendente" {'selected' if restante_pago == 'Pendente' else ''}>Restante: Pendente</option>
                            <option value="Pago" {'selected' if restante_pago == 'Pago' else ''}>Restante: Pago</option>
                        </select>
                    </form>
                </td>
                <td data-label="Ações" style="text-align: right;">
                    <div style="display: flex; align-items: center; gap: 14px; justify-content: flex-end;">
                        <button type="button" onclick="abrirModalEditarEncomenda('{id_enc}', '{cliente_val}', '{nome_val}', '{data_entrega}', {valor_val}, '{notas_val.replace("'", "\\'")}')" style="background: none; border: none; cursor: pointer; font-size: 16px;" title="Editar">✏️</button>
                        <button type="button" onclick="abrirModalDetalhes('{cliente_val}', '{nome_val}', '{data_entrega}', '{estado_atual}', {valor_val}, {itens_json}, '{notas_val.replace("'", "\\'")}')" style="background: none; border: none; cursor: pointer; font-size: 16px;" title="Ver Detalhes">👁️</button>
                        <a href="/encomendas/duplicar/{id_enc}" style="color: #b85d19; text-decoration: none; font-size: 16px;" title="Repetir / Duplicar Encomenda">🔄</a>
                        <a href="/encomendas/apagar/{id_enc}" onclick="return confirm('Tem certeza que deseja apagar esta encomenda?')" style="color: #c0392b; text-decoration: none; font-size: 16px;" title="Apagar">🗑️</a>
                    </div>
                </td>
            </tr>
            """
    else:
        linhas_enc = '<tr><td colspan="7" style="text-align:center; color:#888; padding: 30px;">Nenhuma encomenda encontrada.</td></tr>'

    opcoes_receitas_html = "<option value=''>Selecione a Receita...</option>"
    for _, r in df_rec.iterrows():
        nome_r = obter_primeiro_valor_valido(r, ['Nome da Receita', 'Nome', 'Receita', 'Título'])
        tempo_min = float(r.get('Tempo', r.get('Tempo (min)', 30)) or 30)
        custo_mo_calc = (tempo_min / 60) * 8.00
        opcoes_receitas_html += f"<option value='{nome_r}'>{nome_r} (€{custo_mo_calc:.2f} Mão de Obra)</option>"

    opcoes_ingredientes_html = "<option value=''>Selecione o Material...</option>"
    for _, i in df_ing.iterrows():
        nome_i = obter_primeiro_valor_valido(i, ['Nome', 'Ingrediente', 'Designação', 'Material', 'Artigo', 'Descrição'])
        opcoes_ingredientes_html += f"<option value='{nome_i}'>{nome_i}</option>"

    opcoes_mao_obra_html = "<option value=''>Selecione Mão de Obra...</option>"
    for _, mo in df_mob.iterrows():
        nome_mo = obter_primeiro_valor_valido(mo, ['Descricao', 'Descrição', 'Funcao', 'Função', 'Nome', 'Tipo'])
        opcoes_mao_obra_html += f"<option value='{nome_mo}'>{nome_mo}</option>"

    opcoes_clientes_html = "<option value=''>Selecione o Cliente...</option>"
    for _, c in df_cli.iterrows():
        nome_c = obter_primeiro_valor_valido(c, ['Nome', 'Cliente', 'Nome do Cliente'])
        opcoes_clientes_html += f"<option value='{nome_c}'>{nome_c}</option>"

    filtros_estado_html = ""
    if tab == 'encomendas':
        estados_possiveis = [('todos', 'Todos'), ('Aceite', 'Aceite'), ('Em Produção', 'Em Produção'), ('Concluído', 'Concluído'), ('Cancelado', 'Cancelado')]
        botoes_estado = ""
        for val_st, label_st in estados_possiveis:
            ativo = 'background: #dfa673; color: white;' if filtro_estado.lower() == val_st.lower() else 'background: #fff; color: #444;'
            botoes_estado += f'<a href="/encomendas?tab=encomendas&estado={val_st}&q={pesquisa}" style="{ativo} text-decoration: none; padding: 6px 14px; border-radius: 8px; border: 1px solid #f3e5d0; font-size: 13px; font-weight: bold;">{label_st}</a>'
        
        filtros_estado_html = f"""
        <div style="display: flex; gap: 8px; align-items: center; margin-bottom: 15px; flex-wrap: wrap;">
            <span style="font-size: 13px; font-weight: bold; color: #666;">Filtrar por Estado:</span>
            {botoes_estado}
        </div>
        """

    conteudo_tab = f"""
    <div class="content-box">
        <table class="table-custom">
            <thead>
                <tr>
                    <th>Nome do Orçamento</th>
                    <th>Cliente</th>
                    <th>Nº Pessoas</th>
                    <th>Valor Total</th>
                    <th>Estado</th>
                    <th style="text-align: right;">Ações</th>
                </tr>
            </thead>
            <tbody>
                {linhas_orc}
            </tbody>
        </table>
    </div>
    """ if tab == 'orcamentos' else f"""
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; gap: 15px; flex-wrap: wrap;">
        <div class="search-bar-container" style="flex: 1; margin-bottom: 0;">
            <form method="GET" action="/encomendas" style="display: flex; gap: 10px; width: 100%;">
                <input type="hidden" name="tab" value="encomendas">
                <input type="hidden" name="estado" value="{filtro_estado}">
                <input type="text" name="q" class="form-control" placeholder="🔍 Pesquisar por cliente..." value="{pesquisa}" style="background: #fffdf9;">
            </form>
        </div>
        <a href="/encomendas/folha_cozinha" class="btn-outline" style="background: #fffdf9; border: 2px solid #dfa673; color: #b85d19; padding: 10px 16px; border-radius: 8px; text-decoration: none; font-weight: bold; display: flex; align-items: center; gap: 6px;" target="_blank">🖨️ Folha de Cozinha</a>
    </div>
    {filtros_estado_html}
    <div class="content-box">
        <table class="table-custom">
            <thead>
                <tr>
                    <th>Cliente</th>
                    <th>Data de Entrega</th>
                    <th>Nº Pessoas</th>
                    <th>Estado</th>
                    <th>Valor</th>
                    <th style="text-align: right;">Pagamentos (Sinal / Restante)</th>
                    <th style="text-align: right;">Ações</th>
                </tr>
            </thead>
            <tbody>
                {linhas_enc}
            </tbody>
        </table>
    </div>
    """

    return render_template_string("""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Orçamentos e Encomendas</title>
        <style>
            body { font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; }
            .container { max-width: 1100px; margin: auto; }
            .header-bar { display: flex; justify-content: space-between; align-items: center; background: #fffdf9; padding: 20px 30px; border: 2px solid #f3e5d0; border-radius: 12px; margin-bottom: 25px; gap: 15px; flex-wrap: wrap; }
            .header-left { display: flex; align-items: center; gap: 20px; }
            .back-btn { text-decoration: none; color: #b85d19; font-size: 22px; font-weight: bold; background: #f9f0e6; padding: 8px 14px; border-radius: 8px; }
            h1 { margin: 0; color: #111; font-size: 24px; }
            p.sub { margin: 5px 0 0 0; color: #666; font-size: 14px; }
            .btn-primary { background: #dfa673; color: #fff; border: none; padding: 12px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; width: fit-content; }
            .btn-primary:hover { background: #c8915f; }
            .top-actions { display: flex; gap: 15px; align-items: center; flex-wrap: wrap; }
            .btn-outline { background: #fff; border: 2px solid #f3e5d0; color: #444; padding: 10px 16px; border-radius: 8px; cursor: pointer; font-weight: bold; text-decoration: none; font-size: 14px; display: flex; align-items: center; gap: 6px; }
            .tabs-nav { display: flex; gap: 10px; background: #f3e5d0; padding: 6px; border-radius: 12px; width: fit-content; margin-bottom: 20px; flex-wrap: wrap; }
            .tab-btn { background: transparent; border: none; padding: 10px 20px; border-radius: 8px; font-weight: bold; color: #666; cursor: pointer; text-decoration: none; font-size: 14px; display: flex; align-items: center; gap: 8px; }
            .tab-btn.active { background: #fffdf9; color: #111; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
            .search-bar-container { margin-bottom: 20px; }
            .content-box { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
            table.table-custom { width: 100%; border-collapse: collapse; }
            table.table-custom th { background-color: #f9f0e6; color: #8c4410; padding: 12px; text-align: left; border-bottom: 2px solid #f3e5d0; }
            table.table-custom td { padding: 12px; border-bottom: 1px solid #f3e5d0; color: #444; }
            
            .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.4); overflow-y: auto; padding: 40px 0; }
            .modal-content { background-color: #fffdf9; margin: auto; padding: 30px; border: 2px solid #f3e5d0; width: 90%; max-width: 850px; border-radius: 14px; box-shadow: 0 5px 15px rgba(0,0,0,0.1); position: relative; box-sizing: border-box; }
            .close-btn { position: absolute; right: 20px; top: 20px; font-size: 18px; cursor: pointer; color: #888; font-weight: bold; background: #f9f0e6; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 1px solid #f3e5d0; }
            .form-group { margin-bottom: 18px; }
            .form-group label { display: block; font-weight: bold; color: #222; margin-bottom: 6px; font-size: 13px; }
            .form-control { width: 100%; padding: 10px 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; box-sizing: border-box; outline: none; }
            .form-control:focus { border-color: #dfa673; }
            
            .opcao-card { background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 20px; margin-bottom: 20px; position: relative; }
            .grid-4 { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 15px; }
            
            .badge-tipo-item { display: inline-flex; align-items: center; gap: 5px; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: bold; }
            .badge-despesa { background: #eafaf1; color: #27ae60; border: 1px solid #a9dfbf; }
            .badge-material { background: #fef9e7; color: #d4ac0d; border: 1px solid #f9e79f; }
            .badge-maoobra { background: #e8f8f5; color: #117864; border: 1px solid #a3e4d7; }
            
            .summary-box { display: flex; justify-content: space-between; align-items: flex-start; border-top: 1px solid #f3e5d0; padding-top: 15px; margin-top: 15px; flex-wrap: wrap; gap: 15px; }
            .summary-left { font-size: 13px; color: #555; line-height: 1.6; flex: 1; }
            .modal-footer { display: flex; justify-content: flex-end; gap: 12px; border-top: 2px solid #f3e5d0; padding-top: 20px; margin-top: 20px; flex-wrap: wrap; }
            .btn-secondary { background: #fff; border: 2px solid #f3e5d0; color: #444; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: bold; }
            .btn-small-acao { background: #f9f0e6; color: #8c4410; border: 1px solid #f3e5d0; padding: 6px 14px; border-radius: 8px; font-weight: bold; font-size: 13px; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; }
            .btn-small-acao:hover { background: #f3e5d0; }
            .btn-ver-mais { background: #fff; border: 1px solid #f3e5d0; color: #333; padding: 6px 14px; border-radius: 8px; cursor: pointer; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 6px; margin-top: 10px; }
            .btn-ver-mais:hover { background: #f9f0e6; }

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {
                body { padding: 15px; }
                .header-bar { flex-direction: column; text-align: center; }
                .grid-4 { grid-template-columns: 1fr; }
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
                <div class="header-left">
                    <a href="/" class="back-btn">&#8592;</a>
                    <div>
                        <h1>Encomendas</h1>
                        <p class="sub">Gerencie as suas encomendas</p>
                    </div>
                </div>
                <div class="top-actions">
                    <button class="btn-primary" onclick="abrirModalCriar()">+ Novo Orçamento</button>
                </div>
            </div>

            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 15px;">
                <div class="tabs-nav" style="margin-bottom: 0;">
                    <a href="/encomendas?tab=orcamentos" class="tab-btn {% if tab == 'orcamentos' %}active{% endif %}">📄 Orçamentos ({{num_orc}})</a>
                    <a href="/encomendas?tab=encomendas" class="tab-btn {% if tab == 'encomendas' %}active{% endif %}">🛍 Encomendas ({{num_enc}})</a>
                </div>
                <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                    <a href="/encomendas/cronograma" class="btn-outline">📋 Cronograma de Cozinha</a>
                    <a href="/encomendas/calendario" class="btn-outline">📅 Calendário</a>
                </div>
            </div>

            {{conteudo_tab | safe}}
        </div>

        <!-- MODAL EDITAR ENCOMENDA -->
        <div id="modalEditarEncomenda" class="modal">
            <div class="modal-content" style="max-width: 500px;">
                <span class="close-btn" onclick="fecharModalEditarEncomenda()">&times;</span>
                <h3 style="margin-top:0; color:#8c4410; font-size:20px;">Editar Encomenda</h3>
                <form id="formEditarEncomenda" method="POST">
                    <div class="form-group">
                        <label>Cliente</label>
                        <input type="text" name="cliente" id="edit_enc_cliente" class="form-control" required>
                    </div>
                    <div class="form-group">
                        <label>Nome / Tema</label>
                        <input type="text" name="nome" id="edit_enc_nome" class="form-control" required>
                    </div>
                    <div class="form-group">
                        <label>Data de Entrega</label>
                        <input type="date" name="data_entrega" id="edit_enc_data" class="form-control" required>
                    </div>
                    <div class="form-group">
                        <label>Valor (€)</label>
                        <input type="text" inputmode="decimal" name="total" id="edit_enc_valor" class="form-control" required>
                    </div>
                    <div class="form-group">
                        <label>⚠️ Notas Especiais / Alergénios</label>
                        <textarea name="notas_alergenios" id="edit_enc_notas" class="form-control" rows="2" placeholder="Ex: Alergia a frutos secos, menos açúcar..."></textarea>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModalEditarEncomenda()">Cancelar</button>
                        <button type="submit" class="btn-primary">Guardar Alterações</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL DETALHES DA ENCOMENDA -->
        <div id="modalDetalhes" class="modal">
            <div class="modal-content" style="max-width: 600px;">
                <span class="close-btn" onclick="fecharModalDetalhes()">&times;</span>
                <h3 id="detalheNomeTema" style="margin:0 0 10px 0; color:#8c4410; font-size:20px;">Detalhes da Encomenda</h3>
                <p style="margin:4px 0; color:#333;"><strong>Cliente:</strong> <span id="detalheCliente"></span></p>
                <p style="margin:4px 0; color:#333;"><strong>Data de Entrega:</strong> <span id="detalheData"></span></p>
                <p style="margin:4px 0 10px 0; color:#333;"><strong>Estado:</strong> <span id="detalheEstado" style="background:#27ae60; color:white; padding:3px 8px; border-radius:6px; font-size:12px;"></span></p>
                <div id="detalheNotasBox" style="background: #fdf3e3; border: 1px solid #f3e5d0; padding: 10px; border-radius: 8px; margin-bottom: 15px; display: none;">
                    <strong style="color: #b85d19; font-size: 13px;">⚠️ Notas / Alergénios:</strong>
                    <p id="detalheNotasTexto" style="margin: 4px 0 0 0; font-size: 13px; color: #444;"></p>
                </div>
                
                <h4 style="border-bottom: 1px solid #f3e5d0; padding-bottom: 6px; color: #444; margin-bottom: 10px;">Itens da Encomenda</h4>
                <div id="detalheItensLista" style="max-height: 200px; overflow-y: auto; margin-bottom: 15px;"></div>
                
                <div style="border-top: 1px solid #f3e5d0; padding-top: 10px; font-size: 15px;">
                    <p style="margin:4px 0;">Custo Total: <strong id="detalheCusto">€0.00</strong></p>
                    <p style="margin:4px 0; font-size:18px; color:#b85d19;">Total a Receber: <strong id="detalheTotal">€0.00</strong></p>
                    <p style="margin:4px 0; color:#27ae60;">Lucro: <strong id="detalheLucro">€0.00</strong></p>
                </div>
            </div>
        </div>

        <!-- MODAL NOVO / EDITAR ORÇAMENTO -->
        <div id="modalOrcamento" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModalCriar()">&times;</span>
                
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 10px;">
                    <h2 id="modalTitulo" style="margin:0; color:#111; font-size:22px;">Novo Orçamento</h2>
                    <button type="button" class="btn-small-acao" onclick="adicionarOpcaoOrcamento()">+ Opção</button>
                </div>
                
                <form action="/encomendas" method="POST">
                    <input type="hidden" name="id_editar" id="id_editar">
                    
                    <div style="display: flex; gap: 15px; flex-wrap: wrap;">
                        <div class="form-group" style="flex: 2; min-width: 200px;">
                            <label>Nome do Orçamento / Tema *</label>
                            <input type="text" name="nome" id="input_nome" class="form-control" placeholder="Ex: Bolo de aniversário" required style="border-color: #dfa673;">
                        </div>
                        <div class="form-group" style="flex: 1; min-width: 150px;">
                            <label>👥 Nº de Pessoas / Fatias</label>
                            <input type="text" name="num_pessoas" id="input_num_pessoas" class="form-control" placeholder="Ex: 15-20 pessoas">
                        </div>
                    </div>

                    <div class="form-group">
                        <label>👤 Cliente</label>
                        <select name="cliente" id="select_cliente" class="form-control">
                            {{ opcoes_clientes_html | safe }}
                        </select>
                    </div>

                    <div style="display: flex; gap: 15px; flex-wrap: wrap;">
                        <div class="form-group" style="flex:1; min-width: 200px;">
                            <label>Novo Cliente (se não estiver na lista)</label>
                            <input type="text" name="novo_cliente" class="form-control" placeholder="Nome de novo cliente...">
                        </div>
                        <div class="form-group" style="flex:1; min-width: 200px;">
                            <label>Contacto</label>
                            <input type="text" name="contacto" class="form-control" placeholder="Telemóvel / Email...">
                        </div>
                    </div>

                    <div class="form-group">
                        <label>Data de Entrega</label>
                        <input type="date" name="data_entrega" id="input_data" class="form-control">
                    </div>

                    <div class="form-group">
                        <label>⚠️ Notas Especiais / Alergénios (Geral para o Orçamento/PDF)</label>
                        <textarea name="notas_alergenios" id="input_notas" class="form-control" rows="2" placeholder="Ex: Alergia a frutos secos, sem açúcar..."></textarea>
                    </div>

                    <div id="contentorOpcoes"></div>

                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModalCriar()">Cancelar</button>
                        <button type="submit" id="btnSalvarOrcamento" class="btn-primary">Criar Orçamento</button>
                    </div>
                </form>
            </div>
        </div>

        <!-- MODAL DISTRIBUIÇÃO DE CUSTOS -->
        <div id="modalDistribuicao" class="modal">
            <div class="modal-content" style="max-width: 600px;">
                <span class="close-btn" onclick="fecharModalDistribuicao()">&times;</span>
                <h2 style="margin-top:0; color:#111; font-size:20px; font-family: serif;">Distribuição de Custos</h2>
                
                <div style="text-align: center; margin: 20px 0;">
                    <div id="graficoDonutContainer" style="position: relative; width: 220px; height: 220px; margin: auto;">
                        <svg width="220" height="220" viewBox="0 0 42 42" style="transform: rotate(-90deg);">
                            <circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#f3e5d0" stroke-width="7"></circle>
                            <circle id="circuloIngredientes" cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#f1c40f" stroke-width="7" stroke-dasharray="0 100"></circle>
                            <circle id="circuloFixos" cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="#3498db" stroke-width="7" stroke-dasharray="0 100" stroke-dashoffset="0"></circle>
                        </svg>
                        <div id="centroDonutTexto" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); font-weight: bold; font-size: 16px; color: #333;"></div>
                    </div>
                </div>

                <div style="border-top: 1px solid #f3e5d0; padding-top: 15px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; font-size: 14px;">
                        <span style="display: flex; align-items: center; gap: 8px;"><span style="width: 12px; height: 12px; border-radius: 50%; background: #f1c40f; display: inline-block;"></span> Ingredientes & Materiais</span>
                        <strong id="valIngredientesModal">€0.00 (0.0%)</strong>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; font-size: 14px;">
                        <span style="display: flex; align-items: center; gap: 8px;"><span style="width: 12px; height: 12px; border-radius: 50%; background: #3498db; display: inline-block;"></span> Custos Fixos & Mão de Obra</span>
                        <strong id="valFixosModal">€0.00 (0.0%)</strong>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; border-top: 1px solid #f3e5d0; padding-top: 12px; font-size: 16px;">
                        <strong>Total:</strong>
                        <strong id="valTotalModal" style="color: #b85d19;">€0.00</strong>
                    </div>
                </div>
            </div>
        </div>

        <script>
            let opcoesOrcamento = [];
            let dadosOpcaoAtualModal = null;

            const htmlReceitas = `{{ opcoes_receitas_html | safe }}`;
            const htmlIngredientes = `{{ opcoes_ingredientes_html | safe }}`;
            const htmlMaoObra = `{{ opcoes_mao_obra_html | safe }}`;

            function abrirModalEditarEncomenda(id, cliente, nome, data, valor, notas) {
                document.getElementById('formEditarEncomenda').action = "/encomendas/editar_encomenda/" + id;
                document.getElementById('edit_enc_cliente').value = cliente;
                document.getElementById('edit_enc_nome').value = nome;
                document.getElementById('edit_enc_data').value = data;
                document.getElementById('edit_enc_valor').value = valor;
                document.getElementById('edit_enc_notas').value = notas || '';
                document.getElementById('modalEditarEncomenda').style.display = 'block';
            }

            function fecharModalEditarEncomenda() {
                document.getElementById('modalEditarEncomenda').style.display = 'none';
            }

            function abrirModalDetalhes(cliente, nome, data, estado, valor, itens, notas) {
                document.getElementById('detalheCliente').innerText = cliente;
                document.getElementById('detalheNomeTema').innerText = nome || "Encomenda CáDoces";
                document.getElementById('detalheData').innerText = data;
                document.getElementById('detalheEstado').innerText = estado;
                document.getElementById('detalheTotal').innerText = '€' + valor.toFixed(2);
                
                if (notas && notas.trim() !== '') {
                    document.getElementById('detalheNotasTexto').innerText = notas;
                    document.getElementById('detalheNotasBox').style.display = 'block';
                } else {
                    document.getElementById('detalheNotasBox').style.display = 'none';
                }
                
                let itensHtml = '';
                let custoTotal = 0;
                if (itens && itens.length > 0) {
                    itens.forEach(item => {
                        let sub = (item.custo || 0) * (item.qtd || 1);
                        custoTotal += sub;
                        itensHtml += `<div style="background:#fff; border:1px solid #f3e5d0; padding:6px 10px; border-radius:6px; margin-bottom:6px; display:flex; justify-content:space-between; font-size:13px;"><span>${item.nome || 'Item'} × ${item.qtd || 1}</span><strong>€${sub.toFixed(2)}</strong></div>`;
                    });
                } else {
                    itensHtml = '<p style="color:#777; font-size:13px;">Sem itens registados detalhadamente.</p>';
                }
                
                document.getElementById('detalheItensLista').innerHTML = itensHtml;
                document.getElementById('detalheCusto').innerText = '€' + custoTotal.toFixed(2);
                document.getElementById('detalheLucro').innerText = '€' + (valor - custoTotal).toFixed(2);
                document.getElementById('modalDetalhes').style.display = 'block';
            }

            function fecharModalDetalhes() {
                document.getElementById('modalDetalhes').style.display = 'none';
            }

            function abrirModalCriar() { 
                document.getElementById('modalTitulo').innerText = "Novo Orçamento";
                document.getElementById('btnSalvarOrcamento').innerText = "Criar Orçamento";
                document.getElementById('id_editar').value = "";
                document.getElementById('input_nome').value = "";
                document.getElementById('input_num_pessoas').value = "";
                document.getElementById('select_cliente').value = "";
                document.getElementById('input_data').value = "";
                document.getElementById('input_notas').value = "";
                
                document.getElementById('modalOrcamento').style.display = 'block'; 
                opcoesOrcamento = [];
                adicionarOpcaoOrcamento();
            }

            function abrirModalEditar(id, nome, cliente, dataEntrega, totalVenda, notas, numPessoas) {
                document.getElementById('modalTitulo').innerText = "Editar Orçamento #" + id;
                document.getElementById('btnSalvarOrcamento').innerText = "Guardar Alterações";
                document.getElementById('id_editar').value = id;
                document.getElementById('input_nome').value = nome;
                document.getElementById('input_num_pessoas').value = numPessoas || '';
                document.getElementById('select_cliente').value = cliente;
                document.getElementById('input_data').value = dataEntrega;
                document.getElementById('input_notas').value = notas || '';
                
                document.getElementById('modalOrcamento').style.display = 'block';
                opcoesOrcamento = [{
                    itens: [],
                    precoVenda: parseFloat(totalVenda) || 0
                }];
                renderizarOpcoes();
            }

            function fecharModalCriar() { document.getElementById('modalOrcamento').style.display = 'none'; }

            function abrirModalDistribuicao(opIdx) {
                let opcao = opcoesOrcamento[opIdx];
                let totalIngredientes = 0;
                let totalOutros = 0;

                opcao.itens.forEach(item => {
                    let sub = (item.custo + item.maoObra) * item.qtd;
                    if (item.tipo === 'receita' || item.tipo === 'material') {
                        totalIngredientes += sub;
                    } else {
                        totalOutros += sub;
                    }
                });

                let custoTotal = totalIngredientes + totalOutros;
                let pctIng = custoTotal > 0 ? (totalIngredientes / custoTotal) * 100 : 0;
                let pctOutros = custoTotal > 0 ? (totalOutros / custoTotal) * 100 : 0;

                document.getElementById('valIngredientesModal').innerText = `€${totalIngredientes.toFixed(2)} (${pctIng.toFixed(1)}%)`;
                document.getElementById('valFixosModal').innerText = `€${totalOutros.toFixed(2)} (${pctOutros.toFixed(1)}%)`;
                document.getElementById('valTotalModal').innerText = `€${custoTotal.toFixed(2)}`;

                let circIng = document.getElementById('circuloIngredientes');
                let circFix = document.getElementById('circuloFixos');

                circIng.setAttribute('stroke-dasharray', `${pctIng} ${100 - pctIng}`);
                circFix.setAttribute('stroke-dasharray', `${pctOutros} ${100 - pctOutros}`);
                circFix.setAttribute('stroke-dashoffset', `-${pctIng}`);

                document.getElementById('modalDistribuicao').style.display = 'block';
            }

            function fecharModalDistribuicao() {
                document.getElementById('modalDistribuicao').style.display = 'none';
            }

            function adicionarOpcaoOrcamento() {
                opcoesOrcamento.push({
                    itens: [],
                    precoVenda: 0
                });
                renderizarOpcoes();
            }

            function removerOpcao(index) {
                if (opcoesOrcamento.length === 1) {
                    alert("O orçamento deve ter pelo menos uma opção.");
                    return;
                }
                opcoesOrcamento.splice(index, 1);
                renderizarOpcoes();
            }

            async function selecionarItemOpcao(opcaoIndex, tipo, selectElement) {
                let nome = selectElement.value;
                if (!nome) return;
                try {
                    let response = await fetch(`/api/item/detalhes?tipo=${tipo}&nome=${encodeURIComponent(nome)}`);
                    let data = await response.json();
                    
                    opcoesOrcamento[opcaoIndex].itens.push({
                        tipo: tipo,
                        nome: nome,
                        custo: data.custo || 5.00,
                        maoObra: data.mao_obra || 0.00,
                        qtd: 1
                    });
                    selectElement.value = "";
                    renderizarOpcoes();
                } catch (e) {
                    console.error("Erro ao adicionar item:", e);
                }
            }

            function adicionarDespesaOpcao(opcaoIndex, valor) {
                let val = parseFloat(valor) || 0;
                if (val > 0) {
                    opcoesOrcamento[opcaoIndex].itens.push({
                        tipo: 'despesa',
                        nome: 'Despesa Extra',
                        custo: val,
                        maoObra: 0.00,
                        qtd: 1
                    });
                    renderizarOpcoes();
                }
            }

            function removerItemOpcao(opcaoIndex, itemIndex) {
                opcoesOrcamento[opcaoIndex].itens.splice(itemIndex, 1);
                renderizarOpcoes();
            }

            function atualizarQtdOpcao(opcaoIndex, itemIndex, novaQtd) {
                opcoesOrcamento[opcaoIndex].itens[itemIndex].qtd = parseFloat(novaQtd) || 1;
                renderizarOpcoes();
            }

            function atualizarPrecoVenda(opcaoIndex, novoPreco) {
                let valorLimpo = novoPreco.replace(',', '.');
                opcoesOrcamento[opcaoIndex].precoVenda = valorLimpo === '' ? '' : (parseFloat(valorLimpo) || 0);
                
                let opcao = opcoesOrcamento[opcaoIndex];
                let totalIngredientes = 0;
                let totalDespesas = 0;
                let totalMaoObra = 0;

                opcao.itens.forEach(item => {
                    let sub = item.custo * item.qtd;
                    let subMO = item.maoObra * item.qtd;
                    if (item.tipo === 'receita') {
                        totalIngredientes += sub;
                        totalMaoObra += subMO;
                    } else if (item.tipo === 'material') {
                        totalIngredientes += sub;
                    } else if (item.tipo === 'despesa') {
                        totalDespesas += sub;
                    } else if (item.tipo === 'mao_obra') {
                        totalMaoObra += sub;
                    }
                });

                let precoVendaNum = typeof opcao.precoVenda === 'string' ? (parseFloat(opcao.precoVenda.replace(',', '.')) || 0) : (opcao.precoVenda || 0);
                let custoTotal = totalIngredientes + totalDespesas + totalMaoObra;
                let lucro = precoVendaNum - custoTotal;
                let margem = precoVendaNum > 0 ? (lucro / precoVendaNum) * 100 : 0;

                let card = document.querySelectorAll('.opcao-card')[opcaoIndex];
                if (card) {
                    let summaryLeft = card.querySelector('.summary-left');
                    if (summaryLeft) {
                        let alertaHtml = '';
                        if (precoVendaNum > 0 && margem < 35) {
                            alertaHtml = `<div style="background: #fdedec; color: #c0392b; border: 1px solid #f5b7b1; padding: 6px 10px; border-radius: 6px; font-size: 12px; font-weight: bold; margin-top: 8px;">⚠️ Atenção: Margem de lucro abaixo de 35% (${margem.toFixed(1)}%)!</div>`;
                        }
                        
                        summaryLeft.innerHTML = `
                            Ingredientes & Materiais: €${totalIngredientes.toFixed(2)}<br>
                            Custos Fixos & Despesas: €${totalDespesas.toFixed(2)}<br>
                            Mão de Obra (Automática): €${totalMaoObra.toFixed(2)}<br><br>
                            <strong>Custo Total: €${custoTotal.toFixed(2)}</strong><br>
                            Lucro: €${lucro.toFixed(2)} <span style="color: ${lucro >= 0 ? '#27ae60' : '#c0392b'};">(${margem.toFixed(1)}%)</span><br>
                            % Margem: <span style="color: ${margem >= 0 ? '#27ae60' : '#c0392b'};">${margem.toFixed(1)}%</span>
                            ${alertaHtml}
                            <div style="margin-top: 8px;">
                                <button type="button" class="btn-ver-mais" onclick="abrirModalDistribuicao(${opcaoIndex})">
                                    ⏱ Ver mais
                                </button>
                            </div>
                        `;
                    }
                }
            }

            function renderizarOpcoes() {
                let html = '';
                opcoesOrcamento.forEach((opcao, opIdx) => {
                    let totalIngredientes = 0;
                    let totalDespesas = 0;
                    let totalMaoObra = 0;

                    let itensHtml = '';
                    opcao.itens.forEach((item, itemIdx) => {
                        let subtotal = item.custo * item.qtd;
                        let subtotalMO = item.maoObra * item.qtd;
                        
                        if (item.tipo === 'receita') {
                            totalIngredientes += subtotal;
                            totalMaoObra += subtotalMO;
                        } else if (item.tipo === 'material') {
                            totalIngredientes += subtotal;
                        } else if (item.tipo === 'despesa') {
                            totalDespesas += subtotal;
                        } else if (item.tipo === 'mao_obra') {
                            totalMaoObra += subtotal;
                        }

                        let badgeClass = 'badge-material';
                        let badgeLabel = '📦 ' + item.tipo.toUpperCase();
                        if (item.tipo === 'despesa') badgeClass = 'badge-despesa';
                        if (item.tipo === 'receita') badgeClass = 'badge-maoobra';

                        itensHtml += `
                        <div style="display: flex; justify-content: space-between; align-items: center; background: #fff; border: 1px solid #f3e5d0; border-radius: 8px; padding: 8px 12px; margin-bottom: 8px; font-size: 13px; flex-wrap: wrap; gap: 10px;">
                            <div style="display: flex; align-items: center; gap: 12px;">
                                <span class="badge-tipo-item ${badgeClass}">${badgeLabel}</span>
                                <strong>${item.nome}</strong>
                            </div>
                            <div style="display: flex; align-items: center; gap: 15px;">
                                <span style="font-size: 13px; color: #555;">Qtd: <input type="number" value="${item.qtd}" style="width: 50px; padding: 4px; border: 1px solid #f3e5d0; border-radius: 6px; text-align: center;" onchange="atualizarQtdOpcao(${opIdx}, ${itemIdx}, this.value)"></span>
                                <strong style="color: #b85d19;">€${(subtotal + subtotalMO).toFixed(2)}</strong>
                                <button type="button" onclick="removerItemOpcao(${opIdx}, ${itemIdx})" style="background: none; border: none; cursor: pointer; font-size: 16px; color: #888;" title="Apagar">🗑️</button>
                            </div>
                        </div>
                        `;
                    });

                    let precoVendaNum = typeof opcao.precoVenda === 'string' ? (parseFloat(opcao.precoVenda.replace(',', '.')) || 0) : (opcao.precoVenda || 0);
                    let custoTotal = totalIngredientes + totalDespesas + totalMaoObra;
                    let lucro = precoVendaNum - custoTotal;
                    let margem = precoVendaNum > 0 ? (lucro / precoVendaNum) * 100 : 0;

                    let alertaMargemHtml = '';
                    if (precoVendaNum > 0 && margem < 35) {
                        alertaMargemHtml = `<div style="background: #fdedec; color: #c0392b; border: 1px solid #f5b7b1; padding: 6px 10px; border-radius: 6px; font-size: 12px; font-weight: bold; margin-top: 8px;">⚠️ Atenção: Margem de lucro abaixo de 35% (${margem.toFixed(1)}%)!</div>`;
                    }

                    html += `
                    <div class="opcao-card">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                            <h4 style="margin:0; color:#8c4410; font-size:15px;">Opção ${opIdx + 1}</h4>
                            ${opcoesOrcamento.length > 1 ? `<button type="button" onclick="removerOpcao(${opIdx})" style="background: #fdedec; color: #c0392b; border: 1px solid #f5b7b1; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: bold; cursor: pointer;">Remover Opção</button>` : ''}
                        </div>
                        
                        <div class="grid-4">
                            <div class="form-group" style="margin:0;">
                                <label>🧁 Receita</label>
                                <select class="form-control" onchange="selecionarItemOpcao(${opIdx}, 'receita', this)">
                                    ${htmlReceitas}
                                </select>
                            </div>
                            <div class="form-group" style="margin:0;">
                                <label>📦 Materiais</label>
                                <select class="form-control" onchange="selecionarItemOpcao(${opIdx}, 'material', this)">
                                    ${htmlIngredientes}
                                </select>
                            </div>
                            <div class="form-group" style="margin:0;">
                                <label>⏱ Mão de Obra</label>
                                <select class="form-control" onchange="selecionarItemOpcao(${opIdx}, 'mao_obra', this)">
                                    ${htmlMaoObra}
                                </select>
                            </div>
                            <div class="form-group" style="margin:0;">
                                <label>💵 Despesa (€)</label>
                                <input type="number" step="0.01" class="form-control" value="0.00" onchange="adicionarDespesaOpcao(${opIdx}, this.value)">
                            </div>
                        </div>

                        <div style="margin-bottom: 15px;">
                            ${itensHtml}
                        </div>

                        <div class="summary-box">
                            <div class="summary-left">
                                Ingredientes & Materiais: €${totalIngredientes.toFixed(2)}<br>
                                Custos Fixos & Despesas: €${totalDespesas.toFixed(2)}<br>
                                Mão de Obra (Automática): €${totalMaoObra.toFixed(2)}<br><br>
                                <strong>Custo Total: €${custoTotal.toFixed(2)}</strong><br>
                                Lucro: €${lucro.toFixed(2)} <span style="color: ${lucro >= 0 ? '#27ae60' : '#c0392b'};">(${margem.toFixed(1)}%)</span><br>
                                % Margem: <span style="color: ${margem >= 0 ? '#27ae60' : '#c0392b'};">${margem.toFixed(1)}%</span>
                                ${alertaMargemHtml}
                                <div style="margin-top: 8px;">
                                    <button type="button" class="btn-ver-mais" onclick="abrirModalDistribuicao(${opIdx})">
                                        ⏱ Ver mais
                                    </button>
                                </div>
                            </div>
                            <div style="width: 200px; min-width: 150px;">
                                <label style="font-weight:bold; margin-bottom:5px; display:block; font-size:13px;">Preço de Venda (€)</label>
                                <input type="text" inputmode="decimal" name="total_${opIdx}" class="form-control" value="${opcao.precoVenda !== undefined ? opcao.precoVenda : ''}" oninput="atualizarPrecoVenda(${opIdx}, this.value)" required style="font-size: 18px; font-weight: bold; color: #b85d19;">
                            </div>
                        </div>
                    </div>
                    `;
                });
                document.getElementById('contentorOpcoes').innerHTML = html;
            }
        </script>
    </body>
    </html>
    """, num_orc=num_orc, num_enc=num_enc, conteudo_tab=conteudo_tab, tab=tab,
       opcoes_receitas_html=opcoes_receitas_html, opcoes_ingredientes_html=opcoes_ingredientes_html,
       opcoes_mao_obra_html=opcoes_mao_obra_html, opcoes_clientes_html=opcoes_clientes_html)