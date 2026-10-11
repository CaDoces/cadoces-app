from flask import Blueprint, render_template_string, request, redirect, url_for, send_file, session
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
import json

financeira_bp = Blueprint('financeira_bp', __name__)

EXCEL_FINANCEIRA = 'financeira_dados.xlsx'

def obter_ficheiro_utilizador(nome_base):
    email = session.get('user_email', 'admin')
    if email == 'cadoces023@gmail.com':
        return nome_base
    else:
        email_limpo = email.replace('@', '_at_').replace('.', '_')
        nome, ext = os.path.splitext(nome_base)
        return f"{nome}_{email_limpo}{ext}"

def carregar_dados_financeiros():
    ficheiro = obter_ficheiro_utilizador(EXCEL_FINANCEIRA)
    if not os.path.exists(ficheiro):
        df = pd.DataFrame(columns=['ID', 'Data', 'Orcamento', 'Cliente', 'Ingredientes', 'MaoDeObra', 'Despesas', 'Receita', 'Notas'])
        df.to_excel(ficheiro, index=False)
    return pd.read_excel(ficheiro)

def guardar_dados_financeiros(df):
    ficheiro = obter_ficheiro_utilizador(EXCEL_FINANCEIRA)
    df.to_excel(ficheiro, index=False)

def filtrar_dados_por_periodo(df, tipo_periodo, mes_selecionado, ano_selecionado, meses_nomes, mes_atual_idx):
    if df.empty or 'Data' not in df.columns:
        return pd.DataFrame()
    
    df['Data_dt'] = pd.to_datetime(df['Data'], errors='coerce')
    if tipo_periodo == "Mensal":
        mes_idx = meses_nomes.index(mes_selecionado) + 1 if mes_selecionado in meses_nomes else mes_atual_idx
        return df[(df['Data_dt'].dt.year == ano_selecionado) & (df['Data_dt'].dt.month == mes_idx)]
    else:
        return df[df['Data_dt'].dt.year == ano_selecionado]

@financeira_bp.route("/financeira/exportar_excel")
def exportar_relatorio_excel():
    tipo_periodo = request.args.get("tipo_periodo", "Mensal")
    mes_selecionado = request.args.get("mes", "Janeiro")
    ano_selecionado = int(request.args.get("ano", datetime.now().year))
    
    meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                   "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    
    df = carregar_dados_financeiros()
    df_filtrado = filtrar_dados_por_periodo(df, tipo_periodo, mes_selecionado, ano_selecionado, meses_nomes, datetime.now().month)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df_filtrado.empty:
            df_export = df_filtrado.drop(columns=['Data_dt'], errors='ignore')
        else:
            df_export = pd.DataFrame(columns=['ID', 'Data', 'Orcamento', 'Cliente', 'Ingredientes', 'MaoDeObra', 'Despesas', 'Receita', 'Notas'])
        
        df_export.to_excel(writer, index=False, sheet_name='Relatorio_Financeiro')
        
    output.seek(0)
    nome_ficheiro = f"relatorio_financeiro_{mes_selecionado.lower()}_{ano_selecionado}.xlsx" if tipo_periodo == "Mensal" else f"relatorio_financeiro_anual_{ano_selecionado}.xlsx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=nome_ficheiro
    )

@financeira_bp.route("/financeira", methods=["GET", "POST"])
def analise_financeira():
    df = carregar_dados_financeiros()
    
    ano_atual = datetime.now().year
    mes_atual_idx = datetime.now().month
    
    meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                   "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    mes_atual_nome = meses_nomes[mes_atual_idx - 1]
    
    tipo_periodo = request.args.get("tipo_periodo", "Mensal")
    mes_selecionado = request.args.get("mes", mes_atual_nome)
    ano_selecionado = int(request.args.get("ano", ano_atual))
    
    anos_disponiveis = list(range(2026, max(ano_atual + 10, 2041)))
    
    df_filtrado = filtrar_dados_por_periodo(df, tipo_periodo, mes_selecionado, ano_selecionado, meses_nomes, mes_atual_idx)

    total_encomendas = len(df_filtrado)
    custo_ingredientes = df_filtrado['Ingredientes'].sum() if not df_filtrado.empty and 'Ingredientes' in df_filtrado.columns else 0.0
    custo_mao_obra = df_filtrado['MaoDeObra'].sum() if not df_filtrado.empty and 'MaoDeObra' in df_filtrado.columns else 0.0
    custo_despesas = df_filtrado['Despesas'].sum() if not df_filtrado.empty and 'Despesas' in df_filtrado.columns else 0.0
    
    custo_total = custo_ingredientes + custo_mao_obra + custo_despesas
    receita_total = df_filtrado['Receita'].sum() if not df_filtrado.empty and 'Receita' in df_filtrado.columns else 0.0
    lucro_total = receita_total - custo_total
    
    margem_lucro = (lucro_total / receita_total * 100) if receita_total > 0 else 0.0
    
    pct_ingredientes = (custo_ingredientes / custo_total * 100) if custo_total > 0 else 0.0
    pct_mao_obra = (custo_mao_obra / custo_total * 100) if custo_total > 0 else 0.0
    pct_despesas = (custo_despesas / custo_total * 100) if custo_total > 0 else 0.0

    labels_encomendas = []
    dados_custos = []
    dados_receitas = []
    dados_lucros = []

    linhas_tabela = ""
    if not df_filtrado.empty:
        for idx, row in df_filtrado.iterrows():
            id_reg = row.get('ID', idx)
            data_str = str(row.get('Data', ''))
            orcamento = str(row.get('Orcamento', ''))
            cliente = str(row.get('Cliente', ''))
            ing = float(row.get('Ingredientes', 0) or 0)
            mo = float(row.get('MaoDeObra', 0) or 0)
            desp = float(row.get('Despesas', 0) or 0)
            c_tot = ing + mo + desp
            rec = float(row.get('Receita', 0) or 0)
            luc = rec - c_tot
            mar = (luc / rec * 100) if rec > 0 else 0.0
            
            labels_encomendas.append(orcamento if orcamento else f"Reg. {id_reg}")
            dados_custos.append(round(c_tot, 2))
            dados_receitas.append(round(rec, 2))
            dados_lucros.append(round(luc, 2))

            linhas_tabela += f"""
            <tr>
                <td data-label="Data">{data_str}</td>
                <td data-label="Orçamento" style="font-weight: 600;">{orcamento}</td>
                <td data-label="Cliente">{cliente}</td>
                <td data-label="Ingredientes">{ing:.2f} €</td>
                <td data-label="Mão de Obra">{mo:.2f} €</td>
                <td data-label="Despesas">{desp:.2f} €</td>
                <td data-label="Custo Total" style="font-weight: bold;">{c_tot:.2f} €</td>
                <td data-label="Receita">{rec:.2f} €</td>
                <td data-label="Lucro" style="color: {'#2e7d32' if luc >= 0 else '#c62828'}; font-weight: bold;">{luc:.2f} €</td>
                <td data-label="Margem" style="color: {'#2e7d32' if mar >= 0 else '#c62828'};">{mar:.1f}%</td>
                <td data-label="Ações" style="text-align: right;"><a href="/financeira/apagar/{id_reg}" onclick="return confirm('Apagar este registo?')" style="text-decoration:none; color:#c62828;">🗑️</a></td>
            </tr>
            """
    if not linhas_tabela:
        linhas_tabela = '<tr><td colspan="11" style="text-align:center; color:#888; padding: 25px;">Nenhum registo financeiro encontrado para este período.</td></tr>'

    chart_labels_json = json.dumps(labels_encomendas)
    chart_custos_json = json.dumps(dados_custos)
    chart_receitas_json = json.dumps(dados_receitas)
    chart_lucros_json = json.dumps(dados_lucros)

    return render_template_string(f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Análise Financeira</title>
        <!-- Chart.js CDN -->
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #fdfbf7; margin: 0; padding: 30px; color: #333; }}
            .container {{ max-width: 1200px; margin: auto; }}
            
            .top-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; flex-wrap: wrap; gap: 15px; }}
            .title-area h1 {{ margin: 0; color: #111; font-size: 24px; }}
            .title-area p {{ margin: 5px 0 0 0; color: #666; font-size: 14px; }}
            .btn-export {{ background: #fff; border: 2px solid #f3e5d0; color: #333; padding: 10px 18px; border-radius: 8px; cursor: pointer; font-weight: bold; display: flex; align-items: center; gap: 6px; text-decoration: none; font-size: 14px; width: fit-content; }}
            .btn-export:hover {{ background: #f9f0e6; }}
            
            .box-card {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 25px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }}
            .box-title {{ font-size: 16px; font-weight: bold; color: #111; margin-bottom: 15px; display: flex; align-items: center; gap: 8px; }}
            .filters-grid {{ display: grid; grid-template-columns: repeat(3, 1fr) 2fr; gap: 20px; }}
            .filter-group label {{ display: block; font-size: 13px; font-weight: bold; color: #555; margin-bottom: 6px; }}
            .filter-select {{ width: 100%; padding: 10px 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; font-size: 14px; outline: none; box-sizing: border-box; }}
            
            .metrics-grid-4 {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 25px; }}
            .metric-card {{ background: #fffdf9; border: 2px solid #f3e5d0; border-radius: 12px; padding: 20px; position: relative; }}
            .metric-label {{ font-size: 13px; color: #666; font-weight: bold; margin-bottom: 8px; }}
            .metric-value {{ font-size: 26px; font-weight: bold; color: #111; }}
            .metric-sub {{ font-size: 12px; color: #888; margin-top: 5px; }}
            .metric-value.green {{ color: #2e7d32; }}
            .icon-corner {{ position: absolute; right: 20px; top: 20px; background: #f3e5d0; width: 35px; height: 35px; display: flex; align-items: center; justify-content: center; border-radius: 8px; font-size: 16px; }}

            .metrics-grid-3 {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 25px; }}
            
            .notes-textarea {{ width: 100%; height: 90px; padding: 12px; border: 2px solid #f3e5d0; border-radius: 8px; background: #fff; resize: vertical; font-size: 14px; outline: none; box-sizing: border-box; color: #444; }}

            .charts-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 25px; }}
            table.table-custom {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            table.table-custom th {{ background: transparent; color: #888; font-size: 12px; font-weight: bold; padding: 14px 10px; text-align: left; border-bottom: 2px solid #f3e5d0; text-transform: uppercase; }}
            table.table-custom td {{ padding: 14px 10px; border-bottom: 1px solid #f3e5d0; color: #444; font-size: 13px; }}

            .back-link {{ display: inline-block; margin-bottom: 15px; color: #b85d19; text-decoration: none; font-weight: bold; font-size: 14px; }}
            .back-link:hover {{ text-decoration: underline; }}

            /* Estilos Mobile-First / Responsivo para Telemóveis */
            @media screen and (max-width: 768px) {{
                body {{ padding: 15px; }}
                .filters-grid {{ grid-template-columns: 1fr; gap: 15px; }}
                .metrics-grid-4 {{ grid-template-columns: 1fr; gap: 15px; }}
                .metrics-grid-3 {{ grid-template-columns: 1fr; gap: 15px; }}
                .charts-grid {{ grid-template-columns: 1fr; gap: 15px; }}
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
            <a href="/" class="back-link">← Voltar ao Painel Principal</a>
            
            <div class="top-header">
                <div class="title-area">
                    <h1>Análise Financeira</h1>
                    <p>{mes_selecionado} {ano_selecionado}</p>
                </div>
                <div>
                    <a href="/financeira/exportar_excel?tipo_periodo={tipo_periodo}&mes={mes_selecionado}&ano={ano_selecionado}" class="btn-export">⬇️ Exportar Relatório Excel</a>
                </div>
            </div>

            <div class="box-card">
                <div class="box-title">📅 Período</div>
                <form method="GET" action="/financeira">
                    <div class="filters-grid">
                        <div class="filter-group">
                            <label>Tipo de Período</label>
                            <select name="tipo_periodo" class="filter-select" onchange="this.form.submit()">
                                <option value="Mensal" {'selected' if tipo_periodo=='Mensal' else ''}>Mensal</option>
                                <option value="Anual" {'selected' if tipo_periodo=='Anual' else ''}>Anual</option>
                            </select>
                        </div>
                        <div class="filter-group">
                            <label>Mês Final</label>
                            <select name="mes" class="filter-select" onchange="this.form.submit()">
                                {''.join([f'<option value="{m}" {"selected" if mes_selecionado==m else ""}>{m}</option>' for m in meses_nomes])}
                            </select>
                        </div>
                        <div class="filter-group">
                            <label>Ano</label>
                            <select name="ano" class="filter-select" onchange="this.form.submit()">
                                {''.join([f'<option value="{a}" {"selected" if ano_selecionado==a else ""}>{a}</option>' for a in anos_disponiveis])}
                            </select>
                        </div>
                    </div>
                </form>
            </div>

            <div class="metrics-grid-4">
                <div class="metric-card">
                    <div class="metric-label">Encomendas</div>
                    <div class="metric-value">{total_encomendas}</div>
                    <div class="metric-sub">Concluídas no período</div>
                    <div class="icon-corner">🛒</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Custo Total</div>
                    <div class="metric-value">{custo_total:.2f} €</div>
                    <div class="metric-sub">Custo interno</div>
                    <div class="icon-corner">€</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Receita Total</div>
                    <div class="metric-value">{receita_total:.2f} €</div>
                    <div class="metric-sub">Valor faturado</div>
                    <div class="icon-corner">📈</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Lucro Total</div>
                    <div class="metric-value green">{lucro_total:.2f} €</div>
                    <div class="metric-sub">{margem_lucro:.1f}% margem</div>
                    <div class="icon-corner">📊</div>
                </div>
            </div>

            <div class="metrics-grid-3">
                <div class="metric-card">
                    <div class="metric-label">Custo Ingredientes</div>
                    <div class="metric-value" style="font-size: 22px;">{custo_ingredientes:.2f} €</div>
                    <div class="metric-sub">{pct_ingredientes:.1f}% do custo total</div>
                    <div class="icon-corner" style="font-size: 14px;">📦</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Custo Mão de Obra</div>
                    <div class="metric-value" style="font-size: 22px;">{custo_mao_obra:.2f} €</div>
                    <div class="metric-sub">{pct_mao_obra:.1f}% do custo total</div>
                    <div class="icon-corner" style="font-size: 14px;">⏱️</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Custo Despesas</div>
                    <div class="metric-value" style="font-size: 22px;">{custo_despesas:.2f} €</div>
                    <div class="metric-sub">{pct_despesas:.1f}% do custo total</div>
                    <div class="icon-corner" style="font-size: 14px;">🧾</div>
                </div>
            </div>

            <div class="box-card">
                <div class="box-title">📄 Notas do Período</div>
                <textarea class="notes-textarea" placeholder="Adicione notas sobre este período (faturas, observações, etc.)"></textarea>
            </div>

            <div class="charts-grid">
                <div class="box-card" style="margin-bottom:0; display: flex; flex-direction: column;">
                    <div class="box-title">Análise por Encomenda</div>
                    <div style="flex: 1; position: relative; min-height: 220px; display: flex; align-items: center; justify-content: center;">
                        <canvas id="barChart" style="max-height: 220px;"></canvas>
                    </div>
                </div>
                <div class="box-card" style="margin-bottom:0; display: flex; flex-direction: column;">
                    <div class="box-title">Composição do Custo</div>
                    <div style="flex: 1; position: relative; min-height: 220px; display: flex; align-items: center; justify-content: center;">
                        <canvas id="doughnutChart" style="max-height: 220px;"></canvas>
                    </div>
                </div>
            </div>

            <div class="box-card" style="margin-top: 25px;">
                <div class="box-title">Detalhes por Encomenda</div>
                <table class="table-custom">
                    <thead>
                        <tr>
                            <th>Data</th>
                            <th>Orçamento</th>
                            <th>Cliente</th>
                            <th>Ingredientes</th>
                            <th>Mão de Obra</th>
                            <th>Despesas</th>
                            <th>Custo Total</th>
                            <th>Receita</th>
                            <th>Lucro</th>
                            <th>Margem</th>
                            <th style="text-align: right;">Ações</th>
                        </tr>
                    </thead>
                    <tbody>
                        {linhas_tabela}
                    </tbody>
                </table>
            </div>
        </div>

        <script>
            // Gráfico de Barras - Custo vs Receita vs Lucro por Encomenda
            const ctxBar = document.getElementById('barChart').getContext('2d');
            const barChart = new Chart(ctxBar, {{
                type: 'bar',
                data: {{
                    labels: {chart_labels_json},
                    datasets: [
                        {{
                            label: 'Custo Total',
                            data: {chart_custos_json},
                            backgroundColor: '#d99b6a',
                            borderRadius: 6
                        }},
                        {{
                            label: 'Receita',
                            data: {chart_receitas_json},
                            backgroundColor: '#4a3b32',
                            borderRadius: 6
                        }},
                        {{
                            label: 'Lucro',
                            data: {chart_lucros_json},
                            backgroundColor: '#2e7d32',
                            borderRadius: 6
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }}
                    }},
                    scales: {{
                        y: {{ beginAtZero: true, grid: {{ color: '#f3e5d0' }} }},
                        x: {{ grid: {{ display: false }} }}
                    }}
                }}
            }});

            // Gráfico de Rosca - Composição do Custo
            const ctxDoughnut = document.getElementById('doughnutChart').getContext('2d');
            const doughnutChart = new Chart(ctxDoughnut, {{
                type: 'doughnut',
                data: {{
                    labels: ['Ingredientes', 'Mão de Obra', 'Despesas'],
                    datasets: [{{
                        data: [{custo_ingredientes:.2f}, {custo_mao_obra:.2f}, {custo_despesas:.2f}],
                        backgroundColor: ['#d99b6a', '#4a3b32', '#c8915f'],
                        borderWidth: 2,
                        borderColor: '#fffdf9'
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }}
                    }}
                }}
            }});
        </script>
    </body>
    </html>
    """)

@financeira_bp.route("/financeira/apagar/<int:id_registo>")
def apagar_registo_financeiro(id_registo):
    df = carregar_dados_financeiros()
    if not df.empty and 'ID' in df.columns:
        df = df[df['ID'] != id_registo]
        guardar_dados_financeiros(df)
    return redirect(url_for('financeira_bp.analise_financeira'))