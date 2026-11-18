from flask import Blueprint, render_template_string, request, redirect, url_for, session
import pandas as pd
import os

maodeobra_bp = Blueprint('maodeobra_bp', __name__)

EXCEL_MAO_OBRA = 'mao_de_obra.xlsx'
EXCEL_CONFIG = 'config_mao_obra.xlsx'

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
    return pd.read_excel(ficheiro)

def guardar_dados(ficheiro_base, df):
    ficheiro = obter_ficheiro_utilizador(ficheiro_base)
    df.to_excel(ficheiro, index=False)

def obter_valor_hora():
    ficheiro = obter_ficheiro_utilizador(EXCEL_CONFIG)
    if os.path.exists(ficheiro):
        try:
            df = pd.read_excel(ficheiro)
            if not df.empty and 'ValorHora' in df.columns:
                return float(df.iloc[0]['ValorHora'] or 8.00)
        except:
            pass
    return 8.00

@maodeobra_bp.route("/maodeobra")
def gerir_mao_obra():
    df = carregar_dados(EXCEL_MAO_OBRA, ['ID', 'Tarefa', 'TempoMinutos'])
    valor_hora = obter_valor_hora()
    
    total_minutos = 0
    custo_total = 0.0
    linhas_tabela = ""
    
    if not df.empty:
        for idx, row in df.iterrows():
            id_tarefa = row.get('ID', idx)
            tarefa = str(row.get('Tarefa', ''))
            minutos = float(row.get('TempoMinutos', 10) or 10)
            total_minutos += minutos
            
            custo = (minutos / 60) * valor_hora
            custo_total += custo
            
            if minutos >= 60:
                h = int(minutos // 60)
                m = int(minutos % 60)
                tempo_str = f"{h}h {m}min" if m > 0 else f"{h}h"
            else:
                tempo_str = f"{int(minutos)}min"
                
            linhas_tabela += f"""
            <tr>
                <td data-label="Tarefa" style="font-weight: 600; color: #4A3B32;">{tarefa}</td>
                <td data-label="Tempo" style="color: #8C7A6B;">{tempo_str}</td>
                <td data-label="Custo" style="font-weight: bold; color: #D99B6A;">€{custo:.2f}</td>
                <td data-label="Ações" style="text-align: right;">
                    <button onclick="abrirModalEditar('{id_tarefa}', '{tarefa}', '{minutos}')" style="background:none; border:none; cursor:pointer; font-size: 16px; margin-right: 8px;" title="Editar">✏️</button>
                    <a href="/maodeobra/apagar/{id_tarefa}" onclick="return confirm('Tens a certeza que queres apagar esta tarefa?')" style="text-decoration: none; font-size: 16px;" title="Apagar">🗑️</a>
                </td>
            </tr>
            """
            
    if not linhas_tabela:
        linhas_tabela = '<tr><td colspan="4" style="text-align:center; color:#8C7A6B; padding: 30px;">Nenhuma tarefa registada.</td></tr>'

    total_h = int(total_minutos // 60)
    total_m = int(total_minutos % 60)
    tempo_total_str = f"{total_h}h {total_m}min" if total_h > 0 else f"{total_m}min"

    return render_template_string(f"""
    <!DOCTYPE html>
    <html lang="pt">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>CáDoces - Mão de Obra</title>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #FAF8F5; color: #4A3B32; margin: 0; padding: 30px; }}
            .container {{ max-width: 1200px; margin: auto; }}
            
            .header-bar {{ display: flex; justify-content: space-between; align-items: center; background: #FFFFFF; padding: 20px 25px; border: 1px solid #EEDCCE; border-radius: 14px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); margin-bottom: 25px; flex-wrap: wrap; gap: 15px; }}
            .header-left h1 {{ margin: 0; color: #2C1E1A; font-size: 24px; font-weight: 700; }}
            .header-left p {{ margin: 4px 0 0 0; color: #8C7A6B; font-size: 14px; }}
            
            .header-actions {{ display: flex; gap: 12px; flex-wrap: wrap; }}
            .btn-secondary-top {{ background: #F5EBE0; border: 1px solid #EEDCCE; color: #4A3B32; padding: 10px 18px; border-radius: 10px; cursor: pointer; font-weight: 600; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; transition: all 0.2s; }}
            .btn-secondary-top:hover {{ background: #EEDCCE; }}
            
            .btn-primary {{ background: #D99B6A; color: #FFFFFF; border: none; padding: 10px 20px; border-radius: 10px; cursor: pointer; font-weight: 600; text-decoration: none; display: flex; align-items: center; gap: 6px; font-size: 14px; box-shadow: 0 2px 4px rgba(217,155,106,0.2); transition: background 0.2s; }}
            .btn-primary:hover {{ background: #C48757; }}
            
            .cards-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 35px; }}
            .card-metric {{ background: #FFFFFF; border: 1px solid #EEDCCE; border-radius: 14px; padding: 20px 25px; position: relative; box-shadow: 0 2px 6px rgba(74,59,50,0.02); }}
            .card-metric.highlight {{ background: #FEF9E7; border-color: #F9E79F; }}
            .card-header {{ display: flex; justify-content: space-between; align-items: center; font-size: 14px; font-weight: 600; color: #734D2E; margin-bottom: 12px; }}
            .card-value {{ font-size: 32px; font-weight: bold; color: #2C1E1A; }}
            .icon-badge {{ background: #F5EBE0; width: 35px; height: 35px; display: flex; align-items: center; justify-content: center; border-radius: 8px; font-size: 16px; border: 1px solid #EEDCCE; }}
            
            .section-title {{ font-size: 20px; font-weight: 700; color: #2C1E1A; margin-bottom: 15px; }}
            .content-box {{ background: #FFFFFF; border: 1px solid #EEDCCE; border-radius: 14px; padding: 10px 25px; box-shadow: 0 2px 6px rgba(74,59,50,0.02); }}
            table.table-custom {{ width: 100%; border-collapse: collapse; }}
            table.table-custom th {{ background-color: transparent; color: #734D2E; font-size: 13px; font-weight: 600; padding: 16px 12px; text-align: left; border-bottom: 2px solid #EEDCCE; text-transform: uppercase; }}
            table.table-custom td {{ padding: 16px 12px; border-bottom: 1px solid #F5EBE0; color: #4A3B32; font-size: 14px; }}
            
            .modal {{ display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(74,59,50,0.4); overflow-y: auto; padding: 40px 0; }}
            .modal-content {{ background-color: #FFFFFF; margin: auto; padding: 30px; border: 1px solid #EEDCCE; width: 480px; max-width: 90%; border-radius: 14px; box-shadow: 0 8px 24px rgba(74,59,50,0.1); position: relative; box-sizing: border-box; }}
            .close-btn {{ position: absolute; right: 20px; top: 20px; font-size: 20px; cursor: pointer; color: #8C7A6B; font-weight: bold; background: #F5EBE0; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; border: 1px solid #EEDCCE; }}
            .form-group {{ margin-bottom: 18px; }}
            .form-group label {{ display: block; font-weight: 600; color: #734D2E; margin-bottom: 6px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }}
            .form-control {{ width: 100%; padding: 12px; border: 1px solid #EEDCCE; border-radius: 8px; background: #FAF8F5; font-size: 14px; box-sizing: border-box; outline: none; color: #4A3B32; }}
            .form-control:focus {{ border-color: #D99B6A; background: #FFFFFF; }}
            .modal-footer {{ display: flex; justify-content: flex-end; gap: 12px; border-top: 1px solid #EEDCCE; padding-top: 20px; margin-top: 20px; flex-wrap: wrap; }}
            .btn-secondary {{ background: #F5EBE0; border: 1px solid #EEDCCE; color: #4A3B32; padding: 10px 20px; border-radius: 8px; cursor: pointer; font-weight: 600; }}
            
            .back-link {{ display: inline-block; margin-bottom: 15px; color: #D99B6A; text-decoration: none; font-weight: 600; font-size: 14px; }}
            .back-link:hover {{ text-decoration: underline; }}
        </style>
    </head>
    <body>
        <div class="container">
            <a href="/" class="back-link">← Voltar ao Painel Principal</a>
            
            <div class="header-bar">
                <div class="header-left">
                    <h1>Mão de Obra</h1>
                    <p>Gerencie tarefas, sub-tarefas e tempos de produção</p>
                </div>
                <div class="header-actions">
                    <button class="btn-secondary-top" onclick="abrirModalValorHora()">⚙️ Valor/Hora</button>
                    <button class="btn-primary" onclick="abrirModalNovaTarefa()">+ Adicionar Tarefa</button>
                </div>
            </div>

            <div class="cards-grid">
                <div class="card-metric highlight">
                    <div class="card-header"><span>Valor por Hora</span> <div class="icon-badge">€</div></div>
                    <div class="card-value">€{valor_hora:.2f}</div>
                </div>
                <div class="card-metric">
                    <div class="card-header"><span>Tempo Total</span> <div class="icon-badge">⏱️</div></div>
                    <div class="card-value">{tempo_total_str}</div>
                </div>
                <div class="card-metric">
                    <div class="card-header"><span>Custo Total</span> <div class="icon-badge">€</div></div>
                    <div class="card-value">€{custo_total:.2f}</div>
                </div>
            </div>

            <div class="section-title">Lista de Tarefas e Sub-tarefas</div>
            <div class="content-box">
                <table class="table-custom">
                    <thead>
                        <tr>
                            <th>Tarefa</th>
                            <th>Tempo Total (minutos)</th>
                            <th>Custo</th>
                            <th style="text-align: right;">Ações</th>
                        </tr>
                    </thead>
                    <tbody>
                        {linhas_tabela}
                    </tbody>
                </table>
            </div>
        </div>

        <div id="modalTarefa" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModal()">&times;</span>
                <h2 id="tituloModalTarefa" style="margin-top:0; color:#2C1E1A; font-size:20px;">Adicionar Tarefa</h2>
                
                <form action="/maodeobra/salvar" method="POST">
                    <input type="hidden" name="id" id="tarefaId">
                    
                    <div class="form-group">
                        <label>Nome da Tarefa</label>
                        <input type="text" name="tarefa" id="tarefaNome" class="form-control" placeholder="Ex: Decoração com fondant" required>
                    </div>

                    <div class="form-group">
                        <label>Tempo Total (minutos)</label>
                        <input type="number" step="1" name="tempo" id="tarefaTempo" class="form-control" placeholder="0" required>
                    </div>

                    <div class="form-group">
                        <label style="display: flex; justify-content: space-between; align-items: center;">
                            <span>Sub-tarefas (opcional)</span>
                            <button type="button" class="btn-secondary" style="padding: 4px 10px; font-size: 12px;" onclick="alert('Funcionalidade de sub-tarefas em desenvolvimento')">+ Adicionar Etapa</button>
                        </label>
                    </div>

                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModal()">Cancelar</button>
                        <button type="submit" class="btn-primary">Adicionar</button>
                    </div>
                </form>
            </div>
        </div>

        <div id="modalValorHora" class="modal">
            <div class="modal-content">
                <span class="close-btn" onclick="fecharModalValorHora()">&times;</span>
                <h2 style="margin-top:0; color:#2C1E1A; font-size:20px;">Definir Valor por Hora</h2>
                
                <form action="/maodeobra/valor_hora" method="POST">
                    <div class="form-group">
                        <label>Valor por Hora (€)</label>
                        <input type="number" step="0.01" name="valor_hora" class="form-control" value="{valor_hora:.2f}" required>
                    </div>

                    <div class="modal-footer">
                        <button type="button" class="btn-secondary" onclick="fecharModalValorHora()">Cancelar</button>
                        <button type="submit" class="btn-primary">Atualizar Valor</button>
                    </div>
                </form>
            </div>
        </div>

        <script>
            function abrirModalNovaTarefa() {{
                document.getElementById('tituloModalTarefa').innerText = 'Adicionar Tarefa';
                document.getElementById('tarefaId').value = '';
                document.getElementById('tarefaNome').value = '';
                document.getElementById('tarefaTempo').value = '0';
                document.getElementById('modalTarefa').style.display = 'block';
            }}

            function abrirModalEditar(id, nome, tempo) {{
                document.getElementById('tituloModalTarefa').innerText = 'Editar Tarefa';
                document.getElementById('tarefaId').value = id;
                document.getElementById('tarefaNome').value = nome;
                document.getElementById('tarefaTempo').value = tempo;
                document.getElementById('modalTarefa').style.display = 'block';
            }}

            function fecharModal() {{ document.getElementById('modalTarefa').style.display = 'none'; }}

            function abrirModalValorHora() {{ document.getElementById('modalValorHora').style.display = 'block'; }}
            function fecharModalValorHora() {{ document.getElementById('modalValorHora').style.display = 'none'; }}
        </script>
    </body>
    </html>
    """)

@maodeobra_bp.route("/maodeobra/salvar", methods=["POST"])
def salvar_tarefa():
    id_tarefa = request.form.get("id")
    tarefa = request.form.get("tarefa")
    tempo = float(request.form.get("tempo", 0) or 0)
    
    df = carregar_dados(EXCEL_MAO_OBRA, ['ID', 'Tarefa', 'TempoMinutos'])
    
    if id_tarefa and id_tarefa.isdigit():
        idx = df[df['ID'].astype(str) == str(id_tarefa)].index
        if not idx.empty:
            df.loc[idx, ['Tarefa', 'TempoMinutos']] = [tarefa, tempo]
    else:
        novo_id = int(df['ID'].max()) + 1 if not df.empty and 'ID' in df.columns and not df['ID'].dropna().empty else 1
        nova_linha = {
            'ID': novo_id,
            'Tarefa': tarefa,
            'TempoMinutos': tempo
        }
        df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        
    guardar_dados(EXCEL_MAO_OBRA, df)
    return redirect(url_for('maodeobra_bp.gerir_mao_obra'))

@maodeobra_bp.route("/maodeobra/apagar/<int:id_tarefa>")
def apagar_tarefa(id_tarefa):
    df = carregar_dados(EXCEL_MAO_OBRA, ['ID', 'Tarefa', 'TempoMinutos'])
    if not df.empty:
        df = df[df['ID'] != id_tarefa]
        guardar_dados(EXCEL_MAO_OBRA, df)
    return redirect(url_for('maodeobra_bp.gerir_mao_obra'))

@maodeobra_bp.route("/maodeobra/valor_hora", methods=["POST"])
def atualizar_valor_hora():
    novo_valor = float(request.form.get("valor_hora", 8.00) or 8.00)
    df = pd.DataFrame([{'ValorHora': novo_valor}])
    guardar_dados(EXCEL_CONFIG, df)
    return redirect(url_for('maodeobra_bp.gerir_mao_obra'))