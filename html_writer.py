"""Gerador do Dashboard Visual Interativo HTML autônomo (dashboard_atos.html).

Gera um painel web completo, moderno e responsivo para visualização e verificação
imediata de todos os atos consolidados pelo assessor de magistrado do TJPE (8.17).

Funciona 100% offline (com fallbacks) e abre diretamente em qualquer navegador
(Chrome, Edge, Firefox) sem necessidade de servidor web.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional


def carregar_todos_atos_xlsx(caminho_xlsx: Path) -> List[Dict[str, Any]]:
    """Carrega todos os atos consolidados da aba Lançamento do arquivo Excel."""
    atos: List[Dict[str, Any]] = []
    if not caminho_xlsx.exists():
        return atos
    try:
        import openpyxl  # type: ignore
        wb = openpyxl.load_workbook(str(caminho_xlsx), data_only=True, read_only=True)
        if "Lançamento" in wb.sheetnames:
            ws = wb["Lançamento"]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) > 1:
                for r in rows[1:]:
                    if not r or not r[0]:
                        continue
                    atos.append({
                        "npu": str(r[0] or "").strip(),
                        "classe_judicial": str(r[1] or "").strip(),
                        "assunto_principal": str(r[2] or "").strip(),
                        "data_autuacao": str(r[3] or "").strip(),
                        "valor_causa": str(r[4] or "").strip(),
                        "remessa": str(r[5] or "").strip(),
                        "tipo_ato": str(r[7] or "Indefinido").strip(),
                        "data_ato": str(r[8] or "").strip(),
                        "palavras_chave": str(r[14] or "").strip(),
                        "resultado_sentenca": str(r[15] or "").strip(),
                    })
        wb.close()
    except Exception:
        pass
    return atos


def gerar_dashboard_html(
    caminho_html: Path,
    caminho_xlsx: Path,
    atos_dados: Optional[List[Dict[str, Any]]] = None,
    erros: Optional[List[Any]] = None,
) -> Path:
    """Gera o arquivo dashboard_atos.html com dados consolidados e gráficos interativos."""
    caminho_html = Path(caminho_html)
    caminho_xlsx = Path(caminho_xlsx)
    caminho_html.parent.mkdir(parents=True, exist_ok=True)

    # Se a lista passada estiver vazia ou for parcial, tenta carregar histórico integral da planilha
    todos_da_planilha = carregar_todos_atos_xlsx(caminho_xlsx)
    if len(todos_da_planilha) >= (len(atos_dados) if atos_dados else 0):
        lista_final = todos_da_planilha
    else:
        lista_final = atos_dados or []

    # 1. Agregações para os KPIs e Gráficos
    total_atos = len(lista_final)
    despachos = sum(1 for a in lista_final if a.get("tipo_ato") == "Despacho")
    decisoes = sum(1 for a in lista_final if a.get("tipo_ato") == "Decisão")
    sentencas = sum(1 for a in lista_final if a.get("tipo_ato") == "Sentença")
    indefinidos = sum(1 for a in lista_final if a.get("tipo_ato") == "Indefinido")

    # Contagem de EDs
    eds = sum(
        1
        for a in lista_final
        if "EMBARGOS" in str(a.get("palavras_chave", "")).upper()
        or "(ED)" in str(a.get("resultado_sentenca", ""))
        or "ED" in str(a.get("palavras_chave", "")).split(",")
    )

    # Contagem de resultados de sentença
    resultados_dict: Dict[str, int] = {}
    for a in lista_final:
        res = a.get("resultado_sentenca")
        if res:
            res_str = str(res).strip()
            resultados_dict[res_str] = resultados_dict.get(res_str, 0) + 1

    # Contagem de remessas
    remessas_dict: Dict[str, int] = {}
    for a in lista_final:
        rem = a.get("remessa")
        if rem and rem != "PENDENTE":
            rem_str = str(rem).strip()
            remessas_dict[rem_str] = remessas_dict.get(rem_str, 0) + 1

    top_remessas = sorted(remessas_dict.items(), key=lambda x: x[1], reverse=True)[:8]

    # Serializa dados limpos para injeção no script cliente
    dados_json = json.dumps(lista_final, ensure_ascii=False)
    resultados_json = json.dumps(resultados_dict, ensure_ascii=False)
    top_remessas_labels = json.dumps([r[0] for r in top_remessas], ensure_ascii=False)
    top_remessas_valores = json.dumps([r[1] for r in top_remessas], ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard de Atos Judiciais - TJPE (8.17)</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    :root {{
      --tjpe-navy: #1B365D;
      --tjpe-blue: #0047AB;
      --tjpe-gold: #C5A059;
      --bg-slate: #F8FAFC;
      --border-color: #E2E8F0;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }}
    body {{ background-color: var(--bg-slate); color: #1E293B; padding-bottom: 50px; }}
    
    /* Top Header */
    .header {{
      background: linear-gradient(135deg, #0F172A 0%, #1B365D 50%, #1E3A8A 100%);
      color: white;
      padding: 24px 32px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.1);
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 16px;
    }}
    .header h1 {{ font-size: 22px; font-weight: 700; display: flex; align-items: center; gap: 10px; }}
    .header p {{ font-size: 13px; color: #94A3B8; margin-top: 4px; }}
    .badge-tjpe {{
      background: rgba(197, 160, 89, 0.2);
      border: 1px solid var(--tjpe-gold);
      color: #E2E8F0;
      padding: 6px 14px;
      border-radius: 9999px;
      font-size: 12px;
      font-weight: 600;
    }}

    .container {{ max-width: 1380px; margin: 24px auto; padding: 0 20px; }}

    /* KPI Cards Grid */
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }}
    .kpi-card {{
      background: white;
      border-radius: 12px;
      padding: 18px 20px;
      border: 1px solid var(--border-color);
      box-shadow: 0 1px 3px rgba(0,0,0,0.05);
      transition: transform 0.2s, box-shadow 0.2s;
    }}
    .kpi-card:hover {{ transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.08); }}
    .kpi-title {{ font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; color: #64748B; }}
    .kpi-val {{ font-size: 32px; font-weight: 800; margin-top: 6px; line-height: 1; }}
    .kpi-desc {{ font-size: 11px; color: #94A3B8; margin-top: 8px; }}

    .kpi-total .kpi-val {{ color: #0F172A; }}
    .kpi-despacho .kpi-val {{ color: #2563EB; }}
    .kpi-decisao .kpi-val {{ color: #D97706; }}
    .kpi-sentenca .kpi-val {{ color: #059669; }}
    .kpi-ed .kpi-val {{ color: #7C3AED; }}

    /* Charts Section */
    .charts-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
      gap: 20px;
      margin-bottom: 24px;
    }}
    .chart-box {{
      background: white;
      border-radius: 12px;
      padding: 20px;
      border: 1px solid var(--border-color);
      box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }}
    .chart-box h3 {{
      font-size: 15px;
      font-weight: 700;
      color: #1E293B;
      margin-bottom: 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    .canvas-container {{ position: relative; height: 260px; width: 100%; }}

    /* Filters & Search */
    .controls-panel {{
      background: white;
      border-radius: 12px;
      padding: 16px 20px;
      border: 1px solid var(--border-color);
      margin-bottom: 16px;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
    }}
    .search-box {{
      flex: 1;
      min-width: 260px;
      position: relative;
    }}
    .search-input {{
      width: 100%;
      padding: 10px 16px;
      border: 1px solid #CBD5E1;
      border-radius: 8px;
      font-size: 14px;
      outline: none;
    }}
    .search-input:focus {{ border-color: var(--tjpe-blue); box-shadow: 0 0 0 3px rgba(0, 71, 171, 0.15); }}

    .filter-buttons {{ display: flex; gap: 8px; flex-wrap: wrap; }}
    .btn-filter {{
      background: #F1F5F9;
      border: 1px solid #E2E8F0;
      padding: 8px 14px;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 600;
      color: #475569;
      cursor: pointer;
      transition: all 0.2s;
    }}
    .btn-filter:hover {{ background: #E2E8F0; color: #0F172A; }}
    .btn-filter.active {{ background: #1B365D; color: white; border-color: #1B365D; }}

    /* Table Section */
    .table-card {{
      background: white;
      border-radius: 12px;
      border: 1px solid var(--border-color);
      box-shadow: 0 1px 3px rgba(0,0,0,0.05);
      overflow: hidden;
    }}
    .table-header-bar {{
      padding: 16px 20px;
      border-bottom: 1px solid var(--border-color);
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: #FAFAFA;
    }}
    .table-responsive {{ overflow-x: auto; max-height: 600px; }}
    table {{ width: 100%; border-collapse: collapse; text-align: left; font-size: 13px; }}
    th {{
      background: #F8FAFC;
      color: #475569;
      font-weight: 600;
      padding: 12px 16px;
      border-bottom: 1px solid var(--border-color);
      white-space: nowrap;
      position: sticky;
      top: 0;
      z-index: 10;
    }}
    td {{
      padding: 12px 16px;
      border-bottom: 1px solid #F1F5F9;
      color: #334155;
    }}
    tr:hover td {{ background: #F8FAFC; }}
    
    .badge {{
      display: inline-block;
      padding: 4px 8px;
      border-radius: 6px;
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
    }}
    .badge-despacho {{ background: #DBEAFE; color: #1E40AF; }}
    .badge-decisao {{ background: #FEF3C7; color: #92400E; }}
    .badge-sentenca {{ background: #D1FAE5; color: #065F46; }}
    .badge-indefinido {{ background: #F1F5F9; color: #64748B; }}

    .npu-cell {{ font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-weight: 600; color: #0F172A; white-space: nowrap; }}
    .meta-tag {{ background: #F1F5F9; color: #475569; padding: 2px 6px; border-radius: 4px; font-size: 11px; margin-right: 4px; display: inline-block; margin-top: 2px; }}

    /* Footer Banner */
    .footer-bar {{
      margin-top: 24px;
      background: #EFF6FF;
      border: 1px solid #BFDBFE;
      border-radius: 12px;
      padding: 16px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
    }}
    .footer-bar span {{ font-size: 13px; color: #1E40AF; font-weight: 500; }}
  </style>
</head>
<body>

  <!-- Top Header -->
  <header class="header">
    <div>
      <h1>
        <span>⚖️</span>
        TJPE • Painel de Gestão de Atos Judiciais
      </h1>
      <p>Assessoria de Magistrado • Segmento 8.17 • 100% LGPD Seguro (Sem nomes de partes ou dados sensíveis)</p>
    </div>
    <div>
      <span class="badge-tjpe">Total Consolidado: {total_atos} Atos</span>
    </div>
  </header>

  <div class="container">
    
    <!-- KPI Cards -->
    <div class="kpi-grid">
      <div class="kpi-card kpi-total">
        <div class="kpi-title">Total de Atos</div>
        <div class="kpi-val">{total_atos}</div>
        <div class="kpi-desc">Consolidados na planilha</div>
      </div>
      <div class="kpi-card kpi-despacho">
        <div class="kpi-title">Despachos</div>
        <div class="kpi-val">{despachos}</div>
        <div class="kpi-desc">{(despachos / max(total_atos, 1) * 100):.1f}% do total</div>
      </div>
      <div class="kpi-card kpi-decisao">
        <div class="kpi-title">Decisões</div>
        <div class="kpi-val">{decisoes}</div>
        <div class="kpi-desc">{(decisoes / max(total_atos, 1) * 100):.1f}% do total</div>
      </div>
      <div class="kpi-card kpi-sentenca">
        <div class="kpi-title">Sentenças</div>
        <div class="kpi-val">{sentencas}</div>
        <div class="kpi-desc">{(sentencas / max(total_atos, 1) * 100):.1f}% do total</div>
      </div>
      <div class="kpi-card kpi-ed">
        <div class="kpi-title">Embargos (ED)</div>
        <div class="kpi-val">{eds}</div>
        <div class="kpi-desc">Atos de embargos declaratórios</div>
      </div>
    </div>

    <!-- Charts Section -->
    <div class="charts-grid">
      <div class="chart-box">
        <h3>Distribuição por Tipo de Ato</h3>
        <div class="canvas-container">
          <canvas id="chartTipos"></canvas>
        </div>
      </div>
      <div class="chart-box">
        <h3>Resultados de Sentenças & EDs</h3>
        <div class="canvas-container">
          <canvas id="chartResultados"></canvas>
        </div>
      </div>
      <div class="chart-box">
        <h3>Top Remessas / Unidades</h3>
        <div class="canvas-container">
          <canvas id="chartRemessas"></canvas>
        </div>
      </div>
    </div>

    <!-- Search & Filter Controls -->
    <div class="controls-panel">
      <div class="search-box">
        <input 
          type="text" 
          id="searchInput" 
          class="search-input" 
          placeholder="Filtrar por NPU, palavra-chave, remessa ou assunto..."
          oninput="filtrarAtos()"
        >
      </div>
      <div class="filter-buttons">
        <button class="btn-filter active" onclick="setFiltroTipo('TODOS', this)">Todos ({total_atos})</button>
        <button class="btn-filter" onclick="setFiltroTipo('Sentença', this)">Sentenças ({sentencas})</button>
        <button class="btn-filter" onclick="setFiltroTipo('Decisão', this)">Decisões ({decisoes})</button>
        <button class="btn-filter" onclick="setFiltroTipo('Despacho', this)">Despachos ({despachos})</button>
      </div>
    </div>

    <!-- Table Card -->
    <div class="table-card">
      <div class="table-header-bar">
        <strong id="contadorTabela">Exibindo {total_atos} atos judiciais</strong>
        <span style="font-size: 12px; color: #64748B;">Planilha vinculada: <code>{caminho_xlsx.name}</code></span>
      </div>
      <div class="table-responsive">
        <table id="tabelaAtos">
          <thead>
            <tr>
              <th>Tipo</th>
              <th>NPU (Processo)</th>
              <th>Data do Ato</th>
              <th>Remessa</th>
              <th>Classe Judicial</th>
              <th>Assunto Principal</th>
              <th>Resultado / Palavras-chave</th>
            </tr>
          </thead>
          <tbody id="tabelaCorpo">
            <!-- Linhas inseridas via JavaScript -->
          </tbody>
        </table>
      </div>
    </div>

    <!-- Footer Help Bar -->
    <div class="footer-bar">
      <div>
        <span>📁 Local dos arquivos gerados:</span>
        <div style="font-size: 12px; color: #3B82F6; margin-top: 4px; font-family: monospace;">
          Excel: {caminho_xlsx.resolve()}
        </div>
      </div>
      <div style="font-size: 12px; color: #1E3A8A;">
        💡 Na planilha Excel, clique na aba inferior <strong>"Dashboard"</strong> para ver os gráficos e cartões nativos do Excel.
      </div>
    </div>

  </div>

  <script>
    const ATOS_DADOS = {dados_json};
    const RESULTADOS_DADOS = {resultados_json};
    const REMESSAS_LABELS = {top_remessas_labels};
    const REMESSAS_VALORES = {top_remessas_valores};

    let filtroTipoAtual = 'TODOS';

    // 1. Renderiza Gráficos Chart.js
    function inicializarGraficos() {{
      if (typeof Chart === 'undefined') return;

      // Gráfico Tipos
      new Chart(document.getElementById('chartTipos'), {{
        type: 'doughnut',
        data: {{
          labels: ['Despachos', 'Decisões', 'Sentenças', 'Indefinidos'],
          datasets: [{{
            data: [{despachos}, {decisoes}, {sentencas}, {indefinidos}],
            backgroundColor: ['#2563EB', '#F59E0B', '#10B981', '#94A3B8'],
            borderWidth: 2,
            borderColor: '#FFFFFF'
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{
            legend: {{ position: 'bottom' }}
          }}
        }}
      }});

      // Gráfico Resultados de Sentença
      const resLabels = Object.keys(RESULTADOS_DADOS);
      const resValores = Object.values(RESULTADOS_DADOS);
      new Chart(document.getElementById('chartResultados'), {{
        type: 'bar',
        data: {{
          labels: resLabels.length ? resLabels : ['Nenhum resultado'],
          datasets: [{{
            label: 'Sentenças / EDs',
            data: resValores.length ? resValores : [0],
            backgroundColor: '#059669',
            borderRadius: 6
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          indexAxis: 'y',
          plugins: {{ legend: {{ display: false }} }},
          scales: {{
            x: {{ beginAtZero: true, ticks: {{ stepSize: 1 }} }}
          }}
        }}
      }});

      // Gráfico Remessas
      new Chart(document.getElementById('chartRemessas'), {{
        type: 'bar',
        data: {{
          labels: REMESSAS_LABELS.length ? REMESSAS_LABELS : ['Geral'],
          datasets: [{{
            label: 'Atos',
            data: REMESSAS_VALORES.length ? REMESSAS_VALORES : [0],
            backgroundColor: '#1E3A8A',
            borderRadius: 6
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{ legend: {{ display: false }} }},
          scales: {{
            y: {{ beginAtZero: true, ticks: {{ stepSize: 1 }} }}
          }}
        }}
      }});
    }}

    // 2. Renderiza Linhas da Tabela
    function renderTabela(lista) {{
      const corpo = document.getElementById('tabelaCorpo');
      corpo.innerHTML = '';
      
      document.getElementById('contadorTabela').textContent = `Exibindo ${{lista.length}} atos judiciais`;

      if (lista.length === 0) {{
        corpo.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 24px; color: #94A3B8;">Nenhum ato encontrado para o filtro aplicado.</td></tr>';
        return;
      }}

      lista.forEach(ato => {{
        const tr = document.createElement('tr');
        
        let badgeClass = 'badge-indefinido';
        if (ato.tipo_ato === 'Despacho') badgeClass = 'badge-despacho';
        else if (ato.tipo_ato === 'Decisão') badgeClass = 'badge-decisao';
        else if (ato.tipo_ato === 'Sentença') badgeClass = 'badge-sentenca';

        const palavras = (ato.palavras_chave || '').split(',').map(p => p.trim()).filter(Boolean);
        const tagsHtml = palavras.map(p => `<span class="meta-tag">${{p}}</span>`).join('');
        const resSentencaHtml = ato.resultado_sentenca ? `<div style="font-weight: 600; color: #047857; margin-bottom: 2px;">${{ato.resultado_sentenca}}</div>` : '';

        tr.innerHTML = `
          <td><span class="badge ${{badgeClass}}">${{ato.tipo_ato || 'Indefinido'}}</span></td>
          <td class="npu-cell">${{ato.npu || 'PENDENTE'}}</td>
          <td>${{ato.data_ato || '-'}}</td>
          <td><span class="meta-tag">${{ato.remessa || 'Sem Remessa'}}</span></td>
          <td style="max-width: 200px; font-size: 12px;">${{ato.classe_judicial || 'PENDENTE'}}</td>
          <td style="max-width: 220px; font-size: 12px;">${{ato.assunto_principal || 'PENDENTE'}}</td>
          <td>
            ${{resSentencaHtml}}
            ${{tagsHtml}}
          </td>
        `;
        corpo.appendChild(tr);
      }});
    }}

    // 3. Filtros
    function setFiltroTipo(tipo, btn) {{
      filtroTipoAtual = tipo;
      document.querySelectorAll('.btn-filter').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      filtrarAtos();
    }}

    function filtrarAtos() {{
      const termo = document.getElementById('searchInput').value.toLowerCase().trim();
      const filtrados = ATOS_DADOS.filter(ato => {{
        const matchesTipo = (filtroTipoAtual === 'TODOS') || (ato.tipo_ato === filtroTipoAtual);
        if (!matchesTipo) return false;

        if (!termo) return true;
        
        const textoBusca = [
          ato.npu || '',
          ato.tipo_ato || '',
          ato.remessa || '',
          ato.classe_judicial || '',
          ato.assunto_principal || '',
          ato.palavras_chave || '',
          ato.resultado_sentenca || ''
        ].join(' ').toLowerCase();

        return textoBusca.includes(termo);
      }});

      renderTabela(filtrados);
    }}

    window.addEventListener('DOMContentLoaded', () => {{
      inicializarGraficos();
      renderTabela(ATOS_DADOS);
    }});
  </script>
</body>
</html>
"""

    caminho_html.write_text(html_content, encoding="utf-8")
    return caminho_html
