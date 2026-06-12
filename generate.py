#!/usr/bin/env python3
"""
generate.py - Script de actualización automática del dashboard Jecel
Descarga el Excel de Google Drive, procesa los datos y genera index.html
Ejecutado por GitHub Actions diariamente a las 7am hora Venezuela.
"""

import openpyxl
import json
import re
from collections import defaultdict
from datetime import datetime

MESES = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
         7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
META_MENSUAL = 300000.0


def cargar_datos(archivo='ventas.xlsx'):
    wb = openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    header_skip = True
    for row in ws.iter_rows(values_only=True):
        if header_skip:
            header_skip = False
            continue
        if row[0] is None:
            continue
        ref   = str(row[0]).strip() if row[0] else ''
        fecha = row[1]
        cli   = str(row[2]).strip() if row[2] else ''
        vend  = str(row[3]).strip() if row[3] else ''
        total_raw = row[5]
        estado = str(row[6]).strip() if row[6] else ''

        # parse total
        try:
            if isinstance(total_raw, str):
                total = float(total_raw.replace(',',''))
            else:
                total = float(total_raw) if total_raw else 0.0
        except:
            total = 0.0

        # parse fecha
        try:
            if isinstance(fecha, datetime):
                dt = fecha
            else:
                dt = datetime.strptime(str(fecha)[:19], '%Y-%m-%d %H:%M:%S')
        except:
            continue

        rows.append({'ref':ref,'fecha':dt.strftime('%Y-%m-%d'),'cliente':cli,
                     'vendedor':vend,'total':total,'estado':estado,'dt':dt})
    wb.close()
    return rows


def calcular_kpis(rows):
    now   = datetime.now()
    cy, cm = now.year, now.month

    ytd = [r for r in rows if r['dt'].year == cy]
    mtd = [r for r in rows if r['dt'].year == cy and r['dt'].month == cm]
    ytd_total = sum(r['total'] for r in ytd)
    mtd_total = sum(r['total'] for r in mtd)
    progreso  = min(100, round(mtd_total / META_MENSUAL * 100, 1))

    # ventas por mes
    mes_data = defaultdict(float)
    for r in ytd:
        mes_data[r['dt'].month] += r['total']
    meses_labels = [MESES[m] for m in sorted(mes_data)]
    meses_vals   = [round(mes_data[m], 2) for m in sorted(mes_data)]

    # por vendedor MTD
    vend_mtd = defaultdict(float)
    vend_cnt = defaultdict(int)
    for r in mtd:
        vend_mtd[r['vendedor']] += r['total']
        vend_cnt[r['vendedor']] += 1
    vend_sorted  = sorted(vend_mtd.items(), key=lambda x: -x[1])
    vend_labels  = [v[0].split()[0] for v in vend_sorted]
    vend_vals    = [round(v[1], 2) for v in vend_sorted]
    vend_ordenes = [vend_cnt[v[0]] for v in vend_sorted]

    # top clientes MTD
    cli_data = defaultdict(float)
    for r in mtd:
        cli_data[r['cliente']] += r['total']
    top_cli = sorted(cli_data.items(), key=lambda x: -x[1])[:8]

    # últimas 10 órdenes
    ultimas = sorted(rows, key=lambda x: x['dt'], reverse=True)[:10]

    return {
        "generado"    : now.strftime('%d/%m/%Y %H:%M'),
        "ytd_total"   : round(ytd_total, 2),
        "ytd_ordenes" : len(ytd),
        "mtd_total"   : round(mtd_total, 2),
        "mtd_ordenes" : len(mtd),
        "meta_mensual": META_MENSUAL,
        "progreso_pct": progreso,
        "meses_labels": meses_labels,
        "meses_vals"  : meses_vals,
        "vend_labels" : vend_labels,
        "vend_vals"   : vend_vals,
        "vend_ordenes": vend_ordenes,
        "top_clientes": [[c, round(t,2)] for c, t in top_cli],
        "ultimas"     : [{"ref":r['ref'],"fecha":r['fecha'],
                          "cliente":r['cliente'][:35],"vendedor":r['vendedor'].split()[0],
                          "total":r['total']} for r in ultimas],
    }


def generar_html(data):
    dj = json.dumps(data, ensure_ascii=False)
    progreso_color = '#3fb950' if data['progreso_pct'] >= 80 else '#d29922' if data['progreso_pct'] >= 40 else '#f85149'

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Panel Ventas · Grupo Jecel</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{{color-scheme:dark;
  --bg:#0d1117;--surface:#161b22;--card:#161b22;--elevated:#21262d;
  --border:#30363d;--blue:#58a6ff;--green:#3fb950;--orange:#db6d28;
  --purple:#bc8cff;--red:#f85149;--amber:#d29922;
  --t900:#e6edf3;--t700:#c9d1d9;--t500:#8b949e;--t400:#6e7681;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--t700);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;font-size:14px;line-height:1.5;min-height:100vh}}
::-webkit-scrollbar{{width:6px;height:6px}}::-webkit-scrollbar-track{{background:var(--bg)}}
::-webkit-scrollbar-thumb{{background:var(--border);border-radius:6px}}
.hdr{{background:var(--surface);border-bottom:1px solid var(--border);padding:12px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 1px 3px rgba(0,0,0,.5)}}
.logo{{display:flex;align-items:center;gap:12px}}
.hex{{width:38px;height:38px;background:rgba(88,166,255,.15);clip-path:polygon(50% 0%,100% 25%,100% 75%,50% 100%,0% 75%,0% 25%);display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:900;color:var(--blue);border:1px solid rgba(88,166,255,.3)}}
.logo-text h1{{font-size:15px;font-weight:700;color:var(--t900)}}
.logo-text p{{font-size:10px;color:var(--t500);letter-spacing:.08em;text-transform:uppercase}}
.chip{{font-size:11px;color:var(--t500);background:var(--elevated);padding:4px 10px;border-radius:6px;border:1px solid var(--border)}}
.wrap{{padding:20px 24px;max-width:1400px;margin:0 auto}}
.section-title{{font-size:10px;text-transform:uppercase;letter-spacing:.14em;color:var(--t500);margin:24px 0 14px;display:flex;align-items:center;gap:8px;font-weight:700}}
.section-title::after{{content:"";flex:1;height:1px;background:var(--border)}}
.kgrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:4px}}
.kcard{{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:20px;position:relative;overflow:hidden;transition:border-color .15s,box-shadow .15s}}
.kcard:hover{{border-color:var(--blue);box-shadow:0 0 0 1px rgba(88,166,255,.15)}}
.kcard::before{{content:"";position:absolute;top:0;left:0;bottom:0;width:3px;border-radius:3px 0 0 3px}}
.c-b::before{{background:linear-gradient(180deg,#58a6ff,rgba(88,166,255,.2))}}
.c-g::before{{background:linear-gradient(180deg,#3fb950,rgba(63,185,80,.2))}}
.c-a::before{{background:linear-gradient(180deg,#db6d28,rgba(219,109,40,.2))}}
.c-p::before{{background:linear-gradient(180deg,#bc8cff,rgba(188,140,255,.2))}}
.kcard-icon{{position:absolute;top:14px;right:14px;font-size:20px;opacity:.7}}
.kcard-label{{font-size:11px;color:var(--t500);letter-spacing:.04em;text-transform:uppercase;margin-bottom:6px}}
.kcard-val{{font-size:26px;font-weight:700;color:var(--t900);letter-spacing:-.02em}}
.kcard-sub{{font-size:11px;color:var(--t500);margin-top:4px}}
.meta-bar-wrap{{margin-top:10px}}
.meta-bar-track{{height:6px;background:var(--elevated);border-radius:3px;overflow:hidden}}
.meta-bar-fill{{height:100%;border-radius:3px;transition:width .5s ease;background:{progreso_color}}}
.charts-grid{{display:grid;grid-template-columns:2fr 1fr;gap:16px}}
.chart-card{{background:var(--card);border:1px solid var(--border);border-radius:8px;padding:18px}}
.chart-card h3{{font-size:12px;font-weight:600;color:var(--t500);text-transform:uppercase;letter-spacing:.06em;margin-bottom:14px}}
.chart-wrap{{position:relative;height:220px}}
.grid-3{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}}
.cli-list{{display:flex;flex-direction:column;gap:8px}}
.cli-row{{display:flex;align-items:center;gap:10px;font-size:12px}}
.cli-num{{width:18px;color:var(--t400);font-size:11px;text-align:right;flex-shrink:0}}
.cli-name{{flex:1;color:var(--t700);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.cli-bar-wrap{{width:80px;flex-shrink:0}}
.cli-bar-track{{height:4px;background:var(--elevated);border-radius:2px;overflow:hidden}}
.cli-bar-fill{{height:100%;background:var(--blue);border-radius:2px}}
.cli-val{{width:70px;text-align:right;color:var(--t500);flex-shrink:0}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
th{{text-align:left;padding:8px 10px;color:var(--t500);font-size:10px;text-transform:uppercase;letter-spacing:.08em;border-bottom:1px solid var(--border);font-weight:600}}
td{{padding:9px 10px;border-bottom:1px solid rgba(48,54,61,.5);color:var(--t700)}}
tr:last-child td{{border-bottom:none}}
tr:hover td{{background:rgba(88,166,255,.04)}}
.ref{{color:var(--blue);font-weight:600}}
.total-td{{color:var(--t900);font-weight:600;text-align:right}}
.estado-badge{{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;background:rgba(63,185,80,.1);color:var(--green);border:1px solid rgba(63,185,80,.2)}}
.footer{{text-align:center;padding:20px;font-size:11px;color:var(--t400);border-top:1px solid var(--border);margin-top:24px}}
@media(max-width:768px){{.charts-grid{{grid-template-columns:1fr}}.grid-3{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<header class="hdr">
  <div class="logo">
    <div class="hex">JC</div>
    <div class="logo-text">
      <h1>Grupo Jecel C.A.</h1>
      <p>Panel de Ventas &mdash; Actualización diaria</p>
    </div>
  </div>
  <div class="chip">📅 Actualizado: <strong id="ts"></strong></div>
</header>

<div class="wrap">
  <div class="section-title">KPIs del Mes</div>
  <div class="kgrid">
    <div class="kcard c-a">
      <span class="kcard-icon">🎯</span>
      <div class="kcard-label">Meta Mensual</div>
      <div class="kcard-val" id="k-meta"></div>
      <div class="kcard-sub" id="k-progreso-txt"></div>
      <div class="meta-bar-wrap">
        <div class="meta-bar-track"><div class="meta-bar-fill" id="meta-bar"></div></div>
      </div>
    </div>
    <div class="kcard c-b">
      <span class="kcard-icon">📊</span>
      <div class="kcard-label">Ventas MTD</div>
      <div class="kcard-val" id="k-mtd"></div>
      <div class="kcard-sub" id="k-mtd-ord"></div>
    </div>
    <div class="kcard c-g">
      <span class="kcard-icon">📈</span>
      <div class="kcard-label">Ventas YTD {now.year}</div>
      <div class="kcard-val" id="k-ytd"></div>
      <div class="kcard-sub" id="k-ytd-ord"></div>
    </div>
    <div class="kcard c-p">
      <span class="kcard-icon">💰</span>
      <div class="kcard-label">Ticket Promedio MTD</div>
      <div class="kcard-val" id="k-ticket"></div>
      <div class="kcard-sub">por orden</div>
    </div>
  </div>

  <div class="section-title">Ventas por Mes (YTD)</div>
  <div class="charts-grid">
    <div class="chart-card">
      <h3>Evolución mensual $USD</h3>
      <div class="chart-wrap"><canvas id="chartMes"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>Por vendedor (mes actual)</h3>
      <div class="chart-wrap"><canvas id="chartVend"></canvas></div>
    </div>
  </div>

  <div class="section-title">Top Clientes (mes actual)</div>
  <div class="chart-card" style="margin-bottom:16px">
    <div class="cli-list" id="top-cli"></div>
  </div>

  <div class="section-title">Últimas Órdenes</div>
  <div class="chart-card">
    <table>
      <thead><tr><th>Referencia</th><th>Fecha</th><th>Cliente</th><th>Vendedor</th><th style="text-align:right">Total $USD</th><th>Estado</th></tr></thead>
      <tbody id="tbl-body"></tbody>
    </table>
  </div>
</div>

<footer class="footer">
  Grupo Jecel C.A. · Actualización automática desde Google Drive vía GitHub Actions · <span id="ts2"></span>
</footer>

<script>
const D = {dj};

// Timestamp
document.getElementById('ts').textContent = D.generado;
document.getElementById('ts2').textContent = D.generado;

// KPIs
const fmt = v => '$' + v.toLocaleString('es-VE', {{minimumFractionDigits:0, maximumFractionDigits:0}});
document.getElementById('k-meta').textContent = fmt(D.meta_mensual);
document.getElementById('k-progreso-txt').textContent = D.progreso_pct + '% alcanzado';
document.getElementById('meta-bar').style.width = D.progreso_pct + '%';
document.getElementById('k-mtd').textContent = fmt(D.mtd_total);
document.getElementById('k-mtd-ord').textContent = D.mtd_ordenes + ' órdenes';
document.getElementById('k-ytd').textContent = fmt(D.ytd_total);
document.getElementById('k-ytd-ord').textContent = D.ytd_ordenes + ' órdenes totales';
const ticket = D.mtd_ordenes > 0 ? D.mtd_total / D.mtd_ordenes : 0;
document.getElementById('k-ticket').textContent = fmt(ticket);

// Chart mensual
new Chart(document.getElementById('chartMes'), {{
  type: 'bar',
  data: {{
    labels: D.meses_labels,
    datasets: [{{
      label: 'Ventas $USD',
      data: D.meses_vals,
      backgroundColor: D.meses_labels.map((_,i) => i === D.meses_labels.length-1 ? 'rgba(88,166,255,.9)' : 'rgba(88,166,255,.4)'),
      borderRadius: 4,
      borderSkipped: false
    }}]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{display:false}},
      tooltip: {{callbacks: {{label: ctx => ' $' + ctx.raw.toLocaleString('es-VE', {{minimumFractionDigits:0}})}}}}
    }},
    scales: {{
      x: {{grid:{{color:'rgba(48,54,61,.5)'}}, ticks:{{color:'#8b949e'}}}},
      y: {{grid:{{color:'rgba(48,54,61,.5)'}}, ticks:{{color:'#8b949e', callback: v => '$' + (v/1000).toFixed(0)+'k'}}}}
    }}
  }}
}});

// Chart vendedores
const COLORS = ['rgba(88,166,255,.85)','rgba(63,185,80,.85)','rgba(188,140,255,.85)','rgba(219,109,40,.85)','rgba(248,81,73,.85)'];
new Chart(document.getElementById('chartVend'), {{
  type: 'doughnut',
  data: {{
    labels: D.vend_labels,
    datasets: [{{
      data: D.vend_vals,
      backgroundColor: COLORS,
      borderColor: '#0d1117',
      borderWidth: 2,
      hoverOffset: 6
    }}]
  }},
  options: {{
    responsive: true, maintainAspectRatio: false,
    plugins: {{
      legend: {{position:'bottom', labels:{{color:'#8b949e',padding:12,font:{{size:11}}}}}},
      tooltip: {{callbacks: {{label: ctx => ' $' + ctx.raw.toLocaleString('es-VE', {{minimumFractionDigits:0}})}}}}
    }}
  }}
}});

// Top clientes
const maxCli = D.top_clientes.length > 0 ? D.top_clientes[0][1] : 1;
const cliHtml = D.top_clientes.map((c,i) =>
  `<div class="cli-row">
    <span class="cli-num">${{i+1}}</span>
    <span class="cli-name">${{c[0]}}</span>
    <div class="cli-bar-wrap"><div class="cli-bar-track"><div class="cli-bar-fill" style="width:${{Math.round(c[1]/maxCli*100)}}%"></div></div></div>
    <span class="cli-val">$${{c[1].toLocaleString('es-VE',{{minimumFractionDigits:0}})}}</span>
  </div>`
).join('');
document.getElementById('top-cli').innerHTML = cliHtml;

// Tabla últimas órdenes
const rows = D.ultimas.map(r =>
  `<tr>
    <td class="ref">${{r.ref}}</td>
    <td>${{r.fecha}}</td>
    <td>${{r.cliente}}</td>
    <td>${{r.vendedor}}</td>
    <td class="total-td">$${{r.total.toLocaleString('es-VE',{{minimumFractionDigits:2}})}}</td>
    <td><span class="estado-badge">Confirmada</span></td>
  </tr>`
).join('');
document.getElementById('tbl-body').innerHTML = rows;
</script>
</body>
</html>
"""


if __name__ == '__main__':
    print("📥 Cargando datos del Excel...")
    rows = cargar_datos('ventas.xlsx')
    print(f"✅ {len(rows)} órdenes cargadas")
    print("📊 Calculando KPIs...")
    data = calcular_kpis(rows)
    print(f"   MTD: ${data['mtd_total']:,.2f} ({data['mtd_ordenes']} órdenes)")
    print(f"   YTD: ${data['ytd_total']:,.2f} ({data['ytd_ordenes']} órdenes)")
    print("🔨 Generando index.html...")
    html = generar_html(data)
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print("✅ index.html generado correctamente")
