# -*- coding: utf-8 -*-
"""Dashboard "Capacidade de Empenho" — Atas de Registro de Preços (BCMS).

Lê o consolidado do Robô Extrator de Pregões (Consolidado_Pregoes.xlsx, nome
fixo no Google Drive → link estável), calcula a capacidade de empenho
(Qtd. Saldo × Valor Unitário das atas vigentes, nas linhas da UASG alvo) e gera
um site estático autocontido em site/index.html para o GitHub Pages.

Uso:
  py -3 gerar_dashboard.py                       # baixa do Drive (DRIVE_FILE_ID)
  SOURCE_XLSX=caminho.xlsx py -3 gerar_dashboard.py   # usa arquivo local (teste)

Ambiente:
  DRIVE_FILE_ID  ID do arquivo no Drive (ou usa o padrão embutido)
  SOURCE_XLSX    caminho local (tem precedência; para teste)
  UASG_ALVO      prefixo da UASG participante (padrão: 160329 = BCMS)
"""

from __future__ import annotations

import html
import io
import json
import os
import urllib.request
from datetime import datetime, timedelta

from openpyxl import load_workbook

# ---------------------------------------------------------------- configuração
BASE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(BASE, "site")
DATA = os.path.join(BASE, "data")

# ID do Consolidado_Pregoes.xlsx no Google Drive (Compartilhar → qualquer
# pessoa com o link). Pode ser sobrescrito pelo secret/variável DRIVE_FILE_ID.
DRIVE_FILE_ID_PADRAO = "1-YUgpCqh8N-fZO_rKzyAPQjjxnLFBJ1A"

UASG_ALVO = os.environ.get("UASG_ALVO", "160329").strip()
NOME_UNIDADE = "Batalhão Central de Manutenção e Suprimento"

AUSENTE = "Informação ausente"


# ---------------------------------------------------------------- utilidades
def num(v):
    """Converte célula em float (aceita '1.234,56'); None se não numérico."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s == AUSENTE:
        return None
    try:
        return float(s.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def data_br(v):
    """'dd/mm/aaaa' (ou datetime) → datetime; None se inválida."""
    if isinstance(v, datetime):
        return v
    try:
        return datetime.strptime(str(v).strip(), "%d/%m/%Y")
    except (ValueError, TypeError):
        return None


def fmt_brl(v) -> str:
    if v is None:
        return "—"
    s = f"{v:,.2f}".replace(",", "§").replace(".", ",").replace("§", ".")
    return f"R$ {s}"


def fmt_short(v) -> str:
    """R$ compacto para rótulos de barra (mi/mil)."""
    if v is None:
        return "—"
    if abs(v) >= 1_000_000:
        return f"R$ {v/1_000_000:.2f} mi".replace(".", ",")
    if abs(v) >= 1_000:
        return f"R$ {v/1_000:.0f} mil"
    return fmt_brl(v)


def fmt_int(v) -> str:
    return f"{int(v):,}".replace(",", ".")


def esc(s, limite=None) -> str:
    t = str(s or "").strip()
    if limite and len(t) > limite:
        t = t[: limite - 1] + "…"
    return html.escape(t, quote=True)


# ---------------------------------------------------------------- fonte de dados
def baixar_do_drive(file_id: str) -> io.BytesIO:
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    print(f"[INFO] Baixando consolidado do Drive ({file_id[:8]}…)")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    dados = urllib.request.urlopen(req, timeout=120).read()
    if not dados.startswith(b"PK"):
        # Arquivos grandes ganham página interstitial; tenta o confirm direto.
        req2 = urllib.request.Request(url + "&confirm=t", headers={"User-Agent": "Mozilla/5.0"})
        dados = urllib.request.urlopen(req2, timeout=120).read()
    if not dados.startswith(b"PK"):
        raise RuntimeError(
            "O download não devolveu um .xlsx — confira se o arquivo está "
            "compartilhado como 'Qualquer pessoa com o link' e se o ID está certo.")
    print(f"[INFO] {len(dados)/1024:.0f} KB baixados.")
    return io.BytesIO(dados)


def carregar_linhas(fobj) -> list[dict]:
    wb = load_workbook(fobj, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    linhas, cab = [], None
    for row in ws.iter_rows(values_only=True):
        if cab is None:
            cab = [str(c or "").strip() for c in row]
            continue
        linhas.append(dict(zip(cab, row)))
    wb.close()
    print(f"[INFO] {len(linhas)} linhas lidas do consolidado.")
    return linhas


# ---------------------------------------------------------------- ETL
def classificar_status(r: dict, hoje: datetime) -> str:
    """Bucket do status: vig | v30 | venc | semata | semdados."""
    s = str(r.get("Status_Ata") or "").strip()
    if s.startswith("Sem ata"):
        return "semata"
    if s == "Sem dados":
        return "semdados"
    if s == "Vencida":
        return "venc"
    if s.startswith("Vence em"):
        return "v30"
    if s.startswith("Vigente"):
        return "vig"
    # Sem coluna Status_Ata (planilha antiga): deriva das datas.
    if str(r.get("Nr Ata") or "").strip() in ("", AUSENTE):
        forn = str(r.get("Fornecedor") or "").strip()
        return "semdados" if forn in ("", AUSENTE) else "semata"
    fim = data_br(r.get("Fim Vig Ata"))
    if fim is None:
        return "vig"
    if fim < hoje:
        return "venc"
    if fim < hoje + timedelta(days=30):
        return "v30"
    return "vig"


def etl(linhas: list[dict]) -> dict:
    hoje = datetime.now()
    itens = []          # linhas da UASG alvo com saldo calculável
    datas_coleta = []
    for r in linhas:
        if not str(r.get("UASG") or "").strip().startswith(UASG_ALVO):
            continue
        d = data_br(r.get("Data_Coleta"))
        if d:
            datas_coleta.append(d)
        saldo, vu = num(r.get("Qtd. Saldo")), num(r.get("Val. Unitário"))
        cap = round(saldo * vu, 2) if (saldo is not None and vu is not None) else None
        itens.append({
            "pregao": str(r.get("Pregão") or "").replace("_", "/"),
            "ger": str(r.get("UASG_Gerenciadora") or ""),
            "nr": r.get("Nr Item"),
            "desc": str(r.get("Descrição detalhada") or ""),
            "forn": str(r.get("Fornecedor") or ""),
            "cat": str(r.get("Categoria_Geral") or "").strip() or "Sem categoria",
            "fim": data_br(r.get("Fim Vig Ata")),
            "st": classificar_status(r, hoje),
            "saldo": saldo, "vu": vu, "cap": cap,
            "link": str(r.get("Link_Item") or ""),
        })

    ativos = [i for i in itens if i["st"] in ("vig", "v30") and (i["cap"] or 0) > 0]
    cap_vig = sum(i["cap"] for i in ativos if i["st"] == "vig")
    cap_v30 = sum(i["cap"] for i in ativos if i["st"] == "v30")
    cap_venc = sum(i["cap"] or 0 for i in itens if i["st"] == "venc" and (i["cap"] or 0) > 0)

    por_cat, por_pregao = {}, {}
    for i in ativos:
        por_cat[i["cat"]] = por_cat.get(i["cat"], 0) + i["cap"]
        chave = f"{i['ger']} · {i['pregao']}"
        por_pregao[chave] = por_pregao.get(chave, 0) + i["cap"]

    # Vencimentos ≤60 dias, agrupados por pregão+data.
    venc60 = {}
    for i in ativos:
        if i["fim"] and i["fim"] < hoje + timedelta(days=60):
            k = (i["fim"], i["pregao"], i["ger"])
            v = venc60.setdefault(k, {"cap": 0, "n": 0})
            v["cap"] += i["cap"]
            v["n"] += 1

    return {
        "hoje": hoje,
        "posicao": max(datas_coleta).strftime("%d/%m/%Y") if datas_coleta else hoje.strftime("%d/%m/%Y"),
        "itens": itens, "ativos": ativos,
        "cap_total": cap_vig + cap_v30, "cap_vig": cap_vig, "cap_v30": cap_v30,
        "cap_venc": cap_venc,
        "n_itens": len(ativos),
        "n_pregoes": len({(i["ger"], i["pregao"]) for i in ativos}),
        "por_cat": sorted(por_cat.items(), key=lambda x: -x[1])[:12],
        "por_pregao": sorted(por_pregao.items(), key=lambda x: -x[1])[:12],
        "venc60": sorted(venc60.items())[:12],
    }


# ---------------------------------------------------------------- histórico
def atualizar_historico(m: dict) -> list[dict]:
    os.makedirs(DATA, exist_ok=True)
    caminho = os.path.join(DATA, "history.json")
    hist = []
    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                hist = json.load(f)
        except (json.JSONDecodeError, OSError):
            hist = []
    chave = m["posicao"]  # 1 snapshot por posição de coleta
    hist = [h for h in hist if h.get("posicao") != chave]
    hist.append({"posicao": chave, "cap_total": round(m["cap_total"], 2),
                 "cap_v30": round(m["cap_v30"], 2), "itens": m["n_itens"]})
    hist = hist[-120:]
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=1)
    return hist


def sparkline(hist: list[dict]) -> str:
    if len(hist) < 2:
        return '<p class="muted">O histórico acumula a partir da 2ª atualização.</p>'
    vals = [h["cap_total"] for h in hist]
    vmin, vmax = min(vals), max(vals)
    faixa = (vmax - vmin) or 1
    W, H = 560, 90
    pts = []
    for i, v in enumerate(vals):
        x = 8 + i * (W - 16) / (len(vals) - 1)
        y = H - 12 - (v - vmin) / faixa * (H - 28)
        pts.append(f"{x:.1f},{y:.1f}")
    return (f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="Evolução da capacidade" '
            f'preserveAspectRatio="none" style="width:100%;height:{H}px">'
            f'<polyline points="{" ".join(pts)}" fill="none" stroke="var(--primary)" '
            f'stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/></svg>'
            f'<div class="spark-leg"><span>{hist[0]["posicao"]} · {fmt_short(vals[0])}</span>'
            f'<span>{hist[-1]["posicao"]} · {fmt_short(vals[-1])}</span></div>')


# ---------------------------------------------------------------- render
_ST_LABEL = {"vig": "Vigente", "v30": "≤30 dias", "venc": "Vencida",
             "semata": "Sem ata", "semdados": "Sem dados"}


def _barras(pares: list, cor="var(--primary)") -> str:
    if not pares:
        return '<p class="muted">Sem dados.</p>'
    topo = pares[0][1] or 1
    out = []
    for rotulo, valor in pares:
        pct = max(valor / topo * 100, 1.5)
        out.append(
            f'<div class="brow"><div class="blabel" title="{esc(rotulo)}">{esc(rotulo, 34)}</div>'
            f'<div class="btrack"><div class="bfill" style="width:{pct:.1f}%;background:{cor}"></div></div>'
            f'<div class="bval">{fmt_short(valor)}</div></div>')
    return "".join(out)


def _tabela(itens: list[dict]) -> str:
    linhas = []
    visiveis = [i for i in itens if (i["saldo"] or 0) > 0 and i["st"] != "semdados"]
    visiveis.sort(key=lambda i: -(i["cap"] or 0))
    for i in visiveis:
        fim = i["fim"].strftime("%d/%m/%Y") if i["fim"] else "—"
        link = (f'<a href="{esc(i["link"])}" target="_blank" rel="noopener" '
                f'aria-label="Abrir item no Compras.gov">↗</a>') if i["link"].startswith("http") else ""
        linhas.append(
            f'<tr data-st="{i["st"]}">'
            f'<td>{esc(i["pregao"])}</td><td>{esc(i["ger"])}</td>'
            f'<td class="num">{esc(i["nr"])}</td>'
            f'<td class="desc" title="{esc(i["desc"], 300)}">{esc(i["desc"], 110)}</td>'
            f'<td class="forn" title="{esc(i["forn"], 160)}">{esc(i["forn"], 42)}</td>'
            f'<td data-v="{i["fim"].timestamp() if i["fim"] else 0:.0f}">{fim}</td>'
            f'<td><span class="tag tag-{i["st"]}">{_ST_LABEL[i["st"]]}</span></td>'
            f'<td class="num" data-v="{i["saldo"] or 0}">{fmt_int(i["saldo"] or 0)}</td>'
            f'<td class="num" data-v="{i["vu"] or 0}">{fmt_brl(i["vu"])}</td>'
            f'<td class="num" data-v="{i["cap"] or 0}"><strong>{fmt_brl(i["cap"])}</strong></td>'
            f'<td>{link}</td></tr>')
    return "\n".join(linhas)


def render(m: dict, hist: list[dict]) -> str:
    tpl = _TEMPLATE
    subs = {
        "%%POSICAO%%": m["posicao"],
        "%%GERADO%%": m["hoje"].strftime("%d/%m/%Y %H:%M"),
        "%%UASG%%": UASG_ALVO,
        "%%UNIDADE%%": NOME_UNIDADE,
        "%%CAP_TOTAL%%": fmt_brl(m["cap_total"]),
        "%%CAP_VIG%%": fmt_brl(m["cap_vig"]),
        "%%CAP_V30%%": fmt_brl(m["cap_v30"]),
        "%%CAP_VENC%%": fmt_brl(m["cap_venc"]),
        "%%N_ITENS%%": fmt_int(m["n_itens"]),
        "%%N_PREGOES%%": fmt_int(m["n_pregoes"]),
        "%%PCT_V30%%": f"{(m['cap_v30'] / m['cap_total'] * 100 if m['cap_total'] else 0):.1f}%".replace(".", ","),
        "%%BARRAS_CAT%%": _barras(m["por_cat"]),
        "%%BARRAS_PREGAO%%": _barras(m["por_pregao"], cor="var(--gold)"),
        "%%SPARK%%": sparkline(hist),
        "%%TABELA%%": _tabela(m["itens"]),
        "%%VENC60%%": "".join(
            f'<tr><td>{fim.strftime("%d/%m/%Y")}</td><td>{esc(pregao)}</td>'
            f'<td>{esc(ger)}</td><td class="num">{v["n"]}</td>'
            f'<td class="num"><strong>{fmt_brl(v["cap"])}</strong></td></tr>'
            for (fim, pregao, ger), v in m["venc60"]) or
            '<tr><td colspan="5" class="muted">Nenhuma ata vencendo nos próximos 60 dias.</td></tr>',
    }
    for k, v in subs.items():
        tpl = tpl.replace(k, v)
    return tpl


_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Capacidade de Empenho — Atas BCMS</title>
<style>
:root{--bg:#EEF1F6;--surface:#FFF;--ink:#0F1B2A;--ink-muted:#566374;--border:#D8DEE7;
--primary:#1C4A73;--success:#0F7A5A;--warning:#B5822B;--danger:#B23A2E;--gold:#C8901E;
--track:#E4E8EF;--hero-soft:#E7F1EC;--shadow:0 1px 2px rgba(15,27,42,.06);
--serif:Georgia,"Times New Roman",serif;--sans:-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
[data-theme="dark"]{--bg:#0C1420;--surface:#131E2E;--ink:#E7EDF4;--ink-muted:#97A6B8;
--border:#26374C;--primary:#4E86BD;--success:#35C08F;--warning:#D9A94A;--danger:#E06A5E;
--gold:#D9A94A;--track:#1D2B3F;--hero-soft:#14263A;--shadow:0 1px 2px rgba(0,0,0,.4)}
*{box-sizing:border-box;margin:0}
body{background:var(--bg);color:var(--ink);font:14px/1.5 var(--sans)}
.bcms-bar{height:5px;background:linear-gradient(to bottom,#CE2B2B 50%,#1E6FD0 50%)}
.wrap{max-width:1100px;margin:0 auto;padding:18px 16px 60px}
.topbar{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap}
h1{font-family:var(--serif);font-size:26px;font-weight:600;line-height:1.15}
.subtitle{font-size:12.5px;color:var(--ink-muted);margin-top:2px}
.muted{color:var(--ink-muted)}
button.theme{background:var(--surface);border:1px solid var(--border);border-radius:8px;
padding:6px 12px;cursor:pointer;color:var(--ink);font:inherit}
.hero{background:var(--hero-soft);border:1px solid var(--border);border-radius:14px;
padding:22px 24px;margin:18px 0;box-shadow:var(--shadow)}
.hero .label{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--ink-muted)}
.hero .big{font-family:var(--serif);font-size:44px;font-weight:700;color:var(--success);margin:2px 0 8px}
.eq{font-size:13.5px}.eq b{font-weight:600}
.eq .vig{color:var(--success)}.eq .v30{color:var(--warning)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin:16px 0}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:14px 16px;box-shadow:var(--shadow)}
.kpi .v{font-family:var(--serif);font-size:24px;font-weight:700;margin-top:2px}
.kpi .v.warn{color:var(--warning)}.kpi .v.bad{color:var(--danger)}
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;
padding:18px 20px;margin:16px 0;box-shadow:var(--shadow)}
.card h2{font-family:var(--serif);font-size:18px;font-weight:600;margin-bottom:12px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:820px){.grid2{grid-template-columns:1fr}}
.brow{display:grid;grid-template-columns:180px 1fr 86px;gap:10px;align-items:center;margin:7px 0}
.blabel{font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.btrack{background:var(--track);border-radius:6px;height:14px;overflow:hidden}
.bfill{height:100%;border-radius:6px}
.bval{font-size:12px;text-align:right;font-variant-numeric:tabular-nums}
.spark-leg{display:flex;justify-content:space-between;font-size:11.5px;color:var(--ink-muted)}
.controls{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px}
.controls input{flex:1;min-width:220px;background:var(--bg);color:var(--ink);
border:1px solid var(--border);border-radius:8px;padding:8px 12px;font:inherit}
.chip{background:var(--surface);border:1px solid var(--border);border-radius:999px;
padding:5px 14px;cursor:pointer;font:12.5px var(--sans);color:var(--ink)}
.chip.on{background:var(--primary);border-color:var(--primary);color:#fff}
.tblwrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th,td{padding:7px 8px;border-bottom:1px solid var(--border);text-align:left;vertical-align:top}
th{cursor:pointer;user-select:none;white-space:nowrap;color:var(--ink-muted);font-weight:600}
th:hover{color:var(--ink)}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
td.desc{max-width:330px}td.forn{max-width:190px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tag{border-radius:999px;padding:2px 9px;font-size:11px;white-space:nowrap}
.tag-vig{background:color-mix(in srgb,var(--success) 15%,transparent);color:var(--success)}
.tag-v30{background:color-mix(in srgb,var(--warning) 18%,transparent);color:var(--warning)}
.tag-venc{background:color-mix(in srgb,var(--danger) 15%,transparent);color:var(--danger)}
.tag-semata{background:var(--track);color:var(--ink-muted)}
a{color:var(--primary)}
footer{margin-top:26px;font-size:12px;color:var(--ink-muted);text-align:center}
</style>
</head>
<body>
<div class="bcms-bar" aria-hidden="true"></div>
<div class="wrap">
  <div class="topbar">
    <div>
      <h1>Capacidade de Empenho — Atas de Registro de Preços</h1>
      <div class="subtitle">%%UNIDADE%% · UASG %%UASG%% · Posição: <b>%%POSICAO%%</b></div>
    </div>
    <div><button class="theme" id="btTheme" aria-label="Alternar tema">◐ Tema</button></div>
  </div>

  <div class="hero">
    <div class="label">Capacidade de empenho em atas válidas</div>
    <div class="big">%%CAP_TOTAL%%</div>
    <div class="eq"><b class="vig">%%CAP_VIG%%</b> em atas vigentes
      + <b class="v30">%%CAP_V30%%</b> vencendo em ≤30 dias (%%PCT_V30%% do total)</div>
  </div>

  <div class="kpis">
    <div class="kpi"><div class="muted">Itens com saldo</div><div class="v">%%N_ITENS%%</div></div>
    <div class="kpi"><div class="muted">Pregões com saldo</div><div class="v">%%N_PREGOES%%</div></div>
    <div class="kpi"><div class="muted">Vence em ≤30 dias</div><div class="v warn">%%CAP_V30%%</div></div>
    <div class="kpi"><div class="muted">Perdido em atas vencidas</div><div class="v bad">%%CAP_VENC%%</div></div>
  </div>

  <div class="grid2">
    <div class="card"><h2>Por categoria de despesa</h2>%%BARRAS_CAT%%</div>
    <div class="card"><h2>Por pregão (UASG ger. · nº)</h2>%%BARRAS_PREGAO%%</div>
  </div>

  <div class="grid2">
    <div class="card">
      <h2>⚠ Atas vencendo em até 60 dias</h2>
      <div class="tblwrap"><table>
        <thead><tr><th>Fim vigência</th><th>Pregão</th><th>Ger.</th>
        <th class="num">Itens</th><th class="num">Capacidade</th></tr></thead>
        <tbody>%%VENC60%%</tbody>
      </table></div>
    </div>
    <div class="card"><h2>Evolução da capacidade</h2>%%SPARK%%</div>
  </div>

  <div class="card">
    <h2>Itens com saldo (clique nos títulos para ordenar)</h2>
    <div class="controls">
      <input id="busca" type="search" placeholder="Buscar item, fornecedor, pregão…" aria-label="Buscar">
      <button class="chip on" data-f="ativos">Válidas</button>
      <button class="chip" data-f="v30">≤30 dias</button>
      <button class="chip" data-f="venc">Vencidas</button>
      <button class="chip" data-f="all">Todas</button>
      <span class="muted" id="cont"></span>
    </div>
    <div class="tblwrap">
      <table id="tb">
        <thead><tr>
          <th>Pregão</th><th>Ger.</th><th class="num">Item</th><th>Descrição</th>
          <th>Fornecedor</th><th>Fim vig.</th><th>Status</th>
          <th class="num">Saldo</th><th class="num">Vlr. unit.</th>
          <th class="num">Capacidade</th><th></th>
        </tr></thead>
        <tbody>
%%TABELA%%
        </tbody>
      </table>
    </div>
  </div>

  <footer>%%UNIDADE%% — dados extraídos do Compras.gov.br pelo Robô Extrator de Pregões.<br>
  Gerado em %%GERADO%% · Capacidade = Qtd. Saldo × Valor Unitário das atas registradas.</footer>
</div>
<script>
(function(){
  var root=document.documentElement, KEY='bcms-pregoes-theme';
  try{var t=localStorage.getItem(KEY); if(t) root.dataset.theme=t;
      else if(matchMedia('(prefers-color-scheme: dark)').matches) root.dataset.theme='dark';}catch(e){}
  document.getElementById('btTheme').onclick=function(){
    root.dataset.theme = root.dataset.theme==='dark'?'light':'dark';
    try{localStorage.setItem(KEY, root.dataset.theme);}catch(e){}
  };
  var tb=document.getElementById('tb'), corpo=tb.tBodies[0],
      linhas=[].slice.call(corpo.rows), filtro='ativos', termo='';
  function aplica(){
    var vis=0;
    linhas.forEach(function(tr){
      var st=tr.dataset.st, okF =
        filtro==='all' || (filtro==='ativos' && (st==='vig'||st==='v30')) ||
        (filtro==='v30' && st==='v30') || (filtro==='venc' && st==='venc');
      var okT=!termo || tr.textContent.toLowerCase().indexOf(termo)>-1;
      tr.style.display=(okF&&okT)?'':'none'; if(okF&&okT)vis++;
    });
    document.getElementById('cont').textContent=vis+' itens';
  }
  document.getElementById('busca').addEventListener('input',function(){
    termo=this.value.trim().toLowerCase(); aplica();});
  [].forEach.call(document.querySelectorAll('.chip'),function(c){
    c.onclick=function(){
      [].forEach.call(document.querySelectorAll('.chip'),function(x){x.classList.remove('on')});
      c.classList.add('on'); filtro=c.dataset.f; aplica();};});
  var ord={col:-1,asc:false};
  [].forEach.call(tb.tHead.rows[0].cells,function(th,i){
    th.onclick=function(){
      ord.asc = ord.col===i ? !ord.asc : false; ord.col=i;
      var num = th.classList.contains('num') || i===5;
      linhas.sort(function(a,b){
        var ca=a.cells[i], cb=b.cells[i], va, vb;
        if(num){va=parseFloat(ca.dataset.v||ca.textContent.replace(/[^\d,-]/g,'').replace(',','.'))||0;
                vb=parseFloat(cb.dataset.v||cb.textContent.replace(/[^\d,-]/g,'').replace(',','.'))||0;}
        else {va=ca.textContent.toLowerCase(); vb=cb.textContent.toLowerCase();}
        return (va<vb?-1:va>vb?1:0)*(ord.asc?1:-1);
      });
      linhas.forEach(function(tr){corpo.appendChild(tr)});
    };});
  aplica();
})();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------- main
def main() -> int:
    local = os.environ.get("SOURCE_XLSX", "").strip()
    if local:
        print(f"[INFO] Usando arquivo local: {local}")
        with open(local, "rb") as f:
            fobj = io.BytesIO(f.read())
    else:
        fid = os.environ.get("DRIVE_FILE_ID", "").strip() or DRIVE_FILE_ID_PADRAO
        if "COLOQUE_AQUI" in fid:
            print("[ERRO] Defina DRIVE_FILE_ID (secret/variável) ou edite "
                  "DRIVE_FILE_ID_PADRAO no script.")
            return 2
        fobj = baixar_do_drive(fid)

    m = etl(carregar_linhas(fobj))
    if not m["itens"]:
        print(f"[ERRO] Nenhuma linha da UASG {UASG_ALVO} no consolidado — abortando "
              "para não publicar um painel vazio.")
        return 3

    hist = atualizar_historico(m)
    os.makedirs(SITE, exist_ok=True)
    saida = os.path.join(SITE, "index.html")
    with open(saida, "w", encoding="utf-8") as f:
        f.write(render(m, hist))

    print(f"[OK] {saida} ({os.path.getsize(saida)/1024:.0f} KB)")
    print(f"     Capacidade total: {fmt_brl(m['cap_total'])} "
          f"({m['n_itens']} itens, {m['n_pregoes']} pregões)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
