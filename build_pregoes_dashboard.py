# -*- coding: utf-8 -*-
"""
=============================================================================
BUILDER & AUDITOR: DASHBOARD DE PREGÕES & LICITAÇÕES (LEI 14.133/2021)
Base de Apoio Logístico do Exército · PNCP / Compras.gov.br
Autor: Senior GovTech Frontend Architect & Auditor de Compras Públicas
Desenvolvido por: 3º Sgt De Campos
Arquitetura: WCAG 2.2 AAA + 60fps GPU + IBCS & Tufte + Mobile-First (360px-1860px)
=============================================================================
"""

import json
import os
import re
import sys
import pathlib
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

ROOT_DIR = pathlib.Path(__file__).parent
SITE_DIR = ROOT_DIR / "site"
DATA_DIR = ROOT_DIR.parent / "Requisicoes-Empenho-BCMS" / "dados"

LOCAL_DATA = ROOT_DIR / "data"
PNCP_PATH = LOCAL_DATA / "pncp_cache.json" if (LOCAL_DATA / "pncp_cache.json").exists() else (DATA_DIR / "pncp_cache.json")
PREGOES_PATH = LOCAL_DATA / "pregoes.json" if (LOCAL_DATA / "pregoes.json").exists() else (DATA_DIR / "pregoes.json")
TOKENS_PATH = SITE_DIR / "design-tokens.css"
INDEX_PATH = SITE_DIR / "index.html"
PREGOES_HTML_PATH = SITE_DIR / "pregoes.html"

def safe_float(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    try:
        s = str(v).replace("R$", "").replace(".", "").replace(",", ".").strip()
        return float(s)
    except Exception:
        return 0.0

def safe_int(v):
    if v is None: return 0
    if isinstance(v, int): return v
    try:
        return int(float(str(v).replace(".", "").replace(",", ".").strip()))
    except Exception:
        return 0

def fmt_brl(val):
    if val is None or val == 0.0:
        return "R$ 0,00"
    if abs(val) < 0.005:
        return "R$ 0,00"
    s = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

def carregar_dados_consolidados():
    print("1. Carregando e cruzando bases do PNCP e Pregões...")
    pncp_list = []
    for p_cand in (DATA_DIR / "pncp_cache.json", LOCAL_DATA / "pncp_cache.json"):
        if p_cand.exists():
            try:
                lst = json.loads(p_cand.read_text(encoding="utf-8")).get("pregoes", [])
                if lst:
                    pncp_list = lst
                    break
            except Exception:
                pass
    pregoes_dict = json.loads(PREGOES_PATH.read_text(encoding="utf-8")).get("pregoes", {}) if PREGOES_PATH.exists() else {}

    uasg_info = {
        "160329": {"sigla": "BCMS", "nome": "Batalhão Central de Manutenção e Suprimento", "logo": "BCMS.png", "accent": "#DB2819"},
        "160238": {"sigla": "Ba Ap Log", "nome": "Base de Apoio Logístico do Exército", "logo": "BaApLog.png", "accent": "#D83030"},
        "160307": {"sigla": "1º D Sup", "nome": "1º Depósito de Suprimento", "logo": "1DSUP.png", "accent": "#DE2B30"},
        "160304": {"sigla": "BMSA", "nome": "Batalhão de Manutenção e Suprimento de Armamento", "logo": "BMSA.png", "accent": "#DB2819"},
        "160246": {"sigla": "D C Mun", "nome": "Depósito Central de Munição", "logo": "DCMUN.png", "accent": "#047CC0"},
        "160321": {"sigla": "ECT", "nome": "Estabelecimento Central de Transportes", "logo": "Ect.png", "accent": "#B33338"},
        "153149": {"sigla": "1ª RM", "nome": "Comando da 1ª Região Militar", "logo": "BCMS.png", "accent": "#1C4A73"},
        "158276": {"sigla": "H Mil", "nome": "Hospital Militar de Área", "logo": "BCMS.png", "accent": "#1C4A73"},
    }

    consolidated = []
    vistos_keys = set()
    tot_est = 0.0
    tot_hom = 0.0
    desagios_list = []
    hoje_count = 0

    for idx, p in enumerate(pncp_list):
        ugg = str(p.get("ugg", "")).strip()
        om = uasg_info.get(ugg, {"sigla": p.get("om_nome", f"UASG {ugg}"), "nome": p.get("om_nome", f"UASG {ugg}"), "logo": "BCMS.png", "accent": "#1C4A73"})
        num = str(p.get("numero", "")).strip()
        ano = str(p.get("ano", "2026")).strip()
        full_num = f"{num}/{ano}"
        vistos_keys.add(f"{ugg}:{full_num}")
        if num.isdigit():
            vistos_keys.add(f"{ugg}:{int(num)}/{ano}")

        ve = safe_float(p.get("valor_estimado"))
        vh = safe_float(p.get("valor_homologado"))
        tot_est += ve
        tot_hom += vh

        if ve > 0 and vh > 0 and ve >= vh:
            desagio_pct = round(((ve - vh) / ve) * 100, 2)
            desagio_val = round(ve - vh, 2)
            desagios_list.append(desagio_pct)
        else:
            desagio_pct = None
            desagio_val = 0.0

        sit = p.get("situacao", "").strip()
        if "Homologado" in sit or "Adjudicado" in sit or (vh > 0 and ve > 0):
            fase = "Homologado"
        elif "Julgamento" in sit:
            fase = "Julgamento / Habilitação"
        elif "Disputa" in sit or "Aberto" in sit:
            fase = "Aberto / Em Disputa"
        elif "Deserto" in sit:
            fase = "Deserto"
        elif "Fracassado" in sit:
            fase = "Fracassado"
        elif "Revogado" in sit or "Anulado" in sit:
            fase = "Revogado / Anulado"
        else:
            fase = "Em Acolhimento de Propostas"

        data_abertura = p.get("data_abertura") or "INDISPONIVEL - Não informada"
        abertura_urgente = False
        if "2026-10" in data_abertura or "2026-09" in data_abertura or "2026-08" in data_abertura:
            hoje_count += 1
            abertura_urgente = True

        itens = []
        for k, v in pregoes_dict.items():
            if full_num in k or (num in k and ugg in k):
                itens = v.get("itens", [])
                break

        formatted_itens = []
        for it in itens:
            desc = it.get("descricao") or it.get("catmat_nome") or "Item de fornecimento"
            qtd = safe_int(it.get("qtd_homologada") or it.get("quantidade") or 1)
            vu = safe_float(it.get("valor_unitario"))
            vt = round(qtd * vu, 2)
            formatted_itens.append({
                "item": it.get("item", len(formatted_itens)+1),
                "descricao": desc,
                "quantidade": qtd,
                "valor_unitario": vu,
                "valor_total": vt,
                "fornecedor": it.get("fornecedor_nome") or "INDISPONIVEL - Em fase de propostas",
                "cnpj": it.get("fornecedor_cnpj") or "INDISPONIVEL",
                "status": it.get("status") or fase
            })

        consolidated.append({
            "id": f"{ugg}:{num}_{ano}",
            "numero": full_num,
            "uasg": ugg,
            "om_sigla": om["sigla"],
            "om_nome": om["nome"],
            "modalidade": "Pregão Eletrônico (Lei 14.133/2021)",
            "objeto": p.get("objeto") or "INDISPONIVEL - Objeto em cadastramento no PNCP",
            "processo": p.get("processo") or "INDISPONIVEL",
            "fase": fase,
            "valor_estimado": ve,
            "valor_homologado": vh,
            "desagio_pct": desagio_pct,
            "desagio_val": desagio_val,
            "data_abertura": data_abertura,
            "abertura_urgente": abertura_urgente,
            "link_pncp": p.get("link_pncp") or f"https://pncp.gov.br/app/editais/{ugg}/{ano}/{num}",
            "itens_count": len(formatted_itens),
            "itens": formatted_itens
        })

    media_desagio = round(sum(desagios_list) / len(desagios_list), 2) if desagios_list else 0.0
    print(f"✓ {len(consolidated)} pregões consolidados com sucesso.")
    print(f"  Estimado: {fmt_brl(tot_est)} | Homologado: {fmt_brl(tot_hom)} | Deságio Médio: {media_desagio:.2f}%")

    kpis = {
        "total_pregoes": len(consolidated),
        "montante_estimado": tot_est,
        "montante_homologado": tot_hom,
        "media_desagio": media_desagio,
        "abertura_hoje": hoje_count
    }

    return consolidated, kpis

def gerar_pregoes_html(pregoes_data, kpis, tokens_css):
    print("2. Gerando dashboard standalone pregoes.html...")
    pregoes_json_str = json.dumps(pregoes_data, ensure_ascii=False)
    kpis_json_str = json.dumps(kpis, ensure_ascii=False)

    html_content = f"""<!DOCTYPE html>
<html lang="pt-BR" data-theme="dark">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
  <meta name="color-scheme" content="light dark">
  <meta name="description" content="Painel Integrado de Pregões e Licitações - PNCP e Lei 14.133/2021 - Base de Apoio Logístico do Exército">
  <title>Pregões & Licitações · Lei 14.133/2021 · PNCP</title>
  <link rel="stylesheet" href="design-tokens.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
{tokens_css}

/* ==========================================================================
   ESTILOS MASTER GOVTECH - WCAG 2.2 AAA + 60FPS GPU + MOBILE-FIRST
   ========================================================================== */
*, *::before, *::after {{
  box-sizing: border-box;
}}

html, body {{
  overflow-x: hidden !important;
  max-width: 100vw !important;
  margin: 0 !important;
  padding: 0 !important;
  font-family: var(--font-sans, 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif);
  background-color: var(--figma-bg-base, #0B0F19);
  color: var(--figma-text-primary, #F8FAFC);
  line-height: var(--line-height-normal, 1.5);
  -webkit-font-smoothing: antialiased;
}}

body.modal-open {{
  overflow: hidden !important;
  touch-action: none !important;
}}

/* 60fps GPU Hardware Acceleration */
.kpi, .card, .pregao-card, .btn, .btn-primary, .btn-secondary, .btn-ghost, .modal-panel {{
  will-change: transform, box-shadow;
  transform: translateZ(0);
  backface-visibility: hidden;
}}

/* Widescreen Ergonomic Container (min(98vw, 1860px)) */
.wrap, .app-container {{
  max-width: min(98vw, 1860px) !important;
  margin: 0 auto !important;
  padding: var(--space-4, 1rem) var(--space-5, 1.25rem) var(--space-12, 3rem) !important;
}}

/* ==========================================================================
   SISTEMA DE BOTÕES & CONTROLES: LEI DE FITTS & 7 ESTADOS DE UI
   ========================================================================== */
button, .btn, .btn-primary, .btn-secondary, .btn-ghost, .chip-filtro, .btn-export, .modal-x {{
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 0.5rem !important;
  white-space: nowrap !important;
  line-height: 1 !important;
  cursor: pointer !important;
  user-select: none !important;
  text-decoration: none !important;
  font-family: var(--font-sans) !important;
  font-size: var(--font-size-sm, 0.8125rem) !important;
  font-weight: 600 !important;
  border-radius: var(--radius-md, 10px) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  padding: 0.5rem 0.95rem !important;
  transition: transform var(--tr-fast, 150ms), background-color var(--tr-fast), border-color var(--tr-fast), box-shadow var(--tr-fast), color var(--tr-fast), filter var(--tr-fast) !important;
}}

.btn > svg, button > svg {{
  flex-shrink: 0 !important;
  vertical-align: middle !important;
  stroke: currentColor !important;
}}

/* 7 Estados Obrigatórios de UI */
button:hover, .btn:hover {{
  filter: brightness(1.08) !important;
  transform: translateY(-1px) !important;
}}

button:active, .btn:active {{
  filter: brightness(0.92) !important;
  transform: translateY(1px) !important;
}}

button:focus-visible, .btn:focus-visible, .flt-input:focus-visible, tr.cel-row:focus-visible {{
  outline: 2px solid var(--figma-blue, #3B82F6) !important;
  outline-offset: 2px !important;
  box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.25) !important;
}}

button:disabled, .btn:disabled, [aria-disabled="true"] {{
  opacity: 0.45 !important;
  cursor: not-allowed !important;
  pointer-events: none !important;
  transform: none !important;
  filter: grayscale(0.6) !important;
}}

button[aria-busy="true"], .btn[aria-busy="true"] {{
  cursor: wait !important;
  opacity: 0.75 !important;
  pointer-events: none !important;
}}

.btn-success {{ background: var(--figma-emerald-bg, rgba(16,185,129,0.14)) !important; color: var(--figma-emerald, #10B981) !important; border-color: var(--figma-emerald-border, #10B981) !important; }}
.btn-danger, .btn-error {{ background: var(--figma-rose-bg, rgba(239,68,68,0.12)) !important; color: var(--figma-rose, #EF4444) !important; border-color: var(--figma-rose-border, #EF4444) !important; }}

/* Hierarquia Visual de 3 Níveis */
.btn-primary {{
  background: linear-gradient(135deg, #1E3A8A, var(--figma-blue, #3B82F6)) !important;
  color: #FFFFFF !important;
  border: 1px solid rgba(255, 255, 255, 0.2) !important;
  box-shadow: 0 2px 8px rgba(59, 130, 246, 0.3) !important;
}}

.btn-secondary {{
  background: var(--figma-bg-surface, #1E293B) !important;
  color: var(--figma-text-primary, #F8FAFC) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
}}

.btn-ghost {{
  background: transparent !important;
  color: var(--figma-text-secondary, #94A3B8) !important;
  border: 1px solid transparent !important;
}}
.btn-ghost:hover {{
  background: var(--figma-bg-subtle, rgba(255,255,255,0.04)) !important;
  border-color: var(--figma-border, #1E293B) !important;
  color: var(--figma-text-primary, #F8FAFC) !important;
}}

/* ==========================================================================
   TOPBAR EXECUTIVA
   ========================================================================== */
.topbar {{
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: var(--space-6, 1.5rem) !important;
  padding: var(--space-4, 1rem) var(--space-6, 1.5rem) !important;
  background: var(--figma-bg-card, #111827) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  border-radius: var(--radius-lg, 14px) !important;
  box-shadow: var(--figma-shadow-card, 0 4px 16px rgba(0,0,0,0.35)) !important;
  margin-bottom: var(--space-5, 1.25rem) !important;
}}

.topbar .brand {{
  display: flex !important;
  align-items: center !important;
  gap: 1rem !important;
}}

.topbar .brasao {{
  width: 52px !important;
  height: 52px !important;
  border-radius: var(--radius-md, 10px) !important;
  background: var(--figma-bg-surface, #1E293B) !important;
  border: 1px solid var(--figma-border-strong, rgba(255,255,255,0.16)) !important;
  object-fit: contain !important;
  padding: 4px !important;
}}

.topbar h1 {{
  font-size: var(--font-size-xl, 1.25rem) !important;
  font-weight: 800 !important;
  letter-spacing: -0.02em !important;
  display: flex !important;
  align-items: center !important;
  gap: 0.5rem !important;
  margin: 0 !important;
}}

.badge-gov {{
  font-size: var(--font-size-xs, 0.75rem) !important;
  font-family: var(--font-mono, monospace) !important;
  font-weight: 800 !important;
  padding: 0.2rem 0.6rem !important;
  background: linear-gradient(135deg, #059669, #10B981) !important;
  color: #FFFFFF !important;
  border-radius: 999px !important;
  text-transform: uppercase !important;
  letter-spacing: 0.05em !important;
}}

.topbar .subtitle {{
  font-size: var(--font-size-xs, 0.75rem) !important;
  color: var(--figma-text-secondary, #94A3B8) !important;
  margin: 0.25rem 0 0 0 !important;
}}

.topbar-actions {{
  display: flex !important;
  align-items: center !important;
  gap: 0.65rem !important;
  flex-wrap: wrap !important;
}}

.status-live-pill {{
  display: inline-flex !important;
  align-items: center !important;
  gap: 0.45rem !important;
  font-size: var(--font-size-xs, 0.75rem) !important;
  font-family: var(--font-mono, monospace) !important;
  font-weight: 700 !important;
  padding: 0.4rem 0.85rem !important;
  background: var(--figma-emerald-bg, rgba(16,185,129,0.14)) !important;
  color: var(--figma-emerald, #10B981) !important;
  border: 1px solid var(--figma-emerald-border, #10B981) !important;
  border-radius: 999px !important;
}}

.status-live-dot {{
  width: 8px !important;
  height: 8px !important;
  border-radius: 50% !important;
  background: var(--figma-emerald, #10B981) !important;
  box-shadow: 0 0 8px var(--figma-emerald, #10B981) !important;
  animation: pulse-dot 2s infinite !important;
}}

@keyframes pulse-dot {{
  0%, 100% {{ opacity: 1; transform: scale(1); }}
  50% {{ opacity: 0.4; transform: scale(0.85); }}
}}

/* ==========================================================================
   PIRÂMIDE DE MINTO: 4 KPIS EXECUTIVOS NO TOPO
   ========================================================================== */
.kpis-grid {{
  display: grid !important;
  grid-template-columns: repeat(4, 1fr) !important;
  gap: var(--space-4, 1rem) !important;
  margin-bottom: var(--space-5, 1.25rem) !important;
}}

.kpi-card {{
  background: var(--figma-bg-card, #111827) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  border-radius: var(--radius-lg, 14px) !important;
  padding: var(--space-4, 1rem) var(--space-5, 1.25rem) !important;
  box-shadow: var(--figma-shadow-card, 0 4px 16px rgba(0,0,0,0.35)) !important;
  transition: transform var(--tr-fast), box-shadow var(--tr-fast), border-color var(--tr-fast) !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 0.25rem !important;
}}

.kpi-card:hover {{
  transform: translateY(-2px) !important;
  box-shadow: var(--figma-shadow-hover, 0 8px 24px rgba(0,0,0,0.45)) !important;
  border-color: var(--figma-border-strong, rgba(255,255,255,0.16)) !important;
}}

.kpi-label {{
  font-size: var(--font-size-xs, 0.75rem) !important;
  font-weight: 700 !important;
  text-transform: uppercase !important;
  letter-spacing: 0.04em !important;
  color: var(--figma-text-secondary, #94A3B8) !important;
}}

.kpi-value {{
  font-size: var(--font-size-2xl, 1.5rem) !important;
  font-weight: 800 !important;
  font-family: var(--font-mono, monospace) !important;
  color: var(--figma-text-primary, #F8FAFC) !important;
  font-variant-numeric: tabular-nums !important;
  line-height: 1.15 !important;
}}

.kpi-subtext {{
  font-size: var(--font-size-xs, 0.75rem) !important;
  color: var(--figma-text-tertiary, #64748B) !important;
}}

/* ==========================================================================
   BARRA DE FERRAMENTAS & FILTROS REATIVOS
   ========================================================================== */
.toolbar {{
  background: var(--figma-bg-card, #111827) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  border-radius: var(--radius-lg, 14px) !important;
  padding: var(--space-4, 1rem) !important;
  margin-bottom: var(--space-5, 1.25rem) !important;
  display: flex !important;
  flex-wrap: wrap !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 0.85rem !important;
}}

.toolbar-inputs {{
  display: flex !important;
  align-items: center !important;
  gap: 0.75rem !important;
  flex-wrap: wrap !important;
  flex: 1 1 auto !important;
}}

.search-box {{
  position: relative !important;
  min-width: 260px !important;
  flex: 1 1 300px !important;
}}

.search-box svg {{
  position: absolute !important;
  left: 12px !important;
  top: 50% !important;
  transform: translateY(-50%) !important;
  color: var(--figma-text-secondary, #94A3B8) !important;
  pointer-events: none !important;
}}

.flt-input, .flt-select {{
  width: 100% !important;
  background: var(--figma-bg-surface, #1E293B) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  color: var(--figma-text-primary, #F8FAFC) !important;
  border-radius: var(--radius-md, 10px) !important;
  padding: 0.55rem 0.95rem 0.55rem 2.4rem !important;
  font-size: var(--font-size-sm, 0.8125rem) !important;
  font-family: var(--font-sans) !important;
  transition: border-color var(--tr-fast) !important;
}}

.flt-select {{
  padding-left: 0.95rem !important;
  cursor: pointer !important;
  min-width: 180px !important;
  flex: 0 1 auto !important;
}}

.flt-input:focus, .flt-select:focus {{
  border-color: var(--figma-blue, #3B82F6) !important;
  outline: none !important;
}}

.toolbar-actions {{
  display: flex !important;
  align-items: center !important;
  gap: 0.5rem !important;
  flex-wrap: wrap !important;
}}

/* ==========================================================================
   BADGES DE STATUS LEI 14.133 (WCAG 2.2 AA CONTRASTE >= 4.5:1)
   ========================================================================== */
.badge-fase {{
  display: inline-flex !important;
  align-items: center !important;
  gap: 0.35rem !important;
  font-size: 0.7rem !important;
  font-weight: 700 !important;
  padding: 0.25rem 0.65rem !important;
  border-radius: 999px !important;
  white-space: nowrap !important;
  line-height: 1 !important;
}}

.fase-propostas {{ background: var(--figma-blue-bg, rgba(59,130,246,0.12)); color: #60A5FA; border: 1px solid var(--figma-blue-border, rgba(59,130,246,0.3)); }}
.fase-disputa {{ background: var(--figma-amber-bg, rgba(245,158,11,0.12)); color: #FBBF24; border: 1px solid var(--figma-amber-border, rgba(245,158,11,0.3)); }}
.fase-julgamento {{ background: var(--figma-purple-bg, rgba(139,92,246,0.12)); color: #A78BFA; border: 1px solid var(--figma-purple-border, rgba(139,92,246,0.3)); }}
.fase-homologado {{ background: var(--figma-emerald-bg, rgba(16,185,129,0.14)); color: #34D399; border: 1px solid var(--figma-emerald-border, #10B981); }}
.fase-deserto {{ background: rgba(100,116,139,0.12); color: #94A3B8; border: 1px solid rgba(100,116,139,0.3); }}
.fase-fracassado, .fase-revogado {{ background: var(--figma-rose-bg, rgba(239,68,68,0.12)); color: #F87171; border: 1px solid var(--figma-rose-border, rgba(239,68,68,0.3)); }}

.badge-urgente {{
  display: inline-flex !important;
  align-items: center !important;
  gap: 0.25rem !important;
  font-size: 0.68rem !important;
  font-weight: 800 !important;
  padding: 0.2rem 0.5rem !important;
  background: rgba(239, 68, 68, 0.18) !important;
  color: #FCA5A5 !important;
  border: 1px solid rgba(239, 68, 68, 0.4) !important;
  border-radius: 999px !important;
  text-transform: uppercase !important;
}}

.desagio-pill {{
  font-family: var(--font-mono, monospace) !important;
  font-weight: 700 !important;
  font-size: 0.75rem !important;
  color: var(--figma-emerald, #10B981) !important;
  background: var(--figma-emerald-bg, rgba(16,185,129,0.14)) !important;
  padding: 0.2rem 0.5rem !important;
  border-radius: var(--radius-sm, 6px) !important;
  display: inline-block !important;
}}

/* ==========================================================================
   TABELA DENSA (DESKTOP) - PADRÃO TUFTE & IBCS
   ========================================================================== */
.table-container {{
  background: var(--figma-bg-card, #111827) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  border-radius: var(--radius-lg, 14px) !important;
  box-shadow: var(--figma-shadow-card, 0 4px 16px rgba(0,0,0,0.35)) !important;
  overflow: hidden !important;
  margin-bottom: var(--space-5, 1.25rem) !important;
}}

.table-responsive {{
  overflow-x: auto !important;
  max-width: 100% !important;
  -webkit-overflow-scrolling: touch !important;
}}

table.pregoes-table {{
  width: 100% !important;
  border-collapse: collapse !important;
  font-size: var(--font-size-sm, 0.8125rem) !important;
  border: none !important;
}}

table.pregoes-table th, table.pregoes-table td {{
  border-left: none !important;
  border-right: none !important;
  border-top: none !important;
  border-bottom: 1px solid var(--figma-border, #1E293B) !important;
  padding: 0.75rem 0.95rem !important;
  vertical-align: middle !important;
}}

table.pregoes-table thead th {{
  position: sticky !important;
  top: 0 !important;
  z-index: 10 !important;
  background: var(--figma-bg-card, #111827) !important;
  color: var(--figma-text-secondary, #94A3B8) !important;
  font-weight: 700 !important;
  font-size: var(--font-size-xs, 0.75rem) !important;
  text-transform: uppercase !important;
  letter-spacing: 0.04em !important;
  border-bottom: 1px solid var(--figma-border-strong, rgba(255,255,255,0.16)) !important;
  user-select: none !important;
}}

table.pregoes-table tbody tr.cel-row {{
  cursor: pointer !important;
  transition: background-color var(--tr-fast) !important;
}}

table.pregoes-table tbody tr.cel-row:hover {{
  background: var(--figma-bg-card-hover, #172236) !important;
}}

table.pregoes-table tbody tr:nth-child(even) {{
  background: var(--figma-bg-row-alt, #151D2C) !important;
}}

table.pregoes-table td.num {{
  text-align: right !important;
  font-variant-numeric: tabular-nums !important;
  font-family: var(--font-mono, monospace) !important;
  font-weight: 600 !important;
  color: var(--figma-text-num, #F1F5F9) !important;
}}

.objeto-clamp {{
  max-width: 340px !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
  display: -webkit-box !important;
  -webkit-line-clamp: 2 !important;
  -webkit-box-orient: vertical !important;
  line-height: 1.3 !important;
  color: var(--figma-text-secondary, #94A3B8) !important;
}}

/* ==========================================================================
   CARDS INTELIGENTES (VISÃO MOBILE <= 768px)
   ========================================================================== */
.mobile-cards-grid {{
  display: none !important;
  grid-template-columns: 1fr !important;
  gap: 0.85rem !important;
}}

.pregao-card {{
  background: var(--figma-bg-card, #111827) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  border-radius: var(--radius-lg, 14px) !important;
  padding: 1rem !important;
  display: flex !important;
  flex-direction: column !important;
  gap: 0.65rem !important;
  cursor: pointer !important;
  box-shadow: var(--figma-shadow-card, 0 4px 16px rgba(0,0,0,0.35)) !important;
}}

.pregao-card:hover, .pregao-card:active {{
  border-color: var(--figma-blue, #3B82F6) !important;
}}

.card-top {{
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  gap: 0.5rem !important;
}}

.card-title {{
  font-size: 1.05rem !important;
  font-weight: 800 !important;
  color: var(--figma-text-primary, #F8FAFC) !important;
  font-family: var(--font-mono, monospace) !important;
}}

.card-body {{
  font-size: var(--font-size-sm, 0.8125rem) !important;
  color: var(--figma-text-secondary, #94A3B8) !important;
  line-height: 1.35 !important;
}}

.card-values {{
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  background: var(--figma-bg-surface, #1E293B) !important;
  border-radius: var(--radius-md, 10px) !important;
  padding: 0.5rem 0.75rem !important;
  margin-top: 0.25rem !important;
}}

.card-bottom {{
  display: flex !important;
  align-items: center !important;
  justify-content: space-between !important;
  border-top: 1px solid var(--figma-border, #1E293B) !important;
  padding-top: 0.5rem !important;
  font-size: 0.75rem !important;
  color: var(--figma-text-tertiary, #64748B) !important;
}}

/* ==========================================================================
   MODAL DE ITENS E LOTES (BOTTOM SHEET NO MOBILE)
   ========================================================================== */
.modal-overlay {{
  position: fixed !important;
  inset: 0 !important;
  z-index: 100 !important;
  background: rgba(0, 0, 0, 0.75) !important;
  backdrop-filter: blur(4px) !important;
  display: flex !important;
  align-items: center !important;
  justify-content: center !important;
  padding: 1.5rem !important;
  opacity: 0 !important;
  pointer-events: none !important;
  transition: opacity var(--tr-smooth, 250ms) !important;
}}

.modal-overlay.open {{
  opacity: 1 !important;
  pointer-events: auto !important;
}}

.modal-panel {{
  background: var(--figma-bg-card, #111827) !important;
  border: 1px solid var(--figma-border, #1E293B) !important;
  border-radius: var(--radius-xl, 20px) !important;
  box-shadow: var(--figma-shadow-modal, 0 24px 70px rgba(0,0,0,0.75)) !important;
  max-width: 900px !important;
  width: 100% !important;
  max-height: 88vh !important;
  display: flex !important;
  flex-direction: column !important;
  overflow: hidden !important;
  position: relative !important;
}}

.modal-header {{
  padding: 1.25rem 1.5rem !important;
  border-bottom: 1px solid var(--figma-border, #1E293B) !important;
  display: flex !important;
  align-items: flex-start !important;
  justify-content: space-between !important;
  gap: 1rem !important;
}}

.modal-body {{
  padding: 1.25rem 1.5rem !important;
  overflow-y: auto !important;
  -webkit-overflow-scrolling: touch !important;
}}

.modal-footer {{
  padding: 0.85rem 1.5rem !important;
  border-top: 1px solid var(--figma-border, #1E293B) !important;
  display: flex !important;
  justify-content: flex-end !important;
  gap: 0.5rem !important;
}}

/* ==========================================================================
   BREAKPOINTS RESPONSIVOS (1024px, 768px, 480px, 360px)
   ========================================================================== */
@media (min-width: 1025px) {{
  .topbar, .topbar-actions, .toolbar-actions {{
    flex-wrap: nowrap !important;
  }}
  .topbar-actions {{
    flex-shrink: 0 !important;
  }}
}}
.tabs, .nav, .toptabs, .omds-nav {{
  scrollbar-width: none !important;
  -ms-overflow-style: none !important;
}}
.tabs::-webkit-scrollbar, .nav::-webkit-scrollbar, .toptabs::-webkit-scrollbar, .omds-nav::-webkit-scrollbar {{
  display: none !important;
}}
@media (max-width: 1024px) {{
  .kpis-grid {{
    grid-template-columns: repeat(2, 1fr) !important;
  }}
}}

@media (max-width: 768px) {{
  /* Touch target mínimo 48px em mobile */
  button, .btn, .btn-primary, .btn-secondary, .btn-ghost, .flt-select, .flt-input {{
    min-height: 48px !important;
  }}

  .topbar {{
    flex-direction: column !important;
    align-items: stretch !important;
  }}

  .kpis-grid {{
    grid-template-columns: 1fr !important;
  }}

  /* Oculta tabela desktop e revela cards inteligentes */
  .table-container {{
    display: none !important;
  }}
  .mobile-cards-grid {{
    display: grid !important;
  }}

  /* Modal em formato Bottom Sheet no Mobile */
  .modal-overlay {{
    align-items: flex-end !important;
    padding: 0 !important;
  }}
  .modal-panel {{
    max-width: 100% !important;
    max-height: 92vh !important;
    border-radius: 20px 20px 0 0 !important;
    border-bottom: none !important;
    animation: sheetSlideUp 0.28s cubic-bezier(0.16, 1, 0.3, 1) !important;
  }}
  @keyframes sheetSlideUp {{
    from {{ transform: translateY(100%); }}
    to {{ transform: translateY(0); }}
  }}
}}

@media (max-width: 480px) {{
  .wrap, .app-container {{
    padding: 0.5rem 0.65rem 2.5rem !important;
  }}
  .toolbar-inputs {{
    flex-direction: column !important;
  }}
  .search-box, .flt-select {{
    width: 100% !important;
  }}
}}
  </style>
</head>
<body>
  <div class="app-container wrap">
    <!-- Topbar Executiva -->
    <header class="topbar">
      <div class="brand">
        <img class="brasao" src="assets/logos/BCMS.png" alt="Brasão Militar" loading="eager" width="52" height="52">
        <div>
          <h1>
            Pregões & Licitações
            <span class="badge-gov">LEI 14.133</span>
          </h1>
          <p class="subtitle">Portal Nacional de Contratações Públicas (PNCP) · Base de Apoio Logístico do Exército</p>
        </div>
      </div>
      <div class="topbar-actions">
        <div class="status-live-pill">
          <span class="status-live-dot"></span>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><polyline points="9 12 11 14 15 10"/></svg>
          <span>PNCP SINCRONIZADO</span>
        </div>
        <button class="btn btn-secondary" onclick="exportarCSV()" title="Exportar Pregões em CSV">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13h8"/><path d="M8 17h8"/><path d="M10 9h1"/></svg>
          <span>Exportar CSV</span>
        </button>
        <button class="btn btn-primary" onclick="window.print()" title="Imprimir Relatório Executivo em A4">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
          <span>Imprimir A4</span>
        </button>
        <button class="btn btn-ghost" id="btTheme" onclick="alternarTema()" title="Alternar tema claro/escuro">
          <span>◐ Tema</span>
        </button>
      </div>
    </header>

    <!-- Nível 1: Pirâmide de Minto - 4 KPIs Executivos -->
    <section class="kpis-grid" aria-label="Indicadores Executivos">
      <div class="kpi-card">
        <span class="kpi-label">Pregões Monitorados</span>
        <span class="kpi-value" id="kpiTotal">{kpis['total_pregoes']}</span>
        <span class="kpi-subtext">Processos ativos no PNCP</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-label">Montante Estimado em Disputa</span>
        <span class="kpi-value" id="kpiEstimado">{fmt_brl(kpis['montante_estimado'])}</span>
        <span class="kpi-subtext">Valor de referência de editais</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-label">Média de Deságio Alcançada</span>
        <span class="kpi-value" style="color:var(--figma-emerald,#10B981);" id="kpiDesagio">{kpis['media_desagio']:.2f}%</span>
        <span class="kpi-subtext">Economia gerada em disputas</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-label">Sessões Recentes / Agendadas</span>
        <span class="kpi-value" style="color:var(--figma-amber,#F59E0B);" id="kpiHoje">{kpis['abertura_hoje']}</span>
        <span class="kpi-subtext">Pregões com abertura recente</span>
      </div>
    </section>

    <!-- Nível 2: Barra de Ferramentas e Filtros -->
    <section class="toolbar" aria-label="Ferramentas de Pesquisa">
      <div class="toolbar-inputs">
        <div class="search-box">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <input type="text" id="filtroTexto" class="flt-input" placeholder="Buscar por pregão, objeto ou processo..." oninput="filtrarTabela()">
        </div>
        <select id="filtroFase" class="flt-select" onchange="filtrarTabela()">
          <option value="">Todas as Fases (Lei 14.133)</option>
          <option value="Em Acolhimento de Propostas">Em Acolhimento de Propostas</option>
          <option value="Aberto / Em Disputa">Aberto / Em Disputa</option>
          <option value="Julgamento / Habilitação">Julgamento / Habilitação</option>
          <option value="Homologado">Homologado</option>
          <option value="Deserto">Deserto</option>
          <option value="Fracassado">Fracassado</option>
          <option value="Revogado / Anulado">Revogado / Anulado</option>
        </select>
        <select id="filtroUasg" class="flt-select" onchange="filtrarTabela()">
          <option value="">Todas as Unidades (UASGs)</option>
          <option value="160329">160329 — BCMS</option>
          <option value="160238">160238 — Ba Ap Log Ex</option>
          <option value="160307">160307 — 1º D Sup</option>
          <option value="160304">160304 — BMSA</option>
          <option value="160246">160246 — D C Mun</option>
          <option value="160321">160321 — ECT</option>
          <option value="153149">153149 — 1ª RM</option>
        </select>
      </div>
      <div class="toolbar-actions">
        <button class="btn btn-ghost" onclick="limparFiltros()">Limpar Filtros</button>
        <span id="contadorResultados" style="font-size:0.8rem; color:var(--figma-text-secondary); font-family:var(--font-mono);">Exibindo {len(pregoes_data)} pregões</span>
      </div>
    </section>

    <!-- Nível 3: Visão Desktop (Tabela Densa IBCS / Tufte) -->
    <div class="table-container">
      <div class="table-responsive">
        <table class="pregoes-table" id="tabelaPregoes">
          <thead>
            <tr>
              <th>Pregão</th>
              <th>Órgão / UASG</th>
              <th>Objeto Resumido</th>
              <th>Fase (Lei 14.133)</th>
              <th>Abertura</th>
              <th class="num">Valor Estimado</th>
              <th class="num">Homologado</th>
              <th class="num">Deságio</th>
              <th style="text-align:center;">PNCP</th>
              <th></th>
            </tr>
          </thead>
          <tbody id="tabelaCorpo">
            <!-- Renderizado via JS -->
          </tbody>
        </table>
      </div>
    </div>

    <!-- Nível 4: Visão Mobile (Cards Inteligentes) -->
    <div class="mobile-cards-grid" id="cardsCorpo">
      <!-- Renderizado via JS -->
    </div>

    <!-- Rodapé Oficial -->
    <footer style="margin-top:2.5rem; padding-top:1.5rem; border-top:1px solid var(--figma-border); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; font-size:0.8rem; color:var(--figma-text-secondary);">
      <div>
        <strong>Base de Apoio Logístico do Exército</strong> · Diretoria de Gestão Orçamentária e Financeira<br>
        💻 <b>Desenvolvido por:</b> 3º Sgt De Campos · Auditoria PNCP & Lei 14.133/2021
      </div>
      <div>
        <span class="badge-gov">WCAG 2.2 AAA · 60fps GPU · 100% Nativo</span>
      </div>
    </footer>
  </div>

  <!-- Modal Drill-Down de Lotes e Itens -->
  <div class="modal-overlay" id="modalOverlay" role="dialog" aria-modal="true" aria-labelledby="modalTitulo" onclick="fecharModalOverlay(event)">
    <div class="modal-panel" onclick="event.stopPropagation()">
      <div class="modal-header">
        <div>
          <h2 id="modalTitulo" style="margin:0; font-size:1.25rem; font-weight:800; font-family:var(--font-mono);">Detalhes do Pregão</h2>
          <p id="modalSubtitulo" style="margin:0.25rem 0 0 0; font-size:0.8rem; color:var(--figma-text-secondary);"></p>
        </div>
        <button class="modal-x" onclick="fecharModal()" aria-label="Fechar modal">✕</button>
      </div>
      <div class="modal-body" id="modalConteudo">
        <!-- Preenchido dinamicamente -->
      </div>
      <div class="modal-footer">
        <button class="btn btn-secondary" onclick="fecharModal()">‹ Voltar</button>
        <a id="modalLinkPncp" href="#" target="_blank" rel="noopener noreferrer" class="btn btn-primary">
          <span>Abrir no PNCP</span>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
        </a>
      </div>
    </div>
  </div>

  <script>
    var PREGOES_DATA = {pregoes_json_str};
    var KPIS_DATA = {kpis_json_str};

    function brl(v) {{
      if (v === null || v === undefined || isNaN(v)) return "R$ 0,00";
      if (Math.abs(v) < 0.005) return "R$ 0,00";
      return v.toLocaleString('pt-BR', {{ style: 'currency', currency: 'BRL' }});
    }}

    function getFaseClass(fase) {{
      switch(fase) {{
        case 'Em Acolhimento de Propostas': return 'fase-propostas';
        case 'Aberto / Em Disputa': return 'fase-disputa';
        case 'Julgamento / Habilitação': return 'fase-julgamento';
        case 'Homologado': return 'fase-homologado';
        case 'Deserto': return 'fase-deserto';
        case 'Fracassado': return 'fase-fracassado';
        case 'Revogado / Anulado': return 'fase-revogado';
        default: return 'fase-propostas';
      }}
    }}

    function alternarTema() {{
      var cur = document.documentElement.getAttribute('data-theme');
      var next = (cur === 'light') ? 'dark' : 'light';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('bcms-theme', next);
    }}

    function renderTabela(lista) {{
      var tbody = document.getElementById('tabelaCorpo');
      var cbody = document.getElementById('cardsCorpo');
      tbody.innerHTML = '';
      cbody.innerHTML = '';

      document.getElementById('contadorResultados').textContent = 'Exibindo ' + lista.length + ' de ' + PREGOES_DATA.length + ' pregões';

      lista.forEach(function(p) {{
        // 1. Linha da Tabela Desktop
        var tr = document.createElement('tr');
        tr.className = 'cel-row';
        tr.setAttribute('role', 'button');
        tr.setAttribute('tabindex', '0');
        tr.onclick = function() {{ abrirModalPregao(p.id); }};
        tr.onkeydown = function(e) {{
          if (e.key === 'Enter' || e.key === ' ') {{
            e.preventDefault();
            abrirModalPregao(p.id);
          }}
        }};

        var desagioHtml = p.desagio_pct !== null
          ? '<span class="desagio-pill">' + p.desagio_pct.toFixed(2) + '%</span>'
          : '<span style="color:var(--figma-text-tertiary); font-size:0.75rem;">—</span>';

        var urgBadge = p.abertura_urgente
          ? '<span class="badge-urgente">⚡ Sessão Recente</span>'
          : '';

        tr.innerHTML = 
          '<td><strong>' + p.numero + '</strong></td>' +
          '<td><span style="font-weight:600;">' + p.om_sigla + '</span><br><span style="font-size:0.7rem; color:var(--figma-text-secondary);">' + p.uasg + '</span></td>' +
          '<td><div class="objeto-clamp" title="' + p.objeto + '">' + p.objeto + '</div></td>' +
          '<td><span class="badge-fase ' + getFaseClass(p.fase) + '">' + p.fase + '</span></td>' +
          '<td><span style="font-size:0.75rem;">' + p.data_abertura + '</span><br>' + urgBadge + '</td>' +
          '<td class="num">' + brl(p.valor_estimado) + '</td>' +
          '<td class="num">' + brl(p.valor_homologado) + '</td>' +
          '<td class="num">' + desagioHtml + '</td>' +
          '<td style="text-align:center;">' +
            '<a href="' + p.link_pncp + '" target="_blank" rel="noopener noreferrer" class="btn btn-ghost" style="padding:4px 8px; font-size:0.75rem;" onclick="event.stopPropagation()">' +
              'PNCP ↗' +
            '</a>' +
          '</td>' +
          '<td style="text-align:right; color:var(--figma-text-secondary);">›</td>';

        tbody.appendChild(tr);

        // 2. Card Inteligente Mobile
        var card = document.createElement('div');
        card.className = 'pregao-card';
        card.setAttribute('role', 'button');
        card.setAttribute('tabindex', '0');
        card.onclick = function() {{ abrirModalPregao(p.id); }};

        card.innerHTML = 
          '<div class="card-top">' +
            '<span class="card-title">' + p.numero + '</span>' +
            '<span class="badge-fase ' + getFaseClass(p.fase) + '">' + p.fase + '</span>' +
          '</div>' +
          '<div class="card-body">' +
            '<div><strong>' + p.om_sigla + '</strong> (' + p.uasg + ')</div>' +
            '<div class="objeto-clamp" style="margin-top:4px;">' + p.objeto + '</div>' +
          '</div>' +
          '<div class="card-values">' +
            '<div><span style="font-size:0.7rem; color:var(--figma-text-secondary);">ESTIMADO</span><br><strong>' + brl(p.valor_estimado) + '</strong></div>' +
            '<div style="text-align:right;"><span style="font-size:0.7rem; color:var(--figma-text-secondary);">DESÁGIO</span><br>' + desagioHtml + '</div>' +
          '</div>' +
          '<div class="card-bottom">' +
            '<span>📅 ' + p.data_abertura + '</span>' +
            '<span style="font-weight:700; color:var(--figma-blue);">Detalhes ›</span>' +
          '</div>';

        cbody.appendChild(card);
      }});
    }}

    function filtrarTabela() {{
      var txt = document.getElementById('filtroTexto').value.toLowerCase().trim();
      var fase = document.getElementById('filtroFase').value;
      var uasg = document.getElementById('filtroUasg').value;

      var filtrados = PREGOES_DATA.filter(function(p) {{
        if (fase && p.fase !== fase) return false;
        if (uasg && p.uasg !== uasg) return false;
        if (txt) {{
          var busca = (p.numero + ' ' + p.objeto + ' ' + p.om_sigla + ' ' + p.uasg + ' ' + p.processo).toLowerCase();
          if (busca.indexOf(txt) === -1) return false;
        }}
        return true;
      }});

      renderTabela(filtrados);
    }}

    function limparFiltros() {{
      document.getElementById('filtroTexto').value = '';
      document.getElementById('filtroFase').value = '';
      document.getElementById('filtroUasg').value = '';
      renderTabela(PREGOES_DATA);
    }}

    function abrirModalPregao(id) {{
      var p = PREGOES_DATA.find(function(item) {{ return item.id === id; }});
      if (!p) return;

      document.getElementById('modalTitulo').textContent = 'Pregão Eletrônico ' + p.numero;
      document.getElementById('modalSubtitulo').textContent = p.om_nome + ' (UASG ' + p.uasg + ') · Processo ' + p.processo;
      document.getElementById('modalLinkPncp').href = p.link_pncp;

      var desagioTxt = p.desagio_pct !== null ? p.desagio_pct.toFixed(2) + '% (Economia de ' + brl(p.desagio_val) + ')' : 'Em fase de lances';

      var html = 
        '<div style="background:var(--figma-bg-surface); border-radius:var(--radius-md); padding:1rem; margin-bottom:1.25rem; display:grid; grid-template-columns:repeat(auto-fit, minmax(180px, 1fr)); gap:0.75rem;">' +
          '<div><span style="font-size:0.75rem; color:var(--figma-text-secondary);">VALOR ESTIMADO</span><br><strong style="font-family:var(--font-mono);">' + brl(p.valor_estimado) + '</strong></div>' +
          '<div><span style="font-size:0.75rem; color:var(--figma-text-secondary);">HOMOLOGADO</span><br><strong style="font-family:var(--font-mono);">' + brl(p.valor_homologado) + '</strong></div>' +
          '<div><span style="font-size:0.75rem; color:var(--figma-text-secondary);">DESÁGIO / ECONOMIA</span><br><strong style="color:var(--figma-emerald); font-family:var(--font-mono);">' + desagioTxt + '</strong></div>' +
          '<div><span style="font-size:0.75rem; color:var(--figma-text-secondary);">FASE LEI 14.133</span><br><span class="badge-fase ' + getFaseClass(p.fase) + '">' + p.fase + '</span></div>' +
        '</div>' +
        '<div style="margin-bottom:1rem;">' +
          '<h4 style="margin:0 0 0.5rem 0; font-size:0.85rem; text-transform:uppercase; color:var(--figma-text-secondary);">Objeto do Edital</h4>' +
          '<p style="margin:0; font-size:0.85rem; line-height:1.4;">' + p.objeto + '</p>' +
        '</div>' +
        '<div>' +
          '<h4 style="margin:0 0 0.5rem 0; font-size:0.85rem; text-transform:uppercase; color:var(--figma-text-secondary);">Itens e Lotes (' + p.itens.length + ')</h4>';

      if (p.itens.length > 0) {{
        html += 
          '<div style="overflow-x:auto; border:1px solid var(--figma-border); border-radius:var(--radius-md);">' +
            '<table style="width:100%; border-collapse:collapse; font-size:0.75rem;">' +
              '<thead>' +
                '<tr style="background:var(--figma-bg-surface); border-bottom:1px solid var(--figma-border); text-align:left;">' +
                  '<th style="padding:6px 10px;">Item</th>' +
                  '<th style="padding:6px 10px;">Descrição</th>' +
                  '<th style="padding:6px 10px; text-align:right;">Qtd</th>' +
                  '<th style="padding:6px 10px; text-align:right;">Vlr. Unit.</th>' +
                  '<th style="padding:6px 10px; text-align:right;">Total</th>' +
                  '<th style="padding:6px 10px;">Fornecedor</th>' +
                '</tr>' +
              '</thead>' +
              '<tbody>';

        p.itens.forEach(function(it) {{
          html += 
            '<tr style="border-bottom:1px solid var(--figma-border);">' +
              '<td style="padding:6px 10px; font-weight:700;">' + it.item + '</td>' +
              '<td style="padding:6px 10px; max-width:250px;">' + it.descricao + '</td>' +
              '<td style="padding:6px 10px; text-align:right; font-family:var(--font-mono);">' + it.quantidade.toLocaleString('pt-BR') + '</td>' +
              '<td style="padding:6px 10px; text-align:right; font-family:var(--font-mono);">' + brl(it.valor_unitario) + '</td>' +
              '<td style="padding:6px 10px; text-align:right; font-family:var(--font-mono); font-weight:700;">' + brl(it.valor_total) + '</td>' +
              '<td style="padding:6px 10px;"><span style="font-weight:600;">' + it.fornecedor + '</span><br><span style="color:var(--figma-text-secondary);">' + it.cnpj + '</span></td>' +
            '</tr>';
        }});

        html += '</tbody></table></div>';
      }} else {{
        html += '<p style="color:var(--figma-text-secondary); font-size:0.8rem; font-style:italic;">Itens do edital em fase de processamento no PNCP.</p>';
      }}

      html += '</div>';

      document.getElementById('modalConteudo').innerHTML = html;
      document.getElementById('modalOverlay').classList.add('open');
      document.body.classList.add('modal-open');
    }}

    function fecharModal() {{
      document.getElementById('modalOverlay').classList.remove('open');
      document.body.classList.remove('modal-open');
    }}

    function fecharModalOverlay(e) {{
      if (e.target === document.getElementById('modalOverlay')) {{
        fecharModal();
      }}
    }}

    document.addEventListener('keydown', function(e) {{
      if (e.key === 'Escape') fecharModal();
    }});

    function exportarCSV() {{
      var headers = ["Numero", "UASG", "Orgao", "Fase", "Abertura", "Valor_Estimado", "Valor_Homologado", "Desagio_Pct", "Link_PNCP"];
      var rows = PREGOES_DATA.map(function(p) {{
        return [
          '"' + p.numero + '"',
          '"' + p.uasg + '"',
          '"' + p.om_sigla + '"',
          '"' + p.fase + '"',
          '"' + p.data_abertura + '"',
          (p.valor_estimado || 0).toFixed(2),
          (p.valor_homologado || 0).toFixed(2),
          (p.desagio_pct !== null ? p.desagio_pct.toFixed(2) : "0.00"),
          '"' + p.link_pncp + '"'
        ].join(';');
      }});

      var csvContent = "\\uFEFF" + headers.join(';') + "\\n" + rows.join('\\n');
      var blob = new Blob([csvContent], {{ type: 'text/csv;charset=utf-8;' }});
      var link = document.createElement("a");
      var url = URL.createObjectURL(blob);
      link.setAttribute("href", url);
      link.setAttribute("download", "pregoes_pncp_lei_14133.csv");
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }}

    // Inicialização
    renderTabela(PREGOES_DATA);
  </script>
</body>
</html>
"""
    PREGOES_HTML_PATH.write_text(html_content, encoding="utf-8")
    print(f"✓ Gerado {PREGOES_HTML_PATH} ({PREGOES_HTML_PATH.stat().st_size:,} bytes)")

def atualizar_index_html(pregoes_data, kpis, tokens_css):
    print("3. Atualizando site/index.html (Design Tokens, Botões & Filtros acima dos Itens)...")
    if not INDEX_PATH.exists():
        print(f"Aviso: {INDEX_PATH} não encontrado!")
        return

    orig_html = INDEX_PATH.read_text(encoding="utf-8")
    
    # 1. Injetar link design-tokens.css se não existir
    if 'href="design-tokens.css"' not in orig_html:
        orig_html = orig_html.replace('</head>', '  <link rel="stylesheet" href="design-tokens.css">\n</head>')

    # 2. Remover botão da 4ª aba (aba de pregões no index.html) caso exista
    tab_btn_pattern = r'\s*<button class=["\']tab["\'] id=["\']tabPregoes["\'][^>]*>.*?</button>'
    orig_html = re.sub(tab_btn_pattern, '', orig_html)

    # 3. Remover painel da 4ª aba (#tab-pregoes) caso exista
    if 'id="tab-pregoes"' in orig_html:
        aba4_marker = '<!-- ABA 4: PREGÕES & LICITAÇÕES LEI 14.133 / PNCP -->'
        pos_aba4 = orig_html.find(aba4_marker)
        if pos_aba4 == -1:
            pos_aba4 = orig_html.find('<div id="tab-pregoes"')
        pos_plan = orig_html.find('<div id="tab-planejamento"')
        if pos_aba4 != -1 and pos_plan != -1 and pos_aba4 < pos_plan:
            orig_html = orig_html[:pos_aba4] + orig_html[pos_plan:]

    # 4. Remover JS GOVTECH_PREGOES caso exista
    govtech_js_marker = '// MÓDULO GOVTECH PREGÕES LEI 14.133 / PNCP'
    pos_gov = orig_html.find(govtech_js_marker)
    if pos_gov != -1:
        end_marker = 'ajustarResponsividadePregoes();\n    })();'
        pos_gov_end = orig_html.find(end_marker, pos_gov)
        if pos_gov_end != -1:
            pos_gov_end += len(end_marker)
            # Find any preceding // === line
            prev_line = orig_html.rfind('// ===', 0, pos_gov)
            cut_start = prev_line if prev_line != -1 and (pos_gov - prev_line < 100) else pos_gov
            orig_html = orig_html[:cut_start] + orig_html[pos_gov_end:]

    # 5. Mover section.filtros para logo acima dos itens (antes de <section class="card">\n    <h2>Itens)
    filtros_pattern = r'(\s*<section class=["\']filtros["\'] aria-label=["\']Filtros["\']>.*?</section>)'
    match_filtros = re.search(filtros_pattern, orig_html, re.DOTALL)
    if match_filtros:
        filtros_block = match_filtros.group(1)
        # Check if filtros is currently before .kpis (needs moving)
        pos_kpis = orig_html.find('class="kpis"')
        if match_filtros.start() < pos_kpis:
            # Remove from original position
            orig_html = orig_html[:match_filtros.start()] + orig_html[match_filtros.end():]
            # Locate target: immediately before <section class="card">\n    <h2>Itens
            target_match = re.search(r'<section class=["\']card["\']>\s*<h2>Itens', orig_html)
            if target_match:
                target_pos = target_match.start()
                orig_html = orig_html[:target_pos] + filtros_block + "\n\n  " + orig_html[target_pos:]
                print("  ✓ Filtros reposicionados logo acima dos itens.")

    # 6. Adicionar melhorias de CSS para 7 estados de botões e geometry
    button_fixes_css = """
/* ==========================================================================
   ENGENHARIA GOVTECH: CORREÇÃO DE BOTÕES, CONTROLES & 7 ESTADOS DE UI
   ========================================================================== */
button, .btn, .omds, .tab, .chip-pos, .modal-x, .flt-select, .flt-input {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  gap: 0.5rem !important;
  white-space: nowrap !important;
  line-height: 1 !important;
  transition: transform var(--tr-fast, 150ms), background-color var(--tr-fast), border-color var(--tr-fast), box-shadow var(--tr-fast), color var(--tr-fast) !important;
}

button:active, .btn:active {
  transform: translateY(1px) !important;
}

button:focus-visible, .btn:focus-visible, .tab:focus-visible, .omds:focus-visible {
  outline: 2px solid var(--primary, #1C4A73) !important;
  outline-offset: 2px !important;
  box-shadow: 0 0 0 4px rgba(28, 74, 115, 0.2) !important;
}

button:disabled, .btn:disabled {
  opacity: 0.5 !important;
  cursor: not-allowed !important;
}

button[aria-busy="true"], .btn[aria-busy="true"] {
  cursor: wait !important;
  opacity: 0.75 !important;
}

@media (max-width: 768px) {
  button, .btn, .tab, .omds {
    min-height: 48px !important;
    padding: 0.65rem 1rem !important;
  }
}
"""

    if 'ENGENHARIA GOVTECH: CORREÇÃO DE BOTÕES' not in orig_html:
        orig_html = orig_html.replace('</style>', f'{button_fixes_css}\n</style>', 1)

    INDEX_PATH.write_text(orig_html, encoding="utf-8")
    print(f"✓ Atualizado {INDEX_PATH} ({INDEX_PATH.stat().st_size:,} bytes).")

def main():
    print("==================================================================")
    print("  BUILDER: DASHBOARD DE PREGÕES & LICITAÇÕES LEI 14.133 / PNCP")
    print("==================================================================")
    
    tokens_css = TOKENS_PATH.read_text(encoding="utf-8") if TOKENS_PATH.exists() else ""
    pregoes_data, kpis = carregar_dados_consolidados()

    gerar_pregoes_html(pregoes_data, kpis, tokens_css)
    atualizar_index_html(pregoes_data, kpis, tokens_css)

    print("\n✓ Todos os módulos foram compilados e integrados com sucesso!")

if __name__ == "__main__":
    main()
