"""
╔══════════════════════════════════════════════════════════════════════════════╗
║           MIRAI REPORTS — Gerador de Relatórios Gerenciais                  ║
║           Projeto Mirai · SAP S/4HANA · Sumitomo Rubber do Brasil           ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  Gera dois relatórios HTML a partir do cronograma MS Project (.xml):        ║
║    • report_gerencial_mirai.html  — Curva S · SPI · EVM vs BL5             ║
║    • report_atrasos_mirai.html    — Desvio BL5 (Nº15) e BL0 (Nº14)        ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  NOME PADRÃO DO ARQUIVO XML (sugestão):                                     ║
║    SRB_MIRAI_TP-0500_Cronograma_do_Projeto-DD-MM-AAAA.xml                  ║
║                                                                              ║
║  USO:                                                                        ║
║    # Detecção automática na pasta do script:                                 ║
║    python mirai_reports.py                                                   ║
║                                                                              ║
║    # Caminho explícito:                                                      ║
║    python mirai_reports.py --xml "C:\\Mirai\\Cronograma-07-07-2026.xml"     ║
║                                                                              ║
║    # Pasta de saída personalizada:                                           ║
║    python mirai_reports.py --xml arquivo.xml --out "C:\\Mirai\\Reports"     ║
║                                                                              ║
║    # Sem tarefas futuras no relatório de atrasos (padrão: inclui tudo):     ║
║    python mirai_reports.py --data-status 05/07/2026                         ║
║                                                                              ║
║  DEPENDÊNCIAS: apenas biblioteca padrão do Python (sem pip install)          ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import argparse
import json
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES — ajuste se necessário
# ══════════════════════════════════════════════════════════════════════════════

# Padrão de nome para busca automática na pasta
XML_PATTERN = "SRB_MIRAI_TP-0500_Cronograma*.xml"

# Workstreams principais do projeto
MAIN_FRENTES = {
    "01. Contabilidade", "02. Controladoria", "03. Finanças", "04. Tesouraria",
    "05. Comercial",     "06. Suprimentos",   "07. Logística", "08. Fiscal",
    "09. Produção",      "11. Qualidade",      "12. Manutenção", "14. Recursos Humanos",
    "15. Projetos",      "30. Interfaces",
}

FRENTES_ORDER = [
    "01. Contabilidade", "02. Controladoria", "03. Finanças",  "04. Tesouraria",
    "05. Comercial",     "06. Suprimentos",   "07. Logística", "08. Fiscal",
    "09. Produção",      "11. Qualidade",      "12. Manutenção", "14. Recursos Humanos",
    "15. Projetos",      "30. Interfaces",
]

F_SHORT = {
    "01. Contabilidade":"Contabilidade", "02. Controladoria":"Controladoria",
    "03. Finanças":"Finanças",           "04. Tesouraria":"Tesouraria",
    "05. Comercial":"Comercial",         "06. Suprimentos":"Suprimentos",
    "07. Logística":"Logística",         "08. Fiscal":"Fiscal",
    "09. Produção":"Produção",           "11. Qualidade":"Qualidade",
    "12. Manutenção":"Manutenção",       "14. Recursos Humanos":"RH",
    "15. Projetos":"Projetos",           "30. Interfaces":"Interfaces",
}

FR_COL = {
    "01. Contabilidade":"#854F0B", "02. Controladoria":"#b87a00",
    "03. Finanças":"#2a7098",      "04. Tesouraria":"#639922",
    "05. Comercial":"#EF9F27",     "06. Suprimentos":"#6b3fa0",
    "07. Logística":"#c8002a",     "08. Fiscal":"#3B6D11",
    "09. Produção":"#0077C8",      "11. Qualidade":"#0099b5",
    "12. Manutenção":"#629E16",    "14. Recursos Humanos":"#A32D2D",
    "15. Projetos":"#6e7a90",      "30. Interfaces":"#0044aa",
}

# IDs dos atributos estendidos do MS Project (confirmados no arquivo)
EA = {
    "FRENTE":    "188743734",   # Texto2  — Workstream
    "RESP":      "188743737",   # Texto3  — Responsável (Cast / Sumitomo)
    "FASE":      "188743748",   # Texto8  — Fase SAP Activate
    "SPI":       "188743988",   # Número12 — SPI Proj. Mirai
    "DEV_BL0":   "188743990",   # Número14 — Horas Desvio (vs BL0)
    "DEV_BL5":   "188743991",   # Número15 — Horas Desvio Recuperação (vs BL5)
    "STATUS":    "188744006",   # Texto20  — CAST-Status ('Atrasada', 'Concluída'…)
    "RESP_EX":   "188744015",   # Texto29  — Responsável Execução
}

GF  = ("https://fonts.googleapis.com/css2?"
       "family=IBM+Plex+Mono:wght@400;500"
       "&family=Barlow:wght@300;400;500;600;700;900"
       "&family=Barlow+Condensed:wght@600;700;900&display=swap")
CDN = "https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS DE PARSING
# ══════════════════════════════════════════════════════════════════════════════

def get_ea(task, fid, P):
    """Retorna valor de um ExtendedAttribute pelo FieldID."""
    for ea in task.findall(f"{P}ExtendedAttribute"):
        if ea.findtext(f"{P}FieldID") == fid:
            return ea.findtext(f"{P}Value") or ""
    return ""


def parse_dur(s):
    """Converte duração ISO do Project (PT8H0M0S) em horas float."""
    m = re.match(r"PT(\d+)H(\d+)M", s or "")
    return float(m.group(1)) + float(m.group(2)) / 60 if m else 0.0


def parse_dt(s):
    """Converte string de data ISO em datetime (ignora horário)."""
    if not s:
        return None
    try:
        return datetime.fromisoformat(s[:10])
    except ValueError:
        return None


def get_baseline(task, num, P):
    """Retorna (start, finish, work) de uma baseline específica."""
    for bl in task.findall(f"{P}Baseline"):
        if bl.findtext(f"{P}Number") == str(num):
            return (
                parse_dt(bl.findtext(f"{P}Start")),
                parse_dt(bl.findtext(f"{P}Finish")),
                parse_dur(bl.findtext(f"{P}Work")),
            )
    return (None, None, 0.0)


def classify_resp(s):
    """Classifica responsável como CAST, SRB ou AMBOS."""
    su = s.upper()
    if "CAST" in su and "SUMITOMO" not in su and "CLIENTE" not in su:
        return "CAST"
    if "SUMITOMO" in su or "CLIENTE" in su:
        return "SRB"
    if "CAST" in su:
        return "AMBOS"
    return "SRB"


def wk(dt):
    """Retorna chave de semana no formato 'AAAA-WSS'."""
    iso = dt.isocalendar()
    return f"{iso[0]}-W{iso[1]:02d}"


def pv_linear(start, finish, work, ref_date):
    """Planned Value (distribuição linear) de uma tarefa em ref_date."""
    if not start or not finish or work == 0:
        return 0.0
    if start > ref_date:
        return 0.0
    if finish <= ref_date:
        return work
    elapsed = (ref_date - start).days + 1
    total   = (finish - start).days + 1
    return work * min(elapsed / total, 1.0)


# ══════════════════════════════════════════════════════════════════════════════
# LEITURA DO XML
# ══════════════════════════════════════════════════════════════════════════════

def load_xml(xml_path, status_date):
    """
    Lê o XML do MS Project e retorna dicionários com:
      - leaf_tasks : todas as tarefas folha ativas com baselines
      - delay_tasks: tarefas com CAST-Status = 'Atrasada'
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()
    NS   = "http://schemas.microsoft.com/project"
    P    = f"{{{NS}}}"

    leaf_tasks  = []
    delay_tasks = []

    for t in root.findall(f"{P}Tasks/{P}Task"):
        # Filtros obrigatórios
        if t.findtext(f"{P}IsNull")  == "1": continue
        if t.findtext(f"{P}Summary") == "1": continue   # NUNCA tarefas resumo
        if t.findtext(f"{P}Active")  == "0": continue

        frente = get_ea(t, EA["FRENTE"], P)
        if frente not in MAIN_FRENTES:
            continue

        name    = t.findtext(f"{P}Name")                    or ""
        work    = parse_dur(t.findtext(f"{P}Work"))
        actual  = parse_dur(t.findtext(f"{P}ActualWork"))
        remain  = parse_dur(t.findtext(f"{P}RemainingWork"))
        pct     = float(t.findtext(f"{P}PercentWorkComplete") or 0) / 100
        finish  = parse_dt(t.findtext(f"{P}Finish"))
        start   = parse_dt(t.findtext(f"{P}Start"))
        resp    = get_ea(t, EA["RESP"],    P)
        resp_ex = get_ea(t, EA["RESP_EX"], P)
        status  = get_ea(t, EA["STATUS"],  P)
        spi_s   = get_ea(t, EA["SPI"],     P)

        # Campos de desvio pré-calculados pelo Project
        dev_bl0 = abs(float(get_ea(t, EA["DEV_BL0"], P) or 0))  # Número14
        dev_bl5 = abs(float(get_ea(t, EA["DEV_BL5"], P) or 0))  # Número15

        bl0s, bl0f, bl0w = get_baseline(t, 0, P)
        bl5s, bl5f, bl5w = get_baseline(t, 5, P)

        # Fallbacks: BL5 → BL0 → data corrente
        eff_bl0s = bl0s or start;  eff_bl0f = bl0f or finish; eff_bl0w = bl0w or work
        eff_bl5s = bl5s or bl0s or start
        eff_bl5f = bl5f or bl0f or finish
        eff_bl5w = bl5w or bl0w or work

        row = {
            "name":     name,
            "frente":   frente,
            "resp":     resp,
            "resp_ex":  resp_ex,
            "resp_cat": classify_resp(resp),
            "work":     work,
            "actual":   actual,
            "remain":   remain,
            "pct":      pct,
            "finish":   finish,
            "start":    start,
            "dev_bl0":  dev_bl0,    # Número14 — desvio vs BL0
            "dev_bl5":  dev_bl5,    # Número15 — desvio vs BL5
            "spi":      float(spi_s) if spi_s else None,
            "bl0s": eff_bl0s, "bl0f": eff_bl0f, "bl0w": eff_bl0w,
            "bl5s": eff_bl5s, "bl5f": eff_bl5f, "bl5w": eff_bl5w,
        }
        leaf_tasks.append(row)

        if status == "Atrasada":
            delay_tasks.append(row)

    print(f"  Tarefas folha carregadas : {len(leaf_tasks)}")
    print(f"  Tarefas com status Atrasada: {len(delay_tasks)}")
    return leaf_tasks, delay_tasks


# ══════════════════════════════════════════════════════════════════════════════
# CURVA S
# ══════════════════════════════════════════════════════════════════════════════

def build_scurve(leaf_tasks, status_date):
    """Constrói as séries semanais BL0, BL5 e EV para a Curva S."""
    pw_bl0 = defaultdict(float)
    pw_bl5 = defaultdict(float)
    ev_wk  = defaultdict(float)
    today_wk = wk(status_date)

    for t in leaf_tasks:
        # BL0
        s, f, w = t["bl0s"], t["bl0f"], t["bl0w"]
        if s and f and w:
            n   = max((f - s).days + 1, 1)
            cur = s
            while cur <= f:
                pw_bl0[wk(cur)] += w / n
                cur += timedelta(days=1)

        # BL5
        s, f, w = t["bl5s"], t["bl5f"], t["bl5w"]
        if s and f and w:
            n   = max((f - s).days + 1, 1)
            cur = s
            while cur <= f:
                pw_bl5[wk(cur)] += w / n
                cur += timedelta(days=1)

        # EV — coloca na semana do término (ou status_date, o que vier primeiro)
        if t["actual"] > 0 and t["finish"]:
            ev_wk[min(wk(t["finish"]), today_wk)] += t["actual"]

    all_wks = sorted(set(pw_bl0) | set(pw_bl5))
    sc_labels, sc_bl0, sc_bl5, sc_ev = [], [], [], []
    cum0 = cum5 = cumev = 0

    for w_key in all_wks:
        cum0  += pw_bl0.get(w_key, 0)
        cum5  += pw_bl5.get(w_key, 0)
        if w_key <= today_wk:
            cumev += ev_wk.get(w_key, 0)
        sc_labels.append("S" + w_key.split("-W")[1])
        sc_bl0.append(round(cum0, 1))
        sc_bl5.append(round(cum5, 1))
        sc_ev.append(round(cumev, 1) if w_key <= today_wk else None)

    return sc_labels, sc_bl0, sc_bl5, sc_ev


# ══════════════════════════════════════════════════════════════════════════════
# MÉTRICAS POR FRENTE
# ══════════════════════════════════════════════════════════════════════════════

def compute_frente_metrics(leaf_tasks, delay_tasks, status_date):
    """Calcula SPI vs BL5 e desvios (Nº14 / Nº15) por workstream."""
    # — Métricas gerenciais (SPI) —
    frente_spi = {}
    for fr in FRENTES_ORDER:
        sub = [t for t in leaf_tasks if t["frente"] == fr]
        if not sub:
            continue
        wt  = sum(t["work"]   for t in sub)
        wa  = sum(t["actual"] for t in sub)
        wr  = sum(t["remain"] for t in sub)
        w_bl5 = sum(t["bl5w"] for t in sub)
        w_bl0 = sum(t["bl0w"] for t in sub)
        pv = ev = 0.0
        for t in sub:
            pv += pv_linear(t["bl5s"], t["bl5f"], t["bl5w"], status_date)
            ev += t["actual"]
        spi = round(ev / pv, 2) if pv > 0 else 1.0
        del_ = sum(
            1 for t in sub
            if t["finish"] and t["finish"] < status_date and t["pct"] < 1.0
        )
        pct_bl5 = round(wa / w_bl5 * 100) if w_bl5 else 0
        pct_bl0 = round(wa / w_bl0 * 100) if w_bl0 else 0
        frente_spi[fr] = {
            "n": len(sub), "wt": round(wt, 1), "wa": round(wa, 1),
            "wr": round(wr, 1), "pct": round(wa / wt * 100) if wt else 0,
            "pv": round(pv, 1), "ev": round(ev, 1), "spi": spi, "del": del_,
            "w_bl5": round(w_bl5, 1), "w_bl0": round(w_bl0, 1),
            "pct_bl5": pct_bl5, "pct_bl0": pct_bl0,
        }

    # — Métricas de desvio (Nº14 / Nº15) — apenas tarefas atrasadas
    fr_delay = defaultdict(lambda: {
        "n_del":   0,   "dev_bl0":  0.0, "dev_bl5":  0.0,
        "cast_n":  0,   "srb_n":    0,
        "cast_bl5": 0.0,"srb_bl5":  0.0,
        "n_all":   0,   "w_actual": 0.0, "w_bl5_all": 0.0, "w_bl0_all": 0.0,
    })
    for t in leaf_tasks:
        fr = t["frente"]
        fr_delay[fr]["n_all"]    += 1
        fr_delay[fr]["w_actual"] += t["actual"]
        fr_delay[fr]["w_bl5_all"]+= t["bl5w"]
        fr_delay[fr]["w_bl0_all"]+= t["bl0w"]

    for t in delay_tasks:
        fr = t["frente"]; rc = t["resp_cat"]
        d5 = t["dev_bl5"]; d0 = t["dev_bl0"]
        fr_delay[fr]["n_del"]  += 1
        fr_delay[fr]["dev_bl5"]+= d5
        fr_delay[fr]["dev_bl0"]+= d0
        if rc == "CAST":
            fr_delay[fr]["cast_n"]   += 1
            fr_delay[fr]["cast_bl5"] += d5
        elif rc == "SRB":
            fr_delay[fr]["srb_n"]    += 1
            fr_delay[fr]["srb_bl5"]  += d5
        else:
            fr_delay[fr]["cast_n"]   += 1; fr_delay[fr]["srb_n"]   += 1
            fr_delay[fr]["cast_bl5"] += d5 / 2; fr_delay[fr]["srb_bl5"] += d5 / 2

    # — Overall —
    sub_all = leaf_tasks
    wt_all  = sum(t["work"]   for t in sub_all)
    wa_all  = sum(t["actual"] for t in sub_all)
    w_bl5_all = sum(t["bl5w"] for t in sub_all)
    w_bl0_all = sum(t["bl0w"] for t in sub_all)
    pv_all = sum(pv_linear(t["bl5s"], t["bl5f"], t["bl5w"], status_date) for t in sub_all)
    ev_all = sum(t["actual"] for t in sub_all)
    spi_all = round(ev_all / pv_all, 2) if pv_all > 0 else 1.0
    overall = {
        "wt": round(wt_all, 1), "wa": round(wa_all, 1),
        "pct": round(wa_all / wt_all * 100) if wt_all else 0,
        "pv": round(pv_all, 1), "ev": round(ev_all, 1), "spi": spi_all,
        "w_bl5": round(w_bl5_all, 1), "w_bl0": round(w_bl0_all, 1),
    }

    return frente_spi, dict(fr_delay), overall


# ══════════════════════════════════════════════════════════════════════════════
# FUNÇÕES DE ESTILO (helpers HTML)
# ══════════════════════════════════════════════════════════════════════════════

def spi_col(v):
    return "#A32D2D" if v < 0.8 else ("#854F0B" if v < 0.95 else "#3B6D11")

def spi_bg(v):
    return "#FCEBEB" if v < 0.8 else ("#FAEEDA" if v < 0.95 else "#EAF3DE")

def bar_col(p):
    return "#c8002a" if p < 30 else ("#b87a00" if p < 70 else "#007a52")

def risk_label(v):
    if v < 0.80: return "🔴 CRÍTICO",  "#A32D2D"
    if v < 0.95: return "⚠ ATENÇÃO",  "#854F0B"
    if v < 1.00: return "↗ MONITORAR","#3B6D11"
    return "✓ OK", "#3B6D11"


CSS_BASE = """
*{margin:0;padding:0;box-sizing:border-box}
body{background:#f5f6f8;color:#0d1117;font-family:'Barlow',sans-serif;font-size:13px;line-height:1.5}
.page{max-width:1080px;margin:0 auto;padding:30px}
.hdr{background:#fff;border:1px solid #e2e5eb;border-radius:8px;padding:20px 24px;
     margin-bottom:14px;display:flex;justify-content:space-between;align-items:flex-end}
.eye{font-family:'IBM Plex Mono',monospace;font-size:9px;letter-spacing:.12em;
     text-transform:uppercase;margin-bottom:5px}
.ttl{font-family:'Barlow Condensed',sans-serif;font-size:30px;font-weight:900;
     text-transform:uppercase;line-height:1}
.sub{font-family:'IBM Plex Mono',monospace;font-size:10px;color:#6e7a90;margin-top:4px}
.kpis{display:grid;gap:10px;margin-bottom:14px}
.kpi{background:#fff;border:1px solid #e2e5eb;border-radius:6px;padding:12px 14px;
     position:relative;overflow:hidden}
.kpi::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px}
.kpi.r::after{background:#E24B4A}.kpi.g::after{background:#639922}
.kpi.b::after{background:#0077C8}.kpi.a::after{background:#EF9F27}.kpi.p::after{background:#6b3fa0}
.kpi-lbl{font-size:9px;font-weight:600;color:#6e7a90;text-transform:uppercase;
          letter-spacing:.1em;font-family:'IBM Plex Mono',monospace}
.kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:28px;font-weight:900;margin-top:3px;line-height:1}
.kpi-sub{font-size:10px;color:#6e7a90;margin-top:2px;font-family:'IBM Plex Mono',monospace}
.card{background:#fff;border:1px solid #e2e5eb;border-radius:6px;padding:14px 16px}
.clbl{font-family:'IBM Plex Mono',monospace;font-size:9px;text-transform:uppercase;
      letter-spacing:.1em;color:#6e7a90;margin-bottom:12px;display:flex;align-items:center;gap:6px}
.cdot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.tbl{width:100%;border-collapse:collapse}
.tbl th{font-family:'IBM Plex Mono',monospace;font-size:9px;text-transform:uppercase;
        letter-spacing:.06em;color:#6e7a90;padding:0 9px 8px;text-align:left;
        border-bottom:2px solid #e2e5eb;font-weight:500}
.tbl th.bl5{color:#c8002a!important}.tbl th.bl0{color:#9aa3b5!important}
.tbl tr:hover td{background:#f9fafb}.tbl tr:last-child td{border-bottom:none}
.g2{display:grid;grid-template-columns:1.6fr 1fr;gap:12px;margin-bottom:12px}
.leg{display:flex;align-items:center;gap:5px;font-size:11px;color:#6e7a90}
.leg-sq{width:14px;height:3px;border-radius:2px;flex-shrink:0}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px}
.chip{font-size:10px;padding:3px 10px;border-radius:12px;border:1px solid #e2e5eb;
      cursor:pointer;background:#fff;font-family:monospace;font-weight:600;transition:.12s}
.chip.on{background:#c8002a;color:#fff;border-color:#c8002a}
.footer{margin-top:14px;padding-top:12px;border-top:1px solid #e2e5eb;
        display:flex;justify-content:space-between;font-size:10px;color:#6e7a90;font-family:monospace}
"""


# ══════════════════════════════════════════════════════════════════════════════
# REPORT 1 — GERENCIAL
# ══════════════════════════════════════════════════════════════════════════════

def generate_gerencial(sc_labels, sc_bl0, sc_bl5, sc_ev,
                       frente_spi, fr_delay, overall,
                       status_date, xml_filename, out_path):

    SD   = status_date.strftime("%d/%m/%Y")
    WN   = status_date.isocalendar()[1]
    spi_g= overall["spi"]
    ev_h = overall["ev"];  pv_h = overall["pv"]
    pct_g= overall["pct"]; sv   = ev_h - pv_h
    sc_g = spi_col(spi_g)
    crit = sum(1 for m in frente_spi.values() if m["spi"] < 0.95)
    w_bl0= overall["w_bl0"]; w_bl5 = overall["w_bl5"]

    status_bg = "#FCEBEB" if spi_g < 0.95 else ("#FAEEDA" if spi_g < 1.0 else "#EAF3DE")
    status_fc = "#A32D2D" if spi_g < 0.95 else ("#854F0B" if spi_g < 1.0 else "#3B6D11")
    status_bc = "#f5b8b8" if spi_g < 0.95 else ("#FAC775" if spi_g < 1.0 else "#b5d47a")
    overall_status = (
        "🔴 Projeto em atraso vs BL5" if spi_g < 0.95 else
        "⚠ Atenção vs BL5"           if spi_g < 1.0 else
        "✓ No controle"
    )

    # Workstreams ordenados pior→melhor SPI vs BL5
    ws_sorted = sorted(
        [fr for fr in FRENTES_ORDER if fr in frente_spi],
        key=lambda f: frente_spi[f]["spi"]
    )

    # Gráfico de desvio por frente
    sorted_fr = sorted(fr_delay, key=lambda f: -fr_delay[f].get("dev_bl5", 0))
    chart_fr_lbl = [F_SHORT.get(f, f) for f in sorted_fr if fr_delay[f].get("n_del", 0) > 0]
    chart_d5 = [round(fr_delay[f]["dev_bl5"], 1) for f in sorted_fr if fr_delay[f].get("n_del", 0) > 0]
    chart_d0 = [round(fr_delay[f]["dev_bl0"], 1) for f in sorted_fr if fr_delay[f].get("n_del", 0) > 0]

    # Linhas da tabela de workstreams
    ws_rows = ""
    for fr in ws_sorted:
        m  = frente_spi[fr];  sp = m["spi"]
        sc = spi_col(sp);  sb = spi_bg(sp)
        rl, rc = risk_label(sp)
        fs = F_SHORT.get(fr, fr)
        fd = fr_delay.get(fr, {})
        d5 = fd.get("dev_bl5", 0); d0 = fd.get("dev_bl0", 0)
        p5 = m["pct_bl5"]; p0 = m["pct_bl0"]
        del_txt = (f'<span style="color:#c8002a;font-weight:600">{m["del"]}</span>'
                   if m["del"] else '<span style="color:#6e7a90">—</span>')
        ws_rows += f"""
        <tr style="{'background:#fffbfb' if sp < 0.95 else ''}">
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;font-size:12px;font-weight:500">{fs}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;min-width:170px">
            <div style="margin-bottom:4px">
              <div style="display:flex;justify-content:space-between;font-size:9px;font-family:monospace;color:#6e7a90;margin-bottom:2px">
                <span>BL5</span><span style="color:{bar_col(p5)};font-weight:600">{p5}%</span>
              </div>
              <div style="height:6px;background:#e2e5eb;border-radius:3px;overflow:hidden">
                <div style="height:100%;width:{min(p5,100)}%;background:{bar_col(p5)};border-radius:3px"></div>
              </div>
            </div>
            <div>
              <div style="display:flex;justify-content:space-between;font-size:9px;font-family:monospace;color:#6e7a90;margin-bottom:2px">
                <span>BL0</span><span>{p0}%</span>
              </div>
              <div style="height:4px;background:#e2e5eb;border-radius:2px;overflow:hidden">
                <div style="height:100%;width:{min(p0,100)}%;background:#9aa3b5;border-radius:2px"></div>
              </div>
            </div>
          </td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;font-size:11px;font-family:monospace">{m['wa']:.0f}h / {m['w_bl5']:.0f}h</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;text-align:right;font-size:12px;font-family:monospace;font-weight:700;color:#c8002a">{d5:.0f}h</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;text-align:right;font-size:11px;font-family:monospace;color:#9aa3b5">{d0:.0f}h</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;text-align:center">{del_txt}</td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;text-align:center">
            <span style="font-size:11px;font-weight:700;padding:3px 10px;border-radius:3px;background:{sb};color:{sc}">{sp}</span>
          </td>
          <td style="padding:8px 10px;border-bottom:1px solid #e8edf5;font-size:10px;font-weight:600;color:{rc}">{rl}</td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Report Gerencial — Projeto Mirai · {SD}</title>
<link href="{GF}" rel="stylesheet">
<style>
{CSS_BASE}
.kpis{{grid-template-columns:repeat(5,1fr)}}
</style>
</head><body><div class="page">

<div class="hdr">
  <div>
    <div class="eye" style="color:#0077C8">Sumitomo Rubber do Brasil · SAP S/4HANA · BL0 = Plano Original · BL5 = Plano de Recuperação</div>
    <div class="ttl" style="color:#0d1117">Projeto <span style="color:#0077C8">Mirai</span></div>
    <div class="sub">Report Gerencial · SPI vs BL5 · Dev Nº14 (BL0) e Nº15 (BL5) · Data Status: {SD} · Sem tarefas resumo</div>
  </div>
  <div style="text-align:right">
    <div style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:#6e7a90">Semana {WN} · {status_date.year}</div>
    <div style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:#6e7a90;margin-top:2px">{xml_filename}</div>
    <div style="margin-top:8px;display:inline-block;font-size:10px;font-weight:700;
         padding:4px 14px;border-radius:3px;background:{status_bg};color:{status_fc};border:1px solid {status_bc}">{overall_status}</div>
  </div>
</div>

<div class="kpis">
  <div class="kpi b"><div class="kpi-lbl">Trabalho BL0</div>
    <div class="kpi-val">{w_bl0/1000:.0f}K</div><div class="kpi-sub">{w_bl0:.0f}h plano original</div></div>
  <div class="kpi p"><div class="kpi-lbl">Trabalho BL5</div>
    <div class="kpi-val">{w_bl5/1000:.0f}K</div><div class="kpi-sub">{w_bl5:.0f}h plano recuperação</div></div>
  <div class="kpi g"><div class="kpi-lbl">EV realizado</div>
    <div class="kpi-val">{pct_g}%</div><div class="kpi-sub">{ev_h:.0f}h entregues</div></div>
  <div class="kpi {'r' if spi_g<0.95 else 'a' if spi_g<1.0 else 'g'}">
    <div class="kpi-lbl">SPI vs BL5</div>
    <div class="kpi-val" style="color:{sc_g}">{spi_g}</div>
    <div class="kpi-sub">SV {sv:+.0f}h vs BL5</div></div>
  <div class="kpi r"><div class="kpi-lbl">Críticos BL5</div>
    <div class="kpi-val">{crit}</div><div class="kpi-sub">SPI &lt; 0.95</div></div>
</div>

<div class="g2" style="margin-bottom:12px">
  <div class="card">
    <div class="clbl"><span class="cdot" style="background:#0077C8"></span>Curva S — BL0 Original · BL5 Recuperação · EV Realizado</div>
    <div style="display:flex;gap:14px;margin-bottom:10px;flex-wrap:wrap">
      <div class="leg"><span class="leg-sq" style="background:#c0c8d8"></span>BL0 Original</div>
      <div class="leg"><span class="leg-sq" style="background:#0077C8"></span>BL5 Recuperação</div>
      <div class="leg"><span class="leg-sq" style="background:#639922"></span>EV Realizado</div>
    </div>
    <div style="position:relative;height:210px"><canvas id="curvaS"></canvas></div>
  </div>
  <div class="card">
    <div class="clbl"><span class="cdot" style="background:#c8002a"></span>Desvio por Frente — Nº15 (BL5) e Nº14 (BL0)</div>
    <div style="display:flex;gap:10px;margin-bottom:8px;font-size:10px;color:#6e7a90">
      <span style="display:flex;align-items:center;gap:4px">
        <span style="width:10px;height:10px;background:#c8002a;border-radius:2px;display:inline-block"></span>Dev BL5 (Nº15)
      </span>
      <span style="display:flex;align-items:center;gap:4px">
        <span style="width:10px;height:10px;background:#9aa3b5;border-radius:2px;display:inline-block"></span>Dev BL0 (Nº14)
      </span>
    </div>
    <div style="position:relative;height:220px"><canvas id="barDev"></canvas></div>
  </div>
</div>

<div class="card">
  <div class="clbl"><span class="cdot" style="background:#0077C8"></span>Avanço por Workstream — pior→melhor SPI vs BL5 · % = trabalho realizado / trabalho BL5</div>
  <table class="tbl">
    <thead><tr>
      <th style="width:140px">Workstream</th>
      <th style="min-width:170px">% Conclusão (BL5 / BL0)</th>
      <th>EV / BL5</th>
      <th class="bl5" style="text-align:right;width:90px">Dev BL5 (Nº15)</th>
      <th class="bl0" style="text-align:right;width:90px">Dev BL0 (Nº14)</th>
      <th style="text-align:center;width:65px">Atrasos</th>
      <th style="text-align:center;width:85px">SPI BL5</th>
      <th style="width:110px">Status</th>
    </tr></thead>
    <tbody>{ws_rows}</tbody>
  </table>
  <div style="margin-top:8px;font-size:10px;font-family:monospace;color:#6e7a90">
    <span style="color:#c8002a">■ Dev BL5 (Nº15)</span> = campo Número15 do MS Project (desvio vs plano de recuperação) &nbsp;·&nbsp;
    <span style="color:#9aa3b5">■ Dev BL0 (Nº14)</span> = campo Número14 do MS Project (desvio vs plano original) &nbsp;·&nbsp;
    BL0 e BL5 são paralelas e independentes
  </div>
</div>

<div class="footer">
  <span>Fonte: {xml_filename} · Nº14=Dev BL0 · Nº15=Dev BL5 · sem tarefas resumo · {SD}</span>
  <span style="color:#0077C8;font-weight:600">Projeto Mirai · Sumitomo Rubber do Brasil</span>
</div>
</div>

<script src="{CDN}"></script>
<script>
(function(){{
  const gc='rgba(0,0,0,.04)', tc='#6e7a90';
  new Chart(document.getElementById('curvaS'), {{
    type: 'line',
    data: {{ labels: {json.dumps(sc_labels)}, datasets: [
      {{ label:'BL0', data:{json.dumps(sc_bl0)}, borderColor:'#c0c8d8', borderDash:[2,3], borderWidth:1.5, fill:false, tension:.4, pointRadius:0 }},
      {{ label:'BL5', data:{json.dumps(sc_bl5)}, borderColor:'#0077C8', backgroundColor:'rgba(0,119,200,.07)', borderWidth:2.5, fill:true, tension:.4, pointRadius:0 }},
      {{ label:'EV',  data:{json.dumps(sc_ev)},  borderColor:'#639922', borderDash:[5,4], borderWidth:2.5, fill:false, tension:.4, pointRadius:3, pointBackgroundColor:'#639922', spanGaps:false }}
    ]}},
    options: {{ responsive:true, maintainAspectRatio:false,
      plugins: {{ legend:{{display:false}}, tooltip:{{ mode:'index', intersect:false,
        callbacks:{{ label: c => ` ${{c.dataset.label}}: ${{c.parsed.y?.toLocaleString('pt-BR')}}h` }} }} }},
      scales: {{
        x: {{ grid:{{color:gc}}, ticks:{{color:tc, font:{{size:9}}, maxRotation:0, autoSkip:true, maxTicksLimit:20}} }},
        y: {{ grid:{{color:gc}}, ticks:{{color:tc, font:{{size:9}}, callback: v => v>=1000 ? Math.round(v/1000)+'Kh' : v+'h' }}, beginAtZero:true }}
      }}
    }}
  }});
  new Chart(document.getElementById('barDev'), {{
    type: 'bar',
    data: {{ labels: {json.dumps(chart_fr_lbl)}, datasets: [
      {{ label:'Dev BL5 (Nº15)', data:{json.dumps(chart_d5)}, backgroundColor:'rgba(200,0,42,.80)', borderRadius:3, barPercentage:.65 }},
      {{ label:'Dev BL0 (Nº14)', data:{json.dumps(chart_d0)}, backgroundColor:'rgba(154,163,181,.55)', borderRadius:3, barPercentage:.65 }}
    ]}},
    options: {{ responsive:true, maintainAspectRatio:false,
      plugins: {{ legend:{{display:false}}, tooltip:{{ callbacks:{{ label: c => ` ${{c.dataset.label}}: ${{c.parsed.y}}h` }} }} }},
      scales: {{
        x: {{ grid:{{display:false}}, ticks:{{color:tc, font:{{size:8}}, maxRotation:35}} }},
        y: {{ grid:{{color:gc}}, ticks:{{color:tc, font:{{size:8}}, callback: v => v+'h' }}, beginAtZero:true }}
      }}
    }}
  }});
}})();
</script>
</body></html>"""

    out_path.write_text(html, encoding="utf-8")
    print(f"  ✓ {out_path.name}")


# ══════════════════════════════════════════════════════════════════════════════
# REPORT 2 — ATRASOS
# ══════════════════════════════════════════════════════════════════════════════

def generate_atrasos(delay_tasks, fr_delay, status_date,
                     next_sunday, xml_filename, out_path):

    SD  = status_date.strftime("%d/%m/%Y")
    SUN = next_sunday.strftime("%d/%m/%Y")
    WN  = status_date.isocalendar()[1]

    tot      = len(delay_tasks)
    tot_d5   = sum(t["dev_bl5"] for t in delay_tasks)
    tot_d0   = sum(t["dev_bl0"] for t in delay_tasks)
    h_cast   = sum(fr_delay[f]["cast_bl5"] for f in fr_delay)
    h_srb    = sum(fr_delay[f]["srb_bl5"]  for f in fr_delay)
    n_cast   = sum(fr_delay[f]["cast_n"]   for f in fr_delay)
    n_srb    = sum(fr_delay[f]["srb_n"]    for f in fr_delay)

    sorted_fr = sorted(fr_delay, key=lambda f: -fr_delay[f].get("dev_bl5", 0))
    chart_labels = [F_SHORT.get(f, f) for f in sorted_fr if fr_delay[f].get("n_del", 0) > 0]
    chart_d5     = [round(fr_delay[f]["dev_bl5"], 1) for f in sorted_fr if fr_delay[f].get("n_del", 0) > 0]
    chart_d0     = [round(fr_delay[f]["dev_bl0"], 1) for f in sorted_fr if fr_delay[f].get("n_del", 0) > 0]

    RESP_BG = {"CAST":"rgba(0,119,200,.10)","SRB":"rgba(99,153,34,.10)","AMBOS":"rgba(239,159,39,.10)"}
    RESP_FC = {"CAST":"#0077C8","SRB":"#3B6D11","AMBOS":"#b87a00"}

    # ── Tabela de tarefas: ordenada da mais antiga ────────────────────────────
    task_rows = ""
    for t in sorted(delay_tasks, key=lambda x: (x["finish"] is None, x["finish"] or datetime.max)):
        fr  = t["frente"]; rc = t["resp_cat"]
        fc  = FR_COL.get(fr, "#0077C8"); fs = F_SHORT.get(fr, fr)
        bg  = RESP_BG.get(rc, "#f5f6f8"); rfc = RESP_FC.get(rc, "#6e7a90")
        nm  = t["name"][:65] + "…" if len(t["name"]) > 65 else t["name"]
        pct = round(t["pct"] * 100)
        pct_c = "#c8002a" if pct < 30 else ("#EF9F27" if pct < 80 else "#639922")
        term  = t["finish"].strftime("%d/%m/%y") if t["finish"] else "—"
        rex   = t["resp_ex"][:22] + "…" if len(t["resp_ex"]) > 22 else t["resp_ex"]
        d5 = t["dev_bl5"]; d0 = t["dev_bl0"]
        if t["finish"]:
            days = (status_date - t["finish"]).days
            late_lbl = f"+{days}d" if days > 0 else ("hoje" if days == 0 else "futuro")
            late_col = "#A32D2D" if days > 30 else ("#854F0B" if days > 7 else ("#EF9F27" if days >= 0 else "#6b3fa0"))
        else:
            late_lbl = "—"; late_col = "#6e7a90"

        task_rows += f"""
        <tr data-fr="{fs}" data-rc="{rc}">
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;white-space:nowrap">
            <span style="font-size:9px;padding:1px 5px;border-radius:3px;background:{fc}20;color:{fc};font-weight:700">{fs}</span>
          </td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;font-size:11px;max-width:290px">{nm}</td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;text-align:center">
            <span style="font-size:9px;padding:1px 6px;border-radius:3px;background:{bg};color:{rfc};font-weight:700">{rc}</span>
          </td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;font-family:monospace;font-size:10px;text-align:center;color:#c8002a;font-weight:600;white-space:nowrap">{term}</td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;font-family:monospace;font-size:10px;text-align:center;color:{late_col};font-weight:700">{late_lbl}</td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;min-width:82px">
            <div style="display:flex;align-items:center;gap:4px">
              <div style="flex:1;height:5px;background:#e2e5eb;border-radius:3px;overflow:hidden">
                <div style="height:100%;width:{pct}%;background:{pct_c};border-radius:3px"></div>
              </div>
              <span style="font-size:9px;font-family:monospace;color:{pct_c};min-width:24px">{pct}%</span>
            </div>
          </td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;text-align:right;font-size:11px;font-family:monospace;color:#c8002a;font-weight:700;white-space:nowrap">{d5:.1f}h</td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;text-align:right;font-size:11px;font-family:monospace;color:#9aa3b5;white-space:nowrap">{d0:.1f}h</td>
          <td style="padding:5px 7px;border-bottom:1px solid #f0f2f5;font-size:10px;color:#6e7a90">{rex}</td>
        </tr>"""

    # ── Tabela de resumo por workstream ───────────────────────────────────────
    fr_rows = ""
    for fr in sorted_fr:
        fd = fr_delay.get(fr, {})
        if not fd or fd.get("n_del", 0) == 0:
            continue
        fc = FR_COL.get(fr, "#0077C8"); fs = F_SHORT.get(fr, fr)
        d5 = fd["dev_bl5"]; d0 = fd["dev_bl0"]
        wa = fd.get("w_actual", 0); wb5 = fd.get("w_bl5_all", 0)
        pct5 = round(wa / wb5 * 100) if wb5 else 0
        hc5  = fd["cast_bl5"]; hs5 = fd["srb_bl5"]
        cw = round(hc5 / d5 * 100) if d5 else 0
        sw = round(hs5 / d5 * 100) if d5 else 0
        risk   = "🔴 CRÍTICO" if d5 > 400 else ("⚠ ATENÇÃO" if d5 > 100 else "● MONITORAR")
        risk_c = "#A32D2D" if "🔴" in risk else ("#854F0B" if "⚠" in risk else "#3B6D11")
        pct5_col = bar_col(pct5)

        fr_rows += f"""
        <tr style="border-left:3px solid {fc}">
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;font-size:12px;font-weight:600;color:{fc}">{fs}</td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;text-align:center;font-family:monospace;font-size:12px;font-weight:700;color:#c8002a">{fd['n_del']}</td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;min-width:115px">
            <div style="display:flex;align-items:center;gap:5px">
              <div style="flex:1;height:7px;background:#e2e5eb;border-radius:3px;overflow:hidden">
                <div style="height:100%;width:{min(pct5,100)}%;background:{pct5_col};border-radius:3px"></div>
              </div>
              <span style="font-size:10px;font-family:monospace;font-weight:700;color:{pct5_col};min-width:28px">{pct5}%</span>
            </div>
          </td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;text-align:right;font-family:monospace;font-size:12px;font-weight:700;color:#c8002a">{d5:.0f}h</td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;text-align:right;font-family:monospace;font-size:11px;color:#9aa3b5">{d0:.0f}h</td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;min-width:120px">
            <div style="display:flex;height:14px;border-radius:3px;overflow:hidden">
              <div style="width:{cw}%;background:#0077C8;display:flex;align-items:center;justify-content:center">
                {"<span style='font-size:8px;color:#fff;font-weight:700'>" + str(round(hc5)) + "h</span>" if cw > 15 else ""}
              </div>
              <div style="width:{sw}%;background:#639922;display:flex;align-items:center;justify-content:center">
                {"<span style='font-size:8px;color:#fff;font-weight:700'>" + str(round(hs5)) + "h</span>" if sw > 15 else ""}
              </div>
              <div style="flex:1;background:#e2e5eb"></div>
            </div>
          </td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;font-size:11px;font-family:monospace;color:#0077C8;text-align:right">{hc5:.0f}h ({fd['cast_n']}t)</td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;font-size:11px;font-family:monospace;color:#3B6D11;text-align:right">{hs5:.0f}h ({fd['srb_n']}t)</td>
          <td style="padding:8px 9px;border-bottom:1px solid #e8edf5;font-size:10px;font-weight:700;color:{risk_c}">{risk}</td>
        </tr>"""

    chip_fr = "".join(
        f'<span class="chip" onclick="filterBy(&quot;{F_SHORT.get(f,f)}&quot;,this)" '
        f'style="color:{FR_COL.get(f,"#0077C8")};border-color:{FR_COL.get(f,"#e2e5eb")}">'
        f'{F_SHORT.get(f,f)} ({fr_delay[f]["n_del"]})</span>'
        for f in sorted_fr if fr_delay.get(f, {}).get("n_del", 0) > 0
    )

    html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Desvio vs BL5 — Projeto Mirai · {SD}</title>
<link href="{GF}" rel="stylesheet">
<style>
{CSS_BASE}
.kpis{{grid-template-columns:repeat(5,1fr)}}
.proj-box{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:12px;
           background:linear-gradient(135deg,#f5f0ff,#fdf8ff);
           border:1px solid #d4c6f4;border-radius:8px;padding:14px 16px}}
.pj{{text-align:center}}
.pj-l{{font-size:9px;font-family:monospace;text-transform:uppercase;letter-spacing:.08em;color:#6b3fa0;font-weight:600}}
.pj-v{{font-family:'Barlow Condensed',sans-serif;font-size:26px;font-weight:900;line-height:1.1;margin-top:2px}}
.pj-s{{font-size:10px;font-family:monospace;color:#6e7a90;margin-top:2px}}
.divider{{border-right:1px solid #d4c6f4}}
</style>
</head><body><div class="page" style="max-width:1160px">

<div class="hdr">
  <div>
    <div class="eye" style="color:#c8002a">Projeto Mirai · Desvio vs BL5 (Nº15) e BL0 (Nº14) · Sem tarefas resumo</div>
    <div class="ttl" style="color:#0d1117">Desvio — <span style="color:#c8002a">BL5</span> vs BL0</div>
    <div class="sub">BL0 = plano original · BL5 = plano de recuperação · Paralelas e independentes · Medição: {SD}</div>
  </div>
  <div style="text-align:right">
    <div style="font-family:monospace;font-size:11px;color:#6e7a90">Sem. {WN} · {status_date.year}</div>
    <div style="margin-top:8px;display:inline-block;font-size:10px;font-weight:700;
         padding:4px 16px;border-radius:4px;background:#FCEBEB;color:#c8002a;border:1px solid #f5b8b8">
      🔴 {tot} tarefas · Dev BL5: {tot_d5:.0f}h · Dev BL0: {tot_d0:.0f}h
    </div>
  </div>
</div>

<div class="kpis">
  <div class="kpi" style="border-top-color:#c8002a">
    <div class="kpi-lbl">Tarefas atrasadas</div>
    <div class="kpi-val" style="color:#c8002a">{tot}</div>
    <div class="kpi-sub">status Atrasada · sem resumo</div></div>
  <div class="kpi" style="border-top-color:#c8002a">
    <div class="kpi-lbl">Dev BL5 (Nº15)</div>
    <div class="kpi-val" style="color:#c8002a">{tot_d5:.0f}h</div>
    <div class="kpi-sub">vs plano de recuperação</div></div>
  <div class="kpi" style="border-top-color:#9aa3b5">
    <div class="kpi-lbl">Dev BL0 (Nº14)</div>
    <div class="kpi-val" style="color:#9aa3b5">{tot_d0:.0f}h</div>
    <div class="kpi-sub">vs plano original</div></div>
  <div class="kpi" style="border-top-color:#0077C8">
    <div class="kpi-lbl">CAST · Dev BL5</div>
    <div class="kpi-val" style="color:#0077C8">{h_cast:.0f}h</div>
    <div class="kpi-sub">{n_cast} tarefas</div></div>
  <div class="kpi" style="border-top-color:#639922">
    <div class="kpi-lbl">SRB · Dev BL5</div>
    <div class="kpi-val" style="color:#639922">{h_srb:.0f}h</div>
    <div class="kpi-sub">{n_srb} tarefas</div></div>
</div>

<div class="proj-box">
  <div class="pj divider">
    <div class="pj-l">Dev BL5 hoje ({SD})</div>
    <div class="pj-v" style="color:#c8002a">{tot_d5:.0f}h</div>
    <div class="pj-s">campo Nº15 · {tot} tarefas atrasadas</div>
  </div>
  <div class="pj divider">
    <div class="pj-l">Dev BL0 hoje ({SD})</div>
    <div class="pj-v" style="color:#9aa3b5">{tot_d0:.0f}h</div>
    <div class="pj-s">campo Nº14 · referência plano original</div>
  </div>
  <div class="pj">
    <div class="pj-l">Projeção Dom {SUN}</div>
    <div class="pj-v" style="color:#6b3fa0">{tot_d5:.0f}h+</div>
    <div class="pj-s">tarefas com BL5 fim até Dom adicionam desvio</div>
  </div>
</div>

<div class="g2">
  <div class="card">
    <div class="clbl"><span class="cdot" style="background:#c8002a"></span>Desvio por Workstream — BL5 (Nº15) vs BL0 (Nº14)</div>
    <div style="display:flex;gap:12px;margin-bottom:10px;flex-wrap:wrap">
      <span style="display:flex;align-items:center;gap:5px;font-size:11px;color:#6e7a90">
        <span style="width:12px;height:12px;background:#c8002a;border-radius:2px;display:inline-block"></span>Dev BL5 (Nº15)</span>
      <span style="display:flex;align-items:center;gap:5px;font-size:11px;color:#6e7a90">
        <span style="width:12px;height:12px;background:#9aa3b5;border-radius:2px;display:inline-block"></span>Dev BL0 (Nº14)</span>
    </div>
    <div style="position:relative;height:265px"><canvas id="barWs"></canvas></div>
  </div>
  <div class="card">
    <div class="clbl"><span class="cdot" style="background:#c8002a"></span>CAST vs SRB — Dev BL5 (Nº15)</div>
    <div style="position:relative;height:140px;margin-bottom:14px"><canvas id="donutResp"></canvas></div>
    <div style="padding-top:12px;border-top:1px solid #e2e5eb">
      <div style="font-size:9px;font-family:monospace;text-transform:uppercase;color:#6e7a90;letter-spacing:.06em;margin-bottom:8px">Ranking Dev BL5</div>
      {"".join(f'''<div style="display:flex;align-items:center;gap:7px;margin-bottom:5px">
        <div style="width:7px;height:7px;border-radius:2px;background:{FR_COL.get(f,"#0077C8")};flex-shrink:0"></div>
        <div style="flex:1;font-size:11px">{F_SHORT.get(f,f)}</div>
        <div style="font-size:11px;font-family:monospace;font-weight:700;color:#c8002a">{fr_delay[f]["dev_bl5"]:.0f}h</div>
        <div style="font-size:10px;font-family:monospace;color:#9aa3b5">BL0:{fr_delay[f]["dev_bl0"]:.0f}h</div>
      </div>''' for f in sorted_fr[:7] if fr_delay.get(f, {}).get("n_del", 0) > 0)}
    </div>
  </div>
</div>

<div class="card" style="margin-bottom:12px">
  <div class="clbl"><span class="cdot" style="background:#c8002a"></span>Resumo por Workstream · Dev BL5 (Nº15) · Dev BL0 (Nº14) · % Conclusão · CAST/SRB</div>
  <div style="overflow-x:auto">
    <table class="tbl">
      <thead><tr>
        <th style="text-align:left;width:110px">Workstream</th>
        <th style="text-align:center;width:45px">Qtd</th>
        <th style="min-width:115px">% Conclusão BL5</th>
        <th class="bl5" style="text-align:right">Dev BL5 (Nº15)</th>
        <th class="bl0" style="text-align:right">Dev BL0 (Nº14)</th>
        <th style="text-align:left;min-width:110px">CAST vs SRB</th>
        <th style="text-align:right;color:#0077C8">CAST</th>
        <th style="text-align:right;color:#3B6D11">SRB</th>
        <th>Risco</th>
      </tr></thead>
      <tbody>{fr_rows}</tbody>
    </table>
  </div>
  <div style="margin-top:8px;font-size:10px;font-family:monospace;color:#6e7a90">
    % Conclusão = trabalho realizado / trabalho BL5 &nbsp;·&nbsp;
    <span style="color:#c8002a">Dev BL5 (Nº15)</span> = campo Número15 &nbsp;·&nbsp;
    <span style="color:#9aa3b5">Dev BL0 (Nº14)</span> = campo Número14 &nbsp;·&nbsp;
    BL0 e BL5 são paralelas e independentes
  </div>
</div>

<div class="card">
  <div class="clbl"><span class="cdot" style="background:#c8002a"></span>Detalhamento — {tot} tarefas · sem resumo · ordenadas da mais antiga para a mais recente</div>
  <div class="chips">
    <span class="chip on" onclick="filterBy('ALL',this)">Todos ({tot})</span>
    <span class="chip" onclick="filterBy('CAST',this)" style="color:#0077C8;border-color:#b5d4f4">CAST ({n_cast})</span>
    <span class="chip" onclick="filterBy('SRB',this)"  style="color:#3B6D11;border-color:#b5d4aa">SRB ({n_srb})</span>
    {chip_fr}
  </div>
  <div style="overflow-x:auto;max-height:600px;overflow-y:auto">
    <table class="tbl" id="tblAtr">
      <thead style="position:sticky;top:0;background:#fff;z-index:1">
        <tr>
          <th style="text-align:left;width:80px">WS</th>
          <th style="text-align:left">Atividade</th>
          <th style="width:60px;text-align:center">Resp.</th>
          <th style="width:68px;text-align:center">Término ↑</th>
          <th style="width:50px;text-align:center">Dias</th>
          <th style="width:85px">Progresso</th>
          <th class="bl5" style="width:62px;text-align:right">Dev BL5</th>
          <th class="bl0" style="width:62px;text-align:right">Dev BL0</th>
          <th style="text-align:left">Executor</th>
        </tr>
      </thead>
      <tbody id="tbodyAtr">{task_rows}</tbody>
    </table>
  </div>
</div>

<div class="footer">
  <span>Fonte: {xml_filename} · Nº14=Dev BL0 · Nº15=Dev BL5 · BL0 e BL5 paralelas · sem tarefas resumo · {SD}</span>
  <span style="color:#c8002a;font-weight:600">Projeto Mirai · Sumitomo Rubber do Brasil</span>
</div>
</div>

<script src="{CDN}"></script>
<script>
(function(){{
  const gc='rgba(0,0,0,.04)', tc='#6e7a90';
  new Chart(document.getElementById('barWs'), {{
    type: 'bar',
    data: {{ labels: {json.dumps(chart_labels)}, datasets: [
      {{ label:'Dev BL5 (Nº15)', data:{json.dumps(chart_d5)}, backgroundColor:'rgba(200,0,42,.80)', borderRadius:3, barPercentage:.65 }},
      {{ label:'Dev BL0 (Nº14)', data:{json.dumps(chart_d0)}, backgroundColor:'rgba(154,163,181,.55)', borderRadius:3, barPercentage:.65 }}
    ]}},
    options: {{ responsive:true, maintainAspectRatio:false,
      plugins: {{ legend:{{position:'top',align:'end',labels:{{font:{{size:10}},padding:10,boxWidth:10}}}},
        tooltip:{{ callbacks:{{ label: c => ` ${{c.dataset.label}}: ${{c.parsed.y}}h` }} }} }},
      scales: {{
        x: {{ grid:{{display:false}}, ticks:{{color:tc, font:{{size:9}}, maxRotation:35}} }},
        y: {{ grid:{{color:gc}}, ticks:{{color:tc, font:{{size:9}}, callback: v => v+'h' }}, beginAtZero:true }}
      }}
    }}
  }});
  const hCast={round(h_cast,1)}, hSrb={round(h_srb,1)};
  new Chart(document.getElementById('donutResp'), {{
    type: 'doughnut',
    data: {{ labels:['CAST','SRB/Sumitomo'], datasets:[{{
      data:[hCast,hSrb], backgroundColor:['#0077C8','#639922'], borderWidth:0, hoverOffset:4
    }}]}},
    options: {{ responsive:true, maintainAspectRatio:false, cutout:'62%',
      plugins: {{ legend:{{ position:'right', labels:{{ font:{{size:10}}, padding:8,
        generateLabels: ch => ch.data.labels.map((l,i) => {{
          const v=ch.data.datasets[0].data[i], tot=hCast+hSrb;
          return {{ text:`${{l}}: ${{v}}h (${{Math.round(v/tot*100)}}%)`,
                   fillStyle:ch.data.datasets[0].backgroundColor[i], fontColor:tc }};
        }})
      }} }}, tooltip:{{ callbacks:{{ label: c => ` ${{c.label}}: ${{c.parsed}}h` }} }} }}
    }}
  }});
}})();

function filterBy(key, el) {{
  document.querySelectorAll('.chip').forEach(c => c.classList.remove('on'));
  el.classList.add('on');
  document.querySelectorAll('#tbodyAtr tr').forEach(r => {{
    const fr = r.dataset.fr || '', rc = r.dataset.rc || '';
    if (key === 'ALL') r.style.display = '';
    else if (key === 'CAST' || key === 'SRB') r.style.display = rc === key ? '' : 'none';
    else r.style.display = fr === key ? '' : 'none';
  }});
}}
</script>
</body></html>"""

    out_path.write_text(html, encoding="utf-8")
    print(f"  ✓ {out_path.name}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def find_xml(search_dir):
    """Busca o XML mais recente na pasta pelo padrão de nome."""
    candidates = sorted(search_dir.glob(XML_PATTERN))
    if not candidates:
        return None
    # Ordena pelo nome (que termina com a data) — o último é o mais recente
    return candidates[-1]


def read_status_date(xml_path):
    """Lê a data de status do XML; se não existir, retorna a data atual."""
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        NS = "http://schemas.microsoft.com/project"
        sd_s = root.findtext(f"{{{NS}}}StatusDate") or ""
        return datetime.fromisoformat(sd_s[:10]) if sd_s else datetime.today()
    except Exception:
        return datetime.today()


def next_sunday(ref_date):
    """Retorna o próximo domingo a partir de ref_date (inclusive se já for dom)."""
    days_ahead = (6 - ref_date.weekday()) % 7   # 6 = sunday
    return ref_date + timedelta(days=days_ahead if days_ahead else 7)


def main():
    parser = argparse.ArgumentParser(
        description="Gerador de relatórios gerenciais — Projeto Mirai",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  python mirai_reports.py
  python mirai_reports.py --xml "C:\\Mirai\\SRB_MIRAI_TP-0500_Cronograma-07-07-2026.xml"
  python mirai_reports.py --xml cronograma.xml --out "C:\\Mirai\\Reports"
  python mirai_reports.py --data-status 05/07/2026

Nome sugerido para o arquivo XML:
  SRB_MIRAI_TP-0500_Cronograma_do_Projeto-DD-MM-AAAA.xml
        """,
    )
    parser.add_argument(
        "--xml", type=str, default=None,
        help="Caminho para o arquivo XML do MS Project. "
             "Se omitido, busca automaticamente na pasta do script."
    )
    parser.add_argument(
        "--out", type=str, default=None,
        help="Pasta de saída para os HTMLs. Padrão: mesma pasta do XML."
    )
    parser.add_argument(
        "--data-status", type=str, default=None,
        dest="data_status",
        help="Data de status para análise (dd/mm/aaaa). "
             "Padrão: lê do campo StatusDate do XML."
    )
    args = parser.parse_args()

    print()
    print("═" * 60)
    print("  MIRAI REPORTS — Projeto Mirai · Sumitomo Rubber do Brasil")
    print("═" * 60)

    # 1. Localizar arquivo XML
    if args.xml:
        xml_path = Path(args.xml)
        if not xml_path.exists():
            print(f"\n❌ Arquivo não encontrado: {xml_path}")
            sys.exit(1)
    else:
        search_dir = Path(__file__).parent
        xml_path = find_xml(search_dir)
        if xml_path is None:
            print(f"\n❌ Nenhum arquivo '{XML_PATTERN}' encontrado em:")
            print(f"   {search_dir}")
            print("\n   Use --xml para especificar o caminho, ou coloque o arquivo")
            print(f"   na mesma pasta do script com o padrão:")
            print(f"   SRB_MIRAI_TP-0500_Cronograma_do_Projeto-DD-MM-AAAA.xml")
            sys.exit(1)

    print(f"\n📂 Arquivo XML : {xml_path.name}")
    print(f"   Pasta       : {xml_path.parent}")

    # 2. Pasta de saída
    out_dir = Path(args.out) if args.out else xml_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"   Saída       : {out_dir}")

    # 3. Data de status
    if args.data_status:
        try:
            status_date = datetime.strptime(args.data_status, "%d/%m/%Y")
        except ValueError:
            print(f"\n❌ Formato de data inválido: {args.data_status} (esperado: dd/mm/aaaa)")
            sys.exit(1)
    else:
        status_date = read_status_date(xml_path)

    sun = next_sunday(status_date)
    print(f"   Data status : {status_date.strftime('%d/%m/%Y')} (Sem. {status_date.isocalendar()[1]})")
    print(f"   Próximo Dom : {sun.strftime('%d/%m/%Y')}")

    # 4. Parsing
    print("\n⚙  Lendo XML...")
    leaf_tasks, delay_tasks = load_xml(xml_path, status_date)

    # 5. Curva S
    print("⚙  Calculando Curva S...")
    sc_labels, sc_bl0, sc_bl5, sc_ev = build_scurve(leaf_tasks, status_date)

    # 6. Métricas
    print("⚙  Calculando métricas por workstream...")
    frente_spi, fr_delay, overall = compute_frente_metrics(leaf_tasks, delay_tasks, status_date)

    print(f"\n   SPI Global vs BL5 : {overall['spi']}")
    print(f"   EV realizado      : {overall['ev']:.0f}h ({overall['pct']}%)")
    print(f"   Tarefas atrasadas : {len(delay_tasks)}")
    print(f"   Dev BL5 total     : {sum(fr_delay[f]['dev_bl5'] for f in fr_delay):.0f}h")
    print(f"   Dev BL0 total     : {sum(fr_delay[f]['dev_bl0'] for f in fr_delay):.0f}h")

    # 7. Gerar HTMLs
    print("\n📊 Gerando relatórios...")
    generate_gerencial(
        sc_labels, sc_bl0, sc_bl5, sc_ev,
        frente_spi, fr_delay, overall,
        status_date, xml_path.name,
        out_dir / "report_gerencial_mirai.html",
    )
    generate_atrasos(
        delay_tasks, fr_delay,
        status_date, sun, xml_path.name,
        out_dir / "report_atrasos_mirai.html",
    )

    print()
    print("═" * 60)
    print(f"  ✅ Concluído — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print(f"     report_gerencial_mirai.html")
    print(f"     report_atrasos_mirai.html")
    print(f"  📁 Pasta: {out_dir}")
    print("═" * 60)
    print()


if __name__ == "__main__":
    main()
