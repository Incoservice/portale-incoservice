#!/usr/bin/env python3
"""
aggiorna_portale.py
-------------------
Legge il file Excel di IncoService da OneDrive e aggiorna
automaticamente il file index.html con i dati aggiornati.
Poi esegue git add + commit + push su GitHub Pages.
"""

import os
import re
import sys
import json
import time
import shutil
import datetime
import subprocess
import openpyxl

# ============================================================
# CONFIG — I TUOI PERCORSI CORRETTI
# ============================================================

# Percorso completo al file Excel su OneDrive
EXCEL_PATH = r"C:\Users\Davide\OneDrive - INCOSERVICE SRL\File di Microsoft Copilot Chat\GestionaleIncoservice\011 - Programmazione settimanale LAVORAZIONI\Gestione Produzioni_rev.xlsx"

# Percorso alla cartella del repository GitHub locale sul Desktop
REPO_PATH = r"C:\Users\Davide\OneDrive - INCOSERVICE SRL\Attachments\Desktop\portale-incoservice"

# Nome del file HTML nel repository
HTML_FILENAME = "index.html"

# Quanti minuti tra un aggiornamento e l'altro
INTERVALLO_MINUTI = 5

# ============================================================
# FINE CONFIG
# ============================================================

HTML_PATH = os.path.join(REPO_PATH, HTML_FILENAME)

def log(msg):
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")

def get_row_color(ws, row):
    cell = ws.cell(row, 1)
    fill = cell.fill
    try:
        if fill.fgColor.type == "rgb":
            rgb = fill.fgColor.rgb
            if rgb == "FFFFFF00":
                return "gialla"
            elif rgb in ("00000000", "00FFFFFF", "FFFFFFFF") or not rgb:
                return "bianca"
            else:
                return "colorata"
        elif fill.fgColor.type == "theme":
            return "colorata"
        else:
            return "bianca"
    except Exception:
        return "bianca"

def fmt_date(val):
    if val is None: return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    return s if s else None

def cell_val(ws, row, col):
    v = ws.cell(row, col).value
    if v is None: return None
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v

def js_val(v):
    if v is None: return "null"
    if isinstance(v, bool): return "true" if v else "false"
    if isinstance(v, (int, float)): return str(v)
    s = str(v).replace("\\", "\\\\").replace('"', '\\"')
    return f'"{s}"'

def leggi_ordini(ws):
    ordini = []
    for row in range(2, ws.max_row + 1):
        cliente = cell_val(ws, row, 1)
        if not cliente: continue

        colore = get_row_color(ws, row)
        pronto = (colore == "gialla")

        odl = []
        for c in range(23, 30):
            v = ws.cell(row, c).value
            if v and str(v).strip():
                try: odl.append(int(v))
                except: pass

        ordini.append({
            "cliente": str(cliente),
            "impegno": str(cell_val(ws, row, 2) or ""),
            "tipologia": str(cell_val(ws, row, 7) or ""),
            "qty": str(cell_val(ws, row, 8) or ""),
            "finitura": str(cell_val(ws, row, 9) or "").lower(),
            "ral": str(cell_val(ws, row, 10) or ""),
            "posa": str(cell_val(ws, row, 11) or ""),
            "produzione": str(cell_val(ws, row, 12) or ""),
            "invioZN": fmt_date(cell_val(ws, row, 14)),
            "ritornoZN": fmt_date(cell_val(ws, row, 16)),
            "pulizia": fmt_date(cell_val(ws, row, 17)),
            "invioRAL": fmt_date(cell_val(ws, row, 18)),
            "verniciatura": str(cell_val(ws, row, 19) or ""),
            "imballo": fmt_date(cell_val(ws, row, 20)),
            "consStimata": fmt_date(cell_val(ws, row, 21)),
            "consRichiesta": fmt_date(cell_val(ws, row, 22)),
            "odl": odl,
            "prontoConsegna": pronto,
            "colore": colore
        })
    return ordini

def aggiorna_html(ordini):
    if not os.path.exists(HTML_PATH):
        log(f"ERRORE: Il file {HTML_FILENAME} non esiste in {REPO_PATH}")
        return False

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    # Genera blocco JS
    js_content = "const ORDINI_DATA = [\n"
    for o in ordini:
        odl_js = "[" + ",".join(str(x) for x in o["odl"]) + "]"
        js_content += f'  {{ cliente:{js_val(o["cliente"])}, impegno:{js_val(o["impegno"])}, tipologia:{js_val(o["tipologia"])}, qty:{js_val(o["qty"])}, finitura:{js_val(o["finitura"])}, ral:{js_val(o["ral"])}, posa:{js_val(o["posa"])}, produzione:{js_val(o["produzione"])}, invioZN:{js_val(o["invioZN"])}, ritornoZN:{js_val(o["ritornoZN"])}, pulizia:{js_val(o["pulizia"])}, invioRAL:{js_val(o["invioRAL"])}, verniciatura:{js_val(o["verniciatura"])}, imballo:{js_val(o["imballo"])}, consStimata:{js_val(o["consStimata"])}, consRichiesta:{js_val(o["consRichiesta"])}, odl:{odl_js}, prontoConsegna:{js_val(o["prontoConsegna"])}, colore:{js_val(o["colore"])} }},\n'
    js_content += "];"

    pattern = r"const ORDINI_DATA\s*=\s*\[.*?\];"
    html_new = re.sub(pattern, js_content, html, flags=re.DOTALL)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html_new)
    
    log(f"File HTML aggiornato localmente con {len(ordini)} righe.")
    return True

def git_push():
    try:
        subprocess.run(["git", "-C", REPO_PATH, "add", "."], check=True)
        # Controlla se ci sono modifiche
        status = subprocess.run(["git", "-C", REPO_PATH, "diff", "--cached", "--quiet"])
        if status.returncode == 0:
            log("Nessuna modifica rilevata, salto il caricamento.")
            return
        
        subprocess.run(["git", "-C", REPO_PATH, "commit", "-m", "Auto-update"], check=True)
        subprocess.run(["git", "-C", REPO_PATH, "push"], check=True)
        log("Dati caricati su GitHub con successo!")
    except Exception as e:
        log(f"Errore Git: {e}")

def ciclo():
    log("Inizio ciclo di aggiornamento...")
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb["Programma Produzioni"]
        ordini = leggi_ordini(ws)
        if aggiorna_html(ordini):
            git_push()
    except Exception as e:
        log(f"Errore generale: {e}")

def main():
    log("=== AGGIORNAMENTO AUTOMATICO PORTALE INCOSERVICE AVVIATO ===")
    while True:
        ciclo()
        log(f"In attesa per {INTERVALLO_MINUTI} minuti...")
        time.sleep(INTERVALLO_MINUTI * 60)

if __name__ == "__main__":
    main()