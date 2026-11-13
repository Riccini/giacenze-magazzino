"""
genera_ordini_json.py
Legge la Stampa portafoglio ordini (Stampa_portafoglio*.xlsx) e scrive ordini.json,
che la pagina carico_camion.html legge all'apertura.

Ogni ordine e' identificato da numero + serie (es. "MT") + data: lo stesso numero
puo' esistere in anni diversi, la pagina fa scegliere quale.
Formato di ogni ordine: {"n": numero, "a": serie, "d": "AAAA-MM-GG", "c": cliente, "cc": cod. cliente,
"r": [[codice, descr. anagrafica, descr. riga, unita', quantita', tipo riga], ...]}
"""
import glob
import json
import os
import sys
from datetime import datetime, date, timedelta

from openpyxl import load_workbook

USCITA = "ordini.json"


def trova_file():
    files = glob.glob("Stampa_portafoglio*.xlsx") + glob.glob("**/Stampa_portafoglio*.xlsx", recursive=True)
    files = sorted(set(files), key=lambda f: os.path.getmtime(f), reverse=True)
    return files[0] if files else None


def testo(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def data_iso(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and v > 20000:  # data seriale di Excel
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime("%Y-%m-%d")
    s = testo(v)
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def numero(v):
    if isinstance(v, (int, float)):
        return v
    try:
        return float(testo(v).replace(".", "").replace(",", "."))
    except ValueError:
        return 0


def leggi(percorso):
    wb = load_workbook(percorso, read_only=True, data_only=True)
    for ws in wb.worksheets:
        righe = ws.iter_rows(values_only=True)
        intest = None
        for r in righe:
            nomi = [testo(c).upper() for c in r]
            if "MVNUMDOC" in nomi:
                intest = nomi
                break
        if not intest:
            continue
        col = {n: i for i, n in enumerate(intest) if n}
        if not all(k in col for k in ("MVNUMDOC", "MVDATDOC", "ARCODART", "MVQTAMOV")):
            continue
        val = lambda r, k: r[col[k]] if k in col and col[k] < len(r) else None
        ordini = {}
        for r in righe:
            n = testo(val(r, "MVNUMDOC")).lstrip("0") or ""
            if not n:
                continue
            tipo_conto = testo(val(r, "MVTIPCON")).upper()
            if tipo_conto and tipo_conto != "C":  # solo ordini cliente
                continue
            a = testo(val(r, "MVALFDOC")).upper()
            d = data_iso(val(r, "MVDATDOC"))
            chiave = testo(val(r, "MVSERIAL")) or f"{n}|{a}|{d}"
            o = ordini.get(chiave)
            if o is None:
                o = ordini[chiave] = {"n": n, "a": a, "d": d, "c": testo(val(r, "ANDESCRI")),
                                      "cc": testo(val(r, "MVCODCON")), "r": []}
            q = numero(val(r, "MVQTAMOV"))
            if isinstance(q, float) and q.is_integer():
                q = int(q)
            o["r"].append([testo(val(r, "ARCODART")), testo(val(r, "ARDESART")), testo(val(r, "MVDESART")),
                           testo(val(r, "MVUNIMIS")), q, testo(val(r, "MVTIPRIG")).upper()])
        return list(ordini.values())
    raise ValueError("colonne del portafoglio ordini non trovate")


def main():
    percorso = sys.argv[1] if len(sys.argv) > 1 else trova_file()
    if not percorso:
        print("Nessun Stampa_portafoglio*.xlsx trovato: ordini.json non aggiornato.")
        return
    ordini = leggi(percorso)
    out = {"generato": datetime.now().astimezone().isoformat(timespec="seconds"),
           "file": os.path.basename(percorso), "docs": ordini}
    with open(USCITA, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    mt = sum(1 for o in ordini if o["a"] == "MT")
    print(f"{USCITA}: {len(ordini)} ordini aperti ({mt} MT) da {percorso}")


if __name__ == "__main__":
    main()
