"""
Ordini sospesi ("impegnato virtuale").

Il file ordini_sospesi.json viene scritto dall'app (finestra "Sospesi") e
indica quali righe d'ordine NON devono contare nell'impegnato reale.

Struttura:
{
  "ordini": [
    {
      "num": "2867", "data": "06/07/2026", "cliente": "NUOVA BAI SRL",
      "righe": { "063MF1012": {"sospesa": null}, "31NT063": {"sospesa": 80000} },
      "storico": [ {"quando": "24/11/2026 15:40", "art": "063MF1012", "azione": "sblocca", "qta": 20000} ]
    }
  ]
}

"sospesa" e' la quantita' che RESTA sospesa:
  - null  = tutto il residuo dell'ordine per quell'articolo
  - numero = al massimo quella quantita' (se AdHoc riduce il residuo sotto
             questo valore, si sospende solo il residuo)
Cosi' quando il cliente ritira una tranche e AdHoc abbassa il residuo,
non serve correggere nulla.
"""
import json
from collections import defaultdict
from pathlib import Path

FILE_SOSPESI = "ordini_sospesi.json"


def carica_sospesi(percorso=FILE_SOSPESI):
    p = Path(percorso)
    if not p.exists():
        print("Ordini sospesi: nessun file, impegnato invariato")
        return {"ordini": []}
    try:
        dati = json.loads(p.read_text(encoding="utf-8") or "{}")
    except Exception as e:
        print(f"ATTENZIONE: {percorso} non leggibile ({e}), ignorato")
        return {"ordini": []}
    dati.setdefault("ordini", [])
    return dati


def chiave(num, data_str, art):
    return f"{str(num).strip()}|{str(data_str).strip()}|{str(art).strip()}"


def data_it(v):
    """Data ordine nello stesso formato usato dall'app: gg/mm/aaaa."""
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y")
    s = str(v or "")[:10]
    if len(s) == 10 and s[4] == "-":  # aaaa-mm-gg
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s


def calcola(righe, sospesi):
    """
    righe: lista di dict con num, data (gg/mm/aaaa), art, qta (residuo AdHoc).
    Ritorna:
      per_riga:     lista parallela a righe con la quantita' sospesa di ogni riga
      per_articolo: {codice: totale sospeso}
    """
    regole = {}
    for o in sospesi.get("ordini", []):
        for art, r in (o.get("righe") or {}).items():
            regole[chiave(o.get("num", ""), o.get("data", ""), art)] = (r or {}).get("sospesa")

    # residuo totale per ordine+articolo (lo stesso articolo puo' stare su piu' righe)
    gruppi = defaultdict(list)
    for i, r in enumerate(righe):
        k = chiave(r["num"], r["data"], r["art"])
        if k in regole:
            gruppi[k].append(i)

    per_riga = [0] * len(righe)
    per_articolo = defaultdict(float)
    for k, indici in gruppi.items():
        residuo = sum(max(0, righe[i]["qta"]) for i in indici)
        lim = regole[k]
        da_sospendere = residuo if lim is None else max(0, min(float(lim), residuo))
        for i in indici:
            q = min(max(0, righe[i]["qta"]), da_sospendere)
            per_riga[i] = q
            da_sospendere -= q
            per_articolo[righe[i]["art"]] += q

    trovati = {k.rsplit("|", 1)[0] for k in gruppi}
    mancanti = {f'{o.get("num")}|{o.get("data")}' for o in sospesi.get("ordini", [])} - trovati
    print(f"Ordini sospesi: {len(sospesi.get('ordini', []))} ordini, "
          f"{len(gruppi)} righe trovate in portafoglio, {round(sum(per_articolo.values()))} sospesi in totale")
    if mancanti:
        print(f"  non piu' in portafoglio (evasi o chiusi): {sorted(mancanti)}")
    return per_riga, dict(per_articolo)
