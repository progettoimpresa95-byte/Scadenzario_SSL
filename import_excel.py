#!/usr/bin/env python3
"""
Importa il vecchio file Excel "DVR_SCADENZARIO_CLIENTI.xlsx" nel formato dati
dell'app Scadenzario SSL e Ambiente (una cartella con config.json + aziende/*.json).

Uso:
    python3 import_excel.py  <file.xlsx>  <cartella_destinazione>

La cartella di destinazione è quella da copiare nella cartella condivisa
Dropbox. Lo script NON sovrascrive file esistenti a meno di --force.

Requisiti: Python 3.9+ e openpyxl  (pip install openpyxl)
"""
import json
import re
import sys
import unicodedata
from datetime import datetime, date
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# Catalogo adempimenti: id, nome, periodicità di default (anni, None = nessuna
# scadenza fissa), riferimento normativo e "prassi" = periodicità decisa dallo
# studio (non imposta dalla legge).
# ---------------------------------------------------------------------------
CATALOGO = [
    ("dvr", "DVR", 4, "D.Lgs. 81/08 artt. 17, 28, 29 – rielaborazione entro 30 gg da modifiche significative (art. 29 c.3)", True),
    ("rumore", "Rumore", 4, "D.Lgs. 81/08 art. 181 c.2 e art. 190 – aggiornamento almeno quadriennale", False),
    ("vibrazioni", "Vibrazioni", 4, "D.Lgs. 81/08 art. 181 c.2 e art. 202 – aggiornamento almeno quadriennale", False),
    ("atex", "Atmosfere esplosive (ATEX)", 4, "D.Lgs. 81/08 Titolo XI, artt. 290 e 294 – revisione a modifiche (periodicità di prassi)", True),
    ("roa", "ROA", 4, "D.Lgs. 81/08 art. 181 c.2 e art. 216 – aggiornamento almeno quadriennale", False),
    ("cem", "CEM", 4, "D.Lgs. 81/08 art. 181 c.2 e art. 209 – aggiornamento almeno quadriennale", False),
    ("amianto", "Amianto (stato di conservazione)", 1, "D.Lgs. 81/08 art. 248; DM 06/09/1994 – controllo periodico dei manufatti", False),
    ("vdt", "VDT", 4, "D.Lgs. 81/08 Titolo VII, art. 174 (periodicità di prassi)", True),
    ("chimico", "Rischio chimico", 3, "D.Lgs. 81/08 art. 223 c.7 – aggiornamento periodico (periodicità di prassi)", True),
    ("mmc", "MMC", 4, "D.Lgs. 81/08 Titolo VI, art. 168; ISO 11228-1/2 (periodicità di prassi)", True),
    ("ocra", "OCRA / movimenti ripetitivi", 4, "D.Lgs. 81/08 art. 28; ISO 11228-3 (periodicità di prassi)", True),
    ("slc", "Stress lavoro-correlato", 2, "D.Lgs. 81/08 art. 28 c.1-bis; Indicazioni Commissione consultiva 18/11/2010 (periodicità di prassi)", True),
    ("terra", "Messa a terra (verifica)", 5, "DPR 462/2001 art. 4 – 5 anni; 2 anni per cantieri, locali medici, ambienti a maggior rischio incendio", False),
    ("pe", "Piano di emergenza", None, "DM 02/09/2021 (GSA) – aggiornamento a modifiche, nessuna scadenza fissa", False),
    ("evacuazione", "Prova di evacuazione", 1, "DM 02/09/2021 (GSA), All. II – esercitazione almeno annuale", False),
    ("riunione", "Riunione periodica", 1, "D.Lgs. 81/08 art. 35 – almeno annuale (aziende > 15 lavoratori)", False),
    ("scariche", "Scariche atmosferiche (verifica)", 5, "DPR 462/2001 art. 4 – 5 anni; 2 anni nei casi a maggior rischio", False),
    ("scia", "SCIA VVF", 5, "DPR 151/2011 art. 5 – attestazione di rinnovo periodico ogni 5 anni (10 per alcune attività)", False),
    ("cpi", "CPI", 5, "DPR 151/2011 art. 5 – attestazione di rinnovo periodico ogni 5 anni (10 per alcune attività)", False),
    ("sollevamento", "Apparecchi di sollevamento", 2, "D.Lgs. 81/08 art. 71 c.11 e All. VII – verifiche periodiche annuali o biennali secondo il tipo", False),
    ("haccp", "Manuale HACCP", 3, "Reg. (CE) 852/2004 art. 5 – revisione a modifiche (periodicità di prassi)", True),
    ("aua", "AUA / AIA", 15, "DPR 59/2013 art. 3 c.6 – durata AUA 15 anni; AIA: D.Lgs. 152/06 art. 29-octies", False),
    ("emissioni", "Autocontrolli emissioni in atmosfera", 1, "D.Lgs. 152/06 Parte V, art. 269 – frequenza fissata dall'autorizzazione (AUA/AIA): impostala azienda per azienda", False),
    ("scarichi", "Autocontrolli scarichi idrici", 1, "D.Lgs. 152/06 Parte III, art. 124 – frequenza fissata dall'autorizzazione allo scarico", False),
    ("mud", "MUD (dichiarazione rifiuti)", 1, "L. 70/1994 – presentazione annuale entro il 30 aprile (rifiuti anno precedente)", False),
    ("fgas", "F-gas: controllo perdite", 1, "Reg. (UE) 2024/573 art. 5 – 12, 6 o 3 mesi in base alle t CO2 eq; dichiarazione annuale DPR 146/2018 art. 16", False),
    ("acustica", "Valutazione impatto acustico", None, "L. 447/1995 art. 8 – da rifare a modifiche di impianti/attività, nessuna scadenza fissa", False),
]

# Sezioni dell'app (schede separate) e sezione di ogni adempimento
SEZIONI = [
    {"id": "vr", "nome": "Valutazioni dei rischi", "breve": "Valutazioni rischi"},
    {"id": "ssl", "nome": "Adempimenti SSL", "breve": "Adempimenti SSL"},
    {"id": "amb", "nome": "Adempimenti ambientali", "breve": "Ambiente"},
    {"id": "altro", "nome": "Altri adempimenti", "breve": "Altri"},
]
SEZ_DEFAULT = {
    **{k: "vr" for k in ["dvr", "rumore", "vibrazioni", "atex", "roa", "cem", "amianto", "vdt", "chimico", "mmc", "ocra", "slc"]},
    **{k: "ssl" for k in ["terra", "pe", "evacuazione", "riunione", "scariche", "scia", "cpi", "sollevamento"]},
    **{k: "amb" for k in ["aua", "emissioni", "scarichi", "mud", "fgas", "acustica"]},
    "haccp": "altro",
}

# parola chiave (nel titolo della sezione Excel) -> id catalogo
MAPPA_TITOLI = [
    ("DVR", "dvr"), ("RUMORE", "rumore"), ("VIBRAZIONI", "vibrazioni"),
    ("ATMOSFERE ESPLOSIVE", "atex"), ("ROA", "roa"), ("CEM", "cem"),
    ("AMIANTO", "amianto"), ("VDT", "vdt"), ("RISCHIO CHIMICO", "chimico"),
    ("MMC", "mmc"), ("OCRA", "ocra"), ("STRESS", "slc"), ("MESSA A TERRA", "terra"),
    ("PIANO DI EMERGENZA", "pe"), ("PROVA DI EVACUAZIONE", "evacuazione"),
    ("RIUNIONE PERIODICA", "riunione"), ("SCARICHE", "scariche"),
    ("SCIA", "scia"), ("CPI", "cpi"), ("SOLLEVAMENTO", "sollevamento"),
    ("HACCP", "haccp"), ("AUA", "aua"),
]

FOGLI_DA_SALTARE = {"MASTER", "INDICE", "SITUAZIONE", "Foglio1", "LEGENDA"}


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def slug(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-") or "azienda"


def titolo_pulito(s):
    return " ".join(str(s or "").split())


def a_data(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, str):
        m = re.match(r"^\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})\s*$", v)
        if m:
            d, mth, y = map(int, m.groups())
            try:
                return date(y, mth, d).isoformat()
            except ValueError:
                return None
    return None


def testo(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    s = str(v).strip()
    if s.startswith("="):  # formula: non è un dato
        return ""
    return s


def id_catalogo(titolo):
    t = titolo_pulito(titolo).upper()
    for chiave, cid in MAPPA_TITOLI:
        if re.search(r"\b" + re.escape(chiave) + r"\b", t):
            return cid
    return None


def leggi_regola(g):
    """Interpreta la cella 'Best Before' (colonna G) della riga di testata.
    Ritorna (applicabilita, periodicita_anni|None, nota_extra)."""
    if g is None:
        return "attivo", None, ""
    s = str(g).strip()
    if s.startswith("="):
        m = re.search(r"\(\s*365\s*(?:\*\s*(\d+)\s*)?\)", s)
        anni = int(m.group(1) or 1) if m else None
        da_gestire = "da gestire" in s.lower()
        return ("attivo_dg" if da_gestire else "attivo"), anni, ""
    su = s.upper()
    if su.startswith("NEL DVR"):
        return "nel_dvr", None, ""
    if su in ("NP", "NON PRESENTE", "NON PRESENTI"):
        return "np", None, "" if su == "NP" else s
    if su == "NON SERVE":
        return "non_serve", None, ""
    if isinstance(g, (datetime, date)):
        return "attivo", None, "Scadenza indicata a mano in Excel: " + testo(g)
    return "attivo", None, "Nota da Excel (colonna Best Before): " + s


def importa_foglio(ws, catalogo_ids):
    righe = list(ws.iter_rows(values_only=True))
    pad = lambda r: list(r) + [None] * (11 - len(r))
    righe = [pad(r) for r in righe]

    def cella(etichetta):
        for r in righe[:6]:
            if norm(r[1]) == norm(etichetta):
                return testo(r[2])
        return ""

    az = {
        "ateco": cella("CODICE ATECO"),
        "soggettaCPI": cella("Attività soggetta a CPI?").upper(),
        "rischioIncendio": cella("Livello Rischio Incendio"),
        "classificazionePS": cella("Classificazione PS"),
        "adempimenti": {},
    }
    extra = []  # sezioni non a catalogo
    testate = [i for i, r in enumerate(righe) if str(r[2] or "").strip().lower() == "ultimo aggiornamento"]
    for k, i in enumerate(testate):
        r = righe[i]
        titolo = titolo_pulito(r[1])
        cid = id_catalogo(titolo)
        if cid is None:
            cid = "x-" + slug(titolo)
            extra.append((cid, titolo.capitalize()))
        appl, anni, nota_extra = leggi_regola(r[6])
        if cid == "pe" and appl == "np":
            appl = "attivo"  # per il PE "NP" in Excel = nessuna scadenza fissa
        fine = testate[k + 1] if k + 1 < len(testate) else len(righe)
        storico = []
        for j in range(i + 1, fine):
            rr = righe[j]
            if j > i + 1 and norm(rr[1]) and not any(rr[2:6]):
                break  # tabellina di riepilogo in fondo al foglio
            data = a_data(rr[5])
            note = [testo(x) for x in rr[7:11] if testo(x)]
            prot, rev, stato = testo(rr[2]), testo(rr[3]), testo(rr[4]).upper()
            if not (data or prot or rev or note):
                continue
            if not data and rr[5] is not None and testo(rr[5]):
                note.insert(0, "Data in Excel: " + testo(rr[5]))
            storico.append({
                "id": f"r{len(storico)+1}-{i}",
                "data": data or "",
                "prot": prot,
                "rev": rev,
                "stato": stato if stato in ("C", "DC", "I", "S", "A") else "",
                "note": " | ".join(note),
            })
        nota_testata = " | ".join(x for x in [testo(r[7]), testo(r[8]), nota_extra] if x)
        if appl in ("attivo", "attivo_dg"):
            ha_date = any(s["data"] for s in storico)
            if not ha_date:
                appl = "da_gestire" if appl == "attivo_dg" else "np"
            else:
                appl = "attivo"
        voce = {"applicabilita": appl, "storico": storico, "note": nota_testata}
        default = next((c[2] for c in CATALOGO if c[0] == cid), None)
        if anni is not None and anni != default:
            voce["periodicitaAnni"] = anni
        if cid in az["adempimenti"]:  # sezione duplicata: unisci
            az["adempimenti"][cid]["storico"] += storico
        else:
            az["adempimenti"][cid] = voce
    return az, extra


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    force = "--force" in sys.argv
    if len(args) != 2:
        print(__doc__)
        sys.exit(1)
    src, dst = Path(args[0]), Path(args[1])
    (dst / "aziende").mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(src, data_only=False)

    # Nome "bello" e note dalla SITUAZIONE (col. B nome, col. Y note)
    nomi, note_sit = {}, {}
    if "SITUAZIONE" in wb.sheetnames:
        for r in wb["SITUAZIONE"].iter_rows(min_row=3, values_only=True):
            r = list(r) + [None] * 27
            f = str(r[2] or "")
            m = re.match(r"^='?(.+?)'?!", f)
            if m and r[1]:
                nomi[m.group(1)] = titolo_pulito(r[1])
                if r[24]:
                    note_sit[m.group(1)] = testo(r[24])
    # RSPP e note dall'INDICE (col. C = "X" se RSPP dello studio, col. H note)
    indice = {}
    if "INDICE" in wb.sheetnames:
        for r in wb["INDICE"].iter_rows(min_row=5, values_only=True):
            r = list(r) + [None] * 8
            if r[3]:
                indice[norm(r[3])] = (str(r[2] or "").strip().upper() == "X", testo(r[7]))

    catalogo = [
        {"id": c[0], "nome": c[1], "sezione": SEZ_DEFAULT.get(c[0], "altro"), "periodicitaAnni": c[2], "riferimento": c[3], "prassi": c[4]}
        for c in CATALOGO
    ]
    ids_cat = {c["id"] for c in catalogo}
    scritti = 0
    usati = set()
    for ws in wb.worksheets:
        if ws.title in FOGLI_DA_SALTARE:
            continue
        az, extra = importa_foglio(ws, ids_cat)
        for cid, nome in extra:
            if cid not in ids_cat:
                catalogo.append({"id": cid, "nome": nome, "sezione": "vr", "periodicitaAnni": None,
                                 "riferimento": "", "prassi": True})
                ids_cat.add(cid)
        nome = nomi.get(ws.title) or ws.title.strip().upper()
        rspp, nota_idx = indice.get(norm(nome)) or indice.get(norm(ws.title)) or (False, "")
        note = [x for x in [note_sit.get(ws.title, ""), nota_idx] if x]
        aid = slug(ws.title)
        while aid in usati:
            aid += "-2"
        usati.add(aid)
        record = {
            "id": aid,
            "nome": nome,
            "attiva": True,
            "rsppStudio": rspp,
            "referente": "",
            **{k: az[k] for k in ("ateco", "soggettaCPI", "rischioIncendio", "classificazionePS")},
            "note": "\n".join(note),
            "adempimenti": az["adempimenti"],
            "_meta": {
                "versione": 1,
                "modificatoIl": datetime.now().isoformat(timespec="seconds"),
                "modificatoDa": "Import Excel",
            },
            "registro": [{
                "quando": datetime.now().isoformat(timespec="seconds"),
                "chi": "Import Excel",
                "cosa": f"Importata dal foglio '{ws.title}'",
            }],
        }
        out = dst / "aziende" / f"{aid}.json"
        if out.exists() and not force:
            print(f"  salto {out.name} (esiste già, usa --force per sovrascrivere)")
            continue
        out.write_text(json.dumps(record, ensure_ascii=False, indent=1), encoding="utf-8")
        scritti += 1

    cfg = dst / "config.json"
    if not cfg.exists() or force:
        cfg.write_text(json.dumps({
            "versioneFormato": 2,
            "sezioni": SEZIONI,
            "catalogo": catalogo,
            "sogliaGiorni": {"urgente": 30, "attenzione": 180},
            "_meta": {"versione": 1, "modificatoIl": datetime.now().isoformat(timespec="seconds"),
                      "modificatoDa": "Import Excel"},
        }, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Fatto: {scritti} aziende scritte in {dst / 'aziende'}")


if __name__ == "__main__":
    main()
