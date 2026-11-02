"""Convertitore .pub -> Scribus .sla modificabile come in Publisher.

Publisher (via COM) legge il documento: caselle di testo con colonne, margini,
paragrafi e formattazione, forme, linee e immagini. Il risultato e' un file
Scribus nativo con cornici di testo vere (il testo va a capo e scorre) e una
cartella con le immagini.

Uso:
    python pub2sla.py FILE_O_CARTELLA [...] [--formato scribus|designcraft] [--out CARTELLA] [--verifica]

    --formato   scribus (.sla, default) oppure designcraft (.designcraft + .idml)
    --out       cartella di destinazione (default: accanto al .pub, "<nome>_<formato>")
    --accanto   salva nella stessa cartella del .pub solo il documento (e le immagini per Scribus)
    --elenco    file di testo UTF-8 con un .pub per riga, in aggiunta ai percorsi
    --font      copia accanto al documento i font che usa, nella cartella "Document Fonts"
    --verifica  esporta il risultato in PDF e misura le differenze
                rispetto al PDF esportato da Publisher
"""
import argparse
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path
from xml.sax.saxutils import quoteattr

import pymupdf
import pythoncom
import win32com
import win32com.client
import win32com.gen_py
from win32com.client import gencache

from pub2idml import motivo_tratto, scrivi_idml
from raccogli_font import raccogli


def trova_programma(standard, *modelli):
    """Il percorso abituale o, se non c'e', il programma cercato nelle cartelle di installazione."""
    if Path(standard).exists():
        return standard
    radici = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"),
              Path(os.environ.get("LOCALAPPDATA", "")) / "Programs"]
    for radice in filter(None, radici):
        for modello in modelli:
            trovati = sorted(Path(radice).glob(modello))
            if trovati:
                return str(trovati[-1])
    return standard


SCRIBUS = trova_programma(r"C:\Program Files (x86)\Scribus 1.6.7\Scribus.exe", "Scribus*/Scribus.exe")
DESIGNCRAFT = trova_programma(r"C:\Program Files\ArtCraft\DesignCraft\designcraft.exe",
                              "*/DesignCraft/designcraft.exe", "DesignCraft/designcraft.exe")
DESIGNCRAFT_CLI = str(Path(DESIGNCRAFT).with_name("designcraft-cli.exe"))
LIMITE_SECONDI = 600  # per ogni file: oltre, Publisher e' considerato bloccato
PB_LINE, PB_PICTURE_TYPES = 9, (11, 13, 15)
TIPO_TESTO, TIPO_GRUPPO, TIPO_TABELLA = 17, 6, 19
MSO_RETTANGOLO, MSO_ARROTONDATO, MSO_OVALE, MSO_MISTA, MSO_FOTO = 1, 5, 9, -2, 13
SPESSORE_LINEA = 2  # una forma "mista" piu' sottile di cosi' e' una linea
PB_INTENT_COMMERCIAL, PB_FIXED_FORMAT_PDF, RISOLUZIONE_STAMPA = 4, 2, 3
PB_ELENCO_PUNTATO = 23
STRISCIA_MINIMA = 40  # sotto questa larghezza Publisher non mette testo accanto a un ostacolo
DASH = {1: 1, 2: 3, 3: 3, 4: 2, 5: 4, 6: 5, 7: 2, 8: 4, 9: 5, 10: 2, 11: 4, 12: 5}


def colore(com) -> str:
    v = int(com.RGB) & 0xFFFFFF  # Publisher: 0xBBGGRR
    return f"{v & 255:02x}{(v >> 8) & 255:02x}{(v >> 16) & 255:02x}"


def uniforme(v) -> bool:
    return v in (0, 1, -1)


# ---------------------------------------------------------------- lettura .pub

def segmenti(r, out):
    """Divide un intervallo di testo in pezzi con formattazione uniforme."""
    n = r.Length
    if n <= 0:
        return
    f = r.Font
    # Bold/Italic: -1 = si'; Underline: -1 = misto, >0 = tipo di sottolineatura
    ok = bool(f.Name) and f.Size > 0 and uniforme(f.Bold) and uniforme(f.Italic) and f.Underline >= 0
    if ok and n > 1:
        ok = colore(r.Characters(1, 1).Font.Color) == colore(r.Characters(n, 1).Font.Color)
    if ok or n == 1:
        # su un carattere solo -1 non puo' voler dire "misto": e' sottolineato (es. i link)
        out.append({"start": r.Start, "t": r.Text, "font": f.Name, "size": round(f.Size, 2),
                    "b": f.Bold != 0, "i": f.Italic != 0, "u": f.Underline != 0,
                    "color": colore(f.Color), "scala": f.Scaling})
        return
    meta = n // 2
    segmenti(r.Characters(1, meta), out)
    segmenti(r.Characters(meta + 1, n - meta), out)


def unisci_uguali(runs) -> list:
    """La divisione puo' arrivare al singolo carattere: riunisce i pezzi con lo stesso stile."""
    out = []
    for r in runs:
        if out and all(out[-1][k] == r[k] for k in ("font", "size", "b", "i", "u", "color", "scala")):
            out[-1]["t"] += r["t"]
        else:
            out.append(r)
    return out


def compatta(s: str) -> str:
    return "".join(c for c in s if not c.isspace() and c != "\xad")


class RighePDF:
    """Righe esatte del PDF esportato da Publisher: le coordinate che Publisher
    fornisce via COM sono approssimate a scatti di 2,4 pt."""

    def __init__(self, pagina):
        self.righe = []
        for b in pagina.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                cs = [c for s in l["spans"] for c in s["chars"]]
                pieni = [c for c in cs if not c["c"].isspace()]
                if pieni:
                    y, size = pieni[0]["origin"][1], l["spans"][0]["size"]
                    self.righe.append({"testo": compatta("".join(c["c"] for c in cs)), "pieni": pieni,
                                       "y": y, "x0": pieni[0]["origin"][0], "x1": pieni[-1]["bbox"][2],
                                       # il riquadro del font e' troppo alto: conta l'ingombro che
                                       # Scribus considera attorno alla linea di base
                                       "bbox": (l["bbox"][0], y - 0.8 * size, l["bbox"][2], y + 0.75 * size)})

    def trova(self, testo, top, sinistra):
        t = compatta(testo)
        if not t:
            return None
        t2 = t.rstrip("-")

        def vicino(r):
            return abs(r["y"] - top) < 30 and r["x1"] > sinistra - 5

        uguali = [r for r in self.righe if vicino(r) and r["testo"].rstrip("-") == t2]
        # elenchi: nel PDF il pallino fa parte della riga
        if not uguali and len(t2) >= 5:
            uguali = [r for r in self.righe if vicino(r) and r["testo"].rstrip("-").endswith(t2)]
        if uguali:
            return min(uguali, key=lambda r: abs(r["y"] - top) + abs(r["x0"] - sinistra) / 10)
        return self.ricomponi(t2, vicino)

    def ricomponi(self, t, vicino):
        """Nel PDF una riga puo' essere spezzata in piu' pezzi sulla stessa linea di base
        (es. testo con scala orizzontale): li rimette insieme."""
        file_ = {}
        for r in filter(vicino, self.righe):
            file_.setdefault(round(r["y"]), []).append(r)
        for pezzi in file_.values():
            pezzi.sort(key=lambda r: r["x0"])
            for i in range(len(pezzi)):
                testo = ""
                for k in range(i, len(pezzi)):
                    testo += pezzi[k]["testo"]
                    if testo.rstrip("-") == t:
                        scelti = pezzi[i:k + 1]
                        y0, y1 = scelti[0]["bbox"][1], scelti[0]["bbox"][3]
                        return {"testo": testo, "pieni": [c for p in scelti for c in p["pieni"]],
                                "y": scelti[0]["y"], "x0": scelti[0]["x0"], "x1": scelti[-1]["x1"],
                                "bbox": (scelti[0]["bbox"][0], y0, scelti[-1]["bbox"][2], y1), "parti": scelti}
                    if not t.startswith(testo):
                        break
        return None

    def consuma(self, riga):
        # ogni riga del PDF corrisponde a una sola riga di Publisher
        for r in riga.get("parti", [riga]):
            self.righe.remove(r)


def righe_testo(t, indice):
    """Per ogni riga di Publisher: posizione approssimata (COM) e riga esatta nel PDF."""
    out = []
    for i in range(1, t.LinesCount + 1):
        l = t.Lines(i)
        testo = l.Text
        pdf = indice.trova(testo, l.BoundTop + l.BoundHeight, l.BoundLeft) if indice else None
        out.append({"start": l.Start, "end": l.End, "testo": testo, "left": l.BoundLeft, "top": l.BoundTop,
                    "right": l.BoundLeft + l.BoundWidth, "alta": l.BoundHeight, "pdf": pdf})
        if pdf:
            indice.consuma(pdf)
    return out


def interruzioni(t, righe) -> set:
    """Posizioni in cui Publisher ha sillabato: la riga finisce a meta' parola."""
    testo, inizio, tagli = t.Text, t.Start, set()
    for r in righe[:-1]:
        fine = r["end"] - inizio
        if 0 < fine < len(testo) and testo[fine - 1].isalnum() and testo[fine].isalnum():
            tagli.add(fine + inizio)
    return tagli


def fermate_tab(righe, sinistra) -> list:
    """Tabulazioni: la posizione reale del testo dopo ogni gruppo di tab, dal PDF."""
    fermate = set()
    for r in righe:
        if "\t" not in r["testo"] or not r["pdf"]:
            continue
        pieni, n_prima = r["pdf"]["pieni"], 0
        parti = [p for p in re.split(r"\t+", r["testo"])]
        for parte in parti[:-1]:
            n_prima += len(compatta(parte))
            if n_prima < len(pieni):
                fermate.add(round(pieni[n_prima]["origin"][0] - sinistra, 2))
    return sorted(fermate)


def interlinea_formula(f, size_max) -> float:
    if f["regola"] == 4:  # esatta
        return f["valore"]
    multiplo = {0: 1.0, 1: 1.5, 2: 2.0}.get(f["regola"], f["valore"] if f["regola"] == 5 else 1.0)
    return round(size_max * 1.15 * multiplo, 3)


def leggi_paragrafo(p, tagli, righe, sinistra, inizi_colonna):
    f, runs = p.ParagraphFormat, []
    segmenti(p, runs)
    runs = unisci_uguali(runs)
    mie = [r for r in righe if p.Start <= r["start"] < max(p.End, p.Start + 1)]
    for r in runs:
        testo, nuovo = r["t"], []
        for i, c in enumerate(testo):
            if r["start"] + i in tagli:
                nuovo.append("\xad")
            nuovo.append(c)
        r["t"] = "".join(nuovo)
    size_max = max((r["size"] for r in runs), default=p.Font.Size if p.Font.Size > 0 else 10)
    tabs = []
    try:
        ts = f.Tabs
        tabs = [(ts.Item(i + 1).Position, ts.Item(i + 1).Alignment) for i in range(ts.Count)]
    except Exception:
        pass
    if sinistra is not None:
        misurate = fermate_tab(mie, sinistra)
        if misurate:
            tipo = tabs[0][1] if tabs else 0
            tabs = [(pos, tipo) for pos in misurate]
    bullet = ""
    try:
        if f.ListType == PB_ELENCO_PUNTATO:
            bullet = f.ListBulletText
    except Exception:
        pass
    # distanza fra le righe esatta dal PDF; un paragrafo di una riga la eredita (vedi completa_interlinee)
    basi = [r["pdf"]["y"] for r in mie if r["pdf"]]
    salti = sorted(b - a for a, b in zip(basi, basi[1:]) if 0 < b - a < size_max * 4)
    ls = round(salti[len(salti) // 2], 3) if salti else None
    # accanto a un'immagine Publisher ignora il rientro della prima riga e non giustifica
    first = f.FirstLineIndent
    if first > 0 and mie and mie[0]["pdf"]:
        x0 = mie[0]["pdf"]["x0"]
        if len(mie) >= 2 and mie[1]["pdf"] and abs(x0 - mie[1]["pdf"]["x0"]) < 1:
            first = 0
        elif min(abs(x0 - (c + first)) for c in inizi_colonna) > 4:
            first = 0
    align = f.Alignment
    corpo = mie[:-1]
    if align not in (0, 1, 2) and corpo and all(r["pdf"] for r in corpo) \
            and all(r["right"] - r["pdf"]["x1"] > 4 for r in corpo):
        align = 0
    return {"align": align, "first": first, "left": f.LeftIndent, "right": f.RightIndent,
            "prima": f.SpaceBefore, "dopo": f.SpaceAfter, "ls": ls,
            "regola": f.LineSpacingRule, "valore": f.LineSpacing,
            "tabs": tabs, "bullet": bullet, "runs": runs, "size": size_max}


def completa_interlinee(pagine):
    """Paragrafi di una riga: stessa interlinea dei paragrafi uguali, oppure formula."""
    noti = {}
    for pag in pagine:
        for o in pag["oggetti"]:
            for p in ((o.get("testo") or {}).get("paras") or []):
                if p["ls"]:
                    noti.setdefault((p["regola"], round(p["valore"], 2), p["size"]), p["ls"])
    for pag in pagine:
        for o in pag["oggetti"]:
            for p in ((o.get("testo") or {}).get("paras") or []):
                if not p["ls"]:
                    p["ls"] = noti.get((p["regola"], round(p["valore"], 2), p["size"])) or interlinea_formula(p, p["size"])


def leggi_testo(s, tf, indice):
    t = tf.Story.TextRange if tf.HasNextLink else tf.TextRange
    righe = righe_testo(t, indice)
    tagli = interruzioni(t, righe)
    # con piu' colonne non so a quale colonna riferire la tabulazione
    sinistra = s.Left + tf.MarginLeft if tf.Columns == 1 else None
    interna = s.Width - tf.MarginLeft - tf.MarginRight
    larga = (interna - tf.ColumnSpacing * (tf.Columns - 1)) / tf.Columns
    inizi = [s.Left + tf.MarginLeft + k * (larga + tf.ColumnSpacing) for k in range(tf.Columns)]
    paragrafi = [leggi_paragrafo(t.Paragraphs(i), tagli, righe, sinistra, inizi)
                 for i in range(1, t.ParagraphsCount + 1)]

    def ingombro(r):
        if r["pdf"]:
            return r["pdf"]["bbox"]
        # testo non presente nel PDF (es. con scala orizzontale diventa grafica): stima da COM
        s = r["alta"] / 1.15
        base = r["top"] + 0.85 * r["alta"]
        return (r["left"], base - 0.8 * s, r["right"], base + 0.75 * s)

    return {"cols": tf.Columns, "gap": tf.ColumnSpacing, "valign": tf.VerticalTextAlignment,
            "m": [tf.MarginLeft, tf.MarginTop, tf.MarginRight, tf.MarginBottom], "paras": paragrafi,
            "ingombri": [ingombro(r) for r in righe if compatta(r["testo"])],
            "basi": [(r["pdf"]["x0"], r["pdf"]["y"]) for r in righe if r["pdf"]]}


def ha_testo(s) -> bool:
    try:
        return bool(s.HasTextFrame) and bool(s.TextFrame.TextRange.Text.strip())
    except Exception:
        return False


def base_forma(s):
    o = {"nome": s.Name, "id_pub": s.ID, "x": s.Left, "y": s.Top, "w": s.Width, "h": s.Height, "rot": s.Rotation,
         "flip_h": s.HorizontalFlip != 0, "flip_v": s.VerticalFlip != 0}
    try:
        tw = s.TextWrap
        o["wrap"] = [tw.Type, tw.DistanceLeft, tw.DistanceTop, tw.DistanceRight, tw.DistanceBottom]
    except Exception:
        o["wrap"] = [0, 0, 0, 0, 0]
    try:
        if s.Fill.Visible:
            o["fill"], o["fill_t"] = colore(s.Fill.ForeColor), s.Fill.Transparency
    except Exception:
        pass
    try:
        if s.Line.Visible and s.Line.Weight > 0:
            o["line"], o["line_w"], o["line_t"] = colore(s.Line.ForeColor), s.Line.Weight, s.Line.Transparency
            o["dash"] = DASH.get(s.Line.DashStyle, 1)
            o["motivo"] = motivo_tratto(s.Line.DashStyle, s.Line.Weight)
    except Exception:
        pass
    try:
        sh = s.Shadow
        if sh.Visible:
            o["ombra"] = {"x": sh.OffsetX, "y": sh.OffsetY, "blur": sh.Blur, "t": sh.Transparency,
                          "c": colore(sh.ForeColor)}
    except Exception:
        pass
    return o


class Lettore:
    def __init__(self, doc, immagini: Path):
        self.doc, self.immagini, self.avvisi, self.n = doc, immagini, [], 0
        self.indice = None  # righe del PDF della pagina corrente

    def immagine(self, s, o):
        self.n += 1
        file = self.immagini / f"{self.n:03d}_{s.Name.replace(' ', '_')}.png"
        s.SaveAsPicture(str(file.resolve()), RISOLUZIONE_STAMPA)
        # riempimento e ombra sono gia' nell'immagine salvata; il bordo solo se non e' una fotografia
        for chiave in ("fill", "fill_t", "ombra"):
            o.pop(chiave, None)
        if s.Type != MSO_FOTO:
            for chiave in ("line", "line_w", "line_t", "dash", "motivo"):
                o.pop(chiave, None)
        o.update(tipo="immagine", file=file.name)
        return o

    def forma(self, s, oggetti):
        t = s.Type
        if t == TIPO_GRUPPO:
            for i in range(1, s.GroupItems.Count + 1):
                self.forma(s.GroupItems.Item(i), oggetti)
            return
        o = base_forma(s)
        if t == TIPO_TABELLA or s.HasTable:
            self.avvisi.append(f"{s.Name}: tabella importata come immagine (non modificabile)")
            oggetti.append(self.immagine(s, o))
        elif t in PB_PICTURE_TYPES:
            oggetti.append(self.immagine(s, o))
        elif t == PB_LINE:
            if s.Width or s.Height:
                o["tipo"] = "linea"
                oggetti.append(o)
        elif t == 1 and s.AutoShapeType == MSO_MISTA and min(s.Width, s.Height) <= SPESSORE_LINEA \
                and "line" in o and not ha_testo(s):
            # linee disegnate come AutoShape: una cornice alta 0,4 pt conterrebbe solo un'immagine schiacciata
            o["tipo"] = "linea"
            if s.Height < s.Width:
                o["y"], o["h"] = o["y"] + o["h"] / 2, 0
            else:
                o["x"], o["w"] = o["x"] + o["w"] / 2, 0
            oggetti.append(o)
        elif t == TIPO_TESTO or ha_testo(s):
            self.casella(s, o, oggetti)
        elif t == 1 and s.AutoShapeType in (MSO_RETTANGOLO, MSO_ARROTONDATO, MSO_OVALE):
            o["tipo"], o["auto"] = "forma", s.AutoShapeType
            if s.AutoShapeType == MSO_ARROTONDATO and s.Adjustments.Count:
                o["raggio"] = s.Adjustments.Item(1) * min(s.Width, s.Height)
            if "fill" in o or "line" in o:
                oggetti.append(o)
        else:
            self.avvisi.append(f"{s.Name}: forma tipo {t} importata come immagine")
            oggetti.append(self.immagine(s, o))

    def casella(self, s, o, oggetti):
        tf = s.TextFrame
        o["tipo"] = "testo"
        if s.Type == 1 and s.AutoShapeType == MSO_ARROTONDATO and s.Adjustments.Count:
            o["raggio"] = s.Adjustments.Item(1) * min(s.Width, s.Height)
        elif s.Type == 1 and s.AutoShapeType != MSO_RETTANGOLO:
            self.avvisi.append(f"{s.Name}: forma con testo importata come rettangolo")
        collegata = tf.HasPreviousLink or tf.HasNextLink
        if not ha_testo(s) and not collegata:
            if "fill" not in o and "line" not in o:
                return
            o["testo"] = None
        elif tf.HasPreviousLink:
            o["testo"] = {"cols": tf.Columns, "gap": tf.ColumnSpacing, "valign": tf.VerticalTextAlignment,
                          "m": [tf.MarginLeft, tf.MarginTop, tf.MarginRight, tf.MarginBottom], "paras": None}
        else:
            o["testo"] = leggi_testo(s, tf, self.indice)
        if tf.HasPreviousLink:
            o["precedente"] = tf.PreviousLinkedTextFrame.Parent.ID
        if tf.HasNextLink:
            o["successivo"] = tf.NextLinkedTextFrame.Parent.ID
        oggetti.append(o)

    def sfondo(self, pag, oggetti):
        """Lo sfondo pagina di Publisher non e' una forma: in Scribus diventa un rettangolo."""
        ps = self.doc.PageSetup
        for origine in (pag.Master if not pag.IgnoreMaster else None, pag):
            try:
                fill = origine.Background.Fill
                if not fill.Visible:
                    continue
                if fill.Type != 1:  # sfumatura, trama o immagine: resta da rifare a mano
                    self.avvisi.append(f"sfondo pagina tipo {fill.Type}: importato solo il colore principale")
                oggetti.append({"nome": "Sfondo pagina", "id_pub": -1, "x": 0, "y": 0, "w": ps.PageWidth,
                                "h": ps.PageHeight, "rot": 0, "flip_h": False, "flip_v": False,
                                "wrap": [0, 0, 0, 0, 0], "tipo": "forma", "auto": MSO_RETTANGOLO,
                                "fill": colore(fill.ForeColor)})
                return
            except Exception:
                continue

    def pagina(self, pag, pagina_pdf):
        oggetti = []
        self.indice = RighePDF(pagina_pdf)
        self.sfondo(pag, oggetti)
        if not pag.IgnoreMaster and pag.Master is not None:
            for i in range(1, pag.Master.Shapes.Count + 1):
                s = pag.Master.Shapes.Item(i)
                if ha_testo(s) or s.Type in PB_PICTURE_TYPES:
                    self.forma(s, oggetti)
        forme = [pag.Shapes.Item(i) for i in range(1, pag.Shapes.Count + 1)]
        for s in sorted(forme, key=lambda f: f.ZOrderPosition):
            self.forma(s, oggetti)
        # Page.Width e' arrotondato: le misure esatte sono in PageSetup
        return {"w": self.doc.PageSetup.PageWidth, "h": self.doc.PageSetup.PageHeight, "oggetti": oggetti}


def leggi_pub(app, pub: Path, cartella: Path, pdf: Path):
    immagini = cartella / f"{pub.stem}_immagini"
    immagini.mkdir(parents=True, exist_ok=True)
    doc = app.Open(str(pub.resolve()), True, False)
    try:
        doc.ExportAsFixedFormat(PB_FIXED_FORMAT_PDF, str(pdf.resolve()), PB_INTENT_COMMERCIAL)
        lett = Lettore(doc, immagini)
        with pymupdf.open(pdf) as d:
            pagine = [lett.pagina(doc.Pages(i), d[i - 1]) for i in range(1, doc.Pages.Count + 1)]
        completa_interlinee(pagine)
        return {"pagine": pagine, "immagini": immagini.name}, lett.avvisi
    finally:
        doc.Saved = True  # leggere i dati lo segna "modificato": evita la richiesta invisibile di salvare
        doc.Close()


# ---------------------------------------------------------------- font Scribus

def norm(s: str) -> str:
    return "".join(c for c in s.lower() if c.isalnum())


class Font:
    GRASSETTI = ("bold", "demibold", "semibold", "demi", "heavy", "black")
    NORMALI = ("regular", "roman", "book", "normal", "medium", "")

    def __init__(self, nomi):
        self.nomi = nomi
        self.indice = {norm(n): n for n in nomi}
        self.mancanti = set()

    def cerca(self, nome, b, i):
        base = norm(nome)
        stili = self.GRASSETTI if b else self.NORMALI
        corsivi = ("italic", "oblique") if i else ("",)
        for st in stili:
            for co in corsivi:
                for ordine in (st + co, co + st):
                    if base + ordine in self.indice:
                        return self.indice[base + ordine]
        if b:  # "Lucida Sans Demibold Roman" e varianti: stessa famiglia, grassetto, corsivo giusto
            for nome_font in self.nomi:
                x = norm(nome_font)
                resto = x[len(base):]
                inclinato = "italic" in resto or "oblique" in resto
                if x.startswith(base) and any(resto.startswith(g) for g in self.GRASSETTI) and inclinato == i:
                    return nome_font
        return None

    def risolvi(self, nome, b, i) -> str:
        if nome.startswith("MS Shell Dlg"):  # font di sistema delle caselle vuote
            nome = "Tahoma"
        r = self.cerca(nome, b, i)
        if r is None and i:
            r = self.cerca(nome, b, False)
            if r:
                self.mancanti.add(f"{nome}: nessun corsivo, usato {r}")
        if r is None and b:
            r = self.cerca(nome, False, False)
            if r:
                self.mancanti.add(f"{nome}: nessun grassetto, usato {r}")
        if r is None:  # es. 'AbadiMT-ExtraLight ?': famiglia con stile strano
            r = next((n for n in self.nomi if norm(n).startswith(norm(nome))), None)
        if r is None:  # es. Publisher 'Abadi Extra Light' = Scribus 'AbadiMT-ExtraLight'
            senza_mt = norm(nome).replace("mt", "")
            r = next((n for n in self.nomi if norm(n).replace("mt", "").startswith(senza_mt)), None)
        if r is None:
            self.mancanti.add(f"{nome}: font non trovato, usato Arial")
            r = self.cerca("Arial", b, i) or "Arial Regular"
        return r


def percorso_font() -> Path:
    """Il file dell'elenco font: nell'eseguibile la cartella del programma e' provvisoria."""
    qui = Path(__file__).parent
    dati = qui / "font_scribus.json"
    if getattr(sys, "frozen", False):
        dati = Path(os.environ["LOCALAPPDATA"]) / "Convertitore PUB" / "font_scribus.json"
    if not dati.exists() and not Path(SCRIBUS).exists():  # senza Scribus bastano i font comuni di Windows
        dati.parent.mkdir(parents=True, exist_ok=True)
        dati.write_bytes((qui / "font_base.json").read_bytes())
    return dati


def elenco_font_scribus(dati: Path):
    # senza Scribus (si converte solo per DesignCraft) resta buono l'ultimo elenco salvato
    if dati.exists() and (time.time() - dati.stat().st_mtime < 86400 or not Path(SCRIBUS).exists()):
        return json.loads(dati.read_text(encoding="utf-8"))
    with tempfile.TemporaryDirectory() as d:
        script = Path(d) / "font.py"
        uscita = Path(d) / "font.json"
        script.write_text(f"import scribus, json\nopen(r'{uscita}', 'w', encoding='utf-8')"
                          f".write(json.dumps(scribus.getFontNames()))\n", encoding="utf-8")
        try:
            subprocess.run([SCRIBUS, "-g", "-ns", "-py", str(script)], timeout=180, check=False)
            elenco = uscita.read_text(encoding="utf-8")
        except (OSError, subprocess.TimeoutExpired):  # Scribus non risponde: il documento si converte lo stesso
            riserva = dati if dati.exists() else Path(__file__).parent / "font_base.json"
            return json.loads(riserva.read_text(encoding="utf-8"))
        dati.parent.mkdir(parents=True, exist_ok=True)
        dati.write_text(elenco, encoding="utf-8")
    return json.loads(elenco)


# ---------------------------------------------------------------- scrittura .sla

def n(v) -> str:
    return f"{v:.4f}".rstrip("0").rstrip(".") or "0"


def attr(**kw) -> str:
    return " ".join(f"{k}={quoteattr(str(v))}" for k, v in kw.items())


def rettangolo(x0, y0, x1, y1) -> str:
    return f"M{n(x0)} {n(y0)} L{n(x1)} {n(y0)} L{n(x1)} {n(y1)} L{n(x0)} {n(y1)} L{n(x0)} {n(y0)} Z"


def arrotondato(w, h, r) -> str:
    k = r * 0.5523
    return (f"M{n(r)} 0 L{n(w - r)} 0 C{n(w - r + k)} 0 {n(w)} {n(r - k)} {n(w)} {n(r)} "
            f"L{n(w)} {n(h - r)} C{n(w)} {n(h - r + k)} {n(w - r + k)} {n(h)} {n(w - r)} {n(h)} "
            f"L{n(r)} {n(h)} C{n(r - k)} {n(h)} 0 {n(h - r + k)} 0 {n(h - r)} "
            f"L0 {n(r)} C0 {n(r - k)} {n(r - k)} 0 {n(r)} 0 Z")


def ovale(w, h) -> str:
    a, b = w / 2, h / 2
    kx, ky = a * 0.5523, b * 0.5523
    return (f"M{n(a)} 0 C{n(a + kx)} 0 {n(w)} {n(b - ky)} {n(w)} {n(b)} "
            f"C{n(w)} {n(b + ky)} {n(a + kx)} {n(h)} {n(a)} {n(h)} "
            f"C{n(a - kx)} {n(h)} 0 {n(b + ky)} 0 {n(b)} "
            f"C0 {n(b - ky)} {n(a - kx)} 0 {n(a)} 0 Z")


class Scrittore:
    def __init__(self, dati, font: Font):
        self.dati, self.font, self.colori, self.id = dati, font, {"000000", "ffffff"}, 1000

    def nome_colore(self, c) -> str:
        self.colori.add(c)
        return {"000000": "Black", "ffffff": "White"}.get(c, f"C_{c}")

    def stile_paragrafo(self, p) -> dict:
        allineamenti = {0: 0, 1: 1, 2: 2}
        a = dict(ALIGN=allineamenti.get(p["align"], 3), LINESPMode=0, LINESP=n(p["ls"]), INDENT=n(p["left"]),
                 RMARGIN=n(p["right"]), FIRST=n(p["first"]), VOR=n(p["prima"]), NACH=n(p["dopo"]))
        return a

    def stile_carattere(self, r) -> dict:
        a = dict(FONT=self.font.risolvi(r["font"], r["b"], r["i"]), FONTSIZE=n(r["size"]),
                 FCOLOR=self.nome_colore(r["color"]), FSHADE=100)
        if r["u"]:
            a["FEATURES"] = "underline"
        if r.get("scala") and r["scala"] != 100:
            a["SCALEH"] = n(r["scala"])
        return a

    def storia(self, t) -> str:
        righe = ["<StoryText>", "<DefaultStyle/>"]
        for i, p in enumerate(t["paras"]):
            ultimo_tab = False  # piu' tab di fila vanno a una sola fermata misurata
            if p["bullet"] and p["runs"]:
                # pallino come testo piu' tab: il pallino automatico di Scribus non ha rientro
                simbolo = "•" if p["bullet"] in "·" else p["bullet"]
                righe.append(f'<ITEXT {attr(**self.stile_carattere(p["runs"][0]))} CH={quoteattr(simbolo)}/>')
                righe.append("<tab/>")
                if p["left"] > 0 and not any(abs(pos - p["left"]) < 0.01 for pos, _ in p["tabs"]):
                    p["tabs"] = sorted(p["tabs"] + [(p["left"], 0)])
            for r in p["runs"]:
                a = attr(**self.stile_carattere(r))
                pezzo = []

                def scarica():
                    if pezzo:
                        righe.append(f'<ITEXT {a} CH={quoteattr("".join(pezzo))}/>')
                        pezzo.clear()

                for c in r["t"]:
                    if c == "\r":
                        continue
                    if c == "\t":
                        scarica()
                        if not ultimo_tab:
                            righe.append("<tab/>")
                        ultimo_tab = True
                        continue
                    ultimo_tab = False
                    if c in "\n\x0b":
                        scarica()
                        righe.append("<breakline/>")
                    elif c == "\x0c":
                        scarica()
                        righe.append("<breakcol/>")
                    elif ord(c) >= 32:
                        pezzo.append(c)
                scarica()
            st = self.stile_paragrafo(p)
            if p["runs"]:
                r0 = p["runs"][-1]
                st.update(FONT=self.font.risolvi(r0["font"], r0["b"], r0["i"]), FONTSIZE=n(r0["size"]))
            else:
                st.update(FONTSIZE=n(p["size"]))
            tag = "trail" if i == len(t["paras"]) - 1 else "para"
            tabs = "".join(f'<Tabs Type="{ {0: 0, 1: 1, 2: 4}.get(al, 0) }" Pos="{n(pos)}" Fill=""/>' for pos, al in p["tabs"])
            righe.append(f"<{tag} {attr(**st)}>{tabs}</{tag}>" if tabs else f"<{tag} {attr(**st)}/>")
        righe.append("</StoryText>")
        return "\n".join(righe)

    def comuni(self, o, ox, oy) -> dict:
        w, h, rot = o["w"], o["h"], o["rot"]
        x, y = o["x"], o["y"]
        if rot and o["tipo"] != "immagine":  # Publisher ruota attorno al centro, Scribus attorno all'angolo
            cx, cy, ang = x + w / 2, y + h / 2, math.radians(rot)
            x = cx - w / 2 * math.cos(ang) + h / 2 * math.sin(ang)
            y = cy - w / 2 * math.sin(ang) - h / 2 * math.cos(ang)
        elif rot:  # l'immagine salvata e' gia' ruotata: cornice sul suo ingombro
            ang = math.radians(rot)
            bw = abs(w * math.cos(ang)) + abs(h * math.sin(ang))
            bh = abs(w * math.sin(ang)) + abs(h * math.cos(ang))
            x, y, w, h, rot = x + w / 2 - bw / 2, y + h / 2 - bh / 2, bw, bh, 0
        self.id += 1
        a = dict(XPOS=n(x + ox), YPOS=n(y + oy), OwnPage=o["pagina"], ItemID=self.id, WIDTH=n(w), HEIGHT=n(h),
                 ROT=n(rot), ANNAME=o["nome"], LAYER=0, gXpos=n(x + ox), gYpos=n(y + oy), gWidth=0, gHeight=0,
                 PLINEART=o.get("dash", 1), PWIDTH=n(o.get("line_w", 0)), LOCALSCX=1, LOCALSCY=1, LOCALX=0, LOCALY=0,
                 LOCALROT=0, PICART=1, SCALETYPE=1, RATIO=1, CLIPEDIT=0)
        if o.get("motivo"):
            a.update(NUMDASH=len(o["motivo"]), DASHS=" ".join(n(v) for v in o["motivo"]), DASHOFF=0)
        if "fill" in o:
            a["PCOLOR"] = self.nome_colore(o["fill"])
            if o.get("fill_t"):
                a["TransValue"] = n(o["fill_t"])
        if "line" in o:
            a["PCOLOR2"] = self.nome_colore(o["line"])
            if o.get("line_t"):
                a["TransValueS"] = n(o["line_t"])
        if "ombra" in o:
            s = o["ombra"]
            a.update(HASSOFTSHADOW=1, SOFTSHADOWCOLOR=self.nome_colore(s["c"]), SOFTSHADOWSHADE=100,
                     SOFTSHADOWXOFFSET=n(s["x"]), SOFTSHADOWYOFFSET=n(s["y"]), SOFTSHADOWBLUR=n(s["blur"]),
                     # come Publisher: l'ombra e' trasparente quanto l'oggetto e si vede attraverso il riempimento
                     SOFTSHADOWOPACITY=n(1 - (1 - s["t"]) * (1 - (o.get("fill_t") or 0))), SOFTSHADOWBLENDMODE=0,
                     SOFTSHADOWERASE=0, SOFTSHADOWOBJTRANS=0)
        tw = o["wrap"]
        a["TEXTFLOWMODE"] = 3 if tw[0] in (1, 2, 4) else 0  # 3 = contorno: permette i margini di scontorno
        return a, w, h

    def contorno(self, o, w, h) -> str:
        if "contorno" in o:
            c0, c1, c2, c3 = o["contorno"]
            return rettangolo(c0 - o["x"], c1 - o["y"], c2 - o["x"], c3 - o["y"])
        tipo, l, t, r, b = o["wrap"]
        if tipo == 4:  # sopra e sotto: il contorno copre tutta la larghezza pagina
            l = r = o["pagina_w"]
        return rettangolo(-l, -t, w + r, h + b)

    @staticmethod
    def colonne(f):
        t = f["testo"]
        sx, dx = f["x"] + t["m"][0], f["x"] + f["w"] - t["m"][2]
        larga = (dx - sx - t["gap"] * (t["cols"] - 1)) / t["cols"]
        return [(sx + k * (larga + t["gap"]), sx + k * (larga + t["gap"]) + larga) for k in range(t["cols"])]

    def margine_colonne(self, o) -> float:
        """Scribus accetta in fondo alla colonna una riga che Publisher manda alla
        successiva: alza il margine inferiore fino a riprodurre lo stesso cambio colonna."""
        t = o["testo"]
        basso = t["m"][3]
        basi = t.get("basi") or []
        if t["cols"] < 2 or len(basi) < 2 or not t["paras"]:
            return basso
        ls = max(p["ls"] for p in t["paras"])
        cambi = [a[1] for a, b in zip(basi, basi[1:]) if b[0] > a[0] + 20 and b[1] < a[1]]
        if not cambi:
            return basso
        limite = max(min(cambi) + 0.8 * ls, max(b[1] for b in basi) + 0.25 * ls)
        return max(basso, o["y"] + o["h"] - limite)

    def adatta_contorni(self, oggetti, per_scribus=True):
        """Scribus e' piu' severo di Publisher attorno agli ostacoli: taglia le righe
        anche per un contatto minimo e riempie strisce strettissime. Il layout reale
        di Publisher dice dove il testo stava davvero: il contorno viene ridotto per
        non toccare quelle righe e allargato per coprire le strisce inutilizzabili."""
        for o in oggetti if per_scribus else []:
            t = o.get("testo") if o["tipo"] == "testo" else None
            # righe vuote in testa a una casella il cui testo parte piu' in basso, spinto da un
            # ostacolo: Publisher le mette sotto l'ostacolo, Scribus in cima. Le toglie il contorno
            if t and t.get("paras") and t.get("basi") and len(t["paras"]) > 1:
                ls = t["paras"][0]["ls"]
                if min(b[1] for b in t["basi"]) - o["y"] - t["m"][1] > 2.2 * ls:
                    while len(t["paras"]) > 1 and not any(compatta(r["t"]) for r in t["paras"][0]["runs"]):
                        t["paras"].pop(0)
        # casella con sfondo o bordo troppo bassa per Scribus: sfondo e bordo diventano una
        # forma a parte delle stesse misure, cosi' si puo' alzare solo il testo
        for o in list(oggetti) if per_scribus else []:
            t = o.get("testo") if o["tipo"] == "testo" else None
            if t and t.get("ingombri") and t["cols"] == 1 and not o["rot"] and "successivo" not in o \
                    and ("fill" in o or "line" in o) \
                    and max(b[3] for b in t["ingombri"]) - o["y"] + t["m"][3] + 4 > o["h"]:
                forma = {k: v for k, v in o.items() if k != "testo"}
                forma.update(tipo="forma", auto=MSO_RETTANGOLO, wrap=[0, 0, 0, 0, 0], id_pub=-2,
                             nome=f'{o["nome"]} sfondo')
                forma.pop("successivo", None)
                forma.pop("precedente", None)
                oggetti.insert(oggetti.index(o), forma)
                for chiave in ("fill", "fill_t", "line", "line_w", "line_t", "dash", "ombra"):
                    o.pop(chiave, None)
        for j, o in enumerate(oggetti):
            tipo, l, t, r, b = o["wrap"]
            if tipo not in (1, 2, 4) or o["rot"]:
                continue
            if tipo == 4:
                l = r = o["pagina_w"]
            c = [o["x"] - l, o["y"] - t, o["x"] + o["w"] + r, o["y"] + o["h"] + b]
            sotto = [f for f in oggetti[:j] if f["tipo"] == "testo" and f.get("testo") and f["testo"].get("paras")]
            for f in sotto:
                if not (c[1] < f["y"] + f["h"] and c[3] > f["y"]):
                    continue
                for sx, dx in self.colonne(f):
                    if c[2] > sx and c[0] < dx:
                        if 0 < c[0] - sx < STRISCIA_MINIMA:
                            c[0] = sx - 1
                        if 0 < dx - c[2] < STRISCIA_MINIMA:
                            c[2] = dx + 1
            for f in sotto:
                for x0, y0, x1, y1 in f["testo"]["ingombri"]:
                    if c[0] < x1 and c[2] > x0 and c[1] < y1 and c[3] > y0:
                        tagli = [(y1 - c[1], 1, y1), (c[3] - y0, 3, y0), (x1 - c[0], 0, x1), (c[2] - x0, 2, x0)]
                        _, lato, valore = min(tagli)
                        c[lato] = valore
            # la prima riga sotto l'ostacolo deve cadere dove la metteva Publisher
            for f in sotto:
                if f["y"] + f["testo"]["m"][1] >= c[3]:  # casella tutta sotto: l'ostacolo non la spinge
                    continue
                sotto_ostacolo = []
                for x0, y0, x1, y1 in f["testo"]["ingombri"]:
                    s = (y1 - y0) / 1.25
                    if c[0] < x1 and c[2] > x0 and y0 + 0.8 * s > c[3]:
                        sotto_ostacolo.append((y0 + 0.8 * s, s))
                if sotto_ostacolo:
                    base, s = min(sotto_ostacolo)
                    nuovo = base - 0.95 * s
                    if c[3] < nuovo < c[3] + 2.5 * s:
                        c[3] = nuovo
            if c[0] >= c[2] or c[1] >= c[3]:
                o["wrap"] = [0, 0, 0, 0, 0]
            else:
                o["contorno"] = c

    def oggetto(self, o, ox, oy, ids) -> str:
        t = o.get("testo")
        if o["tipo"] == "testo" and t and t.get("ingombri") and t["cols"] == 1 and not o["rot"] \
                and "fill" not in o and "line" not in o and "successivo" not in o:
            # Scribus nasconde le righe che sporgono anche di poco, Publisher no: alzo la casella.
            # Un testo ancorato in basso o al centro tronca le righe vicino agli ostacoli:
            # senza sfondo l'ancoraggio non si vede, quindi lo ottengo con il margine superiore
            alto, fondo = min(b[1] for b in t["ingombri"]), max(b[3] for b in t["ingombri"])
            if t["valign"]:
                s = (t["ingombri"][0][3] - t["ingombri"][0][1]) / 1.25
                t["m"][1] = max(0, alto + 0.8 * s - 0.95 * s - o["y"])
                t["valign"] = 0
            o["h"] = max(o["h"], fondo - o["y"] + t["m"][3] + 4)
        a, w, h = self.comuni(o, ox, oy)
        ids.append(a["ItemID"])
        tipo = o["tipo"]
        path = rettangolo(0, 0, w, h)
        interno = ""
        if tipo == "immagine":
            a.update(PTYPE=2, FRTYPE=0, SCALETYPE=0, RATIO=0, PFILE=f'{self.dati["immagini"]}/{o["file"]}', IRENDER=1,
                     EMBEDDED=0, COLUMNS=1, COLGAP=0)
        elif tipo == "linea":
            dx = -w if o["flip_h"] else w
            dy = -h if o["flip_v"] else h
            lung = math.hypot(dx, dy)
            x0 = o["x"] + (w if o["flip_h"] else 0)
            y0 = o["y"] + (h if o["flip_v"] else 0)
            ang = math.degrees(math.atan2(dy, dx)) + o["rot"]
            path = f"M0 0 L{n(lung)} 0"
            a.update(PTYPE=5, FRTYPE=0, WIDTH=n(lung), HEIGHT=1, ROT=n(ang), XPOS=n(x0 + ox), YPOS=n(y0 + oy),
                     gXpos=n(x0 + ox), gYpos=n(y0 + oy))
        elif tipo == "forma":
            auto = o["auto"]
            a.update(PTYPE=6, FRTYPE=1 if auto == MSO_OVALE else (0 if auto == MSO_RETTANGOLO else 3))
            if auto == MSO_ARROTONDATO:
                path = arrotondato(w, h, o.get("raggio", 0))
            elif auto == MSO_OVALE:
                path = ovale(w, h)
        else:  # testo
            t = o["testo"]
            a.update(PTYPE=4, FRTYPE=0, FLOP=1, AUTOTEXT=0, PLTSHOW=0, BASEOF=0, textPathType=0, textPathFlipped=0,
                     COLUMNS=1, COLGAP=0)
            if o.get("raggio"):
                a["FRTYPE"] = 3
                path = arrotondato(w, h, o["raggio"])
            if t:
                a.update(COLUMNS=t["cols"], COLGAP=n(t["gap"]), EXTRA=n(t["m"][0]), TEXTRA=n(t["m"][1]),
                         REXTRA=n(t["m"][2]), BEXTRA=n(self.margine_colonne(o)),
                         VAlign=t["valign"] if t["valign"] in (0, 1, 2) else 0)
                if t["paras"] is not None:
                    interno = self.storia(t)
        a["path"], a["copath"] = path, self.contorno(o, w, h) if a["TEXTFLOWMODE"] else path
        a.update(NEXTITEM=o.get("_next", -1), BACKITEM=o.get("_back", -1))
        testa = attr(**a)
        return f"<PAGEOBJECT {testa}>\n{interno}\n</PAGEOBJECT>" if interno else f"<PAGEOBJECT {testa}/>"

    def documento(self) -> str:
        pagine = self.dati["pagine"]
        pw, ph = pagine[0]["w"], pagine[0]["h"]
        gap, sx, sy = 40, 100, 20
        pag_xml, ogg_xml = [], []
        y = sy
        collegamenti = []
        for i, p in enumerate(pagine):
            pag_xml.append(f'<PAGE {attr(PAGEXPOS=sx, PAGEYPOS=n(y), PAGEWIDTH=n(p["w"]), PAGEHEIGHT=n(p["h"]), BORDERLEFT=0, BORDERRIGHT=0, BORDERTOP=0, BORDERBOTTOM=0, NUM=i, NAM="", MNAM="Normal", Size="Custom", Orientation=1, LEFT=0, PRESET=0)}/>')
            for o in p["oggetti"]:
                o["pagina_w"] = p["w"]
            self.adatta_contorni(p["oggetti"])
            for o in p["oggetti"]:
                o.update(pagina=i, pagina_w=p["w"])
                collegamenti.append((o, sx, y))
            y += p["h"] + gap
        # catene di caselle collegate: gli ItemID vengono assegnati nell'ordine di scrittura
        base = self.id + 1
        item_id = {o["id_pub"]: base + k for k, (o, _, _) in enumerate(collegamenti)}
        for o, _, _ in collegamenti:
            if "successivo" in o:
                o["_next"] = item_id[o["successivo"]]
            if "precedente" in o:
                o["_back"] = item_id[o["precedente"]]
        ids = []
        for o, ox, oy in collegamenti:
            ogg_xml.append(self.oggetto(o, ox, oy, ids))
        colori = "\n".join(
            f'<COLOR {attr(NAME={"000000": "Black", "ffffff": "White"}.get(c, "C_" + c), SPACE="RGB", R=int(c[:2], 16), G=int(c[2:4], 16), B=int(c[4:], 16))}/>'
            for c in sorted(self.colori))
        doc = attr(ANZPAGES=len(pagine), PAGEWIDTH=n(pw), PAGEHEIGHT=n(ph), BORDERLEFT=0, BORDERRIGHT=0, BORDERTOP=0,
                   BORDERBOTTOM=0, PRESET=0, BleedTop=0, BleedLeft=0, BleedRight=0, BleedBottom=0, ORIENTATION=1,
                   PAGESIZE="Custom", FIRSTNUM=1, BOOK=0, AUTOSPALTEN=1, ABSTSPALTEN=11, UNITS=1, DFONT="Arial Regular",
                   DSIZE=10, DCOL=1, DGAP=0, TabFill="", TabWidth=36, TextDistLeft=0, TextDistRight=0, TextDistBottom=0,
                   TextDistTop=0, FirstLineOffset=1, AUTOL=100, LANGUAGE="it", ALAYER=0, ScratchBottom=20, ScratchLeft=sx,
                   ScratchRight=sx, ScratchTop=sy, GapHorizontal=0, GapVertical=gap, SHOWGRID=0, SHOWGUIDES=1,
                   SHOWFRAME=1, SHOWMARGIN=0, SHOWPICT=1, rulerMode=1, showrulers=1, PEN="Black", BRUSH="None",
                   PENLINE="Black", PENTEXT="Black", StrokeText="Black", TextBackGround="None", TextLineColor="None")
        return "\n".join([
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<SCRIBUSUTF8NEW Version="1.6.7">',
            f"<DOCUMENT {doc}>",
            colori,
            '<CHARSTYLE CNAME="Default Character Style" DefaultStyle="1" FONT="Arial Regular" FONTSIZE="10" FEATURES="inherit" FCOLOR="Black" FSHADE="100" SCOLOR="None" BGCOLOR="None" BGSHADE="100" SSHADE="100" TXTSHX="5" TXTSHY="-5" TXTOUT="1" TXTULP="-0.1" TXTULW="-0.1" TXTSTP="-0.1" TXTSTW="-0.1" SCALEH="100" SCALEV="100" BASEO="0" KERN="0" LANGUAGE="it"/>',
            '<STYLE NAME="Default Paragraph Style" DefaultStyle="1" ALIGN="0" DIRECTION="0" LINESPMode="0" LINESP="12" INDENT="0" RMARGIN="0" FIRST="0" VOR="0" NACH="0" DROP="0" DROPLIN="2" Bullet="0" Numeration="0" BCOLOR="None" BSHADE="100"/>',
            '<LAYERS NUMMER="0" LEVEL="0" NAME="Sfondo" SICHTBAR="1" DRUCKEN="1" EDIT="1" SELECT="0" FLOW="1" TRANS="1" BLEND="0" OUTL="0" LAYERC="#000000"/>',
            f'<MASTERPAGE {attr(PAGEXPOS=sx, PAGEYPOS=sy, PAGEWIDTH=n(pw), PAGEHEIGHT=n(ph), BORDERLEFT=0, BORDERRIGHT=0, BORDERTOP=0, BORDERBOTTOM=0, NUM=0, NAM="Normal", MNAM="", Size="Custom", Orientation=1, LEFT=0, PRESET=0)}/>',
            "\n".join(pag_xml),
            "\n".join(ogg_xml),
            "</DOCUMENT>",
            "</SCRIBUSUTF8NEW>",
            ""])


# ---------------------------------------------------------------- verifica

def esporta_con_scribus(sla: Path, pdf: Path):
    with tempfile.TemporaryDirectory() as d:
        script = Path(d) / "esporta.py"
        script.write_text(
            "import scribus\n"
            f"scribus.openDoc(r'{sla.resolve()}')\n"
            "p = scribus.PDFfile()\n"
            f"p.file = r'{pdf.resolve()}'\n"
            "p.version = 14\n"  # PDF 1.4: senza, trasparenze e ombre vengono appiattite
            "p.save()\n"
            "scribus.closeDoc()\n", encoding="utf-8")
        subprocess.run([SCRIBUS, "-g", "-ns", "-py", str(script)], timeout=300, check=False)


def designcraft(ingresso: Path, uscita: Path):
    """Con DesignCraft: IDML -> .designcraft nativo, oppure documento -> PDF."""
    r = subprocess.run([DESIGNCRAFT_CLI, "run", "--in", str(ingresso.resolve()), "--export", str(uscita.resolve())],
                       capture_output=True, encoding="utf-8", errors="replace", timeout=300, check=False,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    if r.returncode or not uscita.exists():
        raise RuntimeError(f"DesignCraft non ha creato {uscita.name}: {(r.stderr or r.stdout).strip()[:300]}")


def confronta(originale: Path, nuovo: Path):
    """Per ogni parola dell'originale cerca la stessa parola vicina nel PDF di Scribus."""
    a_posto, totali, dy_tot, fuori = 0, 0, 0.0, []
    with pymupdf.open(originale) as a, pymupdf.open(nuovo) as b:
        if len(a) != len(b):
            print(f"  pagine diverse: {len(a)} contro {len(b)}")
        for pa, pb in zip(a, b):
            per_testo = {}
            for x0, _, _, y1, parola, *_ in pb.get_text("words"):
                per_testo.setdefault(parola, []).append((x0, y1))
            for x0, _, _, y1, parola, *_ in pa.get_text("words"):
                totali += 1
                vicine = [p for p in per_testo.get(parola, []) if abs(p[1] - y1) < 3 and abs(p[0] - x0) < 12]
                if vicine:
                    m = min(vicine, key=lambda p: abs(p[1] - y1) + abs(p[0] - x0))
                    per_testo[parola].remove(m)
                    a_posto += 1
                    dy_tot += abs(m[1] - y1)
                else:
                    fuori.append((pa.number + 1, y1, x0, parola))
    print(f"  parole al loro posto: {a_posto}/{totali} ({100 * a_posto / max(totali, 1):.1f}%), "
          f"scarto verticale medio {dy_tot / max(a_posto, 1):.2f} pt")
    # raggruppa le parole fuori posto per riga, per leggerle meglio
    righe = {}
    for pag, y, x, parola in sorted(fuori):
        righe.setdefault((pag, round(y / 4)), []).append(parola)
    return [(k[0], " ".join(v)) for k, v in righe.items()]


# ---------------------------------------------------------------- avvio

def pid_publisher() -> set:
    r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq MSPUB.EXE", "/FO", "CSV", "/NH"],
                       capture_output=True, text=True, check=False, creationflags=subprocess.CREATE_NO_WINDOW)
    return {int(riga.split('","')[1]) for riga in r.stdout.splitlines() if riga.startswith('"MSPUB')}


def apri_publisher():
    """Publisher con le classi generate da pywin32: senza, assegnare doc.Saved fallisce.

    Le classi stanno in una cartella del programma: quella predefinita e' nei file temporanei,
    che Windows puo' pulire lasciandola a meta'. Se e' danneggiata si ricrea.
    """
    cartella = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "Convertitore PUB" / "gen_py"
    cartella /= f"{sys.version_info.major}.{sys.version_info.minor}"
    try:
        return _publisher_con_classi(cartella)
    except (AttributeError, ImportError, OSError, SyntaxError, EOFError):
        shutil.rmtree(cartella, ignore_errors=True)
        for nome in [m for m in sys.modules if m.startswith("win32com.gen_py.")]:
            del sys.modules[nome]
        return _publisher_con_classi(cartella)


def _publisher_con_classi(cartella: Path):
    cartella.mkdir(parents=True, exist_ok=True)
    win32com.__gen_path__ = str(cartella)
    win32com.gen_py.__path__ = [str(cartella)]
    gencache.clsidToTypelib.clear()
    gencache.__init__()  # rilegge l'indice dalla cartella appena scelta
    app = win32com.client.DispatchEx("Publisher.Application")
    tipo = app._oleobj_.GetTypeInfo().GetContainingTypeLib()[0].GetLibAttr()
    gencache.EnsureModule(str(tipo[0]), tipo[1], tipo[3], tipo[4])  # versione di Publisher installata
    return win32com.client.Dispatch(app)


def trova_pub(percorsi):
    for p in map(Path, percorsi):
        if p.is_dir():
            yield from sorted(p.rglob("*.pub"))
        elif p.suffix.lower() == ".pub":
            yield p


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("percorsi", nargs="*")
    ap.add_argument("--formato", choices=["scribus", "designcraft"], default="scribus")
    ap.add_argument("--out")
    ap.add_argument("--accanto", action="store_true")
    ap.add_argument("--elenco")
    ap.add_argument("--font", action="store_true")
    ap.add_argument("--pdf", action="store_true", help="tiene il PDF esportato da Publisher accanto al risultato")
    ap.add_argument("--verifica", action="store_true")
    args = ap.parse_args()
    if args.elenco:
        args.percorsi += Path(args.elenco).read_text(encoding="utf-8").splitlines()
    files = list(trova_pub(args.percorsi))
    if not files:
        sys.exit("Nessun file .pub trovato.")

    gia_aperti = pid_publisher()

    def termina_publisher():
        for pid in pid_publisher() - gia_aperti:  # solo l'istanza nascosta avviata da questo script
            subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, check=False,
                           creationflags=subprocess.CREATE_NO_WINDOW)

    def scaduto():
        print("Publisher non risponde: interrotto.", flush=True)
        termina_publisher()
        os._exit(2)

    def avvia_guardia():
        g = threading.Timer(LIMITE_SECONDI, scaduto)
        g.daemon = True
        g.start()
        return g

    guardia = avvia_guardia()
    font = Font(elenco_font_scribus(percorso_font()))
    pythoncom.CoInitialize()
    try:
        app = apri_publisher()
    except Exception:
        termina_publisher()
        raise
    codice = 0
    try:
        for i, pub in enumerate(files, 1):
            guardia.cancel()
            guardia = avvia_guardia()
            if args.accanto:
                cartella = pub.parent
            else:
                cartella = Path(args.out) if args.out else pub.parent / f"{pub.stem}_{args.formato}"
            t0 = time.time()
            # PDF e IDML servono solo a lavorare: si preparano fuori dalla cartella del documento
            provvisoria = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
            lavoro = Path(provvisoria.name)
            pdf_orig = lavoro / f"{pub.stem}_publisher.pdf"  # riferimento esatto
            try:
                cartella.mkdir(parents=True, exist_ok=True)
                if args.formato == "designcraft":
                    # immagini incorporate nell'IDML: si preparano fuori da OneDrive, che blocca le cartelle
                    dati, avvisi = leggi_pub(app, pub, lavoro, pdf_orig)
                    idml = lavoro / f"{pub.stem}.idml"
                    scrivi_idml(dati, font, lavoro, idml, Scrittore(dati, font))
                    risultato = cartella / f"{pub.stem}.designcraft"
                    designcraft(idml, risultato)
                else:
                    dati, avvisi = leggi_pub(app, pub, cartella, pdf_orig)
                    risultato = cartella / f"{pub.stem}.sla"
                    risultato.write_text(Scrittore(dati, font).documento(), encoding="utf-8")
                print(f"[{i}/{len(files)}] OK  {pub.name} -> {risultato.name} ({time.time() - t0:.1f}s)", flush=True)
                print(f"  creato: {risultato.resolve()}", flush=True)
                if args.pdf:
                    shutil.copy(pdf_orig, cartella / pdf_orig.name)
                for a in avvisi + sorted(font.mancanti):
                    print("  avviso:", a)
                font.mancanti.clear()
                if args.font:
                    try:
                        for messaggio in raccogli(dati, cartella, pub.stem):
                            print(messaggio, flush=True)
                    except OSError as e:  # il documento e' gia' salvato: non e' un errore di conversione
                        print(f"  avviso: font non copiati: {e}", flush=True)
                if args.verifica:
                    pdf_nuovo = (cartella if args.pdf else lavoro) / f"{pub.stem}_{args.formato}.pdf"
                    if args.formato == "designcraft":
                        designcraft(risultato, pdf_nuovo)
                    else:
                        esporta_con_scribus(risultato, pdf_nuovo)
                    for pag, t in confronta(pdf_orig, pdf_nuovo)[:25]:
                        print(f"  diversa p{pag}: {t[:70]}")
            except Exception as e:
                codice = 1
                # negli errori COM il testo di Publisher sta nella terza parte
                dettaglio = e.args[2][2] if len(e.args) > 2 and e.args[2] else e
                print(f"[{i}/{len(files)}] ERR {pub.name}: {dettaglio}", flush=True)
                try:
                    app.Name
                except Exception:  # Publisher si e' chiuso: gli altri file ne hanno bisogno
                    termina_publisher()
                    app = apri_publisher()
            finally:
                provvisoria.cleanup()
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(1)
        termina_publisher()
    sys.stdout.flush()
    os._exit(codice)  # pywin32 puo' restare appeso all'uscita dopo aver letto molti oggetti COM


if __name__ == "__main__":
    main()
