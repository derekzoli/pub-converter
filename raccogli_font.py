"""Copia accanto al documento i file dei font che usa, nella cartella letta da DesignCraft.

I nomi nel registro di Windows sono tradotti e quelli dei font cloud di Office sono
numeri: la famiglia si legge quindi dentro ogni file.
"""
import os
import shutil
import struct
import winreg
from pathlib import Path

CARTELLA_FONT = "Document Fonts"
ESTENSIONI = {".ttf", ".otf", ".ttc", ".otc"}
CHIAVE_REGISTRO = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
_catalogo = None


def norm(s: str) -> str:
    return "".join(c for c in s.lower() if c.isalnum())


def leggi_nomi(dati, offset):
    """Famiglia, famiglia tipografica, nome completo e licenza di un font dentro il file."""
    n_tabelle = struct.unpack_from(">H", dati, offset + 4)[0]
    tabelle = {}
    for i in range(n_tabelle):
        voce = offset + 12 + 16 * i
        tabelle[dati[voce:voce + 4]] = struct.unpack_from(">II", dati, voce + 8)
    nomi = {1: set(), 16: set(), 4: set()}
    inizio = tabelle[b"name"][0]
    quanti, testi = struct.unpack_from(">xxHH", dati, inizio)
    for i in range(quanti):
        piattaforma, _, _, id_nome, lunghezza, posizione = struct.unpack_from(">6H", dati, inizio + 6 + 12 * i)
        if id_nome in nomi:
            grezzo = dati[inizio + testi + posizione:inizio + testi + posizione + lunghezza]
            testo = grezzo.decode("utf-16-be" if piattaforma != 1 else "latin-1", "ignore")
            if norm(testo):
                nomi[id_nome].add(norm(testo))
    licenza = struct.unpack_from(">H", dati, tabelle[b"OS/2"][0] + 8)[0] if b"OS/2" in tabelle else 0
    return nomi, licenza


def leggi_file(percorso):
    dati = percorso.read_bytes()
    if dati[:4] in (b"ttcf",):
        quanti = struct.unpack_from(">I", dati, 8)[0]
        offsets = struct.unpack_from(f">{quanti}I", dati, 12)
    else:
        offsets = (0,)
    return [leggi_nomi(dati, o) for o in offsets]


def file_installati():
    cartelle = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
    locale = os.environ.get("LOCALAPPDATA")
    if locale:
        cartelle += [Path(locale) / "Microsoft/Windows/Fonts", Path(locale) / "Microsoft/FontCache/4/CloudFonts"]
    trovati = {p for c in cartelle if c.is_dir() for p in c.rglob("*") if p.suffix.lower() in ESTENSIONI}
    for radice in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):  # font installati altrove
        try:
            with winreg.OpenKey(radice, CHIAVE_REGISTRO) as chiave:
                for i in range(winreg.QueryInfoKey(chiave)[1]):
                    p = Path(winreg.EnumValue(chiave, i)[1])
                    if p.is_absolute() and p.suffix.lower() in ESTENSIONI and p.exists():
                        trovati.add(p)
        except OSError:
            pass
    return trovati


def catalogo():
    global _catalogo
    if _catalogo is None:
        _catalogo = []
        for percorso in file_installati():
            try:
                _catalogo += [(percorso, nomi, licenza) for nomi, licenza in leggi_file(percorso)]
            except (OSError, struct.error, KeyError):
                pass  # file non leggibile: come se non fosse installato
    return _catalogo


def famiglie_usate(dati) -> set:
    usate = set()
    if isinstance(dati, dict):
        if isinstance(dati.get("font"), str):
            nome = dati["font"]
            usate.add("Tahoma" if nome.startswith("MS Shell Dlg") else nome)
        for valore in dati.values():
            usate |= famiglie_usate(valore)
    elif isinstance(dati, (list, tuple)):
        for valore in dati:
            usate |= famiglie_usate(valore)
    return usate


def file_della_famiglia(famiglia):
    """Tutti gli stili: anche quelli non usati oggi servono quando si modifica il testo."""
    chiave = norm(famiglia)
    for id_nome in (1, 16, 4):
        scelti = [(p, licenza) for p, nomi, licenza in catalogo() if chiave in nomi[id_nome]]
        if scelti:
            return scelti
    return []


def vietato_copiare(licenza) -> bool:
    return licenza & 0x000E == 0x0002  # "Restricted": gli altri bit, se presenti, lo annullano


def raccogli(dati, cartella: Path, nome_documento: str) -> list:
    """Copia i font in <cartella>/Document Fonts; restituisce i messaggi da mostrare."""
    destinazione = cartella / CARTELLA_FONT
    copiati, non_trovati, con_licenza = set(), [], []
    for famiglia in sorted(famiglie_usate(dati)):
        scelti = file_della_famiglia(famiglia)
        if not scelti:
            non_trovati.append(famiglia)
        permessi = [p for p, licenza in scelti if not vietato_copiare(licenza)]
        if scelti and not permessi:
            con_licenza.append(famiglia)
        for percorso in permessi:
            finale = destinazione / percorso.name
            destinazione.mkdir(parents=True, exist_ok=True)
            if not finale.exists():
                shutil.copy2(percorso, finale)
            copiati.add(percorso.name)

    messaggi = [f"  font copiati: {len(copiati)} file in {CARTELLA_FONT}"] if copiati else []
    messaggi += [f"  avviso: {f}: font non trovato sul PC, non copiato" for f in non_trovati]
    messaggi += [f"  avviso: {f}: la licenza del font non permette la copia" for f in con_licenza]
    elenco = cartella / f"{nome_documento}_font_mancanti.txt"
    if non_trovati or con_licenza:
        righe = [f"{f}: not found on this PC / non trovato sul PC" for f in non_trovati]
        righe += [f"{f}: license does not allow copying / la licenza non permette la copia" for f in con_licenza]
        elenco.write_text("\n".join(righe) + "\n", encoding="utf-8")
    else:
        elenco.unlink(missing_ok=True)
    return messaggi
