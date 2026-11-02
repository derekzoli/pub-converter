"""Finestra per convertire file Publisher (.pub) in documenti Scribus o DesignCraft, in italiano e inglese.

Il lavoro vero lo fa pub2sla.py, avviato come processo separato: se Publisher
si blocca, la finestra resta utilizzabile.
"""
import ctypes
import json
import msvcrt
import os
import queue
import re
import stat
import string
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import winreg
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

ESEGUIBILE = getattr(sys, "frozen", False)
CARTELLA = Path(sys.executable if ESEGUIBILE else __file__).resolve().parent
if not ESEGUIBILE:  # l'eseguibile usa solo la sua copia, non i .py accanto
    sys.path.insert(0, str(CARTELLA))
import pub2sla  # noqa: E402
from pub2sla import DESIGNCRAFT, SCRIBUS, pid_publisher  # noqa: E402

VERSIONE = "1.2"
PREFISSO_CREATO = "  creato: "
ARG_CONVERTI = "--converti"
if ESEGUIBILE:
    AVVIO_CONVERSIONE = [sys.executable, ARG_CONVERTI]  # l'eseguibile richiama se stesso
else:
    # pythonw non ha un output da leggere
    AVVIO_CONVERSIONE = [str(Path(sys.executable).with_name("python.exe")), "-u", str(CARTELLA / "pub2sla.py")]

ESTENSIONI = {"designcraft": ".designcraft", "scribus": ".sla"}
IMPOSTAZIONI = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir())) / "Convertitore PUB" / "impostazioni.json"
SECONDI_PER_FILE = 20
RIGA_FILE = re.compile(r"^\[(\d+)/(\d+)\] (OK|ERR) ")
CODICE_BLOCCATO = 2  # uscita di pub2sla quando Publisher non risponde
CONTINUO, SISTEMA_ATTIVO = 0x80000000, 0x00000001

# cartelle di sistema e copie interne dei programmi: i documenti dell'utente non stanno li'
ESCLUSE = {"windows", "program files", "program files (x86)", "programdata", "$recycle.bin",
           "system volume information", "recovery", "perflogs", "$winreagent", "$windows.~bt",
           "$windows.~ws", "msocache", "appdata", ".ssh"}
FIRMA_OLE = bytes.fromhex("D0CF11E0A1B11AE1")  # anche le chiavi SSH hanno estensione .pub
SOLO_ONLINE = 0x00400000 | 0x00040000 | stat.FILE_ATTRIBUTE_OFFLINE
DISCO_RIMOVIBILE, DISCO_FISSO = 2, 3

TESTI = {
    "it": {
        "titolo": "Convertitore Publisher",
        "lingua": "Lingua",
        "info": "Info",
        "formato": "Formato",
        "dc": "DesignCraft (.designcraft)",
        "sc": "Scribus (.sla)",
        "cosa": "Cosa convertire",
        "scegli_file": "Scegli file .pub…",
        "scegli_cartella": "Scegli una cartella…",
        "cerca": "Cerca e converti…",
        "nessuno": "Nessun file scelto",
        "n_file": "{n} file selezionati",
        "cartella_file": "{nome}: {n} file .pub",
        "file_pub": "File Publisher",
        "tutti_file": "Tutti i file",
        "titolo_scegli_file": "Scegli i file Publisher",
        "titolo_scegli_cartella": "Scegli la cartella: verranno convertiti tutti i .pub, sottocartelle comprese",
        "nessun_file": "Nessun file",
        "nessun_pub": "In questa cartella non ci sono file .pub.",
        "dove": "Dove salvare",
        "accanto": "Accanto a ogni file originale, in una cartella <nome>_{formato}",
        "in_cartella": "In questa cartella:",
        "sfoglia": "Sfoglia…",
        "titolo_dest": "Dove salvare i documenti convertiti",
        "verifica": "Confronta il risultato con l'originale (più lento)",
        "raccogli": "Copia i font usati accanto al documento (cartella Document Fonts)",
        "tieni_pdf": "Tieni una copia PDF del documento originale",
        "annulla": "Annulla",
        "converti": "Converti",
        "risultati": "Documenti creati (doppio clic per aprire)",
        "apri_doc": "Apri documento",
        "apri_cart": "Apri cartella",
        "manca_prog": "Attenzione: programma non trovato in {p}\n",
        "manca_pub": "Attenzione: Microsoft Publisher non risulta installato: senza non si può convertire.\n",
        "manca_cart": "Manca la cartella",
        "manca_cart_testo": "Scegli la cartella in cui salvare.",
        "in_corso": "Conversione in corso…",
        "convertiti": "Convertiti: {x}",
        "annullamento": "Annullamento…",
        "annullato": "Annullato. Documenti già creati: {n}.",
        "fatto": "Fatto. Documenti creati: {n}.",
        "con_errori": "Finito con errori. Documenti creati: {n}: vedi i messaggi sopra.",
        "fatto_errori": "Finito. Documenti creati: {n}. File non convertiti: {e}.",
        "elenco_errori": "\nFile non convertiti:\n",
        "bloccato": "Publisher non ha risposto: file saltato.",
        "chiudi_titolo": "Conversione in corso",
        "chiudi_testo": "La conversione è ancora in corso. Vuoi interromperla e chiudere?",
        # ricerca
        "r_titolo": "Cerca e converti",
        "r_spiega": "Cerca i file .pub su dischi interni ed esterni, escluse le cartelle di sistema e i dischi di "
                    "rete. Il documento nuovo viene salvato nella stessa cartella dell'originale, che non viene "
                    "toccato. Per Scribus serve anche la cartella <nome>_immagini creata accanto al documento.",
        "r_spiega_cartella": "Cerca i file .pub in {cartella} e nelle sue sottocartelle. Il documento nuovo viene "
                             "salvato nella stessa cartella dell'originale, che non viene toccato. Per Scribus serve "
                             "anche la cartella <nome>_immagini creata accanto al documento.",
        "a_domanda": "Dove vuoi cercare i file .pub?",
        "a_cartella": "Cartella selezionata",
        "a_tutto": "Tutto il PC",
        "r_cerco": "Ricerca in corso: {n} file trovati. {cartella}",
        "r_fine": "Trovati {n} file .pub. Selezionati: {s}. Solo online: {o}.",
        "r_ferma": "Ferma la ricerca",
        "r_tutti": "Seleziona tutti",
        "r_nessuno": "Deseleziona tutti",
        "r_converti": "Converti i selezionati",
        "r_chiudi": "Chiudi",
        "r_nome": "File",
        "r_cartella": "Cartella",
        "r_stato": "Stato",
        "r_gia": "già convertito",
        "r_online": "solo online",
        "r_nuovo": "da convertire",
        "r_conferma_titolo": "Conferma",
        "r_conferma": "Convertire {n} file in {formato}?\n\nTempo stimato: circa {t}. Publisher resta occupato "
                      "per tutta la durata.\nGli originali non vengono modificati.",
        "r_conf_online": "\n\n{o} file sono solo su OneDrive e verranno scaricati.",
        "r_conf_gia": "\n\n{g} file già convertiti verranno sovrascritti.",
        "ore": "{o} ore e {m} minuti",
        "minuti": "{m} minuti",
        "r_scelta": "Cerca e converti: {n} file",
        # pagina info
        "i_titolo": "Informazioni",
        "i_chiudi": "Chiudi",
        "i_testo": [
            ("h", "Convertitore Publisher"),
            ("p", f"Versione {VERSIONE}"),
            ("h", "Il progetto"),
            ("p", "Microsoft ha terminato il supporto a Publisher, e chi ha molti documenti .pub rischia di non "
                  "poterli più modificare. Questo programma li converte in documenti che si lavorano con Scribus, "
                  "gratuito e open source, oppure con il nuovo DesignCraft: il testo resta testo vero, che va a capo "
                  "e scorre, e cornici, linee e bordi restano oggetti modificabili."),
            ("p", "Gli originali non vengono mai modificati."),
            ("h", "Chi sono"),
            ("p", "Derek Zoli."),
            ("h", "Come è stata fatta"),
            ("p", "L'applicazione è stata ideata e collaudata da me e scritta con Claude Code, l'assistente di "
                  "programmazione di Anthropic. È scritta in Python: Publisher apre il documento e fornisce testi, "
                  "forme e immagini; il PDF esportato da Publisher dà le posizioni esatte delle righe; infine "
                  "vengono scritti i formati nativi: .sla per Scribus e .designcraft per DesignCraft."),
            ("h", "Cosa serve installato"),
            ("p", "• Windows\n"
                  "• Microsoft Publisher funzionante: è indispensabile, è lui che legge i file .pub\n"
                  "• Scribus (provato con la 1.6.7) oppure DesignCraft (provato con la 0.14), secondo il formato "
                  "scelto, installati nelle cartelle dei programmi\n"
                  "• I font usati nei documenti: se mancano, l'aspetto cambia. Il programma può copiarli nella "
                  "cartella Document Fonts accanto al documento: DesignCraft li carica da solo, per Scribus vanno "
                  "installati. Rispetta le licenze dei font se condividi quella cartella\n"
                  "Python non serve: l'eseguibile lo contiene già."),
            ("h", "Limiti da conoscere"),
            ("p", "• Tabelle, WordArt e forme complesse diventano immagini, quindi non modificabili\n"
                  "• Gli a capo possono cambiare rispetto all'originale, soprattutto in DesignCraft: "
                  "ricontrolla sempre i documenti importanti\n"
                  "• L'eseguibile non ha una firma digitale: Windows o l'antivirus possono segnalarlo"),
            ("h", "Codice e licenza"),
            ("p", "Codice sorgente: github.com/derekzoli/pub-converter\n"
                  "Licenza: GNU AGPL-3.0, la stessa di PyMuPDF, la libreria usata per leggere i PDF."),
            ("p", "Progetto indipendente, non collegato a Microsoft, Scribus o ArtCraft."),
        ],
    },
    "en": {
        "titolo": "Publisher Converter",
        "lingua": "Language",
        "info": "About",
        "formato": "Format",
        "dc": "DesignCraft (.designcraft)",
        "sc": "Scribus (.sla)",
        "cosa": "What to convert",
        "scegli_file": "Choose .pub files…",
        "scegli_cartella": "Choose a folder…",
        "cerca": "Search and convert…",
        "nessuno": "No file selected",
        "n_file": "{n} files selected",
        "cartella_file": "{nome}: {n} .pub files",
        "file_pub": "Publisher files",
        "tutti_file": "All files",
        "titolo_scegli_file": "Choose the Publisher files",
        "titolo_scegli_cartella": "Choose the folder: all .pub files will be converted, subfolders included",
        "nessun_file": "No files",
        "nessun_pub": "There are no .pub files in this folder.",
        "dove": "Where to save",
        "accanto": "Next to each original file, in a <name>_{formato} folder",
        "in_cartella": "In this folder:",
        "sfoglia": "Browse…",
        "titolo_dest": "Where to save the converted documents",
        "verifica": "Compare the result with the original (slower)",
        "raccogli": "Copy the fonts used next to the document (Document Fonts folder)",
        "tieni_pdf": "Keep a PDF copy of the original document",
        "annulla": "Cancel",
        "converti": "Convert",
        "risultati": "Documents created (double-click to open)",
        "apri_doc": "Open document",
        "apri_cart": "Open folder",
        "manca_prog": "Warning: program not found in {p}\n",
        "manca_pub": "Warning: Microsoft Publisher does not appear to be installed: conversion needs it.\n",
        "manca_cart": "Folder missing",
        "manca_cart_testo": "Choose the folder to save to.",
        "in_corso": "Converting…",
        "convertiti": "Converted: {x}",
        "annullamento": "Cancelling…",
        "annullato": "Cancelled. Documents already created: {n}.",
        "fatto": "Done. Documents created: {n}.",
        "con_errori": "Finished with errors. Documents created: {n}: see the messages above.",
        "fatto_errori": "Finished. Documents created: {n}. Files not converted: {e}.",
        "elenco_errori": "\nFiles not converted:\n",
        "bloccato": "Publisher did not respond: file skipped.",
        "chiudi_titolo": "Conversion running",
        "chiudi_testo": "The conversion is still running. Stop it and close?",
        "r_titolo": "Search and convert",
        "r_spiega": "Searches for .pub files on internal and external drives, excluding system folders and network "
                    "drives. The new document is saved in the same folder as the original, which is left untouched. "
                    "Scribus also needs the <name>_immagini folder created next to the document.",
        "r_spiega_cartella": "Searches for .pub files in {cartella} and its subfolders. The new document is saved in "
                             "the same folder as the original, which is left untouched. Scribus also needs the "
                             "<name>_immagini folder created next to the document.",
        "a_domanda": "Where do you want to search for .pub files?",
        "a_cartella": "Selected folder",
        "a_tutto": "Entire PC",
        "r_cerco": "Searching: {n} files found. {cartella}",
        "r_fine": "Found {n} .pub files. Selected: {s}. Online only: {o}.",
        "r_ferma": "Stop searching",
        "r_tutti": "Select all",
        "r_nessuno": "Select none",
        "r_converti": "Convert selected",
        "r_chiudi": "Close",
        "r_nome": "File",
        "r_cartella": "Folder",
        "r_stato": "Status",
        "r_gia": "already converted",
        "r_online": "online only",
        "r_nuovo": "to convert",
        "r_conferma_titolo": "Confirm",
        "r_conferma": "Convert {n} files to {formato}?\n\nEstimated time: about {t}. Publisher stays busy the whole "
                      "time.\nThe originals are not changed.",
        "r_conf_online": "\n\n{o} files are on OneDrive only and will be downloaded.",
        "r_conf_gia": "\n\n{g} files already converted will be overwritten.",
        "ore": "{o} hours and {m} minutes",
        "minuti": "{m} minutes",
        "r_scelta": "Search and convert: {n} files",
        "i_titolo": "About",
        "i_chiudi": "Close",
        "i_testo": [
            ("h", "Publisher Converter"),
            ("p", f"Version {VERSIONE}"),
            ("h", "The project"),
            ("p", "Microsoft has ended support for Publisher, and anyone with many .pub documents risks no longer "
                  "being able to edit them. This program converts them into documents you can work on with Scribus, "
                  "which is free and open source, or with the new DesignCraft: text stays real text that wraps and "
                  "flows, and frames, lines and borders stay editable objects."),
            ("p", "The originals are never changed."),
            ("h", "Who I am"),
            ("p", "Derek Zoli."),
            ("h", "How it was made"),
            ("p", "I came up with and tested the application, and it was written with Claude Code, Anthropic's "
                  "coding assistant. It is written in Python: Publisher opens the document and provides text, shapes "
                  "and images; the PDF exported by Publisher gives the exact position of each line; finally the "
                  "native formats are written: .sla for Scribus, .designcraft for DesignCraft."),
            ("h", "What you need installed"),
            ("p", "• Windows\n"
                  "• A working copy of Microsoft Publisher: this is essential, as it is what reads the .pub files\n"
                  "• Scribus (tested with 1.6.7) or DesignCraft (tested with 0.14), depending on the chosen format, "
                  "installed in the program folders\n"
                  "• The fonts used in your documents: if they are missing, the look changes. The program can copy "
                  "them to a Document Fonts folder next to the document: DesignCraft loads them by itself, for "
                  "Scribus they must be installed. Respect the font licenses if you share that folder\n"
                  "You do not need Python: the executable already includes it."),
            ("h", "Known limits"),
            ("p", "• Tables, WordArt and complex shapes become images, so they cannot be edited\n"
                  "• Line breaks may differ from the original, especially in DesignCraft: always check important "
                  "documents\n"
                  "• The executable has no digital signature: Windows or your antivirus may flag it"),
            ("h", "Source code and license"),
            ("p", "Source code: github.com/derekzoli/pub-converter\n"
                  "License: GNU AGPL-3.0, the same as PyMuPDF, the library used to read PDFs."),
            ("p", "Independent project, not affiliated with Microsoft, Scribus or ArtCraft."),
        ],
    },
}

# i messaggi di pub2sla nascono in italiano
MESSAGGI_EN = [
    ("  avviso:", "  warning:"),
    ("Publisher non risponde: interrotto.", "Publisher is not responding: stopped."),
    ("tabella importata come immagine (non modificabile)", "table imported as an image (not editable)"),
    ("importata come immagine", "imported as an image"),
    ("forma con testo importata come rettangolo", "shape with text imported as a rectangle"),
    ("forma tipo", "shape type"),
    ("sfondo pagina tipo", "page background type"),
    ("importato solo il colore principale", "only the main color imported"),
    ("nessun corsivo, usato", "no italic, using"),
    ("nessun grassetto, usato", "no bold, using"),
    ("font non trovato, usato Arial", "font not found, using Arial"),
    ("font non trovato sul PC, non copiato", "font not found on this PC, not copied"),
    ("la licenza del font non permette la copia", "the font license does not allow copying"),
    ("font copiati:", "fonts copied:"),
    ("file in Document Fonts", "files in Document Fonts"),
    ("font non copiati:", "fonts not copied:"),
    ("parole al loro posto", "words in place"),
    ("scarto verticale medio", "average vertical offset"),
    ("  diversa p", "  different p"),
    ("pagine diverse", "different pages"),
    ("DesignCraft non ha creato", "DesignCraft did not create"),
]


def lingua_iniziale():
    try:
        return json.loads(IMPOSTAZIONI.read_text(encoding="utf-8"))["lingua"]
    except (OSError, ValueError, KeyError):
        return "it" if ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0xFF == 0x10 else "en"


def salva_lingua(lingua):
    try:
        IMPOSTAZIONI.parent.mkdir(parents=True, exist_ok=True)
        IMPOSTAZIONI.write_text(json.dumps({"lingua": lingua}), encoding="utf-8")
    except OSError:
        pass  # alla prossima apertura si riparte dalla lingua di Windows


def publisher_installato():
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Publisher.Application"))
        return True
    except OSError:
        return False


def dischi():
    """Dischi interni ed esterni: niente rete ne' lettori ottici."""
    presenti = ctypes.windll.kernel32.GetLogicalDrives()
    radici = [f"{l}:\\" for i, l in enumerate(string.ascii_uppercase) if presenti >> i & 1]
    return [r for r in radici if ctypes.windll.kernel32.GetDriveTypeW(r) in (DISCO_RIMOVIBILE, DISCO_FISSO)]


def controlla_pub(percorso):
    """Dati del file se e' un documento Publisher; un file solo online non si apre per non scaricarlo."""
    try:
        attributi = os.stat(percorso).st_file_attributes
        online = bool(attributi & SOLO_ONLINE)
        if not online:
            with open(percorso, "rb") as f:
                if f.read(len(FIRMA_OLE)) != FIRMA_OLE:
                    return None
    except OSError:
        return None
    return {"percorso": percorso, "online": online, "gia": False, "scelto": True}


def cerca_pub(radici, fermo, avanza):
    for radice in radici:
        for cartella, sotto, nomi in os.walk(radice):
            if fermo.is_set():
                return
            sotto[:] = [d for d in sotto if d.lower() not in ESCLUSE]
            avanza(cartella, None)
            for nome in nomi:
                if nome.lower().endswith(".pub") and not nome.startswith("~$"):
                    voce = controlla_pub(os.path.join(cartella, nome))
                    if voce:
                        avanza(cartella, voce)


def converti_da_eseguibile():
    """Processo di conversione dell'eseguibile: senza console l'output va collegato a mano."""
    if sys.stdout is None:
        ctypes.windll.kernel32.GetStdHandle.restype = ctypes.c_void_p
        maniglia = ctypes.windll.kernel32.GetStdHandle(-11)
        sys.stdout = sys.stderr = open(msvcrt.open_osfhandle(maniglia, os.O_WRONLY), "w",
                                       encoding="utf-8", buffering=1)
    del sys.argv[1]
    pub2sla.main()


class Ricerca:
    """Finestra di Cerca e converti: trova i file, li mostra e fa togliere quelli da non convertire."""

    def __init__(self, app, cartella=None):
        self.app = app
        self.radici = [cartella] if cartella else dischi()
        self.voci = []
        self.fermo = threading.Event()
        self.coda = queue.Queue()
        t = app.t
        self.win = win = tk.Toplevel(app.root)
        win.title(t("r_titolo"))
        win.geometry("860x560")
        win.minsize(640, 420)
        win.transient(app.root)
        win.grab_set()
        win.protocol("WM_DELETE_WINDOW", self.chiudi)
        pad = {"padx": 10, "pady": 4}

        spiegazione = t("r_spiega_cartella", cartella=cartella) if cartella else t("r_spiega")
        ttk.Label(win, text=spiegazione, wraplength=830, justify="left").pack(fill="x", **pad)
        riga = ttk.Frame(win)
        riga.pack(fill="x", **pad)
        ttk.Label(riga, text=t("formato") + ":").pack(side="left")
        ttk.Radiobutton(riga, text="DesignCraft", variable=app.formato, value="designcraft",
                        command=self.ricalcola).pack(side="left", padx=8)
        ttk.Radiobutton(riga, text="Scribus", variable=app.formato, value="scribus",
                        command=self.ricalcola).pack(side="left")
        self.ferma_btn = ttk.Button(riga, text=t("r_ferma"), command=self.fermo.set)
        self.ferma_btn.pack(side="right")

        self.stato = ttk.Label(win, text="")
        self.stato.pack(fill="x", padx=10)

        cornice = ttk.Frame(win)
        cornice.pack(fill="both", expand=True, **pad)
        self.tabella = ttk.Treeview(cornice, columns=("nome", "cartella", "stato"), selectmode="none")
        self.tabella.heading("#0", text="")
        self.tabella.column("#0", width=34, stretch=False, anchor="center")
        self.tabella.heading("nome", text=t("r_nome"), anchor="w")
        self.tabella.column("nome", width=250)
        self.tabella.heading("cartella", text=t("r_cartella"), anchor="w")
        self.tabella.column("cartella", width=420)
        self.tabella.heading("stato", text=t("r_stato"), anchor="w")
        self.tabella.column("stato", width=120, stretch=False)
        scorri = ttk.Scrollbar(cornice, command=self.tabella.yview)
        self.tabella.configure(yscrollcommand=scorri.set)
        self.tabella.pack(side="left", fill="both", expand=True)
        scorri.pack(side="right", fill="y")
        self.tabella.bind("<Button-1>", self.clic)

        ttk.Checkbutton(win, text=t("raccogli"), variable=app.raccogli_lista).pack(anchor="w", **pad)
        ttk.Checkbutton(win, text=t("tieni_pdf"), variable=app.tieni_pdf).pack(anchor="w", padx=10)

        azioni = ttk.Frame(win)
        azioni.pack(fill="x", **pad)
        ttk.Button(azioni, text=t("r_tutti"), command=lambda: self.segna(True)).pack(side="left")
        ttk.Button(azioni, text=t("r_nessuno"), command=lambda: self.segna(False)).pack(side="left", padx=6)
        ttk.Button(azioni, text=t("r_chiudi"), command=self.chiudi).pack(side="right")
        self.converti_btn = ttk.Button(azioni, text=t("r_converti"), command=self.conferma, state="disabled")
        self.converti_btn.pack(side="right", padx=6)

        self.cartella = ""
        self.lavoro = threading.Thread(target=self.cerca, daemon=True)
        self.lavoro.start()
        win.after(100, self.controlla)

    def cerca(self):
        cerca_pub(self.radici, self.fermo, lambda cartella, voce: self.coda.put((cartella, voce)))

    def controlla(self):
        if not self.win.winfo_exists():
            return
        while not self.coda.empty():
            self.cartella, voce = self.coda.get()
            if voce:
                self.aggiungi(voce)
        if self.lavoro.is_alive() or not self.coda.empty():
            self.stato.configure(text=self.app.t("r_cerco", n=len(self.voci), cartella=self.cartella[-90:]))
            self.win.after(100, self.controlla)
        else:
            self.ferma_btn.configure(state="disabled")
            self.aggiorna_stato()

    def aggiungi(self, voce):
        voce["gia"] = self.convertito(voce)
        voce["scelto"] = not voce["gia"]
        self.voci.append(voce)
        percorso = Path(voce["percorso"])
        self.tabella.insert("", "end", iid=str(len(self.voci) - 1), text=self.casella(voce),
                            values=(percorso.name, str(percorso.parent), self.descrivi(voce)))

    def convertito(self, voce):
        return Path(voce["percorso"]).with_suffix(ESTENSIONI[self.app.formato.get()]).exists()

    @staticmethod
    def casella(voce):
        return "☑" if voce["scelto"] else "☐"

    def descrivi(self, voce):
        chiave = "r_gia" if voce["gia"] else "r_online" if voce["online"] else "r_nuovo"
        return self.app.t(chiave)

    def ricalcola(self):
        for i, voce in enumerate(self.voci):
            voce["gia"] = self.convertito(voce)
            voce["scelto"] = not voce["gia"]
            self.ridisegna(i)
        self.aggiorna_stato()

    def ridisegna(self, i):
        voce = self.voci[i]
        self.tabella.item(str(i), text=self.casella(voce))
        self.tabella.set(str(i), "stato", self.descrivi(voce))

    def clic(self, evento):
        riga = self.tabella.identify_row(evento.y)
        if riga:
            voce = self.voci[int(riga)]
            voce["scelto"] = not voce["scelto"]
            self.ridisegna(int(riga))
            self.aggiorna_stato()

    def segna(self, valore):
        for i, voce in enumerate(self.voci):
            voce["scelto"] = valore
            self.ridisegna(i)
        self.aggiorna_stato()

    def scelti(self):
        return [v for v in self.voci if v["scelto"]]

    def aggiorna_stato(self):
        if self.lavoro.is_alive():
            return
        scelti = self.scelti()
        online = sum(v["online"] for v in scelti)
        self.stato.configure(text=self.app.t("r_fine", n=len(self.voci), s=len(scelti), o=online))
        self.converti_btn.configure(state="normal" if scelti else "disabled")

    def durata(self, quanti):
        minuti = max(round(quanti * SECONDI_PER_FILE / 60), 2)  # per pochi file conta anche l'avvio di Publisher
        if minuti < 60:
            return self.app.t("minuti", m=minuti)
        return self.app.t("ore", o=minuti // 60, m=minuti % 60)

    def conferma(self):
        t, scelti = self.app.t, self.scelti()
        testo = t("r_conferma", n=len(scelti), formato=self.app.nome_formato(), t=self.durata(len(scelti)))
        online = sum(v["online"] for v in scelti)
        gia = sum(v["gia"] for v in scelti)
        if online:
            testo += t("r_conf_online", o=online)
        if gia:
            testo += t("r_conf_gia", g=gia)
        if messagebox.askyesno(t("r_conferma_titolo"), testo, parent=self.win):
            percorsi = [v["percorso"] for v in scelti]
            self.chiudi()
            self.app.avvia_lista(percorsi)

    def chiudi(self):
        self.fermo.set()
        self.win.grab_release()
        self.win.destroy()


class App:
    def __init__(self, root):
        self.root = root
        self.lingua = lingua_iniziale()
        self.scelti = []
        self.descrizione = ("nessuno", {})
        self.creati = []
        self.processo = None
        self.annullato = False
        self.gia_aperti = set()
        self.righe = queue.Queue()
        self.lista = None  # file ancora da convertire con Cerca e converti
        self.parte = []
        self.fatti_parte = 0
        self.totale = 0
        self.fatti = 0
        self.errori = []
        self.file_elenco = None
        self.testo_log = ""
        self.testo_stato = ("", {})
        self.formato = tk.StringVar(value="designcraft")
        self.accanto = tk.BooleanVar(value=True)
        self.destinazione = tk.StringVar()
        self.verifica = tk.BooleanVar(value=False)
        self.raccogli = tk.BooleanVar(value=True)
        self.raccogli_lista = tk.BooleanVar(value=False)  # con Cerca e converti aggiunge una cartella per ogni posto
        self.tieni_pdf = tk.BooleanVar(value=False)
        self.nome_lingua = tk.StringVar()

        root.geometry("720x680")
        root.minsize(600, 560)
        root.protocol("WM_DELETE_WINDOW", self.chiudi)
        self.costruisci()
        if not publisher_installato():
            self.scrivi(self.t("manca_pub"))

    def t(self, chiave, **valori):
        testo = TESTI[self.lingua][chiave]
        return testo.format(**valori) if valori else testo

    def nome_formato(self):
        return "DesignCraft" if self.formato.get() == "designcraft" else "Scribus"

    def in_corso(self):
        return self.processo is not None and self.processo.poll() is None

    def costruisci(self):
        """Disegna la finestra nella lingua scelta; anche cambiando lingua le scelte fatte restano."""
        for figlio in self.root.winfo_children():
            figlio.destroy()
        t, root = self.t, self.root
        root.title(t("titolo"))
        pad = {"padx": 10, "pady": 4}

        testata = ttk.Frame(root)
        testata.pack(fill="x", **pad)
        ttk.Label(testata, text=t("lingua") + ":").pack(side="left")
        self.nome_lingua.set("Italiano" if self.lingua == "it" else "English")
        scelta = ttk.Combobox(testata, textvariable=self.nome_lingua, values=("Italiano", "English"),
                              state="readonly", width=10)
        scelta.pack(side="left", padx=6)
        scelta.bind("<<ComboboxSelected>>", self.cambia_lingua)
        self.scelta_lingua = scelta
        ttk.Button(testata, text=t("info"), command=self.mostra_info).pack(side="right")

        formato = ttk.LabelFrame(root, text=t("formato"))
        formato.pack(fill="x", **pad)
        riga = ttk.Frame(formato)
        riga.pack(fill="x", padx=6, pady=6)
        ttk.Radiobutton(riga, text=t("dc"), variable=self.formato, value="designcraft",
                        command=self.aggiorna_formato).pack(side="left")
        ttk.Radiobutton(riga, text=t("sc"), variable=self.formato, value="scribus",
                        command=self.aggiorna_formato).pack(side="left", padx=16)

        sorgente = ttk.LabelFrame(root, text=t("cosa"))
        sorgente.pack(fill="x", **pad)
        riga = ttk.Frame(sorgente)
        riga.pack(fill="x", padx=6, pady=(6, 2))
        ttk.Button(riga, text=t("scegli_file"), command=self.scegli_file).pack(side="left")
        ttk.Button(riga, text=t("scegli_cartella"), command=self.scegli_cartella).pack(side="left", padx=6)
        self.cerca_btn = ttk.Button(riga, text=t("cerca"), command=self.cerca)
        self.cerca_btn.pack(side="left")
        self.etichetta = ttk.Label(sorgente, foreground="#666666" if not self.scelti else "#000000")
        self.etichetta.pack(fill="x", padx=8, pady=(0, 6))

        dest = ttk.LabelFrame(root, text=t("dove"))
        dest.pack(fill="x", **pad)
        self.opzione_accanto = ttk.Radiobutton(dest, variable=self.accanto, value=True, command=self.aggiorna)
        self.opzione_accanto.pack(anchor="w", padx=6)
        riga = ttk.Frame(dest)
        riga.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Radiobutton(riga, text=t("in_cartella"), variable=self.accanto, value=False,
                        command=self.aggiorna).pack(side="left")
        self.campo = ttk.Entry(riga, textvariable=self.destinazione)
        self.campo.pack(side="left", fill="x", expand=True, padx=6)
        self.sfoglia = ttk.Button(riga, text=t("sfoglia"), command=self.scegli_destinazione)
        self.sfoglia.pack(side="left")

        ttk.Checkbutton(root, text=t("raccogli"), variable=self.raccogli).pack(anchor="w", padx=10, pady=(4, 0))
        ttk.Checkbutton(root, text=t("tieni_pdf"), variable=self.tieni_pdf).pack(anchor="w", padx=10, pady=(4, 0))
        azioni = ttk.Frame(root)
        azioni.pack(fill="x", **pad)
        ttk.Checkbutton(azioni, text=t("verifica"), variable=self.verifica).pack(side="left")
        self.annulla_btn = ttk.Button(azioni, text=t("annulla"), command=self.annulla, state="disabled")
        self.annulla_btn.pack(side="right")
        self.converti_btn = ttk.Button(azioni, text=t("converti"), command=self.avvia,
                                       state="normal" if self.scelti else "disabled")
        self.converti_btn.pack(side="right", padx=6)

        self.barra = ttk.Progressbar(root, mode="determinate")
        self.barra.pack(fill="x", **pad)
        self.stato = ttk.Label(root, text="")
        self.stato.pack(fill="x", padx=10)

        cornice = ttk.Frame(root)
        cornice.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(cornice, height=8, wrap="word", state="disabled", font=("Consolas", 9))
        scorri = ttk.Scrollbar(cornice, command=self.log.yview)
        self.log.configure(yscrollcommand=scorri.set)
        self.log.pack(side="left", fill="both", expand=True)
        scorri.pack(side="right", fill="y")

        risultati = ttk.LabelFrame(root, text=t("risultati"))
        risultati.pack(fill="x", **pad)
        self.elenco = tk.Listbox(risultati, height=4)
        self.elenco.pack(fill="x", padx=6, pady=(6, 2))
        self.elenco.bind("<Double-Button-1>", lambda _: self.apri_documento())
        for percorso in self.creati:
            self.elenco.insert("end", Path(percorso).name)
        riga = ttk.Frame(risultati)
        riga.pack(fill="x", padx=6, pady=(2, 6))
        ttk.Button(riga, text=t("apri_doc"), command=self.apri_documento).pack(side="left")
        ttk.Button(riga, text=t("apri_cart"), command=self.apri_cartella).pack(side="left", padx=6)

        self.etichetta.configure(text=t(self.descrizione[0], **self.descrizione[1]))
        self.stato.configure(text=t(self.testo_stato[0], **self.testo_stato[1]) if self.testo_stato[0] else "")
        self.log.configure(state="normal")
        self.log.insert("end", self.testo_log)
        self.log.configure(state="disabled")
        self.aggiorna()
        self.aggiorna_formato()

    def cambia_lingua(self, _evento=None):
        lingua = "it" if self.nome_lingua.get() == "Italiano" else "en"
        if lingua != self.lingua:
            self.lingua = lingua
            salva_lingua(lingua)
            self.costruisci()

    def mostra_info(self):
        t = self.t
        win = tk.Toplevel(self.root)
        win.title(t("i_titolo"))
        win.geometry("620x560")
        win.transient(self.root)
        cornice = ttk.Frame(win)
        cornice.pack(fill="both", expand=True, padx=10, pady=(10, 4))
        testo = tk.Text(cornice, wrap="word", relief="flat", padx=8, pady=8, font=("Segoe UI", 10),
                        background=win.cget("background"))
        scorri = ttk.Scrollbar(cornice, command=testo.yview)
        testo.configure(yscrollcommand=scorri.set)
        testo.pack(side="left", fill="both", expand=True)
        scorri.pack(side="right", fill="y")
        testo.tag_configure("h", font=("Segoe UI", 11, "bold"), spacing1=10, spacing3=3)
        testo.tag_configure("p", spacing3=6)
        for tipo, contenuto in TESTI[self.lingua]["i_testo"]:
            testo.insert("end", contenuto + "\n", tipo)
        testo.configure(state="disabled")
        ttk.Button(win, text=t("i_chiudi"), command=win.destroy).pack(pady=(4, 10))

    # ------------------------------------------------------------ scelte

    def imposta(self, percorsi, chiave, **valori):
        self.scelti = percorsi
        self.descrizione = (chiave, valori)
        self.etichetta.configure(text=self.t(chiave, **valori), foreground="#000000")
        self.converti_btn.configure(state="normal")

    def scegli_file(self):
        t = self.t
        files = filedialog.askopenfilenames(title=t("titolo_scegli_file"),
                                            filetypes=[(t("file_pub"), "*.pub"), (t("tutti_file"), "*.*")])
        if files:
            self.imposta(list(files), "n_file", n=len(files))

    def scegli_cartella(self):
        cartella = filedialog.askdirectory(title=self.t("titolo_scegli_cartella"))
        if not cartella:
            return
        quanti = sum(1 for _ in Path(cartella).rglob("*.pub"))
        if not quanti:
            messagebox.showinfo(self.t("nessun_file"), self.t("nessun_pub"))
            return
        self.imposta([cartella], "cartella_file", nome=Path(cartella).name, n=quanti)

    def scegli_destinazione(self):
        cartella = filedialog.askdirectory(title=self.t("titolo_dest"))
        if cartella:
            self.destinazione.set(cartella)

    def cerca(self):
        if len(self.scelti) == 1 and Path(self.scelti[0]).is_dir():
            self.chiedi_dove(self.scelti[0])
        else:
            Ricerca(self)

    def chiedi_dove(self, cartella):
        t = self.t
        win = tk.Toplevel(self.root)
        win.title(t("r_titolo"))
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        ttk.Label(win, text=t("a_domanda")).pack(anchor="w", padx=14, pady=(14, 2))
        ttk.Label(win, text=str(Path(cartella)), foreground="#666666", wraplength=460).pack(anchor="w", padx=14)
        riga = ttk.Frame(win)
        riga.pack(padx=14, pady=14)

        def scegli(*argomenti):
            win.destroy()
            Ricerca(self, *argomenti)

        ttk.Button(riga, text=t("a_cartella"), command=lambda: scegli(cartella)).pack(side="left")
        ttk.Button(riga, text=t("a_tutto"), command=scegli).pack(side="left", padx=6)
        ttk.Button(riga, text=t("annulla"), command=win.destroy).pack(side="left")

    def aggiorna_formato(self):
        scelto = self.formato.get()
        self.opzione_accanto.configure(text=self.t("accanto", formato=scelto))
        programma = DESIGNCRAFT if scelto == "designcraft" else SCRIBUS
        avviso = self.t("manca_prog", p=programma)
        if not Path(programma).exists() and avviso not in self.testo_log:
            self.scrivi(avviso)

    def aggiorna(self):
        stato = "disabled" if self.accanto.get() else "normal"
        self.campo.configure(state=stato)
        self.sfoglia.configure(state=stato)

    # ------------------------------------------------------------ conversione

    def avvia(self):
        comando = [*AVVIO_CONVERSIONE, *self.scelti, "--formato", self.formato.get()]
        if not self.accanto.get():
            if not self.destinazione.get():
                messagebox.showwarning(self.t("manca_cart"), self.t("manca_cart_testo"))
                return
            comando += ["--out", self.destinazione.get()]
        if self.verifica.get():
            comando.append("--verifica")
        if self.raccogli.get():
            comando.append("--font")
        if self.tieni_pdf.get():
            comando.append("--pdf")
        self.lista = None
        self.prepara()
        self.lancia(comando)

    def avvia_lista(self, percorsi):
        """Cerca e converti: ogni documento va nella cartella del suo originale."""
        self.lista = percorsi
        self.totale = len(percorsi)
        self.fatti = 0
        self.errori = []
        self.imposta([], "r_scelta", n=len(percorsi))
        self.prepara()
        self.barra.stop()
        self.barra.configure(mode="determinate", maximum=self.totale, value=0)
        self.lancia_parte()

    def lancia_parte(self):
        self.parte = self.lista
        self.fatti_parte = 0
        descrittore, self.file_elenco = tempfile.mkstemp(suffix=".txt", prefix="convertitore-pub-")
        with os.fdopen(descrittore, "w", encoding="utf-8") as f:
            f.write("\n".join(self.parte))
        font = ["--font"] if self.raccogli_lista.get() else []
        pdf = ["--pdf"] if self.tieni_pdf.get() else []
        self.lancia([*AVVIO_CONVERSIONE, "--elenco", self.file_elenco, "--accanto", "--formato", self.formato.get(), *font, *pdf])

    def prepara(self):
        self.creati.clear()
        self.elenco.delete(0, "end")
        self.testo_log = ""
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.annullato = False
        self.gia_aperti = pid_publisher()
        self.converti_btn.configure(state="disabled")
        self.cerca_btn.configure(state="disabled")
        self.scelta_lingua.configure(state="disabled")
        self.annulla_btn.configure(state="normal")
        self.mostra_stato("in_corso")
        self.barra.configure(mode="indeterminate", value=0)
        self.barra.start(12)
        ctypes.windll.kernel32.SetThreadExecutionState(CONTINUO | SISTEMA_ATTIVO)  # il PC non va in sospensione

    def lancia(self, comando):
        ambiente = dict(os.environ, PYTHONIOENCODING="utf-8")
        self.processo = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                         encoding="utf-8", errors="replace", env=ambiente, cwd=CARTELLA,
                                         creationflags=subprocess.CREATE_NO_WINDOW)
        threading.Thread(target=self.leggi_output, args=(self.processo,), daemon=True).start()
        self.root.after(100, self.controlla)

    def leggi_output(self, processo):
        for riga in processo.stdout:
            self.righe.put(riga)
        self.righe.put(None)

    def controlla(self):
        finito = False
        while not self.righe.empty():
            riga = self.righe.get()
            if riga is None:
                finito = True
            elif riga.startswith(PREFISSO_CREATO):
                self.aggiungi_risultato(riga[len(PREFISSO_CREATO):].strip())
            else:
                self.mostra(riga)
        if finito:
            self.fine()
        else:
            self.root.after(100, self.controlla)

    def mostra(self, riga):
        if "FutureWarning" in riga or riga.startswith("  warnings.warn"):
            return
        m = RIGA_FILE.match(riga)
        if m:
            if self.lista is None:
                self.mostra_stato("convertiti", x=f"{m[1]}/{m[2]}")
            else:
                if m[3] == "ERR":
                    self.errori.append(self.parte[self.fatti_parte])
                self.fatti_parte += 1
                self.fatti += 1
                self.barra.configure(value=self.fatti)
                self.mostra_stato("convertiti", x=f"{self.fatti}/{self.totale}")
                riga = RIGA_FILE.sub(f"[{self.fatti}/{self.totale}] {m[3]} ", riga)
        self.scrivi(self.traduci(riga))

    def traduci(self, riga):
        if self.lingua == "en":
            for italiano, inglese in MESSAGGI_EN:
                riga = riga.replace(italiano, inglese)
        return riga

    def mostra_stato(self, chiave, **valori):
        self.testo_stato = (chiave, valori)
        self.stato.configure(text=self.t(chiave, **valori))

    def scrivi(self, testo):
        self.testo_log += testo
        self.log.configure(state="normal")
        self.log.insert("end", testo)
        self.log.see("end")
        self.log.configure(state="disabled")

    def aggiungi_risultato(self, percorso):
        self.creati.append(percorso)
        self.elenco.insert("end", Path(percorso).name)
        self.elenco.selection_clear(0, "end")
        self.elenco.selection_set("end")
        self.elenco.see("end")

    def annulla(self):
        if self.in_corso():
            self.annullato = True
            self.annulla_btn.configure(state="disabled")
            self.mostra_stato("annullamento")
            subprocess.run(["taskkill", "/T", "/F", "/PID", str(self.processo.pid)],
                           capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, check=False)

    def chiudi_publisher_nascosto(self):
        # Publisher nascosto non e' figlio del convertitore: va chiuso a parte
        for pid in pid_publisher() - self.gia_aperti:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True,
                           creationflags=subprocess.CREATE_NO_WINDOW, check=False)

    def togli_elenco(self):
        if self.file_elenco:
            Path(self.file_elenco).unlink(missing_ok=True)
            self.file_elenco = None

    def fine(self):
        codice = self.processo.wait()
        self.togli_elenco()
        if self.lista is not None and not self.annullato and self.fatti_parte < len(self.parte):
            # Publisher si e' bloccato o il processo e' caduto: si salta il file e si continua coi successivi
            self.errori.append(self.parte[self.fatti_parte])
            self.fatti += 1
            self.scrivi(f"[{self.fatti}/{self.totale}] ERR {Path(self.parte[self.fatti_parte]).name}: "
                        f"{self.t('bloccato')}\n")
            self.barra.configure(value=self.fatti)
            self.lista = self.parte[self.fatti_parte + 1:]
            if self.lista:
                self.lancia_parte()
                return
        self.barra.stop()
        self.barra.configure(mode="determinate", value=0)
        self.processo = None
        self.converti_btn.configure(state="normal" if self.scelti else "disabled")
        self.cerca_btn.configure(state="normal")
        self.scelta_lingua.configure(state="readonly")
        self.annulla_btn.configure(state="disabled")
        ctypes.windll.kernel32.SetThreadExecutionState(CONTINUO)
        if self.annullato:
            self.chiudi_publisher_nascosto()
            self.mostra_stato("annullato", n=len(self.creati))
        elif self.lista is not None:
            if self.errori:
                self.scrivi(self.t("elenco_errori") + "".join(f"  {e}\n" for e in self.errori))
                self.mostra_stato("fatto_errori", n=len(self.creati), e=len(self.errori))
            else:
                self.mostra_stato("fatto", n=len(self.creati))
        elif codice == 0:
            self.mostra_stato("fatto", n=len(self.creati))
        else:
            self.mostra_stato("con_errori", n=len(self.creati))

    def chiudi(self):
        if self.in_corso():
            if not messagebox.askyesno(self.t("chiudi_titolo"), self.t("chiudi_testo")):
                return
            subprocess.run(["taskkill", "/T", "/F", "/PID", str(self.processo.pid)],
                           capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, check=False)
            self.chiudi_publisher_nascosto()
        self.togli_elenco()
        self.root.destroy()

    # ------------------------------------------------------------ risultati

    def selezionato(self):
        scelta = self.elenco.curselection()
        return self.creati[scelta[0]] if scelta else None

    def apri_documento(self):
        documento = self.selezionato()
        if documento:
            programma = SCRIBUS if documento.lower().endswith(".sla") else DESIGNCRAFT
            subprocess.Popen([programma, documento])

    def apri_cartella(self):
        documento = self.selezionato()
        if documento:
            os.startfile(Path(documento).parent)


if __name__ == "__main__":
    if ESEGUIBILE:
        # Scribus e DesignCraft non devono ereditare le librerie dell'eseguibile
        ctypes.windll.kernel32.SetDllDirectoryW(None)
    if sys.argv[1:2] == [ARG_CONVERTI]:
        converti_da_eseguibile()
    finestra = tk.Tk()
    finestra.iconbitmap(default=str(Path(getattr(sys, "_MEIPASS", CARTELLA)) / "icona.ico"))
    App(finestra)
    finestra.mainloop()
