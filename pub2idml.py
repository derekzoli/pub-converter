"""Scrive in IDML (formato di scambio di InDesign, lo legge DesignCraft) i dati letti
da pub2sla.py. Le immagini sono incorporate: il file .idml basta da solo.
"""
import base64
import math
import struct
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

PKG = 'xmlns:idPkg="http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging" DOMVersion="16.0"'
TESTA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
GRADI_CORSIVO_SIMULATO = 18.4  # l'inclinazione che Publisher usa sui font senza corsivo
AGGANCIO = {0: "TopAlign", 1: "CenterAlign", 2: "BottomAlign"}
INVOLUCRO = {1: "BoundingBoxTextWrap", 2: "BoundingBoxTextWrap", 4: "JumpObjectTextWrap"}


def n(v) -> str:
    return f"{v:.4f}".rstrip("0").rstrip(".") or "0"


def motivo_tratto(stile_mso, peso):
    """Tratto e spazio in punti dei tratteggi di Publisher, proporzionali allo spessore."""
    unita = {2: [1, 1], 3: [1, 1], 4: [4, 3], 5: [4, 3, 1, 3], 6: [4, 3, 1, 3, 1, 3], 7: [8, 3], 8: [8, 3, 1, 3],
             9: [8, 3, 1, 3, 1, 3], 10: [3, 1], 11: [3, 1, 1, 1], 12: [3, 1, 1, 1, 1, 1]}.get(stile_mso)
    return [round(v * peso, 3) for v in unita] if unita else None


def lato_png(file: Path):
    with open(file, "rb") as f:
        return struct.unpack(">II", f.read(24)[16:])


class ScrittoreIdml:
    def __init__(self, dati, font, cartella: Path):
        self.dati, self.font, self.cartella = dati, font, cartella
        self.colori, self.famiglie, self.storie, self.conta = set(), {}, {}, 0x100
        self.tratteggi = {}  # motivo in punti -> nome dello stile di tratto

    def nuovo_id(self) -> str:
        self.conta += 1
        return f"u{self.conta:x}"

    def colore(self, c) -> str:
        if c == "000000":
            return "Color/Black"
        if c == "ffffff":
            return "Color/Paper"
        self.colori.add(c)
        return f"Color/C_{c}"

    # ------------------------------------------------------------ testo

    def stile_font(self, r):
        """(famiglia, stile, inclinazione): se il font non ha il corsivo lo simula come Publisher."""
        nome = "Tahoma" if r["font"].startswith("MS Shell Dlg") else r["font"]
        corsivo_vero = not r["i"] or self.font.cerca(nome, r["b"], True) is not None
        scelto = self.font.risolvi(nome, r["b"], r["i"] and corsivo_vero)
        parole, chiave = scelto.split(), "".join(c for c in nome.lower() if c.isalnum())
        famiglia, stile = nome, "Regular"
        for k in range(len(parole), 0, -1):
            if "".join(c for c in " ".join(parole[:k]).lower() if c.isalnum()) == chiave:
                famiglia, stile = " ".join(parole[:k]), " ".join(parole[k:]) or "Regular"
                break
        else:  # es. 'AbadiMT-ExtraLight ☞': famiglia registrata con un nome diverso
            famiglia, stile = parole[0], " ".join(p for p in parole[1:] if p.isascii()) or "Regular"
        self.famiglie.setdefault(famiglia, set()).add(stile)
        return famiglia, stile, 0 if corsivo_vero else GRADI_CORSIVO_SIMULATO

    def intervallo(self, r, ls, a_capo) -> str:
        famiglia, stile, inclina = self.stile_font(r)
        a = f'AppliedCharacterStyle="CharacterStyle/$ID/[No character style]" PointSize="{n(r["size"])}"'
        a += f' FontStyle={quoteattr(stile)} FillColor={quoteattr(self.colore(r["color"]))}'
        if r["u"]:
            a += ' Underline="true"'
        if r.get("scala") and r["scala"] != 100:
            a += f' HorizontalScale="{n(r["scala"])}"'
        if inclina:
            a += f' Skew="{n(inclina)}"'
        nome = "Tahoma" if r["font"].startswith("MS Shell Dlg") else r["font"]
        if r["b"] and self.font.cerca(nome, True, False) is None:
            # font senza grassetto: Publisher lo simula, qui con un contorno sottile del carattere
            a += f' StrokeColor={quoteattr(self.colore(r["color"]))} StrokeWeight="{n(r["size"] * 0.025)}"'
        testo, prec_tab, pulito = r["t"], False, []
        for c in testo:
            if c == "\r":
                continue
            if c == "\t":
                if not prec_tab:  # piu' tab di fila vanno a una sola fermata misurata
                    pulito.append("\t")
                prec_tab = True
                continue
            prec_tab = False
            if c in "\n\x0b\x0c":
                pulito.append(" ")  # a capo forzato
            elif ord(c) >= 32:
                pulito.append(c)
        contenuto = f"<Content>{escape(''.join(pulito))}</Content>" if pulito else ""
        br = "<Br/>" if a_capo else ""
        return (f"<CharacterStyleRange {a}><Properties><AppliedFont type=\"string\">{escape(famiglia)}</AppliedFont>"
                f"<Leading type=\"unit\">{n(ls)}</Leading></Properties>{contenuto}{br}</CharacterStyleRange>")

    def paragrafo(self, p, ultimo) -> str:
        allinea = {0: "LeftAlign", 1: "CenterAlign", 2: "RightAlign"}.get(p["align"], "LeftJustified")
        a = (f'AppliedParagraphStyle="ParagraphStyle/$ID/NormalParagraphStyle" Justification="{allinea}" '
             f'FirstLineIndent="{n(p["first"])}" LeftIndent="{n(p["left"])}" RightIndent="{n(p["right"])}" '
             f'SpaceBefore="{n(p["prima"])}" SpaceAfter="{n(p["dopo"])}"')
        prop = ""
        if p["tabs"]:
            voci = "".join(
                '<ListItem type="record"><Alignment type="enumeration">'
                f'{ {1: "CenterAlign", 2: "RightAlign"}.get(al, "LeftAlign") }</Alignment>'
                '<AlignmentCharacter type="string">.</AlignmentCharacter><Leader type="string"></Leader>'
                f'<Position type="unit">{n(pos)}</Position></ListItem>' for pos, al in p["tabs"])
            prop += f'<TabList type="list">{voci}</TabList>'
        if p["bullet"]:
            a += ' BulletsAndNumberingListType="BulletList"'
            codice = 8226 if p["bullet"] in "·" else ord(p["bullet"][0])
            prop += f'<BulletChar BulletCharacterType="UnicodeOnly" BulletCharacterValue="{codice}"/>'
        interno = f"<Properties>{prop}</Properties>" if prop else ""
        runs = p["runs"] or [{"t": "", "font": "Arial", "size": p["size"], "b": False, "i": False, "u": False,
                              "color": "000000", "scala": 100}]
        for i, r in enumerate(runs):
            interno += self.intervallo(r, p["ls"], not ultimo and i == len(runs) - 1)
        return f"<ParagraphStyleRange {a}>{interno}</ParagraphStyleRange>"

    def storia(self, id_storia, t) -> str:
        corpo = "".join(self.paragrafo(p, i == len(t["paras"]) - 1) for i, p in enumerate(t["paras"]))
        return (f'{TESTA}<idPkg:Story {PKG}><Story Self="{id_storia}" AppliedTOCStyle="n" TrackChanges="false" '
                'StoryTitle="$ID/"><StoryPreference OpticalMarginAlignment="false" OpticalMarginSize="12" '
                'FrameType="TextFrameType" StoryOrientation="Horizontal" StoryDirection="LeftToRightDirection"/>'
                f"{corpo}</Story></idPkg:Story>")

    # ------------------------------------------------------------ oggetti

    @staticmethod
    def geometria(x0, y0, x1, y1, aperto=False) -> str:
        punti = [(x0, y0), (x1, y1)] if aperto else [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
        pp = "".join(f'<PathPointType Anchor="{n(x)} {n(y)}" LeftDirection="{n(x)} {n(y)}" '
                     f'RightDirection="{n(x)} {n(y)}"/>' for x, y in punti)
        return (f'<Properties><PathGeometry><GeometryPathType PathOpen="{"true" if aperto else "false"}">'
                f"<PathPointArray>{pp}</PathPointArray></GeometryPathType></PathGeometry></Properties>")

    @staticmethod
    def trasformazione(o, meta_h, cx=None, cy=None) -> str:
        rot = o["rot"] if o["tipo"] != "immagine" else 0
        if not rot:
            return f"1 0 0 1 0 {n(-meta_h)}"
        ang = math.radians(rot)
        a, b = math.cos(ang), math.sin(ang)
        cx, cy = o["x"] + o["w"] / 2, o["y"] + o["h"] / 2
        tx = cx - (a * cx - b * cy)
        ty = cy - (b * cx + a * cy) - meta_h
        return f"{a:.6f} {b:.6f} {-b:.6f} {a:.6f} {n(tx)} {n(ty)}"

    def su_fondo(self, c, trasparenza) -> str:
        """DesignCraft ignora la trasparenza del solo riempimento: colore equivalente sul fondo pagina."""
        if not trasparenza:
            return c
        return "".join(f"{round(int(c[i:i + 2], 16) * (1 - trasparenza) + int(self.fondo[i:i + 2], 16) * trasparenza):02x}"
                       for i in (0, 2, 4))

    def tipo_tratto(self, o) -> str:
        if not o.get("motivo"):
            return "StrokeStyle/$ID/Solid"
        chiave = tuple(o["motivo"])
        nome = self.tratteggi.setdefault(chiave, f"Tratteggio {len(self.tratteggi) + 1}")
        return f"DashedStrokeStyle/{nome}"

    def colore_riempimento(self, o) -> str:
        """Colore finale di un riempimento trasparente sul fondo pagina. Come in Publisher,
        l'ombra e' trasparente quanto l'oggetto e si vede attraverso il riempimento."""
        t, fondo = o.get("fill_t") or 0, self.fondo
        if "ombra" in o and t:
            scuro = (1 - o["ombra"]["t"]) * (1 - t)
            fondo = "".join(f"{round(int(fondo[i:i + 2], 16) * (1 - scuro)):02x}" for i in (0, 2, 4))
        return "".join(f"{round(int(o['fill'][i:i + 2], 16) * (1 - t) + int(fondo[i:i + 2], 16) * t):02x}"
                       for i in (0, 2, 4))

    def aspetto(self, o, meta_h) -> str:
        riempi = self.colore(self.colore_riempimento(o)) if "fill" in o else None
        tratto = self.colore(self.su_fondo(o["line"], o.get("line_t"))) if "line" in o else None
        return (f'ItemLayer="u1" Visible="true" Locked="false" Name={quoteattr(o["nome"])} '
                f'AppliedObjectStyle="ObjectStyle/$ID/[None]" ItemTransform="{self.trasformazione(o, meta_h)}" '
                f'FillColor="{riempi or "Swatch/None"}" FillTint="100" '
                f'StrokeColor="{tratto or "Swatch/None"}" StrokeWeight="{n(o.get("line_w", 0))}" StrokeTint="100" '
                f'StrokeType="{self.tipo_tratto(o)}" StrokeAlignment="CenterAlignment"')

    @staticmethod
    def effetti(o) -> str:
        pezzi = ""
        if "ombra" in o:
            s = o["ombra"]
            distanza = math.hypot(s["x"], s["y"])
            opacita = (1 - s["t"]) * (1 - (o.get("fill_t") or 0)) * 100
            pezzi += ('<TransparencySetting><DropShadowSetting Mode="Drop" BlendMode="Multiply" '
                      f'Opacity="{n(opacita)}" XOffset="{n(s["x"])}" YOffset="{n(s["y"])}" '
                      f'Size="{n(s["blur"])}" Spread="0" Noise="0" KnockedOut="false" '
                      f'EffectColor="Color/Black" Distance="{n(distanza)}"/></TransparencySetting>')
        return pezzi

    @staticmethod
    def scontorno(o) -> str:
        tipo, l, t, r, b = o["wrap"]
        modo = INVOLUCRO.get(tipo, "None")
        if modo == "None":
            l = t = r = b = 0
        elif "contorno" in o:  # gli scostamenti non possono entrare nell'oggetto
            c = o["contorno"]
            l, t = max(0, o["x"] - c[0]), max(0, o["y"] - c[1])
            r, b = max(0, c[2] - o["x"] - o["w"]), max(0, c[3] - o["y"] - o["h"])
        return (f'<TextWrapPreference Inverse="false" ApplyToMasterPageOnly="false" TextWrapSide="BothSides" '
                f'TextWrapMode="{modo}"><Properties><TextWrapOffset Top="{n(t)}" Left="{n(l)}" Bottom="{n(b)}" '
                f'Right="{n(r)}"/></Properties><ContourOption ContourType="SameAsClipping" '
                'IncludeInsideEdges="false"/></TextWrapPreference>')

    def immagine(self, o, meta_h) -> str:
        w, h, x, y = o["w"], o["h"], o["x"], o["y"]
        if o["rot"]:  # l'immagine salvata e' gia' ruotata: cornice sul suo ingombro
            ang = math.radians(o["rot"])
            bw = abs(w * math.cos(ang)) + abs(h * math.sin(ang))
            bh = abs(w * math.sin(ang)) + abs(h * math.cos(ang))
            x, y, w, h = x + w / 2 - bw / 2, y + h / 2 - bh / 2, bw, bh
        file = self.cartella / self.dati["immagini"] / o["file"]
        pw, ph = lato_png(file)
        contenuto = base64.b64encode(file.read_bytes()).decode("ascii")
        sx, sy = w / pw, h / ph
        ppi_x, ppi_y = 72 / sx if sx else 72, 72 / sy if sy else 72  # le linee hanno un lato nullo
        return (f'<Rectangle Self="{self.nuovo_id()}" ContentType="GraphicType" {self.aspetto(o, meta_h)}>'
                f"{self.geometria(x, y, x + w, y + h)}{self.scontorno(o)}"
                f'<Image Self="{self.nuovo_id()}" ActualPpi="{n(ppi_x)} {n(ppi_y)}" '
                f'ItemTransform="{sx:.6f} 0 0 {sy:.6f} {n(x)} {n(y)}"><Properties><Contents><![CDATA[{contenuto}]]></Contents>'
                f'<GraphicBounds Left="0" Top="0" Right="{pw}" Bottom="{ph}"/></Properties>'
                f'<Link Self="{self.nuovo_id()}" AssetURL="$ID/" AssetID="$ID/" LinkResourceURI={quoteattr("file:" + o["file"])} '
                'LinkResourceFormat="$ID/Portable Network Graphics (PNG)" StoredState="Embedded" LinkClassID="35906" '
                'LinkClientID="257" LinkResourceModified="false" LinkObjectModified="false" ShowInUI="true" CanEmbed="true" '
                'CanUnembed="true" CanPackage="true" ImportPolicy="NoAutoImport" ExportPolicy="NoAutoExport"/></Image>'
                "</Rectangle>")

    def casella(self, o, meta_h, cornici) -> str:
        t = o["testo"]
        io = cornici[o["id_pub"]]
        if t and t["paras"] is not None:
            # la catena di caselle collegate condivide una sola storia: la crea la prima
            self.storie[io["storia"]] = self.storia(io["storia"], t)
        m = t["m"] if t else [0, 0, 0, 0]
        colonne = t["cols"] if t else 1
        gap = t["gap"] if t else 0
        aggancio = AGGANCIO.get(t["valign"] if t else 0, "TopAlign")
        inset = "".join(f'<ListItem type="unit">{n(v)}</ListItem>' for v in (m[1], m[0], m[3], m[2]))
        return (f'<TextFrame Self="{io["id"]}" ParentStory="{io["storia"]}" PreviousTextFrame="{io["prec"]}" '
                f'NextTextFrame="{io["succ"]}" ContentType="TextType" {self.aspetto(o, meta_h)}{self.angoli(o)}>'
                f'{self.geometria(o["x"], o["y"], o["x"] + o["w"], o["y"] + o["h"])}{self.effetti(o)}'
                f'<TextFramePreference TextColumnCount="{colonne}" TextColumnGutter="{n(gap)}" '
                f'FirstBaselineOffset="AscentOffset" MinimumFirstBaselineOffset="0" VerticalJustification="{aggancio}" '
                f'IgnoreWrap="{"true" if o.get("ignora_scontorni") else "false"}"><Properties><InsetSpacing type="list">{inset}</InsetSpacing></Properties>'
                f"</TextFramePreference>{self.scontorno(o)}</TextFrame>")

    def forma(self, o, meta_h) -> str:
        x0, y0, x1, y1 = o["x"], o["y"], o["x"] + o["w"], o["y"] + o["h"]
        corpo = f'{self.geometria(x0, y0, x1, y1)}{self.effetti(o)}{self.scontorno(o)}'
        if o["auto"] == 9:
            return f'<Oval Self="{self.nuovo_id()}" ContentType="Unassigned" {self.aspetto(o, meta_h)}>{corpo}</Oval>'
        return f'<Rectangle Self="{self.nuovo_id()}" ContentType="Unassigned" {self.aspetto(o, meta_h)}{self.angoli(o)}>{corpo}</Rectangle>'

    @staticmethod
    def angoli(o) -> str:
        if not o.get("raggio"):
            return ""
        return "".join(f' {lato}CornerOption="RoundedCorner" {lato}CornerRadius="{n(o["raggio"])}"'
                       for lato in ("TopLeft", "TopRight", "BottomLeft", "BottomRight"))

    def linea(self, o, meta_h) -> str:
        x0 = o["x"] + (o["w"] if o["flip_h"] else 0)
        x1 = o["x"] + (0 if o["flip_h"] else o["w"])
        y0 = o["y"] + (o["h"] if o["flip_v"] else 0)
        y1 = o["y"] + (0 if o["flip_v"] else o["h"])
        return (f'<GraphicLine Self="{self.nuovo_id()}" ContentType="Unassigned" {self.aspetto(o, meta_h)}>'
                f'{self.geometria(x0, y0, x1, y1, aperto=True)}{self.effetti(o)}{self.scontorno(o)}</GraphicLine>')

    def oggetto(self, o, meta_h, cornici) -> str:
        if o["tipo"] == "immagine":
            return self.immagine(o, meta_h)
        if o["tipo"] == "linea":
            return self.linea(o, meta_h)
        if o["tipo"] == "forma":
            return self.forma(o, meta_h)
        return self.casella(o, meta_h, cornici)

    def cornici(self) -> dict:
        """Id di ogni casella di testo e, per le caselle collegate, della storia comune."""
        tutte = {o["id_pub"]: o for pag in self.dati["pagine"] for o in pag["oggetti"] if o["tipo"] == "testo"}
        ids = {k: {"id": self.nuovo_id(), "storia": None, "prec": "n", "succ": "n"} for k in tutte}
        for k, o in tutte.items():
            if "successivo" in o:
                ids[k]["succ"] = ids[o["successivo"]]["id"]
            if "precedente" in o:
                ids[k]["prec"] = ids[o["precedente"]]["id"]
        for k, o in tutte.items():
            if "precedente" not in o:
                id_storia = self.nuovo_id()
                if not o["testo"]:  # forma vuota con testo vuoto: storia vuota
                    self.storie[id_storia] = self.storia(id_storia, {"paras": [{
                        "align": 0, "first": 0, "left": 0, "right": 0, "prima": 0, "dopo": 0, "ls": 12, "tabs": [],
                        "bullet": "", "runs": [], "size": 10}]})
                attuale = k
                while True:
                    ids[attuale]["storia"] = id_storia
                    if "successivo" not in tutte[attuale]:
                        break
                    attuale = tutte[attuale]["successivo"]
        return ids

    # ------------------------------------------------------------ pacchetto

    def scrivi(self, dest: Path):
        cornici = self.cornici()
        pagine = self.dati["pagine"]
        pw, ph = pagine[0]["w"], pagine[0]["h"]
        spread = {}
        for i, p in enumerate(pagine):
            meta_h = p["h"] / 2
            id_spread = self.nuovo_id()
            sfondo = next((o for o in p["oggetti"] if o["nome"] == "Sfondo pagina"), None)
            self.fondo = sfondo["fill"] if sfondo else "ffffff"
            oggetti = "".join(self.oggetto(o, meta_h, cornici) for o in p["oggetti"])
            spread[id_spread] = (
                f'{TESTA}<idPkg:Spread {PKG}><Spread Self="{id_spread}" ShowMasterItems="true" PageCount="1" '
                f'BindingLocation="0" AllowPageShuffle="true" ItemTransform="1 0 0 1 0 0">'
                f'<Page Self="{self.nuovo_id()}" Name="{i + 1}" AppliedMaster="n" LayoutRule="Off" OverrideList="" '
                f'GeometricBounds="0 0 {n(p["h"])} {n(p["w"])}" ItemTransform="1 0 0 1 0 {n(-meta_h)}">'
                f'<MarginPreference ColumnCount="1" ColumnGutter="12" Top="0" Bottom="0" Left="0" Right="0" '
                f'ColumnDirection="Horizontal" ColumnsPositions="0 {n(p["w"])}"/></Page>{oggetti}</Spread></idPkg:Spread>')

        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("mimetype", "application/vnd.adobe.indesign-idml-package", compress_type=zipfile.ZIP_STORED)
            z.writestr("META-INF/container.xml", TESTA + '<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="designmap.xml" media-type="text/xml"/></rootfiles></container>')
            z.writestr("designmap.xml", self.designmap(spread))
            z.writestr("Resources/Graphic.xml", self.grafica())
            z.writestr("Resources/Fonts.xml", self.font_xml())
            z.writestr("Resources/Styles.xml", self.stili())
            z.writestr("Resources/Preferences.xml", self.preferenze(pw, ph, len(pagine)))
            for k, v in spread.items():
                z.writestr(f"Spreads/Spread_{k}.xml", v)
            for k, v in self.storie.items():
                z.writestr(f"Stories/Story_{k}.xml", v)

    def designmap(self, spread) -> str:
        fogli = "".join(f'<idPkg:Spread src="Spreads/Spread_{k}.xml"/>' for k in spread)
        storie = "".join(f'<idPkg:Story src="Stories/Story_{k}.xml"/>' for k in self.storie)
        return (f'{TESTA}<?aid style="50" type="document" readerVersion="6.0" featureSet="257" product="16.0(1)" ?>\n'
                f'<Document xmlns:idPkg="http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging" DOMVersion="16.0" Self="d" '
                f'StoryList="{" ".join(self.storie)}" ZeroPoint="0 0" ActiveLayer="u1">'
                '<Language Self="Language/$ID/[No Language]" Name="$ID/[No Language]" SingleQuotes="\'\'" DoubleQuotes="&quot;&quot;" PrimaryLanguageName="$ID/[No Language]" SublanguageName="$ID/[No Language]" Id="0" HyphenationVendor="$ID/" SpellingVendor="$ID/"/>'
                '<Language Self="Language/$ID/Italian" Name="$ID/Italian" SingleQuotes="‘’" DoubleQuotes="“”" PrimaryLanguageName="$ID/Italian" SublanguageName="$ID/" Id="16" HyphenationVendor="Hunspell" SpellingVendor="Hunspell"/>'
                '<idPkg:Graphic src="Resources/Graphic.xml"/><idPkg:Fonts src="Resources/Fonts.xml"/>'
                '<idPkg:Styles src="Resources/Styles.xml"/>'
                '<NumberingList Self="NumberingList/$ID/[Default]" Name="$ID/[Default]" ContinueNumbersAcrossStories="false" ContinueNumbersAcrossDocuments="false"/>'
                '<idPkg:Preferences src="Resources/Preferences.xml"/>'
                '<Layer Self="u1" Name="Livello 1" Visible="true" Locked="false" IgnoreWrap="false" ShowGuides="true" LockGuides="false" UI="true" Expendable="true" Printable="true"><Properties><LayerColor type="enumeration">LightBlue</LayerColor></Properties></Layer>'
                f"{fogli}{storie}</Document>")

    def grafica(self) -> str:
        def colore(self_, nome, valori, spazio="RGB", extra=""):
            return (f'<Color Self="{self_}" Model="Process" Space="{spazio}" ColorValue="{valori}" {extra}'
                    f'AlternateSpace="NoAlternateColor" AlternateColorValue="" Name="{nome}" ColorEditable="true" '
                    'ColorRemovable="true" Visible="true" SwatchCreatorID="7937"/>')

        righe = [colore("Color/Black", "Black", "0 0 0 100", "CMYK", 'ColorOverride="Specialblack" '),
                 colore("Color/Paper", "Paper", "0 0 0 0", "CMYK", 'ColorOverride="Specialpaper" ')]
        for c in sorted(self.colori):
            righe.append(colore(f"Color/C_{c}", f"C_{c}", f"{int(c[:2], 16)} {int(c[2:4], 16)} {int(c[4:], 16)}"))
        righe.append('<Swatch Self="Swatch/None" Name="None" ColorEditable="false" ColorRemovable="false" Visible="true" SwatchCreatorID="7937"/>')
        righe.append('<StrokeStyle Self="StrokeStyle/$ID/Solid" Name="$ID/Solid"/>')
        for motivo, nome in self.tratteggi.items():
            righe.append(f'<DashedStrokeStyle Self="DashedStrokeStyle/{nome}" Name="{nome}" '
                         f'DashArray="{" ".join(n(v) for v in motivo)}" StrokeCornerAdjustment="None" EndCap="ButtEndCap"/>')
        return f"{TESTA}<idPkg:Graphic {PKG}>{''.join(righe)}</idPkg:Graphic>"

    def font_xml(self) -> str:
        fam = []
        for i, (famiglia, stili) in enumerate(sorted(self.famiglie.items())):
            voci = "".join(
                f'<Font Self={quoteattr(f"di{i}Fontn{famiglia} {s}")} FontFamily={quoteattr(famiglia)} '
                f'Name={quoteattr(f"{famiglia} {s}")} PostScriptName="$ID/" Status="Installed" '
                f'FontStyleName={quoteattr(s)} FontType="OpenTypeTT" WritingScript="0" '
                f'FullName={quoteattr(f"{famiglia} {s}")} FullNameNative={quoteattr(f"{famiglia} {s}")} '
                f'FontStyleNameNative={quoteattr(s)} PlatformName="$ID/" Version="$ID/" TypekitID="$ID/"/>'
                for s in sorted(stili))
            fam.append(f'<FontFamily Self="di{i}" Name={quoteattr(famiglia)}>{voci}</FontFamily>')
        return f"{TESTA}<idPkg:Fonts {PKG}>{''.join(fam)}</idPkg:Fonts>"

    @staticmethod
    def stili() -> str:
        # MinimumWordSpacing 100: come Publisher allarga gli spazi ma non li stringe.
        # DesignCraft usa i punti di sillabazione (U+00AD) solo col compositore di paragrafo
        base = ('Justification="LeftAlign" LeftIndent="0" RightIndent="0" FirstLineIndent="0" SpaceBefore="0" SpaceAfter="0" '
                'Composer="HL Composer" Hyphenation="true" MinimumWordSpacing="100" DesiredWordSpacing="100" '
                'MaximumWordSpacing="1000" MinimumLetterSpacing="0" DesiredLetterSpacing="0" MaximumLetterSpacing="0" '
                'MinimumGlyphScaling="100" DesiredGlyphScaling="100" MaximumGlyphScaling="100" AutoLeading="120" '
                'SingleWordJustification="LeftAlign" KeepFirstLines="1" KeepLastLines="1" KeepLinesTogether="false" '
                'BulletsAndNumberingListType="NoList" BulletsTextAfter="^t" '
                'AppliedLanguage="$ID/Italian" PointSize="10" FontStyle="Regular" FillColor="Color/Black" FillTint="100"')
        return (f'{TESTA}<idPkg:Styles {PKG}>'
                '<RootCharacterStyleGroup Self="u81"><CharacterStyle Self="CharacterStyle/$ID/[No character style]" Imported="false" Name="$ID/[No character style]"/></RootCharacterStyleGroup>'
                f'<RootParagraphStyleGroup Self="u82"><ParagraphStyle Self="ParagraphStyle/$ID/[No paragraph style]" Name="$ID/[No paragraph style]" Imported="false" {base}><Properties><AppliedFont type="string">Arial</AppliedFont><Leading type="enumeration">Auto</Leading></Properties></ParagraphStyle>'
                '<ParagraphStyle Self="ParagraphStyle/$ID/NormalParagraphStyle" Name="$ID/NormalParagraphStyle" Imported="false"><Properties><BasedOn type="string">$ID/[No paragraph style]</BasedOn></Properties></ParagraphStyle></RootParagraphStyleGroup>'
                '<RootObjectStyleGroup Self="u83"><ObjectStyle Self="ObjectStyle/$ID/[None]" Name="$ID/[None]" EnableFill="false" EnableStroke="false" EnableParagraphStyle="false" EnableTextFrameGeneralOptions="false"/></RootObjectStyleGroup>'
                "</idPkg:Styles>")

    @staticmethod
    def preferenze(pw, ph, pagine) -> str:
        return (f'{TESTA}<idPkg:Preferences {PKG}>'
                f'<DocumentPreference PageHeight="{n(ph)}" PageWidth="{n(pw)}" PagesPerDocument="{pagine}" FacingPages="false" '
                'DocumentBleedTopOffset="0" DocumentBleedBottomOffset="0" DocumentBleedInsideOrLeftOffset="0" DocumentBleedOutsideOrRightOffset="0" '
                'DocumentBleedUniformSize="true" PageBinding="LeftToRight" ColumnDirection="Horizontal" Intent="PrintIntent" CreatePrimaryTextFrame="false"/>'
                f'<MarginPreference ColumnCount="1" ColumnGutter="12" Top="0" Bottom="0" Left="0" Right="0" ColumnDirection="Horizontal" ColumnsPositions="0 {n(pw)}"/>'
                '<ViewPreference HorizontalMeasurementUnits="Millimeters" VerticalMeasurementUnits="Millimeters" RulerOrigin="SpreadOrigin" PointsPerInch="72"/>'
                '<TextPreference SuperscriptSize="58.3" SuperscriptPosition="33.3" SubscriptSize="58.3" SubscriptPosition="33.3" ZOrderTextWrap="true"/>'
                "</idPkg:Preferences>")


def scrivi_idml(dati, font, cartella: Path, dest: Path, adattatore):
    """adattatore: lo Scrittore di pub2sla, per riusare il calcolo dei contorni di scontorno."""
    for p in dati["pagine"]:
        for o in p["oggetti"]:
            o["pagina_w"] = p["w"]
        adattatore.adatta_contorni(p["oggetti"], per_scribus=False)
        solo_scontorni_sopra(p["oggetti"])
        for o in p["oggetti"]:
            ancora_in_alto(o)
    ScrittoreIdml(dati, font, cartella).scrivi(dest)


def si_toccano(a, b) -> bool:
    return a["x"] < b["x"] + b["w"] and b["x"] < a["x"] + a["w"] and a["y"] < b["y"] + b["h"] and b["y"] < a["y"] + a["h"]


def solo_scontorni_sopra(oggetti):
    """In Publisher il testo gira solo attorno a cio' che gli sta sopra, in DesignCraft
    anche attorno a cio' che gli sta sotto. Toglie lo scontorno agli oggetti che non
    hanno testo sotto di se'; una casella che ha sotto un ostacolo ancora necessario ad
    altri e sopra nessuno ignora gli scontorni."""
    testo = [o.get("testo") and o["testo"].get("paras") is not None for o in oggetti]
    for j, o in enumerate(oggetti):
        if o["wrap"][0] and not any(testo[i] and si_toccano(oggetti[i], o) for i in range(j)):
            o["wrap"] = [0, 0, 0, 0, 0]
            o.pop("contorno", None)
    for j, o in enumerate(oggetti):
        if not testo[j]:
            continue
        sotto = any(x["wrap"][0] and si_toccano(x, o) for x in oggetti[:j])
        sopra = any(x["wrap"][0] and si_toccano(x, o) for x in oggetti[j + 1:])
        if sotto and not sopra:
            o["ignora_scontorni"] = True


def ancora_in_alto(o):
    """Un testo ancorato in basso o al centro sparisce se non entra nella casella:
    in quel caso lo ancoro in alto e lo tengo al suo posto con il margine superiore."""
    t = o.get("testo") if o["tipo"] == "testo" else None
    if not t or not t.get("ingombri") or not t.get("paras") or t["cols"] != 1:
        return
    if not t["valign"]:
        # DesignCraft nasconde le righe che non entrano, Publisher le lascia sporgere. La casella
        # si puo' alzare solo se la cosa non si vede: senza sfondo, bordo o scontorno
        if not (o["wrap"][0] or "fill" in o or "line" in o or "successivo" in o or o["rot"]):
            o["h"] = max(o["h"], max(b[3] for b in t["ingombri"]) - o["y"] + t["m"][3] + 4)
        return
    ls = max(p["ls"] for p in t["paras"])
    if o["h"] - t["m"][1] - t["m"][3] >= ls * len(t["ingombri"]):
        return
    primo = t["ingombri"][0]
    s = (primo[3] - primo[1]) / 1.55  # l'ingombro va da 0,8 a 0,75 corpi attorno alla linea di base
    t["m"][1] = max(0, primo[1] + 0.8 * s - 0.95 * s - o["y"])
    t["valign"] = 0
