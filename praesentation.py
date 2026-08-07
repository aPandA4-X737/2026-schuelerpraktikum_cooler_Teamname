"""
Erzeugt die Projektpraesentation als PowerPoint-Datei.

    venv/bin/python praesentation.py

Erzeugt: Satelliten-Telemetrie.pptx

Die Folien werden vollstaendig aus Vektorformen gebaut, es sind keine Bilder
eingebettet. Schrift und Farben stehen als Konstanten oben und lassen sich an
einer Stelle aendern.
"""

import random

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ----------------------------------------------------------------------------
# Gestaltung
# ----------------------------------------------------------------------------

FONT = "Helvetica Neue"   # auf Windows ggf. auf "Segoe UI" oder "Arial" aendern
MONO = "Menlo"            # auf Windows ggf. auf "Consolas" aendern

INK = RGBColor(0x07, 0x0B, 0x18)      # Folienhintergrund, fast schwarzes Marineblau
PANEL = RGBColor(0x0F, 0x17, 0x29)    # Karten
PANEL2 = RGBColor(0x14, 0x1D, 0x33)   # Karten, eine Stufe heller
BORDER = RGBColor(0x22, 0x30, 0x4C)   # Rahmen und Trennlinien
TXT = RGBColor(0xE8, 0xED, 0xF7)      # Haupttext
DIM = RGBColor(0x8E, 0x9D, 0xBB)      # Nebentext
FAINT = RGBColor(0x5A, 0x6A, 0x89)    # sehr zurueckhaltender Text

ACCENT = RGBColor(0x38, 0xBD, 0xF8)   # Hauptakzent, Himmelblau
AMBER = RGBColor(0xFB, 0xBF, 0x24)    # Hinweis, Warnung
VIOLET = RGBColor(0xA7, 0x8B, 0xFA)   # Verwaltung
GREEN = RGBColor(0x34, 0xD3, 0x99)    # erledigt, Datenbank
RED = RGBColor(0xF8, 0x71, 0x71)      # Fehler

# Farben der echten Webapp. Die Oberflaeche ist hell, deshalb heben sich die
# Mockups auf den dunklen Folien deutlich ab und sehen aus wie ein Screenshot.
# Werte uebernommen aus webapp/src/stylesheet.css.
APP_BG = RGBColor(0xF8, 0xFA, 0xFC)      # --grau-50, Seitenhintergrund
APP_CARD = RGBColor(0xFF, 0xFF, 0xFF)    # --weiss, Karten und Tabelle
APP_LINE = RGBColor(0xE8, 0xED, 0xF3)    # --grau-200, Rahmen
APP_TXT = RGBColor(0x0F, 0x17, 0x2A)     # --grau-900, Text
APP_DIM = RGBColor(0x64, 0x74, 0x8B)     # --grau-500, Nebentext
APP_MUTE = RGBColor(0x94, 0xA3, 0xB8)    # --grau-400, Label
APP_BLUE = RGBColor(0x25, 0x63, 0xEB)    # --blau-600, Akzent
APP_BLUE_D = RGBColor(0x1D, 0x4E, 0xD8)  # --blau-700, Badge-Text
APP_BLUE_L = RGBColor(0xEF, 0xF6, 0xFF)  # --blau-50, Badge-Flaeche
APP_BAR = RGBColor(0x3B, 0x82, 0xF6)     # --blau-500, Balkenfuellung
APP_BAR_BG = RGBColor(0xE8, 0xED, 0xF3)  # Balkenschiene
APP_MARK = RGBColor(0xDB, 0xEA, 0xFE)    # --blau-100, Suchtreffer-Hervorhebung
CHROME = RGBColor(0xDD, 0xE4, 0xEE)      # Browser-Titelleiste

SW, SH = 13.333, 7.5                  # Foliengroesse in Zoll (16:9)
LM = 0.85                             # linker und rechter Rand
CW = SW - 2 * LM                      # nutzbare Breite

TITEL = "Satelliten-Telemetrie"
TEAM = "Team „cooler Teamname“"


# ----------------------------------------------------------------------------
# Bausteine
# ----------------------------------------------------------------------------

def rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75, shape=MSO_SHAPE.RECTANGLE,
         radius=None):
    box = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    box.shadow.inherit = False
    if fill is None:
        box.fill.background()
    else:
        box.fill.solid()
        box.fill.fore_color.rgb = fill
    if line is None:
        box.line.fill.background()
    else:
        box.line.color.rgb = line
        box.line.width = Pt(line_w)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        box.adjustments[0] = radius
    box.text_frame.word_wrap = True
    return box


def panel(slide, x, y, w, h, fill=PANEL, line=BORDER, radius=0.045):
    return rect(slide, x, y, w, h, fill=fill, line=line,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)


def text(slide, x, y, w, h, content, size=14, color=TXT, bold=False, font=FONT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.2, space_after=0,
         italic=False, caps=False):
    """Setzt Text. content ist ein String oder eine Liste von Absaetzen."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0

    absaetze = content if isinstance(content, (list, tuple)) else [content]
    for i, inhalt in enumerate(absaetze):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.alignment = align
        para.line_spacing = spacing
        if space_after:
            para.space_after = Pt(space_after)
        run = para.add_run()
        run.text = (inhalt.upper() if caps else inhalt)
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font
        run.font.color.rgb = color
        if caps:
            run.font._rPr.set("spc", "160")  # Laufweite fuer Kapitaelchen-Optik
    return box


def line_h(slide, x, y, w, color=BORDER, thickness=0.011):
    return rect(slide, x, y, w, thickness, fill=color)


def arrow(slide, x1, y1, x2, y2, color=ACCENT, width=1.25, dashed=False, head=False):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    ln = conn.line._get_or_add_ln()
    if dashed:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": "sysDash"}))
    if head:
        ln.append(ln.makeelement(qn("a:headEnd"),
                                 {"type": "triangle", "w": "med", "len": "med"}))
    ln.append(ln.makeelement(qn("a:tailEnd"),
                             {"type": "triangle", "w": "med", "len": "med"}))
    return conn


def stars(slide, anzahl=70, seed=7, xmax=SW, ymax=SH):
    """Sparsames Sternenfeld: kleine, unterschiedlich helle Punkte."""
    rnd = random.Random(seed)
    farben = [RGBColor(0x2A, 0x3A, 0x5A), RGBColor(0x3D, 0x51, 0x7A),
              RGBColor(0x5B, 0x72, 0xA0), RGBColor(0x8B, 0xA3, 0xCC)]
    for _ in range(anzahl):
        d = rnd.choice([0.018, 0.022, 0.028, 0.038])
        farbe = farben[min(3, [0.018, 0.022, 0.028, 0.038].index(d))]
        rect(slide, rnd.uniform(0.1, xmax - 0.1), rnd.uniform(0.1, ymax - 0.1),
             d, d, fill=farbe, shape=MSO_SHAPE.OVAL)


def neue_folie(prs, notiz=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SW, SH, fill=INK)
    if notiz:
        slide.notes_slide.notes_text_frame.text = notiz
    return slide


def kopf(slide, eyebrow, titel, farbe=ACCENT):
    text(slide, LM, 0.52, CW, 0.24, eyebrow, size=9.5, color=farbe, bold=True, caps=True)
    text(slide, LM, 0.82, CW, 0.5, titel, size=27, color=TXT, bold=True)
    line_h(slide, LM, 1.44, 1.5, color=farbe, thickness=0.028)


def fuss(slide, nummer):
    text(slide, LM, SH - 0.52, CW * 0.6, 0.22, f"{TITEL}  ·  {TEAM}",
         size=8.5, color=FAINT)
    text(slide, SW - LM - 1.0, SH - 0.52, 1.0, 0.22, f"{nummer:02d}",
         size=8.5, color=FAINT, align=PP_ALIGN.RIGHT)


def karte(slide, x, y, w, h, titel, zeilen, farbe=ACCENT, titel_size=13, zeile_size=10.5):
    """Karte mit farbigem Streifen links, Titel und Aufzaehlung."""
    panel(slide, x, y, w, h)
    rect(slide, x + 0.001, y + 0.18, 0.055, h - 0.36, fill=farbe)
    text(slide, x + 0.28, y + 0.2, w - 0.5, 0.3, titel, size=titel_size, color=TXT, bold=True)
    if zeilen:
        text(slide, x + 0.28, y + 0.62, w - 0.5, h - 0.8, list(zeilen),
             size=zeile_size, color=DIM, spacing=1.3, space_after=5)


def gitter(slide, x, y, breiten, kopfzeile, zeilen, zeilen_h=0.4, size=10.5,
           kopf_farbe=ACCENT, farben=None):
    """Eigene Tabelle aus Formen, damit sie zum dunklen Layout passt.

    farben: optionale Liste, eine Farbe je Datenzeile fuer die erste Spalte.
    """
    gesamt = sum(breiten)
    text_y = y
    # Kopfzeile
    xi = x
    for spalte, bw in zip(kopfzeile, breiten):
        text(slide, xi, text_y, bw - 0.12, 0.26, spalte, size=8.5, color=kopf_farbe,
             bold=True, caps=True)
        xi += bw
    y_cursor = y + 0.34
    line_h(slide, x, y_cursor, gesamt, color=kopf_farbe, thickness=0.014)
    y_cursor += 0.06

    for i, zeile in enumerate(zeilen):
        if i % 2 == 0:
            rect(slide, x - 0.12, y_cursor - 0.03, gesamt + 0.24, zeilen_h, fill=PANEL2)
        xi = x
        for j, (zelle, bw) in enumerate(zip(zeile, breiten)):
            farbe = TXT if j == 0 else DIM
            if farben and j == 0:
                farbe = farben[i]
            text(slide, xi, y_cursor + 0.04, bw - 0.12, zeilen_h - 0.06, zelle,
                 size=size, color=farbe, bold=(j == 0), spacing=1.05)
            xi += bw
        y_cursor += zeilen_h
    return y_cursor


def code(slide, x, y, w, zeilen, size=10, farbe=ACCENT):
    h = 0.32 + len(zeilen) * (size * 1.55 / 72)
    panel(slide, x, y, w, h, fill=RGBColor(0x0A, 0x11, 0x21))
    rect(slide, x + 0.001, y + 0.14, 0.045, h - 0.28, fill=farbe)
    text(slide, x + 0.26, y + 0.16, w - 0.45, h - 0.3, list(zeilen), size=size,
         color=RGBColor(0xC7, 0xD6, 0xEE), font=MONO, spacing=1.35)
    return h


def code_farbig(slide, x, y, w, zeilen, size=10, streifen=ACCENT):
    """Wie code(), aber jede Zeile kann ihre eigene Farbe haben.

    zeilen: Liste von (text, farbe). Wird gebraucht, um Verschachtelungsebenen
    farblich mit dem Code zu verknuepfen, der sie durchlaeuft.
    """
    zh = size * 1.62 / 72
    h = 0.3 + len(zeilen) * zh
    panel(slide, x, y, w, h, fill=RGBColor(0x0A, 0x11, 0x21))
    rect(slide, x + 0.001, y + 0.13, 0.045, h - 0.26, fill=streifen)
    for i, (inhalt, farbe) in enumerate(zeilen):
        text(slide, x + 0.26, y + 0.15 + i * zh, w - 0.45, zh + 0.05, inhalt,
             size=size, color=farbe, font=MONO, spacing=1.0)
    return h


def kennzahl(slide, x, y, w, wert, label, farbe=ACCENT):
    panel(slide, x, y, w, 1.15)
    text(slide, x, y + 0.16, w, 0.5, wert, size=26, color=farbe, bold=True,
         align=PP_ALIGN.CENTER)
    text(slide, x, y + 0.72, w, 0.3, label, size=9, color=DIM, align=PP_ALIGN.CENTER,
         caps=True)


# ----------------------------------------------------------------------------
# Bausteine fuer die Frontend-Folien
#
# Die Oberflaeche wird aus Formen nachgebaut statt als Screenshot eingebettet.
# Vorteil: sie bleibt scharf, laesst sich beschriften und passt farblich zum
# Rest der Praesentation.
# ----------------------------------------------------------------------------

def browser(slide, x, y, w, h, url, chrome=0.36):
    """Browserfenster mit Titelleiste und Adresszeile.

    Gibt die Inhaltsflaeche als (x, y, breite, hoehe) zurueck, damit die
    aufrufende Folie darin weiterzeichnen kann.
    """
    rect(slide, x, y, w, h, fill=CHROME, line=RGBColor(0x2E, 0x3E, 0x5E),
         line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.028)

    for i, punkt in enumerate((RGBColor(0xF2, 0x8B, 0x82), RGBColor(0xF5, 0xC2, 0x64),
                               RGBColor(0x86, 0xCF, 0x8B))):
        rect(slide, x + 0.14 + i * 0.155, y + 0.135, 0.095, 0.095, fill=punkt,
             shape=MSO_SHAPE.OVAL)

    pille_x = x + 0.66
    pille_w = max(1.2, min(w - 0.66 - 0.16, 4.4))
    rect(slide, pille_x, y + 0.095, pille_w, 0.185, fill=APP_CARD,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.45)
    text(slide, pille_x + 0.13, y + 0.122, pille_w - 0.22, 0.16, url, size=7,
         color=APP_DIM, font=MONO)

    ix, iy = x + 0.07, y + chrome
    iw, ih = w - 0.14, h - chrome - 0.07
    rect(slide, ix, iy, iw, ih, fill=APP_BG)
    return ix, iy, iw, ih


def app_karte(slide, x, y, w, h, fill=APP_CARD, line=APP_LINE, radius=0.06):
    return rect(slide, x, y, w, h, fill=fill, line=line, line_w=0.6,
                shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=radius)


def app_balken(slide, x, y, w, anteil, hoehe=0.045):
    """Mini-Balken wie in der Tabelle: Schiene plus Fuellung nach Anteil."""
    rect(slide, x, y, w, hoehe, fill=APP_BAR_BG,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    breite = max(0.03, w * min(max(anteil, 0.0), 1.0))
    rect(slide, x, y, breite, hoehe, fill=APP_BAR,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)


def app_kennzahl(slide, x, y, w, label, wert, einheit=None, h=0.62):
    """Kennzahlen-Karte: Label in Kapitaelchen, darunter der Wert."""
    app_karte(slide, x, y, w, h)
    text(slide, x + 0.14, y + 0.11, w - 0.28, 0.16, label, size=6.5, color=APP_MUTE,
         bold=True, caps=True)
    text(slide, x + 0.14, y + 0.28, w - 0.28, 0.28, wert, size=14, color=APP_TXT,
         bold=True)
    if einheit:
        # Die Einheit steht direkt hinter der Zahl, ihre Breite wird aus der
        # Zeichenzahl geschaetzt (Schriftgroesse 14, rund 0,108" je Zeichen).
        versatz = 0.14 + len(wert) * 0.108 + 0.04
        text(slide, x + versatz, y + 0.365, 0.6, 0.16, einheit, size=7, color=APP_MUTE)


def app_suchfeld(slide, x, y, w, inhalt=None, platzhalter="Sensor, Typ oder Zeitstempel suchen…",
                 h=0.34, kbd=True):
    """Suchfeld mit Lupen-Chip links und Tastenkuerzel rechts."""
    app_karte(slide, x, y, w, h, radius=0.14)
    rect(slide, x + 0.08, y + 0.06, h - 0.12, h - 0.12, fill=APP_BLUE_L,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
    lupe = rect(slide, x + 0.145, y + 0.115, 0.105, 0.105, shape=MSO_SHAPE.OVAL)
    lupe.line.color.rgb = APP_BLUE
    lupe.line.width = Pt(0.9)
    rect(slide, x + 0.238, y + 0.208, 0.055, 0.02, fill=APP_BLUE)

    if inhalt:
        text(slide, x + h + 0.04, y + 0.105, w - h - 0.9, 0.2, inhalt, size=9,
             color=APP_TXT, font=MONO)
    else:
        text(slide, x + h + 0.04, y + 0.105, w - h - 0.9, 0.2, platzhalter, size=8.5,
             color=APP_MUTE)
    if kbd:
        rect(slide, x + w - 0.52, y + 0.085, 0.42, h - 0.17, fill=APP_BG, line=APP_LINE,
             line_w=0.6, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
        text(slide, x + w - 0.52, y + 0.125, 0.42, 0.16, "⌘K", size=7, color=APP_DIM,
             align=PP_ALIGN.CENTER)


def app_tabelle(slide, x, y, w, zeilen, zeilen_h=0.34, kopf_h=0.3, size=8,
                treffer=(), sortiert=4, richtung="↓"):
    """Sensortabelle im Look der Webapp.

    zeilen: (name, typ, druck, temperatur, zeit) - Druck und Temperatur als Zahl,
    damit der Mini-Balken die Lage im gueltigen Bereich zeigen kann.
    treffer: Indizes der Zeilen, die als Suchtreffer hervorgehoben werden.
    """
    anteile = (0.24, 0.145, 0.20, 0.205, 0.21)
    breiten = [w * a for a in anteile]
    kopf = ("Name", "Typ", "Druck (bar)", "Temperatur (K)", "Zeitpunkt")
    bereiche = {2: (0.5, 9.0, 2), 3: (200.0, 500.0, 1)}

    app_karte(slide, x, y, w, kopf_h + len(zeilen) * zeilen_h + 0.1)

    xi = x + 0.12
    for j, (spalte, bw) in enumerate(zip(kopf, breiten)):
        pfeil = f"  {richtung}" if j == sortiert else "  ↕"
        rechts = j in bereiche
        text(slide, xi, y + 0.1, bw - 0.1, 0.18, spalte + pfeil, size=6.5,
             color=APP_BLUE if j == sortiert else APP_MUTE, bold=True, caps=True,
             align=PP_ALIGN.RIGHT if rechts else PP_ALIGN.LEFT)
        xi += bw
    rect(slide, x + 0.12, y + kopf_h, w - 0.24, 0.008, fill=APP_LINE)

    yc = y + kopf_h + 0.04
    for i, zeile in enumerate(zeilen):
        if i in treffer:
            rect(slide, x + 0.06, yc, w - 0.12, zeilen_h - 0.02, fill=APP_MARK,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
        xi = x + 0.12
        for j, (zelle, bw) in enumerate(zip(zeile, breiten)):
            if j in bereiche:
                unten, oben, stellen = bereiche[j]
                zahl = f"{zelle:.{stellen}f}".replace(".", ",")
                text(slide, xi, yc + 0.05, bw - 0.14, 0.18, zahl, size=size,
                     color=APP_TXT, font=MONO, align=PP_ALIGN.RIGHT)
                balken_w = min(0.62, bw - 0.2)
                app_balken(slide, xi + bw - 0.14 - balken_w, yc + 0.235, balken_w,
                           (zelle - unten) / (oben - unten))
            elif j == 1:
                rect(slide, xi, yc + 0.06, min(bw - 0.16, 0.05 + len(zelle) * 0.062),
                     0.19, fill=APP_BLUE_L, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                     radius=0.32)
                text(slide, xi + 0.06, yc + 0.085, bw - 0.2, 0.16, zelle, size=6.5,
                     color=APP_BLUE_D)
            else:
                text(slide, xi, yc + 0.09, bw - 0.1, 0.2, zelle, size=size,
                     color=APP_TXT if j == 0 else APP_DIM,
                     font=MONO if j == 0 else FONT)
            xi += bw
        if i < len(zeilen) - 1:
            rect(slide, x + 0.12, yc + zeilen_h - 0.02, w - 0.24, 0.006, fill=APP_LINE)
        yc += zeilen_h
    return yc + 0.06


def app_knopf(slide, x, y, w, h, label, gefuellt=False):
    rect(slide, x, y, w, h, fill=APP_BLUE if gefuellt else APP_CARD,
         line=None if gefuellt else APP_LINE, line_w=0.6,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
    text(slide, x, y + (h - 0.16) / 2, w, 0.18, label, size=7.5,
         color=APP_CARD if gefuellt else APP_DIM, align=PP_ALIGN.CENTER)


def knoten(slide, x, y, w, h, titel, unter=None, farbe=ACCENT, route=None):
    """Kasten im Komponentenbaum: Dateiname, Rolle und optional die Route."""
    panel(slide, x, y, w, h, fill=PANEL2)
    rect(slide, x + 0.001, y + 0.1, 0.045, h - 0.2, fill=farbe)
    text(slide, x + 0.2, y + 0.1, w - 0.3, 0.2, titel, size=9.5, color=TXT, bold=True,
         font=MONO)
    if unter:
        text(slide, x + 0.2, y + 0.33, w - 0.3, 0.36, unter, size=7.5, color=DIM,
             spacing=1.15)
    if route:
        text(slide, x + 0.2, y + h - 0.28, w - 0.3, 0.2, route, size=7, color=farbe,
             font=MONO)
    return x + w


def ast(slide, x, y_von, y_bis, abzweige, stub, farbe=BORDER):
    """Senkrechte Sammellinie mit waagerechten Abzweigen - fuer Baumdiagramme."""
    rect(slide, x, y_von, 0.011, y_bis - y_von, fill=farbe)
    for y in abzweige:
        rect(slide, x, y, stub, 0.011, fill=farbe)


def marker(slide, x, y, nummer, farbe=ACCENT, d=0.28):
    """Nummerierte Sprechblase, um ein Mockup zu beschriften."""
    rect(slide, x, y, d, d, fill=farbe, shape=MSO_SHAPE.OVAL)
    text(slide, x, y + 0.055, d, 0.18, str(nummer), size=9, color=INK, bold=True,
         align=PP_ALIGN.CENTER)


def legende(slide, x, y, w, nummer, titel, zeilen=None, farbe=ACCENT, abstand=0.0):
    """Zeile zur Legende eines Mockups: Nummer, Stichwort, kurze Erklaerung."""
    marker(slide, x, y + abstand, nummer, farbe, d=0.24)
    text(slide, x + 0.36, y + abstand + 0.015, w - 0.36, 0.2, titel, size=9.5,
         color=TXT, bold=True)
    if zeilen:
        text(slide, x + 0.36, y + abstand + 0.24, w - 0.36, 0.4, list(zeilen), size=8.5,
             color=DIM, spacing=1.2)


def chip(slide, x, y, label, farbe=ACCENT, size=8, breite=None, h=0.24, mono=True):
    """Kleines Etikett mit Rahmen in der Akzentfarbe."""
    w = breite if breite else 0.22 + len(label) * (0.062 if mono else 0.056)
    rect(slide, x, y, w, h, fill=PANEL2, line=farbe, line_w=0.75,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
    text(slide, x, y + (h - 0.15) / 2, w, 0.18, label, size=size, color=farbe,
         align=PP_ALIGN.CENTER, font=MONO if mono else FONT)
    return x + w


def zustand(slide, x, y, w, h, label, unter=None, farbe=ACCENT):
    """Zustandskasten fuer den Zustandsautomaten."""
    rect(slide, x, y, w, h, fill=PANEL2, line=farbe, line_w=1.25,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.22)
    text(slide, x, y + (0.16 if unter else (h - 0.2) / 2), w, 0.22, label, size=11,
         color=farbe, bold=True, align=PP_ALIGN.CENTER, font=MONO)
    if unter:
        text(slide, x, y + 0.44, w, 0.2, unter, size=7.5, color=DIM,
             align=PP_ALIGN.CENTER)


# ----------------------------------------------------------------------------
# Folien
# ----------------------------------------------------------------------------

def folie_titel(prs):
    slide = neue_folie(prs, notiz=(
        "Begruessung. Ziel der Praesentation: den Weg eines einzelnen Messwerts "
        "vom Satelliten bis auf den Bildschirm zeigen und erklaeren, welche "
        "Komponenten dabei voneinander abhaengen."))
    stars(slide, anzahl=90, seed=11)

    # Planet unten rechts, teilweise ausserhalb der Folie
    planet = rect(slide, 8.7, 3.9, 5.6, 5.6, fill=RGBColor(0x11, 0x1B, 0x30),
                  shape=MSO_SHAPE.OVAL)
    planet.line.color.rgb = RGBColor(0x27, 0x3A, 0x5E)
    planet.line.width = Pt(1.0)
    rect(slide, 9.35, 4.55, 4.3, 4.3, fill=RGBColor(0x14, 0x21, 0x3A), shape=MSO_SHAPE.OVAL)

    # Umlaufbahn
    bahn = rect(slide, 6.9, 4.55, 8.6, 3.3, shape=MSO_SHAPE.OVAL)
    bahn.line.color.rgb = RGBColor(0x24, 0x3B, 0x63)
    bahn.line.width = Pt(1.0)
    bahn.rotation = 338

    # Satellit auf der Bahn
    sx, sy = 8.05, 3.42
    rect(slide, sx + 0.52, sy + 0.14, 0.46, 0.42, fill=PANEL2, line=ACCENT, line_w=1.0,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15)
    rect(slide, sx - 0.22, sy + 0.22, 0.66, 0.26, fill=RGBColor(0x1B, 0x2C, 0x4A),
         line=RGBColor(0x33, 0x4C, 0x7A), line_w=0.75)
    rect(slide, sx + 1.06, sy + 0.22, 0.66, 0.26, fill=RGBColor(0x1B, 0x2C, 0x4A),
         line=RGBColor(0x33, 0x4C, 0x7A), line_w=0.75)
    rect(slide, sx + 0.73, sy + 0.58, 0.04, 0.2, fill=ACCENT)
    rect(slide, sx + 0.63, sy + 0.76, 0.24, 0.09, fill=ACCENT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
    # Funkstrecke zum Boden
    arrow(slide, sx + 0.75, sy + 0.92, 2.35, 5.55, color=RGBColor(0x2C, 0x6E, 0x92),
          width=1.0, dashed=True)

    # Bodenstation als Schuessel
    rect(slide, 1.95, 5.62, 0.78, 0.2, fill=RGBColor(0x1B, 0x2C, 0x4A),
         line=RGBColor(0x33, 0x4C, 0x7A), line_w=0.75,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    rect(slide, 2.3, 5.8, 0.05, 0.28, fill=RGBColor(0x33, 0x4C, 0x7A))
    text(slide, 1.5, 6.13, 1.7, 0.24, "Bodenstation", size=8.5, color=FAINT,
         align=PP_ALIGN.CENTER, caps=True)

    # Titelblock
    text(slide, LM, 1.62, 7.6, 0.26, "QAware Schülerpraktikum 2026", size=10,
         color=ACCENT, bold=True, caps=True)
    text(slide, LM, 2.08, 7.8, 1.5, TITEL, size=50, color=TXT, bold=True, spacing=0.95)
    line_h(slide, LM, 3.42, 1.7, color=ACCENT, thickness=0.032)
    text(slide, LM, 3.72, 6.9, 1.1,
         "Vom Messwert im Orbit bis zur Anzeige am Boden – "
         "vier Komponenten, eine Datenstrecke.",
         size=17, color=DIM, spacing=1.35)
    text(slide, LM, 6.36, 6.0, 0.5, [TEAM, "August 2026"], size=10.5, color=FAINT,
         spacing=1.35)
    return slide


def folie_aufgabe(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Die Aufgabenstellung in eigenen Worten. Wichtig: die Daten kommen als "
        "Dateien an, nicht als Live-Stream. Genau daraus ergibt sich der Aufbau "
        "des Systems. Der Space Controller ist unser Nutzer."))
    kopf(slide, "Ausgangspunkt", "Die Aufgabe")
    stars(slide, anzahl=22, seed=3, ymax=1.2)

    karte(slide, LM, 1.85, 3.7, 2.15, "Was ankommt", [
        "Ein Satellit sendet Telemetrie in", "unregelmäßigen Abständen.",
        "Jede Datei enthält die Messung", "genau eines Sensors."], farbe=ACCENT)
    karte(slide, LM + 3.95, 1.85, 3.7, 2.15, "Was zu tun ist", [
        "Daten entgegennehmen, prüfen", "und den Zustand jedes Sensors",
        "aktuell halten – ohne eine", "Datei doppelt zu verarbeiten."], farbe=VIOLET)
    karte(slide, LM + 7.9, 1.85, 3.73, 2.15, "Wer es braucht", [
        "Ein Space Controller sieht alle", "aktuellen Werte auf einen Blick",
        "und kann den Verlauf eines", "Sensors nachvollziehen."], farbe=AMBER)

    text(slide, LM, 4.35, CW, 0.3, "Ein Datensatz, wie er im Aufgabenblatt steht",
         size=9.5, color=FAINT, bold=True, caps=True)
    code(slide, LM, 4.68, 6.3, [
        "{",
        '    "type":        "thruster",',
        '    "name":        "thruster_1.a",',
        '    "pressure":    3.345,',
        '    "temperature": 432.21',
        "}"], size=10.5)

    karte(slide, LM + 6.6, 4.68, 5.03, 1.72, "Der Zeitstempel fehlt im Inhalt", [
        "Er steckt im Dateinamen:",
        "TM_20260806_143000.json",
        "Die Bodenstation muss ihn herauslösen."], farbe=AMBER, titel_size=11.5)
    fuss(slide, nr)
    return slide


def folie_architektur(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Der Ueberblick. Fuenf Stationen in einer Reihe, das Frontend haengt "
        "unten am Backend. Wichtig zu betonen: die Uebergabe zwischen Satellit "
        "und Bodenstation laeuft ueber das Dateisystem, alles danach ueber HTTP. "
        "Jede Komponente laeuft als eigener Prozess und kann einzeln neu "
        "gestartet werden."))
    kopf(slide, "Überblick", "Architektur")

    stationen = [
        ("Satellit", "tools/DataGenerator.py", "erzeugt Messwerte", ACCENT),
        ("Dateiablage", "data/*.json", "Übergabepunkt", ACCENT),
        ("Bodenstation", "Bodenstation.py", "liest, prüft, sendet", ACCENT),
        ("Verwaltung", "BeispielVerwaltung.py", "FastAPI, 4 Routen", VIOLET),
        ("Datenbank", "MongoDB", "2 Collections", GREEN),
    ]
    bw, gap = 2.04, 0.34
    bx, by, bh = LM + 0.05, 2.1, 1.55

    for i, (name, datei, rolle, farbe) in enumerate(stationen):
        x = bx + i * (bw + gap)
        panel(slide, x, by, bw, bh)
        rect(slide, x + 0.001, by + 0.16, 0.05, bh - 0.32, fill=farbe)
        text(slide, x + 0.24, by + 0.22, bw - 0.4, 0.28, name, size=12.5, color=TXT, bold=True)
        text(slide, x + 0.24, by + 0.58, bw - 0.4, 0.24, datei, size=8.5, color=farbe, font=MONO)
        text(slide, x + 0.24, by + 0.92, bw - 0.4, 0.45, rolle, size=9.5, color=DIM,
             spacing=1.15)
        if i < len(stationen) - 1:
            x2 = x + bw
            arrow(slide, x2 + 0.06, by + bh / 2, x2 + gap - 0.06, by + bh / 2,
                  color=RGBColor(0x3A, 0x5A, 0x86), width=1.5)

    for i, label in enumerate(["JSON-Datei", "liest & prüft", "HTTP POST", "Motor, async"]):
        x = bx + i * (bw + gap) + bw - 0.28
        text(slide, x, by + bh / 2 - 0.42, gap + 0.56, 0.24, label, size=8,
             color=FAINT, align=PP_ALIGN.CENTER)

    # Frontend unterhalb der Verwaltung, mit seinen beiden Ansichten
    fx = bx + 3 * (bw + gap)
    fw = bw + gap + bw
    panel(slide, fx, 4.72, fw, 1.3)
    rect(slide, fx + 0.001, 4.88, 0.05, 0.98, fill=AMBER)
    text(slide, fx + 0.24, 4.86, 2.4, 0.28, "Frontend", size=12.5, color=TXT, bold=True)
    text(slide, fx + 0.24, 5.2, 2.8, 0.24, "webapp/  ·  React 19 + Vite", size=8.5,
         color=AMBER, font=MONO)
    xc = chip(slide, fx + 0.24, 5.52, "Startseite", farbe=AMBER, size=8, mono=False)
    chip(slide, xc + 0.14, 5.52, "Datenansicht", farbe=AMBER, size=8, mono=False)
    text(slide, fx + 2.94, 4.94, 1.4, 0.6, ["fragt alle", "10 Sekunden ab"], size=9.5,
         color=DIM, spacing=1.2)

    arrow(slide, fx + 1.0, 4.68, fx + 1.0, by + bh + 0.06,
          color=RGBColor(0x3A, 0x5A, 0x86), width=1.5)
    text(slide, fx + 1.14, 3.78, 2.2, 0.7, [
        "GET /data/current", "GET /data_wsi/", "GET /data_wsi/{name}"],
        size=7.5, color=FAINT, font=MONO, spacing=1.3)

    line_h(slide, LM, 6.22, CW)
    text(slide, LM, 6.38, CW, 0.5,
         "Vier eigenständige Prozesse plus die Datenbank im Container. Der Übergang "
         "Satellit → Bodenstation läuft über das Dateisystem, alles danach über HTTP – "
         "das Frontend nutzt drei der vier Routen.",
         size=10.5, color=DIM, spacing=1.3)
    fuss(slide, nr)
    return slide


def folie_satellit(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Der Datengenerator steht fuer den Satelliten. Entscheidend fuer das "
        "Verstaendnis des ganzen Projekts: er liefert absichtlich zu 40 Prozent "
        "unbrauchbare Daten. Genau dafuer braucht die Bodenstation ihre "
        "Pruefungen. Die Wahrscheinlichkeiten stehen direkt im Code."))
    kopf(slide, "Backend · Komponente 1", "Der Satellit")
    stars(slide, anzahl=18, seed=5, ymax=1.2)

    karte(slide, LM, 1.85, 5.6, 1.62, "13 Sensoren, zwei Typen", [
        "9 × thruster  –  thruster_1.a bis thruster_3.c",
        "4 × gas_valve  –  Sauerstoff- und Wasserstofftanks",
        "Alle 3 Sekunden meldet sich ein zufällig gewählter Sensor."],
        farbe=ACCENT, zeile_size=10.5)

    karte(slide, LM + 5.9, 1.85, 5.73, 1.62, "Jede Meldung wird eine Datei", [
        "Dateiname:  TM_JJJJMMTT_HHMMSS.json",
        "Inhalt:  name, type, pressure, temperature",
        "Abgelegt im Ordner data/ – dem Übergabepunkt."],
        farbe=ACCENT, zeile_size=10.5)

    text(slide, LM, 3.72, CW, 0.3,
         "Der Generator stört absichtlich – die Wahrscheinlichkeiten stehen im Code",
         size=9.5, color=AMBER, bold=True, caps=True)
    gitter(slide, LM, 4.06, [1.3, 3.5, 4.4, 2.43],
           ["Anteil", "Was passiert", "Wie es aussieht", "Wer es abfängt"],
           [["60 %", "Alles in Ordnung", "Werte in gültigen Bereichen", "–"],
            ["10 %", "Unplausible Messwerte", "Druck ~2 Mio. bar, −220 K", "Wertebereichsprüfung"],
            ["10 %", "Kaputter Datensatz", 'überall der Text "Error"', "Datentypprüfung"],
            ["20 %", "Falsches Dateiformat", ".txt, .yaml, .xml, .pdf", "Endungsprüfung"]],
           zeilen_h=0.44, farben=[GREEN, AMBER, AMBER, AMBER], kopf_farbe=AMBER)

    text(slide, LM, 6.4, CW, 0.4,
         "Rund 40 Prozent der eintreffenden Dateien sind unbrauchbar. "
         "Die Bodenstation muss sie erkennen, ohne dabei stehenzubleiben.",
         size=10.5, color=DIM, spacing=1.3)
    fuss(slide, nr)
    return slide


def folie_bodenstation(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Die Bodenstation ist die erste Komponente aus dem Aufgabenblatt. "
        "Fuenf Pruefstufen hintereinander, jede kann die Datei aussortieren. "
        "Erst danach geht der Datensatz per HTTP raus. Die Datei wird in jedem "
        "Fall geloescht, damit sie nicht doppelt verarbeitet wird."))
    kopf(slide, "Backend · Komponente 2", "Die Bodenstation")

    stufen = [
        ("1", "Endung", ".json?"),
        ("2", "Inhalt", "beginnt mit {?"),
        ("3", "Zeitstempel", "Dateiname lesbar?"),
        ("4", "Datentypen", "3 Texte, 2 Zahlen?"),
        ("5", "Wertebereich", "0,5–9 bar / 200–500 K?"),
    ]
    bw, gap = 2.04, 0.34
    bx, by, bh = LM + 0.05, 1.85, 1.28
    for i, (num, name, frage) in enumerate(stufen):
        x = bx + i * (bw + gap)
        panel(slide, x, by, bw, bh)
        text(slide, x + 0.24, by + 0.18, 0.5, 0.3, num, size=13, color=ACCENT, bold=True)
        text(slide, x + 0.24, by + 0.52, bw - 0.4, 0.26, name, size=11.5, color=TXT, bold=True)
        text(slide, x + 0.24, by + 0.83, bw - 0.4, 0.3, frage, size=8.5, color=DIM)
        if i < len(stufen) - 1:
            arrow(slide, x + bw + 0.06, by + bh / 2, x + bw + gap - 0.06, by + bh / 2,
                  color=RGBColor(0x3A, 0x5A, 0x86), width=1.4)
        arrow(slide, x + bw / 2, by + bh + 0.05, x + bw / 2, by + bh + 0.42,
              color=RGBColor(0x6B, 0x3C, 0x4A), width=1.1, dashed=True)

    text(slide, LM, 3.62, CW, 0.26, "Aussortiert – mit Begründung im Protokoll",
         size=9, color=RED, bold=True, caps=True)

    karte(slide, LM, 4.0, 5.6, 1.62, "Was danach passiert", [
        "Zeitstempel aus dem Dateinamen wird ISO-Format",
        "Datensatz geht per HTTP POST an die Verwaltung",
        "Die Datei wird gelöscht – keine Doppelverarbeitung"],
        farbe=GREEN, zeile_size=10.5)

    karte(slide, LM + 5.9, 4.0, 5.73, 1.62, "Dauerbetrieb", [
        "Endlosschleife, halbe Sekunde Pause wenn nichts da ist",
        "Dateien werden sortiert gelesen – Namen sind chronologisch",
        "Eine kaputte Datei stoppt die Schleife nicht"],
        farbe=ACCENT, zeile_size=10.5)

    text(slide, LM, 5.88, CW, 0.5,
         "Ergebnis: von 100 eintreffenden Dateien erreichen etwa 60 die Verwaltung – "
         "und zwar nur solche, die vollständig und plausibel sind.",
         size=10.5, color=DIM, spacing=1.3)
    fuss(slide, nr)
    return slide


def folie_verwaltung(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Die Verwaltung ist das Herz des Systems: FastAPI mit vier Routen. "
        "Sie ist die einzige Komponente, die mit der Datenbank spricht. "
        "Wer Daten will, fragt hier und nicht direkt in MongoDB."))
    kopf(slide, "Backend · Komponente 3", "Die Verwaltung", farbe=VIOLET)

    gitter(slide, LM, 1.82, [1.15, 3.3, 4.4, 2.78],
           ["Methode", "Route", "Zweck", "Aufrufer"],
           [["POST", "/data/", "Messwert aufnehmen und speichern", "Bodenstation.py"],
            ["GET", "/data/current", "Aktueller Wert aller Sensoren", "Startseite"],
            ["GET", "/data_wsi/{name}", "Komplette Historie eines Sensors", "Detailfenster"],
            ["GET", "/data_wsi/", "Historie aller Sensoren", "Datenansicht"]],
           zeilen_h=0.46, kopf_farbe=VIOLET,
           farben=[GREEN, VIOLET, VIOLET, VIOLET])

    karte(slide, LM, 4.02, 5.6, 1.75, "Ein POST, zwei Schreibvorgänge", [
        "Historie: der Messwert wird hinten angehängt",
        "Aktueller Wert: wird nur ersetzt, wenn der neue",
        "Messwert tatsächlich jünger ist",
        "Beides in einem try/except – Fehler wird 503, nicht 500"],
        farbe=VIOLET, zeile_size=10)

    karte(slide, LM + 5.9, 4.02, 5.73, 1.75, "Warum FastAPI hier passt", [
        "Pydantic prüft jeden eingehenden Datensatz automatisch",
        "Unpassende Daten werden mit 422 abgelehnt, bevor",
        "sie die Datenbank erreichen",
        "Swagger-Doku unter /docs entsteht ohne Zusatzarbeit"],
        farbe=ACCENT, zeile_size=10)
    fuss(slide, nr)
    return slide


def folie_datenmodell(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Das Datenmodell ist der Tuersteher. Drei Klassen: DataModel fuer das, "
        "was hereinkommt, MeasurementModel fuer einen einzelnen Messwert, "
        "GroupedDataModel fuer einen Sensor mit seiner Historie. Der Typ von "
        "time ist entscheidend - dazu kommt spaeter noch eine eigene Folie."))
    kopf(slide, "Backend · Schnittstelle", "Das Datenmodell")

    code(slide, LM, 1.82, 6.3, [
        "class DataModel(BaseModel):",
        "    time:        datetime",
        "    name:        str = Field(min_length=1)",
        "    type:        str = Field(min_length=1)",
        "    pressure:    float",
        "    temperature: float",
        "",
        "class MeasurementModel(BaseModel):",
        "    time:        datetime",
        "    pressure:    float",
        "    temperature: float",
    ], size=10.5, farbe=ACCENT)

    karte(slide, LM + 6.6, 1.82, 5.03, 1.35, "DataModel", [
        "Was die Bodenstation sendet. Fehlt ein Feld",
        "oder passt ein Typ nicht, antwortet die",
        "Verwaltung mit 422."], farbe=ACCENT, titel_size=12, zeile_size=10)

    karte(slide, LM + 6.6, 3.32, 5.03, 1.35, "MeasurementModel", [
        "Ein Messwert ohne Sensornamen – der steht",
        "schon im Dokument darüber. Wird für current",
        "und für die Historie verwendet."], farbe=VIOLET, titel_size=12, zeile_size=10)

    karte(slide, LM + 6.6, 4.82, 5.03, 1.35, "GroupedDataModel", [
        "Ein Sensor mit seiner Liste von Messwerten –",
        "genau die Form, in der ein Dokument in der",
        "Datenbank liegt."], farbe=GREEN, titel_size=12, zeile_size=10)

    karte(slide, LM, 4.82, 6.3, 1.35, "time ist ein datetime, kein Text", [
        "Damit kann die Datenbank Zeitpunkte vergleichen",
        "und sortieren. Ein Text ließe sich nur alphabetisch",
        "sortieren."],
        farbe=AMBER, titel_size=12, zeile_size=10)
    fuss(slide, nr)
    return slide


def folie_mongodb(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Zwei Collections mit unterschiedlicher Aufgabe. Die eine bleibt klein "
        "und schnell, weil sie pro Sensor nur ein Dokument hat und der alte Wert "
        "ersetzt wird. Die andere waechst, weil sie nichts wegwirft. Der "
        "Vergleich der Dokumentzahlen macht das deutlich: 13 gegen 13 Dokumente, "
        "aber 0 gegen 2315 Messwerte."))
    kopf(slide, "Backend · Komponente 4", "Die Datenbank", farbe=GREEN)

    panel(slide, LM, 1.8, 5.6, 2.5)
    rect(slide, LM + 0.001, 1.98, 0.055, 2.14, fill=ACCENT)
    text(slide, LM + 0.28, 1.96, 3.0, 0.3, "Collection  data", size=13.5, color=TXT,
         bold=True, font=MONO)
    text(slide, LM + 0.28, 2.32, 5.0, 0.26, "Nur der aktuelle Zustand", size=10,
         color=ACCENT, caps=True, bold=True)
    text(slide, LM + 0.28, 2.68, 5.05, 1.5, [
        "Ein Dokument pro Sensor – immer 13.",
        "Ein neuer Messwert ersetzt den alten vollständig.",
        "Bleibt dadurch dauerhaft klein und schnell.",
        "Beantwortet: „Wie geht es dem Satelliten gerade?“"],
        size=10, color=DIM, spacing=1.35, space_after=4)

    panel(slide, LM + 5.9, 1.8, 5.73, 2.5)
    rect(slide, LM + 5.901, 1.98, 0.055, 2.14, fill=GREEN)
    text(slide, LM + 6.18, 1.96, 3.2, 0.3, "Collection  data_wsi", size=13.5, color=TXT,
         bold=True, font=MONO)
    text(slide, LM + 6.18, 2.32, 5.0, 0.26, "Die komplette Historie", size=10,
         color=GREEN, caps=True, bold=True)
    text(slide, LM + 6.18, 2.68, 5.2, 1.5, [
        "Ein Dokument pro Sensor mit einer Liste aller Messwerte.",
        "Es wird nur angehängt, nie gelöscht.",
        "Wächst mit jeder Meldung weiter.",
        "Beantwortet: „Wie war der Verlauf?“"],
        size=10, color=DIM, spacing=1.35, space_after=4)

    text(slide, LM, 4.52, CW, 0.28, "Stand der Datenbank", size=9.5, color=FAINT,
         bold=True, caps=True)
    kennzahl(slide, LM, 4.86, 2.72, "13", "Sensoren in data", farbe=ACCENT)
    kennzahl(slide, LM + 2.97, 4.86, 2.72, "13", "Dokumente in data_wsi", farbe=GREEN)
    kennzahl(slide, LM + 5.94, 4.86, 2.72, "2 315", "Messwerte in der Historie", farbe=GREEN)
    kennzahl(slide, LM + 8.91, 4.86, 2.72, "1", "Datenbank in Docker", farbe=VIOLET)

    text(slide, LM, 6.22, CW, 0.4,
         "Dieselbe Frage zweimal getrennt beantwortet: die eine Collection bleibt "
         "klein, weil sie vergisst – die andere wächst, weil sie sich alles merkt.",
         size=10.5, color=DIM, spacing=1.3)
    fuss(slide, nr)
    return slide


def folie_fe_ueberblick(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Einstieg in den Frontend-Teil. Die Anwendung hat genau zwei Ansichten, "
        "umgeschaltet wird ueber eine einzige Zustandsvariable in App.jsx. Links "
        "die Startseite mit den aktuellen Werten, nach Tanktyp gruppiert. Rechts "
        "die Datenansicht mit Suche, Sortierung und Kennzahlen. Wichtig: die "
        "beiden Ansichten holen ihre Daten von unterschiedlichen Routen - die "
        "Startseite die aktuellen Werte, die Datenansicht die Historie."))
    kopf(slide, "Frontend · Überblick", "Zwei Ansichten, ein Umschalter", farbe=AMBER)

    bw, bh, by = 5.15, 4.4, 1.8
    bx2 = SW - LM - bw

    # ---- links: Startseite ------------------------------------------------
    ix, _, iw, _ = browser(slide, LM, by, bw, bh, "localhost:5173")
    mitte = ix + iw / 2

    bahn = rect(slide, mitte - 0.34, 2.28, 0.68, 0.3, shape=MSO_SHAPE.OVAL)
    bahn.line.color.rgb = APP_BLUE
    bahn.line.width = Pt(1.0)
    bahn.rotation = 330
    rect(slide, mitte - 0.085, 2.355, 0.17, 0.17, fill=APP_BLUE, shape=MSO_SHAPE.OVAL)

    text(slide, ix, 2.6, iw, 0.28, "Satelliten Daten", size=13, color=APP_TXT,
         bold=True, align=PP_ALIGN.CENTER)
    text(slide, ix, 2.85, iw, 0.2, "Live-Telemetrie der Bodenstation", size=7.5,
         color=APP_DIM, align=PP_ALIGN.CENTER)
    app_suchfeld(slide, ix + 0.72, 3.06, iw - 1.44, platzhalter="Daten durchsuchen…",
                 h=0.32)

    gruppen = [
        ("Thruster", [("thruster_1.a", "3,35", "432,2"), ("thruster_1.b", "3,12", "441,7")]),
        ("Oxygen Tanks", [("o_tank_1", "5,10", "268,4"), ("o_tank_2", "4,87", "271,9")]),
        ("Hydrogen Tanks", [("h_tank_1", "6,02", "240,1"), ("h_tank_2", "5,74", "236,8")]),
    ]
    gy = 3.56
    for titel, reihen in gruppen:
        text(slide, ix + 0.16, gy, 2.4, 0.18, titel, size=8, color=APP_TXT, bold=True)
        app_karte(slide, ix + 0.16, gy + 0.21, iw - 0.32, 0.54)
        for k, (name, druck, temp) in enumerate(reihen):
            zy = gy + 0.27 + k * 0.23
            text(slide, ix + 0.28, zy, 2.0, 0.17, name, size=7, color=APP_BLUE,
                 font=MONO)
            text(slide, ix + 2.3, zy, 1.0, 0.17, druck, size=7, color=APP_TXT,
                 font=MONO, align=PP_ALIGN.RIGHT)
            text(slide, ix + 3.45, zy, 1.2, 0.17, temp, size=7, color=APP_TXT,
                 font=MONO, align=PP_ALIGN.RIGHT)
        gy += 0.86

    text(slide, LM, by + bh + 0.12, bw, 0.2, "Startseite", size=11, color=TXT,
         bold=True, align=PP_ALIGN.CENTER)
    text(slide, LM, by + bh + 0.36, bw, 0.2,
         "Startseite.jsx + AlteTable.jsx  ·  GET /data/current", size=8,
         color=FAINT, align=PP_ALIGN.CENTER, font=MONO)

    # ---- rechts: Datenansicht --------------------------------------------
    jx, jy, jw, _ = browser(slide, bx2, by, bw, bh, "localhost:5173")

    rect(slide, jx, jy, jw, 0.4, fill=APP_CARD)
    rect(slide, jx, jy + 0.4, jw, 0.008, fill=APP_LINE)
    marke = rect(slide, jx + 0.14, jy + 0.13, 0.22, 0.14, shape=MSO_SHAPE.OVAL)
    marke.line.color.rgb = APP_BLUE
    marke.line.width = Pt(0.75)
    marke.rotation = 330
    text(slide, jx + 0.42, jy + 0.12, 1.6, 0.18, "Satelliten Daten", size=8,
         color=APP_TXT, bold=True)
    text(slide, jx + jw - 2.05, jy + 0.135, 0.85, 0.16, "vor 3 s", size=7,
         color=APP_MUTE, align=PP_ALIGN.RIGHT)
    app_knopf(slide, jx + jw - 1.12, jy + 0.09, 1.0, 0.23, "Aktualisieren")

    text(slide, jx + 0.16, 2.66, 2.4, 0.26, "Telemetrie", size=12.5, color=APP_TXT,
         bold=True)
    text(slide, jx + 0.16, 2.92, 4.4, 0.18, "Live-Sensordaten der Bodenstation",
         size=7.5, color=APP_DIM)

    kw = (jw - 0.32 - 2 * 0.15) / 3
    for i, (label, wert, einheit) in enumerate(
            [("Datensätze", "2 315", None), ("Ø Druck", "4,71", "bar"),
             ("Ø Temperatur", "351,8", "K")]):
        app_kennzahl(slide, jx + 0.16 + i * (kw + 0.15), 3.14, kw, label, wert, einheit)

    app_suchfeld(slide, jx + 0.16, 3.92, jw - 0.32, h=0.32)
    app_tabelle(slide, jx + 0.16, 4.4, jw - 0.32, [
        ("thruster_1.a", "thruster", 3.35, 432.2, "14:31:02"),
        ("o_tank_1", "gas_valve", 5.10, 268.4, "14:30:53"),
        ("thruster_2.b", "thruster", 2.88, 455.7, "14:30:44"),
        ("h_tank_2", "gas_valve", 6.02, 240.1, "14:30:35"),
    ], zeilen_h=0.32, size=7.5)

    text(slide, bx2, by + bh + 0.12, bw, 0.2, "Datenansicht", size=11, color=TXT,
         bold=True, align=PP_ALIGN.CENTER)
    text(slide, bx2, by + bh + 0.36, bw, 0.2,
         "App.jsx + Suche.jsx + SensorTable.jsx  ·  GET /data_wsi/", size=8,
         color=FAINT, align=PP_ALIGN.CENTER, font=MONO)

    # ---- Umschalter dazwischen ------------------------------------------
    gx = LM + bw
    gm = gx + (bx2 - gx) / 2
    arrow(slide, gx + 0.12, 3.62, bx2 - 0.12, 3.62, color=ACCENT, width=1.75)
    text(slide, gx, 3.24, bx2 - gx, 0.2, "⌘K · Enter · Klick", size=8, color=ACCENT,
         align=PP_ALIGN.CENTER)
    arrow(slide, bx2 - 0.12, 4.34, gx + 0.12, 4.34, color=DIM, width=1.25, dashed=True)
    text(slide, gx, 4.44, bx2 - gx, 0.2, "Escape", size=8, color=DIM,
         align=PP_ALIGN.CENTER)
    chip(slide, gm - 0.62, 2.62, 'ansicht', farbe=AMBER, size=8)

    fuss(slide, nr)
    return slide


def folie_fe_komponenten(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Der Aufbau in React. Von links nach rechts wird es spezieller. App.jsx "
        "ist die einzige Stelle, die entscheidet, welche Ansicht laeuft. Der Hook "
        "useSensorData steht bewusst daneben und nicht in einer Komponente: er "
        "kapselt das Laden, und App.jsx bekommt nur noch das Ergebnis. "
        "Bemerkenswert: das Frontend benutzt drei der vier Backend-Routen - jede "
        "an genau einer Stelle."))
    kopf(slide, "Frontend · Aufbau", "Der Komponentenbaum", farbe=AMBER)

    l0, l1, l2, l3, l4 = LM, 2.62, 4.90, 7.38, 9.86
    w0, w1, w2, w3, w4 = 1.55, 2.05, 2.25, 2.25, 2.62

    knoten(slide, l0, 2.85, w0, 0.6, "main.jsx", "StrictMode", farbe=FAINT)
    knoten(slide, l1, 2.72, w1, 0.85, "App.jsx",
           "hält ansicht und suchbegriff", farbe=AMBER)
    arrow(slide, l0 + w0 + 0.04, 3.15, l1 - 0.04, 3.15, color=BORDER, width=1.25)

    # Verzweigung: senkrechte Sammellinie mit Abzweigen zu beiden Ansichten
    bus = l1 + w1 + 0.08
    zweige = [2.29, 3.52, 4.16, 4.80]
    ast(slide, bus, 2.29, 4.80, zweige, l2 - bus)
    rect(slide, l1 + w1, 3.14, bus - (l1 + w1), 0.011, fill=BORDER)

    chip(slide, l2, 1.52, 'ansicht = "start"', farbe=ACCENT, size=7.5)
    knoten(slide, l2, 1.80, w2, 0.98, "Startseite.jsx",
           "Umlaufbahn-Animation,\n⌘K und Enter führen weiter", farbe=ACCENT)
    knoten(slide, l3, 1.80, w3, 0.98, "AlteTable.jsx",
           "3 Gruppen: Thruster,\nOxygen, Hydrogen", farbe=ACCENT,
           route="GET /data/current")
    knoten(slide, l4, 1.80, w4, 0.98, "SensorDetail",
           "Modal mit der Historie\neines Sensors", farbe=VIOLET,
           route="GET /data_wsi/{name}")
    arrow(slide, l2 + w2 + 0.04, 2.29, l3 - 0.04, 2.29, color=BORDER, width=1.25)
    arrow(slide, l3 + w3 + 0.04, 2.29, l4 - 0.04, 2.29, color=BORDER, width=1.25)
    text(slide, l3 + w3 + 0.02, 2.03, 0.55, 0.2, "Klick", size=7, color=FAINT,
         align=PP_ALIGN.CENTER)

    chip(slide, l2, 2.98, 'ansicht = "daten"', farbe=AMBER, size=7.5)
    knoten(slide, l2, 3.26, w2, 0.52, "Kennzahl ×3", farbe=AMBER)
    knoten(slide, l2, 3.90, w2, 0.52, "Suche.jsx", farbe=AMBER)
    knoten(slide, l2, 4.54, w2, 0.52, "SensorTable.jsx", farbe=AMBER)
    text(slide, l3, 3.30, 3.4, 0.2, "Ø Druck, Ø Temperatur, Anzahl", size=8,
         color=DIM)
    text(slide, l3, 3.94, 3.4, 0.2, "⌘K, Escape, Trefferzähler", size=8, color=DIM)
    text(slide, l3, 4.58, 4.2, 0.2, "5 Spalten sortierbar, Mini-Balken, Treffer",
         size=8, color=DIM)

    # Der Hook haengt an App.jsx und liefert die Daten nach oben zurueck
    knoten(slide, l1, 5.35, w1, 0.92, "useSensorData",
           "eigener Hook,\nkapselt das Laden", farbe=GREEN,
           route="GET /data_wsi/")
    arrow(slide, l1 + 1.0, 5.31, l1 + 1.0, 3.61, color=GREEN, width=1.25, dashed=True)
    text(slide, l1 - 1.55, 4.32, 1.5, 0.4, ["{ daten, status,", "fehler, … }"],
         size=7.5, color=GREEN, font=MONO, align=PP_ALIGN.RIGHT, spacing=1.2)

    line_h(slide, LM, 6.5, CW)
    text(slide, LM, 6.62, CW, 0.24,
         "Drei der vier Backend-Routen sind angebunden – jede an genau einer Stelle.",
         size=10, color=DIM)
    fuss(slide, nr)
    return slide


def folie_fe_ansicht(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Die Datenansicht im Detail. Diese Folie am besten durchklicken und die "
        "sechs Punkte einzeln erklaeren. Punkt 1 ist ein schoenes Beispiel fuer "
        "eine bewusste Entscheidung: die Sekundenanzeige steckt in einer eigenen "
        "kleinen Komponente, damit der Sekundentakt nicht jede Sekunde die ganze "
        "Tabelle neu rendert. Punkt 2 ist ehrlich zu benennen - die Mittelwerte "
        "laufen ueber die gefilterte Liste, und die enthaelt derzeit die ganze "
        "Historie."))
    kopf(slide, "Frontend · Oberfläche", "Die Datenansicht", farbe=AMBER)

    bx, by, bw, bh = LM, 1.75, 7.75, 4.75
    ix, iy, iw, _ = browser(slide, bx, by, bw, bh, "localhost:5173")

    rect(slide, ix, iy, iw, 0.46, fill=APP_CARD)
    rect(slide, ix, iy + 0.46, iw, 0.008, fill=APP_LINE)
    marke = rect(slide, ix + 0.16, iy + 0.15, 0.26, 0.17, shape=MSO_SHAPE.OVAL)
    marke.line.color.rgb = APP_BLUE
    marke.line.width = Pt(0.9)
    marke.rotation = 330
    text(slide, ix + 0.5, iy + 0.13, 2.0, 0.2, "Satelliten Daten", size=9.5,
         color=APP_TXT, bold=True)
    text(slide, ix + iw - 2.4, iy + 0.155, 1.0, 0.18, "vor 3 s", size=8,
         color=APP_MUTE, align=PP_ALIGN.RIGHT)
    app_knopf(slide, ix + iw - 1.3, iy + 0.1, 1.14, 0.26, "Aktualisieren")

    text(slide, ix + 0.2, 2.72, 3.0, 0.3, "Telemetrie", size=15, color=APP_TXT,
         bold=True)
    text(slide, ix + 0.2, 3.02, 5.4, 0.2, "Live-Sensordaten der Bodenstation – "
         "durchsuchbar und sortierbar.", size=8, color=APP_DIM)

    kw = (iw - 0.4 - 2 * 0.25) / 3
    for i, (label, wert, einheit) in enumerate(
            [("Datensätze", "2 315", None), ("Ø Druck", "4,71", "bar"),
             ("Ø Temperatur", "351,8", "K")]):
        app_kennzahl(slide, ix + 0.2 + i * (kw + 0.25), 3.32, kw, label, wert,
                     einheit, h=0.7)

    app_suchfeld(slide, ix + 0.2, 4.16, iw - 0.4, h=0.36)
    app_tabelle(slide, ix + 0.2, 4.66, iw - 0.4, [
        ("thruster_1.a", "thruster", 3.35, 432.2, "06.08.2026, 14:31:02"),
        ("o_tank_1", "gas_valve", 5.10, 268.4, "06.08.2026, 14:30:53"),
        ("thruster_2.b", "thruster", 2.88, 455.7, "06.08.2026, 14:30:44"),
        ("h_tank_2", "gas_valve", 6.02, 240.1, "06.08.2026, 14:30:35"),
    ], zeilen_h=0.34, size=8)

    # Beschriftung: Marker am rechten Fensterrand, Legende auf gleicher Hoehe
    punkte = [
        (2.20, "Aktualisieren & „vor 3 s“",
         ["Eigene Komponente – der Sekundentakt rendert",
          "nur diese Zeile neu, nicht die Tabelle."], AMBER),
        (3.44, "Kennzahlen",
         ["Rechnen über die gefilterten Zeilen,",
          "verändern sich also mit der Suche."], ACCENT),
        (4.18, "Suche",
         ["⌘K springt hinein, Escape leert das Feld."], ACCENT),
        (4.72, "Sortierung",
         ["Alle 5 Spalten, je eigener Vergleich pro Typ."], VIOLET),
        (5.30, "Mini-Balken",
         ["Lage im gültigen Bereich: 0,5–9 bar / 200–500 K."], GREEN),
        (5.92, "Typ als Badge",
         ["thruster oder gas_valve – der Wert kommt",
          "unverändert aus der Datenbank."], GREEN),
    ]
    # Die Legende sitzt auf gleicher Hoehe wie der jeweilige Marker am
    # Fensterrand, deshalb braucht es keine Verbindungslinien.
    lx = bx + bw + 0.35
    for i, (y, titel, zeilen, farbe) in enumerate(punkte, start=1):
        marker(slide, bx + bw - 0.15, y - 0.02, i, farbe)
        legende(slide, lx, y, SW - LM - lx, i, titel, zeilen, farbe)

    fuss(slide, nr)
    return slide


def folie_fe_datenfluss(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Die Stelle, an der Backend und Frontend aufeinandertreffen. Das Backend "
        "liefert dreifach verschachtelt: Typ, dann Name, dann die Liste der "
        "Messwerte. Eine Tabelle braucht aber eine flache Liste. Genau das macht "
        "flachKlopfen mit drei ineinanderliegenden Schleifen - die Farben zeigen, "
        "welche Schleife welche Ebene aufloest. Der Preis dafuer steht unten: aus "
        "13 Sensoren werden 2315 Zeilen, und zwar bei jedem Abruf neu."))
    kopf(slide, "Frontend · Datenfluss", "Von der Antwort zur Tabellenzeile",
         farbe=AMBER)

    text(slide, LM, 1.72, 5.15, 0.22, "Antwort von GET /data_wsi/", size=9,
         color=ACCENT, bold=True, caps=True)
    code_farbig(slide, LM, 2.02, 5.15, [
        ("{", DIM),
        ('  "thruster": {', ACCENT),
        ('      "thruster_1.a": [', VIOLET),
        ('          { "time": "…T14:31:02",', GREEN),
        ('            "pressure": 3.35,', GREEN),
        ('            "temperature": 432.2 },', GREEN),
        ('          { "time": "…T14:29:53", … }', GREEN),
        ("      ],", VIOLET),
        ('      "thruster_1.b": [ … ]', VIOLET),
        ("  },", ACCENT),
        ('  "gas_valve": { … }', ACCENT),
        ("}", DIM),
    ], size=8.5)

    tx = LM + 6.48
    tw = SW - LM - tx
    text(slide, tx, 1.72, tw, 0.22, "Was die Tabelle braucht", size=9, color=GREEN,
         bold=True, caps=True)
    app_tabelle(slide, tx, 2.02, tw, [
        ("thruster_1.a", "thruster", 3.35, 432.2, "14:31:02"),
        ("thruster_1.a", "thruster", 3.41, 430.8, "14:29:53"),
        ("thruster_1.b", "thruster", 3.12, 441.7, "14:31:11"),
        ("o_tank_1", "gas_valve", 5.10, 268.4, "14:30:53"),
    ], zeilen_h=0.33, size=8)

    arrow(slide, LM + 5.28, 3.0, tx - 0.12, 3.0, color=AMBER, width=1.75)
    text(slide, LM + 5.2, 2.62, 1.34, 0.34, ["flachKlopfen()", "3 Schleifen"],
         size=7.5, color=AMBER, align=PP_ALIGN.CENTER, spacing=1.25, font=MONO)

    text(slide, tx, 3.95, tw, 0.22, "eine Zeile pro Messwert", size=9, color=DIM,
         caps=True, bold=True)
    text(slide, tx, 4.24, tw, 0.6, [
        "Der Schlüssel _id wird aus type, name und time",
        "zusammengesetzt – damit React jede Zeile",
        "wiedererkennt, auch wenn sich die Sortierung ändert."],
        size=9, color=DIM, spacing=1.25)

    xk = chip(slide, LM, 4.9, "Ebene 1  ·  type", farbe=ACCENT, size=8, mono=False)
    xk = chip(slide, xk + 0.18, 4.9, "Ebene 2  ·  name", farbe=VIOLET, size=8,
              mono=False)
    chip(slide, xk + 0.18, 4.9, "Ebene 3  ·  Messwerte", farbe=GREEN, size=8,
         mono=False)

    code_farbig(slide, LM, 5.24, CW, [
        ("for (const [type, sensoren] of Object.entries(json))", ACCENT),
        ("  for (const [name, messungen] of Object.entries(sensoren ?? {}))", VIOLET),
        ("    for (const messung of messungen ?? [])", GREEN),
        ("      zeilen.push({ _id: `${type}-${name}-${messung.time}`, type, name, "
         "...messung })", TXT),
    ], size=9.5, streifen=AMBER)

    text(slide, LM, 6.56, CW, 0.24,
         "13 Sensoren × rund 178 Messwerte = 2 315 Zeilen – neu gebaut bei jedem Abruf.",
         size=10, color=AMBER)
    fuss(slide, nr)
    return slide


def folie_fe_abruf(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Wie das Frontend aktuell bleibt. Kein Push vom Server, sondern Abfragen "
        "im Takt. Der wichtige Punkt an der Zeitleiste: der Timer fuer die naechste "
        "Abfrage wird erst gesetzt, nachdem die Antwort da ist. Mit setInterval "
        "waere das anders - dann koennten sich langsame Anfragen stapeln und "
        "ueberholen. Rechts der Aufraeumteil, der leicht vergessen wird: ohne "
        "AbortController wuerde eine Antwort auf eine verlassene Ansicht "
        "zurueckkommen und in einer Komponente landen, die es nicht mehr gibt."))
    kopf(slide, "Frontend · Aktualisierung", "Wie die Daten nachkommen", farbe=AMBER)

    text(slide, LM, 1.76, CW, 0.22, "Der Abruf-Takt", size=9, color=ACCENT,
         bold=True, caps=True)

    achse_y = 3.02
    for i in range(4):
        bx = 1.15 + i * 2.70
        rect(slide, bx, 2.3, 0.62, 0.42, fill=PANEL2, line=ACCENT, line_w=1.25,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
        text(slide, bx, 2.42, 0.62, 0.2, "fetch", size=8, color=ACCENT,
             align=PP_ALIGN.CENTER, font=MONO)
        rect(slide, bx + 0.305, achse_y - 0.07, 0.011, 0.14, fill=BORDER)
        text(slide, bx - 0.3, achse_y + 0.12, 1.22, 0.2, f"{i + 1}. Abruf", size=8,
             color=FAINT, align=PP_ALIGN.CENTER)
        if i < 3:
            arrow(slide, bx + 0.68, 2.51, bx + 2.64, 2.51, color=RGBColor(0x3A, 0x5A, 0x86),
                  width=1.25, dashed=True)
            text(slide, bx + 1.1, 2.6, 0.78, 0.2, "10 s Pause", size=8, color=FAINT,
                 align=PP_ALIGN.CENTER)
    text(slide, 10.5, 2.42, 0.6, 0.2, "…", size=12, color=FAINT)
    line_h(slide, 1.15, achse_y, 10.0)
    text(slide, LM, 1.98, 4.0, 0.2, "GET /data_wsi/", size=8, color=DIM, font=MONO)

    text(slide, LM, 3.44, CW, 0.24,
         "Der Timer für die nächste Abfrage wird erst gesetzt, nachdem die Antwort "
         "da ist – deshalb kann sich nichts stapeln.",
         size=10.5, color=TXT)

    # ---- links: Zustandsautomat -----------------------------------------
    panel(slide, LM, 3.82, 5.6, 2.62)
    rect(slide, LM + 0.001, 4.0, 0.055, 2.26, fill=VIOLET)
    text(slide, LM + 0.28, 3.98, 4.0, 0.26, "Drei Zustände", size=12.5, color=TXT,
         bold=True)

    zustand(slide, LM + 0.3, 4.5, 1.5, 0.62, "laedt", farbe=ACCENT)
    zustand(slide, LM + 3.6, 4.24, 1.5, 0.58, "ok", farbe=GREEN)
    zustand(slide, LM + 3.6, 5.16, 1.5, 0.58, "fehler", farbe=RED)
    arrow(slide, LM + 1.84, 4.75, LM + 3.56, 4.53, color=GREEN, width=1.25)
    arrow(slide, LM + 1.84, 4.87, LM + 3.56, 5.45, color=RED, width=1.25)
    text(slide, LM + 1.9, 4.36, 1.6, 0.2, "Antwort da", size=8, color=GREEN)
    text(slide, LM + 1.9, 5.12, 1.7, 0.2, "Netz / Server", size=8, color=RED)
    text(slide, LM + 0.28, 5.9, 5.1, 0.42,
         "404 heißt „noch nichts empfangen“ und gilt bewusst als ok – "
         "die Tabelle zeigt dann ihren Leer-Hinweis.",
         size=9, color=DIM, spacing=1.25)

    # ---- rechts: Aufraeumen ---------------------------------------------
    rx = LM + 5.9
    panel(slide, rx, 3.82, 5.73, 2.62)
    rect(slide, rx + 0.001, 4.0, 0.055, 2.26, fill=GREEN)
    text(slide, rx + 0.28, 3.98, 4.6, 0.26, "Aufräumen beim Verlassen", size=12.5,
         color=TXT, bold=True)

    schritte = [
        ("Ansicht wird verlassen", "React ruft die Aufräumfunktion"),
        ("controller.abort()", "das laufende fetch wird abgebrochen"),
        ("fetch wirft AbortError", "kein echter Fehler, nur ein Abbruch"),
        ("return", "kein setState auf eine leere Komponente"),
    ]
    rect(slide, rx + 0.39, 4.46, 0.011, 1.56, fill=BORDER)
    for i, (titel, unter) in enumerate(schritte):
        y = 4.4 + i * 0.52
        marker(slide, rx + 0.28, y, i + 1, GREEN, d=0.24)
        text(slide, rx + 0.66, y + 0.01, 4.9, 0.2, titel, size=9.5, color=TXT,
             bold=True, font=MONO)
        text(slide, rx + 0.66, y + 0.21, 4.9, 0.2, unter, size=8, color=DIM)

    fuss(slide, nr)
    return slide


def folie_fe_suche(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Suche und Sortierung, die zwei Funktionen, die der Space Controller "
        "wirklich benutzt. Bei der Suche ist der Kniff die UND-Verknuepfung: "
        "mehrere Woerter verengen, sie erweitern nicht. Und ein Detail, das man "
        "kennen muss: gesucht wird im rohen Zeitstempel, nicht in der angezeigten "
        "Form. Wer '06.08.2026' eingibt, findet nichts - '2026-08-06' schon."))
    kopf(slide, "Frontend · Bedienung", "Suchen und Sortieren", farbe=AMBER)

    # ---- links: der Filter als Ablauf -----------------------------------
    lw = 4.5
    text(slide, LM, 1.72, lw, 0.22, "Eingabe", size=9, color=ACCENT, bold=True,
         caps=True)
    app_suchfeld(slide, LM, 2.0, lw, inhalt="thruster 14:31", h=0.34)

    arrow(slide, LM + lw / 2, 2.42, LM + lw / 2, 2.74, color=AMBER, width=1.4)
    xk = chip(slide, LM + 0.62, 2.8, "thruster", farbe=AMBER, size=8.5)
    text(slide, xk + 0.06, 2.83, 0.55, 0.2, "UND", size=8, color=DIM,
         align=PP_ALIGN.CENTER, bold=True)
    chip(slide, xk + 0.61, 2.8, "14:31", farbe=AMBER, size=8.5)

    arrow(slide, LM + lw / 2, 3.1, LM + lw / 2, 3.42, color=AMBER, width=1.4)
    code_farbig(slide, LM, 3.48, lw, [
        ('text = [name, type, time]', DIM),
        ('         .join(" ").toLowerCase()', DIM),
        ("begriffe.every(b => text.includes(b))", AMBER),
    ], size=8.5, streifen=AMBER)

    karte(slide, LM, 4.52, lw, 1.28, "Gesucht wird im rohen Zeitstempel", [
        "2026-08-06T14:31:02  findet man",
        "06.08.2026, 14:31:02  nicht"],
        farbe=RED, titel_size=10.5, zeile_size=9)

    # ---- rechts: das Ergebnis -------------------------------------------
    rx = LM + 5.15
    rw = SW - LM - rx
    text(slide, rx, 1.72, rw, 0.22, "Ergebnis", size=9, color=GREEN, bold=True,
         caps=True)
    app_tabelle(slide, rx, 2.0, rw, [
        ("thruster_1.a", "thruster", 3.35, 432.2, "2026-08-06T14:31:02"),
        ("o_tank_1", "gas_valve", 5.10, 268.4, "2026-08-06T14:31:11"),
        ("thruster_2.b", "thruster", 2.88, 455.7, "2026-08-06T14:31:20"),
        ("thruster_1.a", "thruster", 3.41, 430.8, "2026-08-06T14:29:53"),
        ("h_tank_2", "gas_valve", 6.02, 240.1, "2026-08-06T14:30:44"),
        ("thruster_3.c", "thruster", 4.15, 398.6, "2026-08-06T14:31:35"),
    ], zeilen_h=0.34, size=8, treffer=(0, 2, 5))

    text(slide, rx, 4.5, rw, 0.22, "3 von 6 Datensätzen", size=9, color=APP_BLUE,
         bold=True)
    text(slide, rx, 4.76, rw, 0.24,
         "Zeile 2 scheitert am Namen, Zeile 4 und 5 an der Uhrzeit.",
         size=9, color=DIM)

    karte(slide, rx, 5.2, 3.1, 1.24, "Sortierung", [
        "Text  ·  localeCompare(de)",
        "Zahl  ·  x − y",
        "Zeit  ·  parseZeit() → ms"],
        farbe=VIOLET, titel_size=10.5, zeile_size=9)

    bkx = rx + 3.4
    bkw = SW - LM - bkx
    karte(slide, bkx, 5.2, bkw, 1.24, "Mini-Balken", [], farbe=GREEN,
          titel_size=10.5)
    for i, (label, wert, unten, oben) in enumerate(
            [("3,35 bar", 3.35, 0.5, 9.0), ("432,2 K", 432.2, 200.0, 500.0)]):
        y = 5.72 + i * 0.3
        text(slide, bkx + 0.28, y, 0.95, 0.18, label, size=8.5, color=TXT, font=MONO)
        app_balken(slide, bkx + 1.3, y + 0.07, bkw - 1.6, (wert - unten) / (oben - unten))
    fuss(slide, nr)
    return slide


def folie_abhaengigkeiten(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Diese Folie beantwortet die Frage aus dem Aufgabenblatt: wer ruft wen "
        "auf. Die Kopplung ist absichtlich lose. Besonders schoen: faellt die "
        "Verwaltung aus, bleiben die Dateien einfach im Ordner liegen und gehen "
        "nicht verloren - allerdings nur, solange die Bodenstation sie nicht "
        "vorher schon geloescht hat. Das ist eine ehrliche Schwachstelle."))
    kopf(slide, "Kopplung", "Wer hängt von wem ab")

    gitter(slide, LM, 1.82, [2.5, 2.5, 2.9, 3.73],
           ["Komponente", "braucht", "Schnittstelle", "Fällt sie aus, dann …"],
           [["Satellit", "nichts", "schreibt Dateien",
             "keine neuen Daten, System bleibt stabil"],
            ["Bodenstation", "Dateiablage, Verwaltung", "liest Dateien, HTTP POST",
             "Dateien stapeln sich im Ordner data/"],
            ["Verwaltung", "MongoDB", "HTTP, Motor-Treiber",
             "Bodenstation bekommt Fehler, Frontend leer"],
            ["MongoDB", "Docker", "Port 27017",
             "Verwaltung antwortet mit 503"],
            ["Frontend", "Verwaltung", "3 × HTTP GET",
             "Tabelle meldet den Fehler, Daten laufen weiter"]],
           zeilen_h=0.5, size=10)

    karte(slide, LM, 4.88, 5.6, 1.8, "Was gut funktioniert", [
        "Jede Komponente kennt nur ihren direkten Nachbarn.",
        "Nur die Verwaltung spricht mit der Datenbank.",
        "Jeder Prozess ist einzeln neu startbar."],
        farbe=GREEN, titel_size=12, zeile_size=10)

    karte(slide, LM + 5.9, 4.88, 5.73, 1.8, "Wo es noch wehtut", [
        "Die Bodenstation löscht die Datei auch dann, wenn der",
        "POST fehlschlägt – der Messwert ist dann verloren.",
        "Ein Wiederholversuch braucht vorher einen Schutz gegen",
        "doppelte Einträge in der Historie."],
        farbe=RED, titel_size=12, zeile_size=10)
    fuss(slide, nr)
    return slide


def folie_fehler(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Fuenf echte Fehler aus der Entwicklung. Diese Folie ist der "
        "interessanteste Teil des Vortrags: nicht was geplant war, sondern was "
        "schiefging und wie man es gemerkt hat. Bei jedem Punkt kurz sagen, "
        "woran man den Fehler bemerkt hat - der Statuscode ist meist der Hinweis."))
    kopf(slide, "Aus der Praxis", "Fünf Fehler, die wir gefunden haben", farbe=AMBER)

    gitter(slide, LM, 1.82, [1.15, 3.5, 3.6, 3.38],
           ["Symptom", "Ursache", "Warum es passierte", "Lösung"],
           [["422", "Zeitstempel im falschen Format",
             "Dateiname 20260806_1102 ist kein Datum", "In ISO-Format umwandeln"],
            ["500", "find_many gibt es nicht",
             "Methode existiert in Motor nicht", "find bzw. aggregate"],
            ["500", "Altdaten ohne Feld current",
             "Ältere Codeversion schrieb anderes Format", "Filter plus Aufräumen"],
            ["still", "Älterer Wert überschrieb den neuen",
             "Reihenfolge der Dateien war zufällig", "Nur ersetzen wenn jünger"],
            ["still", "Historie war nicht chronologisch",
             "Verzeichnis liefert keine Sortierung", "Dateien sortiert einlesen"]],
           zeilen_h=0.5, size=10, kopf_farbe=AMBER,
           farben=[RED, RED, RED, AMBER, AMBER])

    karte(slide, LM, 4.88, 5.6, 1.8, "Die stillen Fehler waren die schlimmeren", [
        "Ein Statuscode fällt sofort auf. Ein falsch sortierter",
        "Verlauf sieht auf den ersten Blick richtig aus –",
        "auffallen würde er erst im Diagramm."],
        farbe=AMBER, titel_size=12, zeile_size=10)

    karte(slide, LM + 5.9, 4.88, 5.73, 1.8, "Was uns beim Suchen geholfen hat", [
        "Die Fehlermeldung des Servers nennt Feld und Grund.",
        "Den gesendeten Datensatz und die Antwort direkt",
        "nebeneinander protokollieren – dann sieht man sofort,",
        "ob der Fehler beim Sender oder beim Empfänger liegt."],
        farbe=ACCENT, titel_size=12, zeile_size=10)
    fuss(slide, nr)
    return slide


def folie_422(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Ein Fehler im Detail, weil er alles zeigt: die Fehlermeldung war "
        "praezise, man musste sie nur lesen. Und die zweite Lehre war noch "
        "wichtiger: der Code war schon repariert, aber der laufende Prozess "
        "hatte die Datei nie neu gelesen. Der Server laedt mit --reload "
        "automatisch neu, eigene Skripte nicht."))
    kopf(slide, "Backend · Im Detail", "Der Fehler 422", farbe=AMBER)

    text(slide, LM, 1.78, 5.6, 0.28, "Was der Server antwortete", size=9.5,
         color=RED, bold=True, caps=True)
    code(slide, LM, 2.1, 5.6, [
        '"loc":   ["body", "time"],',
        '"msg":   "Input should be a valid',
        '          datetime or date",',
        '"input": "20260806_110216"',
    ], size=10, farbe=RED)

    text(slide, LM + 5.9, 1.78, 5.73, 0.28, "Was gesendet wurde – und was nötig war",
         size=9.5, color=GREEN, bold=True, caps=True)
    code(slide, LM + 5.9, 2.1, 5.73, [
        'falsch:  "time": "20260806_110216"',
        'richtig: "time": "2026-08-06T11:02:16"',
        "",
        'datetime.strptime(name, "%Y%m%d_%H%M%S")',
        "        .isoformat()",
    ], size=10, farbe=GREEN)

    line_h(slide, LM, 4.05, CW)

    karte(slide, LM, 4.28, 5.6, 1.85, "Was der Statuscode verrät", [
        "422 heißt: die Anfrage kam an, aber der Inhalt",
        "passt nicht zum Modell. Also kein Netzwerkproblem.",
        "Die Meldung nannte Feld, Grund und den gesendeten",
        "Wert – damit war die Ursache eindeutig."],
        farbe=AMBER, titel_size=12, zeile_size=10)

    karte(slide, LM + 5.9, 4.28, 5.73, 1.85, "Die zweite, wichtigere Lehre", [
        "Nach dem Fix blieb der Fehler – der laufende Prozess",
        "hatte die geänderte Datei nie neu gelesen.",
        "uvicorn lädt mit --reload selbst neu.",
        "Eigene Skripte muss man von Hand neu starten."],
        farbe=VIOLET, titel_size=12, zeile_size=10)
    fuss(slide, nr)
    return slide


def folie_ausblick(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Ehrlicher Ausblick, links Frontend und rechts Backend. Zwei Punkte, die "
        "auf einer aelteren Fassung dieser Folie noch offen standen, sind "
        "inzwischen erledigt - das steht unten als eigene Zeile. Beim ersten "
        "Frontend-Punkt ist der Zusammenhang wichtig: die Datenansicht laedt die "
        "komplette Historie, deshalb rechnen die Mittelwerte ueber alle Messwerte "
        "und nicht ueber den aktuellen Stand. Die Zahlen rechts sind ausgerechnet, "
        "nicht geschaetzt: 16 MB geteilt durch die gemessene Groesse eines "
        "Messwerts ergibt 255.882 Messwerte, bei der gemessenen Rate etwa 190 Tage."))
    kopf(slide, "Ausblick", "Was als Nächstes kommt", farbe=VIOLET)

    text(slide, LM, 1.74, 5.6, 0.24, "Frontend", size=9.5, color=AMBER, bold=True,
         caps=True)
    gitter(slide, LM, 2.04, [2.95, 2.65],
           ["Aufgabe", "Warum"],
           [["Verlaufsdiagramm", "Die Historie ist bisher nur eine Tabelle"],
            ["Ø-Werte auf den aktuellen Stand",
             "sie rechnen heute über die ganze Historie"],
            ["DataFetching.jsx löschen",
             "Restdatei, ruft GET /data/ – gibt es nur als POST"],
            ["Die zwei Abruf-Schleifen zusammenlegen",
             "Startseite und Datenansicht fragen getrennt"]],
           zeilen_h=0.54, size=9.5, kopf_farbe=AMBER, farben=[AMBER] * 4)

    rx = LM + 5.9
    text(slide, rx, 1.74, 5.61, 0.24, "Backend", size=9.5, color=VIOLET, bold=True,
         caps=True)
    gitter(slide, rx, 2.04, [2.95, 2.66],
           ["Thema", "Ab wann es wehtut"],
           [["Ein Dokument pro Messwert",
             "255 882 Messwerte, ca. 190 Tage Dauerbetrieb"],
            ["Seitenweises Abrufen der Historie",
             "ab etwa 40 000 Messwerten pro Sensor"],
            ["Index auf Typ und Name", "ab einigen hundert Sensoren"],
            ["Schutz gegen doppelte Einträge",
             "sobald ein Wiederholversuch dazukommt"]],
           zeilen_h=0.54, size=9.5, kopf_farbe=VIOLET, farben=[VIOLET] * 4)

    panel(slide, LM, 4.66, CW, 0.78)
    rect(slide, LM + 0.001, 4.82, 0.055, 0.46, fill=GREEN)
    text(slide, LM + 0.28, 4.8, 3.2, 0.24, "Seit dem letzten Stand erledigt",
         size=10.5, color=TXT, bold=True)
    xc = chip(slide, LM + 3.7, 4.78, "CORS freigeschaltet", farbe=GREEN, size=8.5,
              mono=False)
    xc = chip(slide, xc + 0.16, 4.78, "Frontend angebunden", farbe=GREEN, size=8.5,
              mono=False)
    chip(slide, xc + 0.16, 4.78, "3 von 4 Routen in Benutzung", farbe=GREEN,
         size=8.5, mono=False)
    text(slide, LM + 0.28, 5.08, 11.0, 0.24,
         "Die Anzeige läuft – offen ist nicht mehr die Verbindung, sondern was sie "
         "genau holen soll.", size=9, color=DIM)

    line_h(slide, LM, 5.72, CW)
    text(slide, LM, 5.9, CW, 0.5, [
        "Woher die Zahl 255 882 kommt: MongoDB erlaubt pro Dokument 16 MB, ein Messwert "
        "braucht gemessene 64 Byte.",
        "Die saubere Lösung wäre ein eigenes Dokument pro Messwert statt einer langen "
        "Liste – gut zu wissen, für dieses Projekt aber nicht nötig."],
        size=10, color=DIM, spacing=1.3, space_after=3)
    fuss(slide, nr)
    return slide


def folie_fazit(prs, nr):
    slide = neue_folie(prs, notiz=(
        "Abschluss. Die drei Punkte sind die eigentlichen Lernergebnisse und "
        "gelten unabhaengig von diesem Projekt. Dann Raum fuer Fragen."))
    kopf(slide, "Abschluss", "Was wir gelernt haben", farbe=GREEN)
    stars(slide, anzahl=26, seed=13, ymax=1.3)

    punkte = [
        ("Schnittstellen zuerst klären",
         ["Der 422er war kein Programmierfehler, sondern eine",
          "Unstimmigkeit über ein Format. Beide Seiten für sich",
          "waren richtig – nur eben nicht zueinander passend."], ACCENT),
        ("Fehlermeldungen genau lesen",
         ["Die Antwort des Servers nannte Feld, Grund und Wert.",
          "Das Suchen bestand darin, die Meldung ernst zu nehmen,",
          "und nicht darin, im Code herumzuprobieren."], AMBER),
        ("Prüfen, was wirklich läuft",
         ["Geänderter Code wirkt erst, wenn der Prozess ihn",
          "gelesen hat. Zweimal haben wir einen Fehler gesucht,",
          "der längst behoben war."], VIOLET),
    ]
    for i, (titel, zeilen, farbe) in enumerate(punkte):
        x = LM + i * (3.95)
        panel(slide, x, 1.9, 3.73, 2.6)
        text(slide, x + 0.28, 2.1, 0.6, 0.4, f"{i + 1:02d}", size=22, color=farbe, bold=True)
        text(slide, x + 0.28, 2.66, 3.2, 0.55, titel, size=13, color=TXT, bold=True,
             spacing=1.15)
        text(slide, x + 0.28, 3.3, 3.25, 1.1, list(zeilen), size=9.5, color=DIM,
             spacing=1.3)

    panel(slide, LM, 4.78, CW, 1.25)
    text(slide, LM, 4.98, CW, 0.4,
         "Ein Messwert legt einen weiten Weg zurück.", size=17, color=TXT, bold=True,
         align=PP_ALIGN.CENTER)
    text(slide, LM, 5.44, CW, 0.4,
         "Satellit  →  Datei  →  Bodenstation  →  Verwaltung  →  Datenbank  →  Anzeige",
         size=12, color=ACCENT, align=PP_ALIGN.CENTER)

    text(slide, LM, 6.15, CW, 0.4, "Vielen Dank – Fragen?", size=14, color=DIM,
         align=PP_ALIGN.CENTER)
    fuss(slide, nr)
    return slide


# ----------------------------------------------------------------------------
# Zusammenbauen
# ----------------------------------------------------------------------------

def main():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)

    folie_titel(prs)

    # Backend und Frontend bekommen gleich viel Platz: je sechs Folien.
    rahmen_vorn = [folie_aufgabe, folie_architektur]
    backend = [
        folie_satellit, folie_bodenstation, folie_verwaltung, folie_datenmodell,
        folie_mongodb,
    ]
    frontend = [
        folie_fe_ueberblick, folie_fe_komponenten, folie_fe_ansicht,
        folie_fe_datenfluss, folie_fe_abruf, folie_fe_suche,
    ]
    rahmen_hinten = [
        folie_abhaengigkeiten, folie_fehler, folie_422, folie_ausblick, folie_fazit,
    ]
    # folie_422 gehoert inhaltlich zum Backend, steht aber bei den Fehlern.
    inhalt = rahmen_vorn + backend + frontend + rahmen_hinten
    for i, bauen in enumerate(inhalt, start=2):
        bauen(prs, i)

    ziel = "Satelliten-Telemetrie.pptx"
    prs.save(ziel)
    anzahl = len(prs.slides._sldIdLst)
    print(f"{ziel} geschrieben – {anzahl} Folien "
          f"({len(backend) + 1} Backend / {len(frontend)} Frontend)")


if __name__ == "__main__":
    main()
