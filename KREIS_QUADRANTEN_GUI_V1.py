# GUI = Graphical User Interface - Deutsch üblich: Grafische Benutzeroberfläche (auch: „grafische Oberfläche“)

# ============================================================
# KREIS_QUADRANTEN_GUI.py
# 3 Reihen à 4 Kreise: Daten aus C5:F6, C9:F10, C13:F14
# Fehlende Buchstaben -> MsgBox nach dem Zeichnen, dann Ende
# ============================================================

# ============================================================
# KREIS_QUADRANTEN_GUI – Kurzüberblick (Spickzettel)
# ============================================================
#
# WICHTIGE IDEE:
# - Layout (Zeilenhöhen) und Geometrie (Kreispositionen) sind gekoppelt.
# - Kreise werden NICHT mehr über START_Y + OFFSET_Y gerechnet,
#   sondern über die Oberkante bestimmter Tabellenzeilen.
#
# ABSCHNITTE:
# 1) KONFIGURATION
#    - CIRCLE_DIAMETER: Kreis-Durchmesser (1/100 mm)
#    - OFFSET_X_BASE: horizontaler Abstand zwischen Kreisen
#    - START_X: linker Startpunkt der Kreisspalten
#    - GROUPS: Zuordnung der Zeilen (EF-Input / Fallback-Zeilen)
#
# 2) INIT (Format / einmalig)
#    - Spaltenbreiten, Zentrierung, Merge E+F, Überschriften, Freeze
#    - Wird automatisch 1x pro Dokument ausgeführt (_ensure_initialized)
#
# 3) ZEILEN AN KREIS ANPASSEN (Geometrie / jedes Mal)
#    - adjust_rows_to_circle_radius(sheet):
#      Wortzeile + Buchstabenzeile = Radius
#      Buchstabenzeile + Wortzeile = Radius (unten spiegelverkehrt)
#
# 4) KREISE AM ZEILENRASTER AUSRICHTEN (Geometrie / jedes Mal)
#    - _row_top_y(sheet, row): Y der Zeilenoberkante
#    - _circle_top_y_for_group(sheet, group_index):
#      Kreis-Top-Y = Oberkante EF-Wortzeile der Gruppe
#
# 5) HAUPT-MAKROS
#    - create_circle_grid(): löscht, initialisiert, passt Zeilen an, zeichnet Kreise
#    - update_texts_only(): aktualisiert nur Texte (kein Neuzeichnen)
#    - reflow_circles_only(): verschiebt bestehende Kreise/Lines/Texte ans Raster
#    - scramble/rotate: mischt/rotiert die Quadranten
#
# MERKSATZ:
# - Init = Darstellung/Format (einmalig)
# - adjust_rows_to_circle_radius + _circle_top_y_for_group = Geometrie (jedes Mal)
#
# ============================================================

import uno
import random
from com.sun.star.awt import Point, Size

# ============================================================
# 1) KONFIGURATION
# ============================================================

SHEET_NAME = "KREIS_GUI"
MARK_DESC  = "KREIS_QUADRANTEN_GUI"   # Marker zum Löschen (Description)

EF_ROWS = (3, 6, 8, 11, 13, 16)
GRID_ROWS = (4, 5, 9, 10, 14, 15)

NUM_COLS = 4            # C..F

CIRCLE_DIAMETER = 3900   # Durchmesser in 1/100 mm
# Radius wäre: CIRCLE_RADIUS = CIRCLE_DIAMETER // 2

OFFSET_X_BASE = 4000  # alt: 4000

START_X = 14000 # horizontaler Abstand vom linken Tabellenrand

ROW_COLORS = [
    0xCCE5FF,  # oben: blau
    0xFBE5B6,  # mitte: beige
    0xCCFFCC,  # unten: grün
]

CHAR_HEIGHT = 16        # Buchstaben kleiner/größer

# Datenblöcke (1-basiert wie in Calc):
# (Titel, obere Zeile, untere Zeile)
# 3 Kreis-Gruppen (je 4 Kreise):
# ef_top  -> Wort für obere Hälfte (liefert C..F in top_row)
# ef_bot  -> Wort für untere Hälfte (liefert C..F in bot_row)

GROUPS = [
    ("Gruppe 1",  3,  4,  5,  6),   # EF3 -> CF4, EF6 -> CF5
    ("Gruppe 2",  8,  9, 10, 11),   # EF8 -> CF9, EF11 -> CF10
    ("Gruppe 3", 13, 14, 15, 16),   # EF13 -> CF14, EF16 -> CF15
]

# ============================================================
# INIT (Konstanten)
# ============================================================
U100MM_PER_CM = 1000
def cm(v): return int(round(v * U100MM_PER_CM))

INIT_CENTER_RANGE = "B2:F50"
INIT_HEADER_RANGE = "B2:F2"
INIT_WORD_ROWS    = (3, 6, 8, 11, 13, 16)
INIT_COL_W_C_TO_G = cm(1.7)
INIT_COL_W_J_L    = cm(4.0)
INIT_COL_W_B_MAX  = cm(4.5)
INIT_ROW_H_17     = cm(1.7)
INIT_ROW_H_12     = cm(1.2)
INIT_ROW_H_09     = cm(0.9)

INIT_ROWS_H17 = (1, 2, 17, 18)             # 1,7 cm

# ============================================================
# 2) UNO / CALC HELFER
# ============================================================

def _get_doc():
    return XSCRIPTCONTEXT.getDocument()

def _get_sheet(doc):
    return doc.Sheets.getByName(SHEET_NAME)

def _get_draw_page(sheet):
    return sheet.DrawPage

def _col_to_name(col0):
    name = ""
    n = col0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        name = chr(65 + r) + name
    return name

def _a1(col0, row0):
    return f"{_col_to_name(col0)}{row0 + 1}"

def _msgbox(doc, title, message):
    try:
        parent = doc.CurrentController.Frame.ContainerWindow
        toolkit = parent.getToolkit()
        buttons = uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK")
        box_type = uno.Enum("com.sun.star.awt.MessageBoxType", "WARNINGBOX")
        box = toolkit.createMessageBox(parent, box_type, buttons, title, message)
        box.execute()
    except Exception:
        pass

def _format_messages(title: str, items: list, max_lines: int = 25) -> str:
    """
    English comment: Make messageboxes shorter & readable.
    Shows only first max_lines entries + summary.
    """
    if not items:
        return ""
    items = [x for x in items if x]
    if not items:
        return ""

    if len(items) <= max_lines:
        body = "\n".join(items)
        return f"{title} ({len(items)}):\n{body}"

    head = "\n".join(items[:max_lines])
    rest = len(items) - max_lines
    return f"{title} ({len(items)}):\n{head}\n...\n(+{rest} weitere)"

# ============================================================
# 3) DATEN LESEN + LESEN/WORTPAARE + VALIDIEREN
# ============================================================


def _pairs_from_word(word, cell_label):
    """
    Wort -> 4 Zweierpäckchen für C..F.
    Regeln:
    - Wird nur aufgerufen, wenn EF >= 4 Zeichen hat
    - >8 => abschneiden (Hinweis)
    - 4..7 => Hinweis, aber NICHT auffüllen (fehlende Quadranten bleiben leer)
    - genau 8 => perfekt
    """
    errors = []
    clean = _normalize_for_crossword(word)   # <-- IMPORTANT: OE/AE/UE/SS for circles
    n = len(clean)

    if n > 8:
        errors.append(f"{cell_label}: '{word}' hat {n} Zeichen → auf 8 gekürzt")
        clean = clean[:8]
    elif n != 8:
        errors.append(f"{cell_label}: '{word}' hat {n} Zeichen → erwartet 8 (keine Auffüllung)")

    # Paare bilden, fehlende Paare als "" (leer) lassen
    pairs = []
    for i in range(0, 8, 2):
        chunk = clean[i:i+2]
        pairs.append(chunk if len(chunk) == 2 else "")
    return pairs, errors


def _pairs_from_row_cells(sheet, row_1based):
    """
    Reads C..F in one row. Each cell is intended to be exactly 2 characters.
    Rules:
      - If >2 chars: cut to first 2
      - If 1 char: KEEP it and pad with a space (so the letter is not lost)
      - If empty: use two spaces
    Returns: (pairs, errors)
    """
    errors = []
    pairs = []
    row0 = row_1based - 1

    for col in range(2, 6):  # C..F
        raw = (sheet.getCellByPosition(col, row0).getString() or "")
        addr = _a1(col, row0)

        # IMPORTANT: don't strip away intentional spaces too early
        raw = raw.strip()

        norm = _normalize_for_crossword(raw)  # OE/AE/UE/SS, only A-Z

        if len(norm) > 2:
            cut = norm[:2]
            errors.append(f"{addr}: >2 Zeichen → auf '{cut}' gekürzt")
            pairs.append(cut)
            continue

        if len(norm) == 2:
            pairs.append(norm)
            continue

        if len(norm) == 1:
            # <-- THIS is the bugfix: keep the letter
            pairs.append(norm + " ")
            errors.append(f"{addr}: fehlt 1 Buchstabe (erwartet 2)")
            continue

        # len == 0
        pairs.append("  ")
        errors.append(f"{addr}: fehlen beide Buchstaben (erwartet 2)")

    return pairs, errors


def _select_cell(doc, sheet, col0, row_1based):
    """English comment: Put cursor into a single cell reliably."""
    try:
        ctrl = doc.CurrentController
        cell = sheet.getCellByPosition(col0, row_1based - 1)
        ctrl.select(cell)
        try:
            ctrl.setActiveCell(cell)
        except Exception:
            pass
        return True
    except Exception:
        return False

def set_cursor_next_input(doc, sheet):
    """
    English comment: Cursor priority:
      1) first empty EF cell (E column) in rows 3,6,8,11,13,16
      2) else first empty cell in C rows 4,5,9,10,14,15
      3) else D, then E, then F in the same rows
    """
    # 1) EF (merged E+F -> we use E)
    for r in EF_ROWS:
        s = (sheet.getCellByPosition(4, r - 1).getString() or "").strip()  # E
        if not s:
            if _select_cell(doc, sheet, 4, r):
                return

    # 2) Grid: C then D then E then F
    for col0 in (2, 3, 4, 5):  # C,D,E,F
        for r in GRID_ROWS:
            s = (sheet.getCellByPosition(col0, r - 1).getString() or "").strip()
            if not s:
                if _select_cell(doc, sheet, col0, r):
                    return



def _get_pairs_for_half(sheet, ef_row_1based, fallback_row_1based):
    """
    Entscheidet:
    - wenn EF <4 Zeichen: nimm C..F aus fallback_row
    - wenn EF >=4 Zeichen: nutze Wort (8 Zeichen Ziel)
    """
    ef_word = _get_ef_word(sheet, ef_row_1based)
    cell_label = f"EF{ef_row_1based}"

    if len("".join(ch for ch in ef_word if not ch.isspace())) >= 4:
        return _pairs_from_word(ef_word, cell_label)
    else:
        return _pairs_from_row_cells(sheet, fallback_row_1based)


def _upper_keep_umlauts(s):
    """For EF cells: uppercase only, keep ÄÖÜ, show ß as ẞ."""
    return (s or "").replace("ß", "ẞ").upper()

def _normalize_for_crossword(s):
    """
    For circles (crossword logic):
    - uppercase
    - ÄÖÜ -> AE/OE/UE
    - ß/ẞ -> SS
    - keep only A-Z
    """
    x = (s or "")
    x = "".join(ch for ch in x if not ch.isspace())
    x = x.replace("ß", "ẞ").upper()
    x = (x.replace("Ä", "AE")
           .replace("Ö", "OE")
           .replace("Ü", "UE")
           .replace("ẞ", "SS"))
    x = "".join(ch for ch in x if "A" <= ch <= "Z")
    return x



# ============================================================
# 4) Helper - UPPERCASE
# ============================================================

def _to_upper_visual(s):
    # Verhindert, dass "ß" zu "SS" wird (würde sonst deine 8-Zeichen-Logik sprengen)
    return (s or "").replace("ß", "ẞ").upper()

def _get_ef_word(sheet, ef_row_1based):
    cell = sheet.getCellByPosition(4, ef_row_1based - 1)  # Spalte E (EF merged)
    raw = cell.getString() or ""

    # EF: keep umlauts visually
    vis = _upper_keep_umlauts(raw)
    if raw != vis:
        cell.setString(vis)

    return vis.strip()  # RETURN: the visual text (with umlauts)

def _upper_cell(cell):
    raw = cell.getString() or ""
    up = _to_upper_visual(raw)
    if raw != up:
        cell.setString(up)

def normalize_input_cells(sheet):
    """
    Schreibt rein optisch Großbuchstaben in:
    - alle EF-Zellen der Gruppen (E-Spalte, da EF verbunden)
    - alle C..F-Zellen der zugehörigen Datenzeilen (Fallback-Zeilen)
    """
    # EF-Zeilen + Datenzeilen sammeln
    ef_rows = []
    data_rows = []
    for (_title, ef_top, top_row, bot_row, ef_bot) in GROUPS:
        ef_rows.extend([ef_top, ef_bot])
        data_rows.extend([top_row, bot_row])  # die C..F-Zeilen

    # EF (verbunden E+F -> wir schreiben in E)
    for r in set(ef_rows):
        cell = sheet.getCellByPosition(4, r - 1)  # E (merged EF)
        raw = cell.getString() or ""
        up  = _upper_keep_umlauts(raw)
        if raw != up:
            cell.setString(up)


    # C..F Zweier-Zellen
    for r in set(data_rows):
        row0 = r - 1
        for col in range(2, 6):  # C..F
            _upper_cell(sheet.getCellByPosition(col, row0))

def disable_spellcheck_for_range(sheet, a1_range: str):
    """
    English comment: Disable spell checking (red wavy underline) for a given cell range
    by setting language to 'None' (CharLocale).
    """
    try:
        rng = sheet.getCellRangeByName(a1_range)
        loc = uno.createUnoStruct("com.sun.star.lang.Locale")
        loc.Language = ""   # empty = "no language"
        loc.Country  = ""
        loc.Variant  = ""
        rng.CharLocale = loc
    except Exception:
        pass

# ============================================================
# 5) ZEICHNEN (PRIMITIVES)
# ============================================================

def _mark_shape(shape, name_suffix):
    try:
        shape.Name = f"{MARK_DESC}_{name_suffix}"
    except Exception:
        pass
    try:
        shape.Description = MARK_DESC
    except Exception:
        pass


def draw_text_center(doc, draw_page, cx, cy, text, name_suffix):
    t = doc.createInstance("com.sun.star.drawing.TextShape")
    _mark_shape(t, name_suffix)
    draw_page.add(t)  # erst adden -> dann zuverlässig setzen

    # ---> HIER: Parameter "text" umwandeln (nicht "txt")
    text = _to_upper_visual(text)

    w, h = 700, 700
    t.Size = Size(w, h)
    t.Position = Point(int(cx - w/2), int(cy - h/2))

    try:
        t.FillStyle = uno.Enum("com.sun.star.drawing.FillStyle", "NONE")
        t.LineStyle = uno.Enum("com.sun.star.drawing.LineStyle", "NONE")
    except Exception:
        pass

    try:
        t.String = text
    except Exception:
        try:
            t.Text.setString(text)
        except Exception:
            pass

    try:
        t.CharHeight = CHAR_HEIGHT
        t.CharWeight = 150
        t.CharColor  = 0x000000
        t.TextHorizontalAdjust = 1  # CENTER
        t.TextVerticalAdjust   = 1  # CENTER
    except Exception:
        pass

    return t

def draw_line(doc, draw_page, x1, y1, x2, y2, name_suffix):
    """
    Zeichnet eine Linie von (x1,y1) nach (x2,y2) auf die DrawPage.
    Wichtig: erst adden, dann Eigenschaften setzen (LO/UNO stabiler).
    """
    line = doc.createInstance("com.sun.star.drawing.LineShape")
    _mark_shape(line, name_suffix)
    draw_page.add(line)  # erst adden

    line.Position = Point(x1, y1)
    line.Size = Size(x2 - x1, y2 - y1)
    line.LineColor = 0x000000


# ============================================================
# 6) KREIS + QUADRANTEN (COMPOSITE)
# ============================================================

def draw_circle_with_quadrants(doc, draw_page, x, y, size, fill_color, label_texts, idx_tag):
    circle = doc.createInstance("com.sun.star.drawing.EllipseShape")
    _mark_shape(circle, f"circle_{idx_tag}")
    draw_page.add(circle)  # erst adden

    circle.Position  = Point(x, y)
    circle.Size      = Size(size, size)
    circle.FillColor = fill_color
    circle.LineColor = 0x000000

    cx = x + size // 2
    cy = y + size // 2

    draw_line(doc, draw_page, cx, y,  cx, y + size, f"vline_{idx_tag}")
    draw_line(doc, draw_page, x,  cy, x + size, cy, f"hline_{idx_tag}")

    # Quadranten-Zentren
    q1x = x + size // 4
    q3x = x + (3 * size) // 4
    q1y = y + size // 4
    q3y = y + (3 * size) // 4

    centers = [
        (q1x, q1y),  # UL
        (q3x, q1y),  # UR
        (q1x, q3y),  # LL
        (q3x, q3y),  # LR
    ]

    text_shapes = []
    for i, ((tcx, tcy), txt) in enumerate(zip(centers, label_texts), start=1):
        sh = draw_text_center(doc, draw_page, tcx, tcy, txt, f"text_{idx_tag}_{i}")
        text_shapes.append(sh)


# ============================================================
# 7) ROTATION (OPTIONAL)
# ============================================================

MAX_SOLVED_CIRCLES = 2  # 0..2 sind ok; >2 vermeiden

def _rotate_right_vals(vals):
    # vals = [UL, UR, LL, LR]
    return [vals[2], vals[0], vals[3], vals[1]]

def _rotate_vals(vals, steps_clockwise):
    v = list(vals)
    for _ in range(steps_clockwise % 4):
        v = _rotate_right_vals(v)
    return v

def _name_map_from_drawpage(dp):
    m = {}
    for i in range(dp.getCount()):
        sh = dp.getByIndex(i)
        nm = getattr(sh, "Name", "") or ""
        if nm.startswith(MARK_DESC + "_"):
            m[nm] = sh
    return m

def scramble_all_circles_no_solution(*args):
    """
    Mischt alle Kreise durch Rotation der Buchstaben.
    Ziel: NICHT die komplette Ausgangsstellung (Lösung) erzeugen.
    Außerdem: versucht, dass max. MAX_SOLVED_CIRCLES Kreise in Ausgangsstellung bleiben.
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    missing = []
    forced_solved = 0
    total = 0
    plan = []  # (idx_tag, base_quad, chosen_step)

    for gi, (title, ef_top, top_row, bot_row, ef_bot) in enumerate(GROUPS):
        top_pairs, _ = _get_pairs_for_half(sheet, ef_top, top_row)
        bot_pairs, _ = _get_pairs_for_half(sheet, ef_bot, bot_row)

        for c in range(NUM_COLS):
            idx_tag = f"b{gi}_c{c}"
            circle_name = f"{MARK_DESC}_circle_{idx_tag}"
            if circle_name not in name_map:
                missing.append(f"{title}: Kreis fehlt ({circle_name})")
                continue

            total += 1

            t = top_pairs[c] if c < len(top_pairs) else ""
            b = bot_pairs[c] if c < len(bot_pairs) else ""

            # Ausgangsstellung (Soll): Leer bleibt leer
            ul = t[0] if len(t) > 0 else " "
            ur = t[1] if len(t) > 1 else " "
            ll = b[0] if len(b) > 0 else " "
            lr = b[1] if len(b) > 1 else " "
            base = [ul, ur, ll, lr]

            # Erlaubte Rotationen: nur solche, die wirklich ändern (wenn möglich)
            candidates = [k for k in (1, 2, 3) if _rotate_vals(base, k) != base]

            if candidates:
                step = random.choice(candidates)  # garantiert "nicht Ausgangsstellung" für diesen Kreis
            else:
                # rotationssymmetrisch (z.B. alles leer oder gleiche Buchstaben) -> lässt sich nicht "entschärfen"
                step = 0
                forced_solved += 1

            plan.append((idx_tag, base, step))

    if missing:
        _msgbox(doc, "Scramble", "Es fehlen Shapes (bitte neu zeichnen):\n" + "\n".join(missing))
        return

    # Anwenden
    solved_after = 0
    for idx_tag, base, step in plan:
        vals = _rotate_vals(base, step)
        if vals == base:
            solved_after += 1

        for q, val in enumerate(vals, start=1):
            nm = f"{MARK_DESC}_text_{idx_tag}_{q}"
            sh = name_map.get(nm)
            if sh is not None:
                sh.String = val

    # Gesamtzustand darf nie komplett „Lösung“ sein
    # (wenn total==0: nichts zu tun)
    if total > 0 and solved_after == total:
        _msgbox(doc, "Scramble", "Alle Kreise sind rotationssymmetrisch/leer – kann nicht mischen, ohne Lösung zu treffen.")
        return

    # Soft-Limit: max. 2 Kreise dürfen in Ausgangsstellung sein
    allowed = min(MAX_SOLVED_CIRCLES, max(total - 1, 0))
    if solved_after > allowed:
        _msgbox(
            doc,
            "Scramble Hinweis",
            f"{solved_after} Kreise stehen in Ausgangsstellung (erlaubt: {allowed}).\n"
            f"Grund: mindestens {forced_solved} Kreise sind rotationssymmetrisch/leer und lassen sich nicht „verdrehen“."
        )

def refresh_and_scramble(*args):
    """
    1) Falls Kreise existieren: nur Buchstaben aktualisieren
       sonst: Kreise neu erzeugen
    2) Danach mischen (Rotation), aber nur wenn keine Fehler gemeldet wurden
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    # Decide if shapes exist (no lock needed for check, but we keep it simple)
    has_any = False
    try:
        for i in range(dp.getCount()):
            sh = dp.getByIndex(i)
            if getattr(sh, "Description", "") == MARK_DESC:
                has_any = True
                break
    except Exception:
        pass

    # Run update/create (they handle locking + messages)
    ok = update_texts_only() if has_any else create_circle_grid()
    if not ok:
        return

    # Scramble only if update/create succeeded
    scramble_all_circles_no_solution()

# ============================================================
# 8) REFLOW-FUNKTION (verschiebt vorhandene Kreise + Linien + Texte)
# ============================================================

def reflow_circles_only(*args):
    """
    Verschiebt NUR bestehende Shapes (Kreis + Kreuzlinien + 4 Texte) auf neue Soll-Positionen.
    Es wird NICHT gelöscht und NICHT neu gezeichnet.
    Voraussetzung: Shapes wurden mit _mark_shape() benannt/markiert.
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    # 1) Alle markierten Shapes in ein Dictionary nach Name sammeln
    name_map = {}
    for i in range(dp.getCount()):
        sh = dp.getByIndex(i)

        # Buttons / Controls niemals anfassen
        try:
            if sh.supportsService("com.sun.star.drawing.ControlShape"):
                continue
        except Exception:
            pass

        nm = getattr(sh, "Name", "") or ""
        if nm.startswith(MARK_DESC + "_"):
            name_map[nm] = sh

    missing = []

    # 2) Für jede Reihe/Spalte Zielposition berechnen und alle zugehörigen Shapes verschieben
    for block_index, (title, _ef_top, _top_row, _bot_row, _ef_bot) in enumerate(GROUPS):

        # Ziel-Y aus dem Zeilenraster (Oberkante der EF-Wortzeile der Gruppe)
        target_y = _circle_top_y_for_group(sheet, block_index)

        for c in range(NUM_COLS):
            target_x = START_X + c * OFFSET_X_BASE

            idx_tag = f"b{block_index}_c{c}"
            circle_name = f"{MARK_DESC}_circle_{idx_tag}"
            circle = name_map.get(circle_name)

            if circle is None:
                missing.append(f"{title}: fehlt {circle_name}")
                continue

            # Delta berechnen (wie weit muss alles verschoben werden?)
            dx = target_x - circle.Position.X
            dy = target_y - circle.Position.Y

            def move(nm):
                sh = name_map.get(nm)
                if sh is None:
                    missing.append(f"{title}: fehlt {nm}")
                    return
                sh.Position = Point(sh.Position.X + dx, sh.Position.Y + dy)

            # Kreis + Linien + 4 Texte gemeinsam verschieben
            move(circle_name)
            move(f"{MARK_DESC}_vline_{idx_tag}")
            move(f"{MARK_DESC}_hline_{idx_tag}")
            for q in range(1, 5):
                move(f"{MARK_DESC}_text_{idx_tag}_{q}")

    if missing:
        _msgbox(doc, "Reflow: fehlende Shapes", "\n".join(missing))


# ============================================================
# 9) ZEILEN AN KREIS ANPASSEN“
# ============================================================

# ============================================================
# - obere 2 Zeilen zusammen = Radius
# - untere 2 Zeilen zusammen = Radius (spiegelverkehrt)
# - beide Zeilen werden proportional skaliert (kein manuelles Suchen)
# ============================================================

# Paare je Gruppe: (Wortzeile, Buchstaben-zeile) oben und (Buchstaben-zeile, Wortzeile) unten
ROW_PAIRS = [
    ((3, 4), (5, 6)),
    ((8, 9), (10, 11)),
    ((13, 14), (15, 16)),
]

def _set_row_height(sheet, row_1based, height_100mm):
    r = sheet.Rows.getByIndex(row_1based - 1)
    try:
        r.OptimalHeight = False
    except Exception:
        pass
    r.Height = int(height_100mm)

def adjust_rows_to_circle_radius(sheet):
    """
    Skaliert die Zeilenhöhen so, dass:
    - Wortzeile + Buchstabenzeile = Kreisradius
    - unten ebenso (spiegelverkehrt)
    Verhältnis bleibt wie INIT_ROW_H_12 : INIT_ROW_H_09 (z.B. 1.2 : 0.9)
    """
    radius = int(CIRCLE_DIAMETER // 2)

    base_word = int(INIT_ROW_H_12)   # z.B. 1,2 cm in 1/100mm
    base_letters = int(INIT_ROW_H_09)  # z.B. 0,9 cm in 1/100mm
    base_sum = base_word + base_letters
    if base_sum <= 0:
        return

    # proportional skalieren
    scale = radius / float(base_sum)
    new_word = int(round(base_word * scale))
    new_letters = radius - new_word  # sorgt dafür, dass Summe exakt passt

    # Anwenden auf alle Paare (oben und unten)
    for (top_word, top_letters), (bot_letters, bot_word) in ROW_PAIRS:
        _set_row_height(sheet, top_word, new_word)
        _set_row_height(sheet, top_letters, new_letters)
        _set_row_height(sheet, bot_letters, new_letters)
        _set_row_height(sheet, bot_word, new_word)


# ============================================================
# 10) KREISE AM ZEILENRASTER AUSRICHTEN
# ============================================================

# ============================================================
# - Kreis-Top-Y = Oberkante der Wort-Zeile (EF-Zeile) der Gruppe
# - dadurch passt Kreis genau über 4 Zeilen (Wort + Buchstaben + Buchstaben + Wort)
# ============================================================

def _row_top_y(sheet, row_1based):
    """Y-Koordinate (1/100 mm) der Oberkante einer Zeile."""
    y = 0
    rows = sheet.Rows
    for i in range(row_1based - 1):
        y += rows.getByIndex(i).Height
    return y

def _circle_top_y_for_group(sheet, group_index):
    """
    GROUPS: (Titel, ef_top, top_row, bot_row, ef_bot)
    ef_top ist die Wort-Zeile (E+F verbunden) -> dort soll der Kreis oben beginnen.
    """
    ef_top_row = GROUPS[group_index][1]  # z.B. 3, 8, 13
    return _row_top_y(sheet, ef_top_row)


# ============================================================
# 11) LÖSCHEN (ALLES MIT MARK_DESC)
# ============================================================

def delete_all_circles(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    def is_ours(sh):
        # 1) NIE Buttons/Form-Controls löschen
        try:
            if sh.supportsService("com.sun.star.drawing.ControlShape"):
                return False
        except Exception:
            pass

        # 2) Optional extra Schutz: alles was an J1 verankert ist, behalten
        # J = Spalte 9, 1 = Zeile 0 (0-basiert)
        try:
            a = sh.Anchor.getRangeAddress()
            if a.StartColumn == 9 and a.StartRow == 0:
                return False
        except Exception:
            pass

        # 3) Unsere alten/neuen Marker/Namen löschen
        nm = getattr(sh, "Name", "") or ""
        desc = getattr(sh, "Description", "") or ""

        return (
            desc == MARK_DESC
            or nm == "KREIS_GUI"
            or nm.startswith(MARK_DESC + "_")
            or nm.startswith("TKQ_")
        )

    removed = 0
    for i in range(dp.getCount() - 1, -1, -1):
        sh = dp.getByIndex(i)
        try:
            if is_ours(sh):
                dp.remove(sh)
                removed += 1
        except Exception:
            pass

    try:
        sheet.getCellByPosition(0, 0).setString(f"Gelöscht: {removed}")
    except Exception:
        pass


# ============================================================
# 12 INIT (läuft automatisch 1x pro Dokument beim Start des Hauptmakros)
# ============================================================

SHEET_INIT_FLAG = "KREIS_GUI_INIT_DONE"

def _doc_user_props(doc):
    return doc.getDocumentProperties().getUserDefinedProperties()

def _get_init_done(doc) -> bool:
    props = _doc_user_props(doc)
    try:
        if props.hasByName(SHEET_INIT_FLAG):
            return str(props.getPropertyValue(SHEET_INIT_FLAG)) == "1"
    except Exception:
        pass
    return False

def _set_init_done(doc):
    props = _doc_user_props(doc)
    try:
        if not props.hasByName(SHEET_INIT_FLAG):
            props.addProperty(SHEET_INIT_FLAG, 0, "0")
        props.setPropertyValue(SHEET_INIT_FLAG, "1")
    except Exception:
        pass

def _set_center(rng):
    rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
    rng.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")

def _set_bold_14(rng):
    rng.CharHeight = 14
    rng.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

def _init_sheet_layout(doc):
    sheet = doc.Sheets.getByName(SHEET_NAME)
    
    # Disable spellcheck for EF input rows (merged E+F, but range is still E:F)
    disable_spellcheck_for_range(sheet, "E3:F3")
    disable_spellcheck_for_range(sheet, "E6:F6")
    disable_spellcheck_for_range(sheet, "E8:F8")
    disable_spellcheck_for_range(sheet, "E11:F11")
    disable_spellcheck_for_range(sheet, "E13:F13")
    disable_spellcheck_for_range(sheet, "E16:F16")

    # Disable spellcheck for the 2-letter cells C..F where people type pairs
    disable_spellcheck_for_range(sheet, "C4:F5")
    disable_spellcheck_for_range(sheet, "C9:F10")
    disable_spellcheck_for_range(sheet, "C14:F15")


    # Überschriften
    sheet.getCellRangeByName("B2").setString("Bezeichnung")
    sheet.getCellRangeByName("C2").setString("Kreis \n-1-")
    sheet.getCellRangeByName("D2").setString("Kreis \n-2-")
    sheet.getCellRangeByName("E2").setString("Kreis \n-3-")
    sheet.getCellRangeByName("F2").setString("Kreis \n-4-")

    hdr = sheet.getCellRangeByName(INIT_HEADER_RANGE)   # "B2:F2"
    _set_center(hdr)
    _set_bold_14(hdr)
    
    # Spalte B Hinweise (Zeilen 3,6,8,11,13,16)
    hint = "Begriffe\nmit 8 Buchstaben"
    for r in (3, 6, 8, 11, 13, 16,):
        cell = sheet.getCellByPosition(1, r - 1)  # B = 1
        cell.setString(hint)
        try:
            cell.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
            cell.CharHeight = 14  # optional: gleich wie Überschrift
        except Exception:
            pass

    # B2:F50 mittig (hori + vert)
    _set_center(sheet.getCellRangeByName(INIT_CENTER_RANGE))  # "B2:F50"

    # Wort-Zeilen: E+F verbinden + zentrieren + fett
    for r in INIT_WORD_ROWS:
        r0 = r - 1
        ef = sheet.getCellRangeByPosition(4, r0, 5, r0)  # E..F
        try:
            ef.merge(True)
        except Exception:
            pass
        _set_center(ef)
        _set_bold_14(ef)

    # Spaltenbreiten
    cols = sheet.Columns

    # C..G = 1,7 cm
    for col in range(2, 7):  # C(2)..G(6)
        cols.getByIndex(col).Width = INIT_COL_W_C_TO_G

    # J und L = 4 cm
    cols.getByIndex(9).Width  = INIT_COL_W_J_L  # J
    cols.getByIndex(11).Width = INIT_COL_W_J_L  # L

    # B: optimal, aber max 4,5 cm
    col_b = cols.getByIndex(1)
    try:
        col_b.OptimalWidth = True
    except Exception:
        pass
    try:
        if col_b.Width > INIT_COL_W_B_MAX:
            col_b.Width = INIT_COL_W_B_MAX
    except Exception:
        pass

    # Zeilenhöhen
    rows = sheet.Rows
    for r in INIT_ROWS_H17:
        rows.getByIndex(r - 1).Height = INIT_ROW_H_17

    adjust_rows_to_circle_radius(sheet)

    # Zeile 1 einfrieren
    try:
        doc.CurrentController.freezeAtPosition(0, 1)
    except Exception:
        pass

def _ensure_initialized(doc):
    if _get_init_done(doc):
        return
    _init_sheet_layout(doc)
    _set_init_done(doc)


def update_texts_only(*args):
    """
    Aktualisiert nur die Texte in bestehenden Kreisen (keine neuen Shapes).
    Nutzt dieselbe EF/Fallback-Logik wie create_circle_grid().
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)

    all_msgs = []
    missing = []
    ok = False

    doc.lockControllers()
    try:
        _ensure_initialized(doc)
        adjust_rows_to_circle_radius(sheet)
        normalize_input_cells(sheet)

        dp = _get_draw_page(sheet)
        name_map = _name_map_from_drawpage(dp)

        for group_index, (title, ef_top, top_row, bot_row, ef_bot) in enumerate(GROUPS):
            top_pairs, m1 = _get_pairs_for_half(sheet, ef_top, top_row)
            bot_pairs, m2 = _get_pairs_for_half(sheet, ef_bot, bot_row)
            all_msgs.extend([f"{title}: {x}" for x in (m1 + m2)])

            for c in range(NUM_COLS):
                idx_tag = f"b{group_index}_c{c}"
                circle_name = f"{MARK_DESC}_circle_{idx_tag}"
                circle = name_map.get(circle_name)
                if circle is None:
                    missing.append(f"{title}: Kreis fehlt ({circle_name}) – bitte neu zeichnen")
                    continue

                t = top_pairs[c] if c < len(top_pairs) else ""
                b = bot_pairs[c] if c < len(bot_pairs) else ""

                ul = t[0] if len(t) > 0 else " "
                ur = t[1] if len(t) > 1 else " "
                ll = b[0] if len(b) > 0 else " "
                lr = b[1] if len(b) > 1 else " "
                quad = [ul, ur, ll, lr]

                x0 = circle.Position.X
                y0 = circle.Position.Y
                d = min(circle.Size.Width, circle.Size.Height)

                q1x = x0 + d // 4
                q3x = x0 + (3 * d) // 4
                q1y = y0 + d // 4
                q3y = y0 + (3 * d) // 4
                centers = [(q1x, q1y), (q3x, q1y), (q1x, q3y), (q3x, q3y)]

                for i, ((cx, cy), txt) in enumerate(zip(centers, quad), start=1):
                    nm = f"{MARK_DESC}_text_{idx_tag}_{i}"
                    sh = name_map.get(nm)
                    if sh is None:
                        missing.append(f"{title}: Text fehlt ({nm}) – bitte neu zeichnen")
                        continue

                    txt = _to_upper_visual(txt)
                    try:
                        sh.String = txt
                    except Exception:
                        try:
                            sh.Text.setString(txt)
                        except Exception:
                            pass

                    # Re-center (optional)
                    w = sh.Size.Width or 700
                    h = sh.Size.Height or 700
                    sh.Position = Point(int(cx - w/2), int(cy - h/2))

        # Debug A1
        try:
            sheet.getCellByPosition(0, 0).setString(f"Hinweise: {len(all_msgs) + len(missing)}")
        except Exception:
            pass

        ok = not missing  # missing shapes -> not ok
        return ok

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

        set_cursor_next_input(doc, sheet)

        # Messages after unlock (short + readable)
        if all_msgs or missing:
            parts = []
            if all_msgs:
                parts.append(_format_messages("Eingabe-Hinweise", all_msgs, max_lines=20))
            if missing:
                parts.append(_format_messages("Fehlende Shapes", missing, max_lines=20))
            _msgbox(doc, "Update Hinweise", "\n\n".join([p for p in parts if p]))

# ============================================================
# 13) HAUPT-MAKRO: 3 BLÖCKE ZEICHNEN + FEHLER MSGBOX + ENDE
# ============================================================

def create_circle_grid(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)

    # We collect messages and show them AFTER unlockControllers (more stable UI).
    all_msgs = []
    ok = False

    doc.lockControllers()
    try:
        if not self_check_kreis_quadranten():
            all_msgs.append("Self-Check fehlgeschlagen.")
            return False

        delete_all_circles()  # delete first (inside lock)

        _ensure_initialized(doc)
        adjust_rows_to_circle_radius(sheet)
        normalize_input_cells(sheet)

        draw_page = _get_draw_page(sheet)

        for group_index, (title, ef_top, top_row, bot_row, ef_bot) in enumerate(GROUPS):
            top_pairs, m1 = _get_pairs_for_half(sheet, ef_top, top_row)
            bot_pairs, m2 = _get_pairs_for_half(sheet, ef_bot, bot_row)

            # softer messages (not too many)
            all_msgs.extend([f"{title}: {x}" for x in (m1 + m2)])

            y = _circle_top_y_for_group(sheet, group_index)
            fill = ROW_COLORS[group_index] if group_index < len(ROW_COLORS) else ROW_COLORS[-1]

            for c in range(NUM_COLS):
                x = START_X + c * OFFSET_X_BASE
                idx_tag = f"b{group_index}_c{c}"

                t = top_pairs[c] if c < len(top_pairs) else ""
                b = bot_pairs[c] if c < len(bot_pairs) else ""

                ul = t[0] if len(t) > 0 else " "
                ur = t[1] if len(t) > 1 else " "
                ll = b[0] if len(b) > 0 else " "
                lr = b[1] if len(b) > 1 else " "
                quad = [ul, ur, ll, lr]

                draw_circle_with_quadrants(doc, draw_page, x, y, CIRCLE_DIAMETER, fill, quad, idx_tag)

        # Debug A1
        try:
            sheet.getCellByPosition(0, 0).setString(f"Hinweise: {len(all_msgs)}")
        except Exception:
            pass

        ok = True
        return True

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

        # Cursor always set after macro
        set_cursor_next_input(doc, sheet)

        # Show messagebox only after unlock
        if all_msgs:
            msg = _format_messages("Hinweise", all_msgs, max_lines=25)
            _msgbox(doc, "Hinweise / Eingaben", msg)


# ============================================================
# 14 SELF-CHECK (Debug): prüft fehlende Makros/Konstanten/Globals
# ============================================================

import builtins
import dis

def _safe_read_source():
    """Versucht den eigenen Quelltext zu lesen (für Duplikat-Checks)."""
    try:
        path = __file__
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""

def _count_def_in_source(source_text, func_name):
    if not source_text:
        return 0
    needle = f"def {func_name}"
    return source_text.count(needle)

def _undefined_globals_in_func(fn, allowed=None):
    """
    Findet echte LOAD_GLOBAL/LOAD_NAME Referenzen, die im Modul nicht existieren.
    (Attribute wie .DrawPage zählen NICHT als fehlende Globals.)
    """
    if allowed is None:
        allowed = set()

    missing = set()
    g = globals()

    for ins in dis.get_instructions(fn):
        if ins.opname in ("LOAD_GLOBAL", "LOAD_NAME"):
            name = ins.argval
            if name in allowed:
                continue
            if name in g:
                continue
            if hasattr(builtins, name):
                continue
            missing.add(name)

    return sorted(missing)

def self_check_kreis_quadranten(*args):
    """
    Prüft:
    - ob wichtige Konstanten existieren
    - ob exportierte Makros existieren und callable sind
    - ob wichtige Funktionen fehlende Global-Namen referenzieren
    - optional: ob bestimmte defs doppelt im Quelltext vorkommen
    """
    issues = []
    notes = []

    # 1) Pflicht-Konstanten (für dein Programm)
    required_consts = [
        "SHEET_NAME", "MARK_DESC", "NUM_COLS",
        "CIRCLE_DIAMETER", "OFFSET_X_BASE", "START_X",
        "GROUPS"
    ]
    for c in required_consts:
        if c not in globals():
            issues.append(f"Konstante fehlt: {c}")

    # 2) Muss-Functions (Kern)
    required_funcs = [
        "create_circle_grid",
        "delete_all_circles",
        "reflow_circles_only",
        "scramble_all_circles_no_solution",
        "refresh_and_scramble",
    ]
    for fn in required_funcs:
        obj = globals().get(fn)
        if obj is None:
            issues.append(f"Funktion fehlt: {fn}")
        elif not callable(obj):
            issues.append(f"Name ist nicht callable: {fn}")

    # 3) Exporte prüfen
    if "g_exportedScripts" not in globals():
        issues.append("g_exportedScripts fehlt (Makros erscheinen dann nicht im Dialog).")
    else:
        exp = globals().get("g_exportedScripts")
        if not isinstance(exp, (tuple, list)):
            issues.append("g_exportedScripts ist nicht tuple/list.")
        else:
            for item in exp:
                if not callable(item):
                    issues.append(f"In g_exportedScripts ist etwas nicht callable: {repr(item)}")

    # 4) Referenz-Check: fehlen Globals, die Funktionen benutzen?
    #    (Findet z.B. update_texts_only, wenn es nur aufgerufen aber nicht definiert ist.)
    allowed = {
        "XSCRIPTCONTEXT", "uno", "random", "Point", "Size",
        # UNO/basics, falls du sie in Code referenzierst:
        "cm", "U100MM_PER_CM",
    }

    funcs_to_scan = [
        "create_circle_grid",
        "refresh_and_scramble",
        "reflow_circles_only",
        "scramble_all_circles_no_solution",
    ]
    for fn_name in funcs_to_scan:
        fn = globals().get(fn_name)
        if callable(fn):
            missing = _undefined_globals_in_func(fn, allowed=allowed)
            # Häufige False-Positives minimieren: manche Namen sind Attribute, die hier nicht auftauchen.
            # LOAD_GLOBAL trifft aber echte Probleme ziemlich gut.
            if missing:
                issues.append(f"{fn_name}: fehlende Global-Namen: {', '.join(missing)}")

    # 5) Optional: Duplikat-Check über Quelltext (hilft gegen versehentlich doppelte defs)
    src = _safe_read_source()
    if src:
        for dup_name in ["_get_ef_word"]:
            n = _count_def_in_source(src, dup_name)
            if n > 1:
                notes.append(f"Hinweis: '{dup_name}' ist {n}x im Quelltext definiert (die letzte Definition gewinnt).")

    # Ausgabe
    doc = None
    try:
        doc = _get_doc()
    except Exception:
        pass

    if not issues and not notes:
        if doc:
            _msgbox(doc, "Self-Check", "OK ✅\nKeine Probleme gefunden.")
        return True

    msg_lines = []
    if issues:
        msg_lines.append("Probleme:")
        msg_lines.extend(f"- {x}" for x in issues)
    if notes:
        msg_lines.append("")
        msg_lines.append("Hinweise:")
        msg_lines.extend(f"- {x}" for x in notes)

    if doc:
        _msgbox(doc, "Self-Check", "\n".join(msg_lines))
    return False


# ============================================================
# 15) EXPORTIERTE MAKROS
# ============================================================

g_exportedScripts = (
    self_check_kreis_quadranten,
    create_circle_grid,
    delete_all_circles,
    update_texts_only,
    reflow_circles_only,
    scramble_all_circles_no_solution,
    refresh_and_scramble,
)
