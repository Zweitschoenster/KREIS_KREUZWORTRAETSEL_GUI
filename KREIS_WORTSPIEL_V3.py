# -*- coding: utf-8 -*-
import uno
import random
from datetime import datetime, timedelta
from com.sun.star.awt import Point, Size

# =========================
# KONFIG
# =========================
SHEET_NAME = "KREIS_WORTSPIEL"
MARK_DESC  = "KREIS_Wortspiel_GUI_V3"

# Kreise: 3 Gruppen, je Gruppe 4 Kreise (C..F)
GROUPS = [
    ("Gruppe 1",  3,  4,  5,  6),
    ("Gruppe 2",  8,  9, 10, 11),
    ("Gruppe 3", 13, 14, 15, 16),
]
NUM_COLS = 4

# Kreise Layout
CIRCLE_DIAMETER = 3900
START_X         = 14000
OFFSET_X_BASE   = 4000
ROW_COLORS = [0xCCE5FF, 0xFBE5B6, 0xCCFFCC]
CHAR_HEIGHT = 16

# Wortspalten (gleicher Sheet) + used-Spalten sind +3/+4
WORDLIST_COLS = ["AG", "AN", "AU", "BB"]
RANDOM_DAYS_LOCK = 30
WORDLIST_ROW_START = 4
WORDLIST_ROW_END   = 600

# Debug: True zeigt eine Info, wenn Random-Kandidaten = 0
DEBUG = False

# =========================
# UNO HELPERS
# =========================
def _get_doc():
    return XSCRIPTCONTEXT.getDocument()

def _msgbox(doc, title, message):
    try:
        parent = doc.CurrentController.Frame.ContainerWindow
        toolkit = parent.getToolkit()
        buttons = uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK")
        box_type = uno.Enum("com.sun.star.awt.MessageBoxType", "INFOBOX")
        box = toolkit.createMessageBox(parent, box_type, buttons, title, message)
        box.execute()
    except Exception:
        pass

def _get_sheet(doc):
    return doc.Sheets.getByName(SHEET_NAME)

def _get_draw_page(sheet):
    return sheet.DrawPage

def _col0_from_letters(col_letters: str) -> int:
    s = col_letters.strip().upper()
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def _cell(sheet, col0, row_1based):
    return sheet.getCellByPosition(col0, row_1based - 1)

def _a1(col0, row0):
    # row0 = 0-based
    col = ""
    n = col0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        col = chr(65 + r) + col
    return f"{col}{row0+1}"

def _row_top_y(sheet, row_1based):
    y = 0
    rows = sheet.Rows
    for i in range(row_1based - 1):
        y += rows.getByIndex(i).Height
    return y

# =========================
# TEXT NORMALISIERUNG (Umlaute bleiben in Sheet, Kreise auch)
# =========================
def _to_upper_visual(s):
    return (s or "").replace("ß", "ẞ").upper()

def _normalize_keep_umlauts_no_spaces(s: str) -> str:
    x = (s or "")
    x = "".join(ch for ch in x if not ch.isspace())
    x = _to_upper_visual(x)
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜẞ")
    return "".join(ch for ch in x if ch in allowed)
    
def _upper_keep_umlauts_and_sz(s: str) -> str:
    """
    Uppercase, aber ß bleibt ß (keine ß->ẞ und keine ß->SS).
    Umlaute werden groß: ä->Ä usw.
    """
    out = []
    for ch in (s or ""):
        if ch == "ß":
            out.append("ß")
        else:
            out.append(ch.upper())
    return "".join(out)

def _normalize_ef_visual_no_spaces(s: str) -> str:
    """
    Für EF (Anzeige): Leerzeichen entfernen, Großschrift,
    Umlaute und ß bleiben als Zeichen erhalten.
    """
    x = (s or "")
    x = "".join(ch for ch in x if not ch.isspace())
    x = _upper_keep_umlauts_and_sz(x)
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜß")
    return "".join(ch for ch in x if ch in allowed)

def _normalize_for_circles(s: str) -> str:
    """
    Für Kreise (Kreuzworträtsel-Logik):
    ÄÖÜ -> AE/OE/UE, ß/ẞ -> SS, nur A-Z.
    """
    x = (s or "")
    x = "".join(ch for ch in x if not ch.isspace())
    x = x.replace("ß", "ẞ").upper()     # hier darf ß in SS laufen
    x = (x.replace("Ä", "AE")
           .replace("Ö", "OE")
           .replace("Ü", "UE")
           .replace("ẞ", "SS"))
    x = "".join(ch for ch in x if "A" <= ch <= "Z")
    return x


# =========================
# J2 lesen (1–6)
# =========================
def _read_j2_target(sheet, default=6):
    try:
        c = sheet.getCellRangeByName("J2")
        v = int(c.getValue())
        if v == 0:
            s = (c.getString() or "").strip()
            v = int(s) if s else default
    except Exception:
        v = default
    if v < 1: v = 1
    if v > 6: v = 6
    return v

# =========================
# Dispatch-Helper + “commit” + Cursor parken
# =========================
from com.sun.star.beans import PropertyValue

def _dispatch(doc, cmd, props=()):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    helper = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    frame = doc.CurrentController.Frame
    helper.executeDispatch(frame, cmd, "", 0, tuple(props))

def _commit_current_cell_input(doc):
    # bestätigt ggf. aktive Eingabe (wie Enter)
    try:
        _dispatch(doc, ".uno:Enter")
    except Exception:
        pass

def _park_cursor_in_empty_cell(doc, sheet):
    """
    Geht ans Ende des verwendeten Bereichs und springt zwei Spalten rechts
    in Zeile 1 (0-based row=0). Das ist i.d.R. leer und stört nicht.
    """
    try:
        cur = sheet.createCursor()
        cur.gotoEndOfUsedArea(True)
        end_col = cur.RangeAddress.EndColumn
        # zwei Spalten rechts außerhalb "Used Area"
        park_col = end_col + 2
        park_row = 0  # Zeile 1
        cell = sheet.getCellByPosition(park_col, park_row)
        doc.CurrentController.select(cell)
    except Exception:
        # Fallback: einfach A1 (falls oben schiefgeht)
        try:
            cell = sheet.getCellRangeByName("A1")
            doc.CurrentController.select(cell)
        except Exception:
            pass



# =========================
# Regeln: EF und GRID
# =========================
def _pairs_from_ef_word(sheet, ef_row_1based, label):
    """
    EF: sobald NICHT leer -> wird verwendet.
    - EF-Zelle wird nur optisch bereinigt (Umlaute/ß bleiben)
    - für Kreise wird separat Kreuzworträtsel-normalisiert
    """
    msgs = []
    ef_cell = _cell(sheet, 4, ef_row_1based)  # E (merged)
    raw = (ef_cell.getString() or "").strip()
    if not raw:
        return None, msgs, True

    # 1) Visual für EF (keine Umlaut/ß-Umwandlung)
    vis = _normalize_ef_visual_no_spaces(raw)

    # >8 Regel bezieht sich auf EF-Visual (wie bisher)
    if len(vis) > 8:
        msgs.append(f"{label}: >8 Zeichen → auf 8 gekürzt")
        vis = vis[:8]

    # zurückschreiben, damit der Nutzer sieht, was verwendet wird
    if vis != raw:
        ef_cell.setString(vis)
        
    # 2) Für Kreise (WICHTIG: norm IMMER setzen)
    norm = _normalize_for_circles(vis)
    if norm is None:
        norm = ""
    # optional (wie besprochen): auf 8 begrenzen
    # norm = norm[:8]

    # LOG: Wort + Timestamp in Y/Z derselben Zeile
    _log_used_word(sheet, ef_row_1based, vis)


    pairs = []
    for i in range(0, 8, 2):
        chunk = norm[i:i+2]
        if len(chunk) == 2:
            pairs.append(chunk)
        elif len(chunk) == 1:
            pairs.append(chunk + " ")
        else:
            pairs.append("  ")

    return pairs, msgs, True


def _pairs_from_grid_row(sheet, row_1based):
    """
    Grid C..F:
    - 0 Zeichen ok
    - 1 Zeichen ok -> Meldung
    - 2 Zeichen ok
    - >2 -> kürzen auf 2 -> Meldung (mit Anzeige der 2 Zeichen)
    """
    msgs = []
    pairs = []
    parts_vis = []   # <- für Y-Spalte (Anzeige)

    row0 = row_1based - 1

    for col0 in range(2, 6):  # C..F
        c = sheet.getCellByPosition(col0, row0)
        raw = (c.getString() or "").strip()

        vis = _normalize_ef_visual_no_spaces(raw)   # Anzeige (Umlaute/ß bleiben)
        norm = _normalize_for_circles(vis)          # für Kreise

        # Für Log: wir nehmen die Anzeige-Teile so, wie sie verwendet/angezeigt werden
        parts_vis.append(vis)

        if len(vis) == 0:
            pairs.append("  ")
        elif len(norm) == 1:
            pairs.append(norm + " ")
            msgs.append(f"{_a1(col0,row0)}: 1 Buchstabe ('{vis}') → 2. fehlt")
            if vis != raw:
                c.setString(vis)
        elif len(norm) == 2:
            pairs.append(norm)
            if vis != raw:
                c.setString(vis)
        else:
            cut = norm[:2]
            pairs.append(cut)
            msgs.append(f"{_a1(col0,row0)}: zu viele Zeichen ('{vis}'→'{norm}') → gekürzt auf '{cut}'")
            c.setString(vis)

    # LOG-Wort bauen (max 8 Zeichen wie beim Kreis-Konzept)
    word_vis = "".join(parts_vis)
    if len(word_vis) > 8:
        word_vis = word_vis[:8]

    return pairs, msgs, word_vis

def _grid_all_empty(pairs):
    for p in pairs:
        if (p or "").strip():
            return False
    return True

def _half_has_manual_input(sheet, ef_row_1based, grid_row_1based):
    # EF nicht leer?
    ef = (_cell(sheet, 4, ef_row_1based).getString() or "").strip()
    if ef:
        return True
    # Grid irgendwas?
    row0 = grid_row_1based - 1
    for col0 in range(2, 6):
        raw = (sheet.getCellByPosition(col0, row0).getString() or "").strip()
        if _normalize_keep_umlauts_no_spaces(raw):
            return True
    return False

# =========================
# RANDOM aus Wortspalten (nur 5..8), 30 Tage Sperre
# =========================
def _parse_ts(ts_str: str):
    ts_str = (ts_str or "").strip()
    if not ts_str:
        return None
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(ts_str, fmt)
        except Exception:
            pass
    return None

def _now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def _mark_word_used(sheet, col_letters: str, row_1based: int, used_word: str):
    col0 = _col0_from_letters(col_letters)
    _cell(sheet, col0 + 3, row_1based).setString(_to_upper_visual(used_word))
    _cell(sheet, col0 + 4, row_1based).setString(_now_ts())

def _pick_random_word(sheet):
    candidates = []
    cutoff = datetime.now() - timedelta(days=RANDOM_DAYS_LOCK)

    for col_letters in WORDLIST_COLS:
        col0 = _col0_from_letters(col_letters)
        for r in range(WORDLIST_ROW_START, WORDLIST_ROW_END):
            w = (_cell(sheet, col0, r).getString() or "").strip()
            if not w:
                continue
            norm = _normalize_keep_umlauts_no_spaces(w)
            if not (5 <= len(norm) <= 8):
                continue

            used_dt = _parse_ts(_cell(sheet, col0 + 4, r).getString())
            if used_dt is None or used_dt < cutoff:
                candidates.append((norm, col_letters, r))

    if DEBUG:
        doc = _get_doc()
        _msgbox(doc, "DEBUG", f"Random-Kandidaten: {len(candidates)}")

    if not candidates:
        return None, None, None

    return random.choice(candidates)
    
    
# =========================
# CURSOR: parken in leere Zelle
# =========================    
def _is_cell_empty_for_parking(cell) -> bool:
    """
    'Leer' bedeutet: String ist leer oder nur Leerzeichen.
    (Formeln o.ä. werden hier bewusst nicht speziell behandelt.)
    """
    try:
        s = cell.getString()
        if s is None:
            s = ""
        return (s.strip() == "")
    except Exception:
        return True


def _park_cursor_priority(doc, sheet):
    """
    Cursor in eine passende leere Zelle setzen:
    1) EF: E3,E6,E8,E11,E13,E16
    2) Grid: C..F in Zeilen 4,5,9,10,14,15
    3) sonst C17 (auch wenn nicht leer)
    """
    # 1) EF-Zellen
    ef_rows = [3, 6, 8, 11, 13, 16]     # 1-based
    ef_col0 = 4                         # E (0-based)

    for r1 in ef_rows:
        cell = _cell(sheet, ef_col0, r1)
        if _is_cell_empty_for_parking(cell):
            doc.CurrentController.select(cell)
            return

    # 2) Grid-Zellen C..F
    grid_rows = [4, 5, 9, 10, 14, 15]   # 1-based
    for r1 in grid_rows:
        row0 = r1 - 1
        for col0 in range(2, 6):        # C(2) .. F(5)
            cell = sheet.getCellByPosition(col0, row0)
            if _is_cell_empty_for_parking(cell):
                doc.CurrentController.select(cell)
                return

    # 3) Fallback C17
    fallback = sheet.getCellRangeByName("C17")
    doc.CurrentController.select(fallback)


# =========================
# ROTATION: Text in einem Kreis drehen (TextShapes 1..4)
# =========================
def _get_circle_text_shapes(draw_page, idx_tag: str):
    """
    Liefert die 4 TextShapes eines Kreises als dict {1:shape,2:shape,3:shape,4:shape}
    idx_tag z.B. "b0_c2"
    """
    out = {}
    prefix = f"{MARK_DESC}_text_{idx_tag}_"
    for i in range(draw_page.getCount()):
        sh = draw_page.getByIndex(i)
        try:
            nm = getattr(sh, "Name", "") or ""
        except Exception:
            continue
        if nm.startswith(prefix):
            # Name endet auf _1.._4
            try:
                n = int(nm.split("_")[-1])
                if 1 <= n <= 4:
                    out[n] = sh
            except Exception:
                pass
    return out


def _get_shape_text(sh) -> str:
    try:
        return sh.String or ""
    except Exception:
        try:
            return sh.Text.getString() or ""
        except Exception:
            return ""


def _set_shape_text(sh, s: str):
    try:
        sh.String = s
    except Exception:
        try:
            sh.Text.setString(s)
        except Exception:
            pass


def _rotate_quad_texts(texts, steps_cw: int):
    """
    texts = [UL, UR, LL, LR] (entspricht Textshape _1,_2,_3,_4)
    steps_cw: 0..3 (90° pro Schritt im Uhrzeigersinn)
    """
    steps_cw %= 4
    if steps_cw == 0:
        return texts[:]

    # Uhrzeigersinn 90°: UL->UR, UR->LR, LR->LL, LL->UL
    # also neues [UL,UR,LL,LR] aus altem:
    # newUL = oldLL, newUR = oldUL, newLL = oldLR, newLR = oldUR  (für 90° CW)
    def rot90(t):
        ul, ur, ll, lr = t
        return [ll, ul, lr, ur]

    out = texts[:]
    for _ in range(steps_cw):
        out = rot90(out)
    return out


def _rotate_one_circle_letters(draw_page, idx_tag: str, steps_cw: int):
    """
    Rotiert die Buchstaben innerhalb eines Kreises, indem die Texte der 4 TextShapes vertauscht werden.
    steps_cw: 0..3
    """
    d = _get_circle_text_shapes(draw_page, idx_tag)
    if len(d) != 4:
        return  # Kreis (noch) nicht vorhanden oder Namen passen nicht

    shapes = [d[1], d[2], d[3], d[4]]
    texts = [_get_shape_text(s) for s in shapes]  # [UL,UR,LL,LR]

    new_texts = _rotate_quad_texts(texts, steps_cw)

    for sh, tx in zip(shapes, new_texts):
        _set_shape_text(sh, tx)


# =========================
# SHAPES: löschen / zeichnen
# =========================
def delete_all_circles(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = sheet.DrawPage

    removed = 0
    skipped = 0
    for i in range(dp.getCount() - 1, -1, -1):
        sh = dp.getByIndex(i)

        # Controls nie anfassen
        try:
            if sh.supportsService("com.sun.star.drawing.ControlShape"):
                skipped += 1
                continue
        except Exception:
            pass

        # Nur eigene Shapes entfernen (Name oder Description)
        try:
            if getattr(sh, "Description", "") != MARK_DESC and not getattr(sh, "Name", "").startswith(MARK_DESC + "_"):
                continue
        except Exception:
            continue

        try:
            dp.remove(sh)
            removed += 1
        except Exception:
            pass

    if DEBUG:
        _msgbox(doc, "Löschen", f"Entfernt: {removed}\nControls: {skipped}")

def _mark_shape(shape, name_suffix):
    try:
        shape.Name = f"{MARK_DESC}_{name_suffix}"
    except Exception:
        pass
    try:
        shape.Description = MARK_DESC
    except Exception:
        pass

def _draw_text_center(doc, draw_page, cx, cy, text, name_suffix):
    t = doc.createInstance("com.sun.star.drawing.TextShape")
    _mark_shape(t, name_suffix)
    draw_page.add(t)

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
        t.TextHorizontalAdjust = 1
        t.TextVerticalAdjust   = 1
    except Exception:
        pass
    return t

def _draw_line(doc, draw_page, x1, y1, x2, y2, name_suffix):
    line = doc.createInstance("com.sun.star.drawing.LineShape")
    _mark_shape(line, name_suffix)
    draw_page.add(line)
    line.Position = Point(x1, y1)
    line.Size = Size(x2 - x1, y2 - y1)
    line.LineColor = 0x000000

def _draw_circle(doc, draw_page, x, y, size, fill_color, quad, idx_tag):
    circle = doc.createInstance("com.sun.star.drawing.EllipseShape")
    _mark_shape(circle, f"circle_{idx_tag}")
    draw_page.add(circle)

    circle.Position  = Point(x, y)
    circle.Size      = Size(size, size)
    circle.FillColor = fill_color
    circle.LineColor = 0x000000

    cx = x + size // 2
    cy = y + size // 2
    _draw_line(doc, draw_page, cx, y,  cx, y + size, f"v_{idx_tag}")
    _draw_line(doc, draw_page, x,  cy, x + size, cy, f"h_{idx_tag}")

    q1x = x + size // 4
    q3x = x + (3 * size) // 4
    q1y = y + size // 4
    q3y = y + (3 * size) // 4

    centers = [(q1x,q1y), (q3x,q1y), (q1x,q3y), (q3x,q3y)]
    for i, ((tcx,tcy), ch) in enumerate(zip(centers, quad), start=1):
        _draw_text_center(doc, draw_page, tcx, tcy, ch, f"text_{idx_tag}_{i}")

# =========================
# Prioritätsprüfung pro Halbkreis (EF -> GRID -> RANDOM)
# =========================
def _pairs_for_half(sheet, title, ef_row, grid_row, state):
    """
    state:
      remaining_random: wie viele Wörter noch automatisch gefüllt werden dürfen
      msgs: Sammelliste
    """
    # 1) EF (wenn nicht leer -> immer nutzen)
    pairs, m, ok = _pairs_from_ef_word(sheet, ef_row, f"EF{ef_row}")
    if m:
        state["msgs"].extend([f"{title}: {x}" for x in m])
    if pairs is not None:
        return pairs, ok, True  # used_word=True

    # 2) Grid
    gpairs, gm, gword = _pairs_from_grid_row(sheet, grid_row)
    if gm:
        state["msgs"].extend([f"{title}: {x}" for x in gm])

    if not _grid_all_empty(gpairs):
        # LOG: Wort + Timestamp in Y/Z derselben Zeile wie das Grid
        _log_used_word(sheet, grid_row, gword)
        return gpairs, True, True

    # 3) Random (NUR wenn erlaubt)
    if state["remaining_random"] <= 0:
        return ["  ","  ","  ","  "], True, False

    w, col_letters, row1 = _pick_random_word(sheet)
    if not w:
        if DEBUG:
            state["msgs"].append(f"{title}: keine Random-Kandidaten gefunden")
        return ["  ","  ","  ","  "], True, False
        
        
    if DEBUG:
        state["msgs"].append(f"{title}: Random gewählt '{w}' aus {col_letters}{row1} (remaining_random vorher={state['remaining_random']})")


    # <<< IMPORTANT: Random NICHT nach EF/Grid zurückschreiben!
    # Nur in Wortliste (+3/+4) markieren:
    _mark_word_used(sheet, col_letters, row1, w)
    state["remaining_random"] -= 1

    # Wort in 4 Paare (8 Plätze) -> Rest leer, letzter Einzelbuchstabe gepadded
    pairs = []
    for i in range(0, 8, 2):
        chunk = w[i:i+2]
        if len(chunk) == 2:
            pairs.append(chunk)
        elif len(chunk) == 1:
            pairs.append(chunk + " ")
        else:
            pairs.append("  ")
    return pairs, True, True
    
# =========================
# LOG: Y/Z (Wort + Timestamp)
# =========================
LOG_WORD_COL = "Y"
LOG_TS_COL   = "Z"

def _log_used_word(sheet, row_1based: int, word_visual: str):
    """
    Schreibt das verwendete Wort parallel in Spalte Y und Timestamp in Z
    in derselben Zeile (1-based).
    """
    try:
        w = (word_visual or "").strip()
        if not w:
            return
        y0 = _col0_from_letters(LOG_WORD_COL)
        z0 = _col0_from_letters(LOG_TS_COL)
        _cell(sheet, y0, row_1based).setString(w)
        _cell(sheet, z0, row_1based).setString(_now_ts())
    except Exception:
        pass

# =========================
# KREIS-INHALT LEEREN (nur TextShapes, Shapes bleiben)
# =========================
def clear_circle_text_only(sheet):
    """
    Löscht NUR die Buchstaben in den Kreisen (TextShapes),
    lässt Ellipsen und Linien stehen.
    """
    dp = sheet.DrawPage

    for i in range(dp.getCount()):
        sh = dp.getByIndex(i)

        # nur unsere Shapes
        try:
            if getattr(sh, "Description", "") != MARK_DESC and not getattr(sh, "Name", "").startswith(MARK_DESC + "_"):
                continue
        except Exception:
            continue

        # nur TextShapes (Quadranten-Buchstaben)
        try:
            if not sh.supportsService("com.sun.star.drawing.TextShape"):
                continue
        except Exception:
            # falls supportsService nicht geht: per Name filtern
            try:
                nm = getattr(sh, "Name", "") or ""
                if "text_" not in nm:
                    continue
            except Exception:
                continue

        # Text leeren
        try:
            sh.String = ""
        except Exception:
            try:
                sh.Text.setString("")
            except Exception:
                pass

# =========================
# INPUT-BEREICH LEEREN (ohne Format zu löschen)
# =========================
def _clear_cells_keep_format(sheet, a1_range: str):
    """
    Löscht Inhalte (Text/Wert/Datum/Formel), aber lässt Formatierungen stehen.
    """
    try:
        rng = sheet.getCellRangeByName(a1_range)
        flags = (1 | 2 | 4 | 16)  # VALUE|DATETIME|STRING|FORMULA
        rng.clearContents(flags)
    except Exception:
        pass


def clear_circle_contents_and_cells(*args):
    """
    Button-Funktion:
    - leert NUR den Inhalt der Kreise (Buchstaben), Kreise bleiben
    - leert C3:F17 (Format bleibt)
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)

    doc.lockControllers()
    try:
        clear_circle_text_only(sheet)
        _clear_cells_keep_format(sheet, "C3:F17")
    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

    return True


# =========================
# KREISE SHUFFLEN (nur Positionen tauschen)
# =========================
def _shapes_for_idx_tag(draw_page, idx_tag: str):
    """
    Findet alle Shapes, die zu einem Kreis gehören (circle, v, h, text_..)
    anhand des idx_tag (z.B. "b0_c2").
    """
    out = []
    needle = f"_{idx_tag}"
    for i in range(draw_page.getCount()):
        sh = draw_page.getByIndex(i)
        try:
            nm = getattr(sh, "Name", "") or ""
        except Exception:
            nm = ""
        if needle in nm:
            out.append(sh)
    return out


def _move_shapes_x(shapes, dx: int):
    for sh in shapes:
        try:
            p = sh.Position
            sh.Position = Point(int(p.X + dx), int(p.Y))
        except Exception:
            pass


def rotate_letters_in_circles_and_clear_cells(*args):
    """
    Button-Funktion:
    - dreht NUR die Buchstaben innerhalb jedes Kreises (keine Positionsänderung der Kreise)
    - 0° (Ausgangsposition) ist erlaubt, aber höchstens 1 Kreis insgesamt darf 0° haben
      (es muss keiner 0° haben)
    - alle Kreise werden behandelt (auch wenn leer/teil-leer)
    - danach C3:F17 leeren (Format bleibt)
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    doc.lockControllers()
    try:
        # Alle Kreis-Tags sammeln (z.B. b0_c0 ... b2_c3)
        tags = []
        for gi in range(len(GROUPS)):
            for c in range(NUM_COLS):
                tags.append(f"b{gi}_c{c}")

        # Standard: alle drehen (1/2/3 = 90/180/270 cw; 3 entspricht 90° ccw)
        steps_map = {tag: random.choice([1, 2, 3]) for tag in tags}

        # Optional: genau 1 Kreis darf 0° behalten
        if random.choice([True, False]):  # 50/50 ob überhaupt einer 0° hat
            keep_zero_tag = random.choice(tags)
            steps_map[keep_zero_tag] = 0

        # Anwenden
        for tag, steps in steps_map.items():
            _rotate_one_circle_letters(dp, tag, steps)

        # Eingabebereich leeren (Format bleibt)
        _clear_cells_keep_format(sheet, "C3:F17")

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

    return True


# =========================
# INIT: Y/Z Format (100 Zeilen)
# =========================
def _init_yz_format(sheet, rows: int = 100):
    """
    Setzt Y/Z (1..rows) horizontal + vertikal zentriert, Schriftgröße 12.
    """
    try:
        rng = sheet.getCellRangeByName(f"Y1:Z{rows}")
        rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
        rng.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
        rng.CharHeight = 12
    except Exception:
        pass


# =========================
# Button: immer neu zeichnen
# =========================
def run_wordspiel_button(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)
    
    _commit_current_cell_input(doc)
    _park_cursor_priority(doc, sheet)
    _init_yz_format(sheet, rows=100)

    # 1) Ziel J2 (wie viele Halbkreise via Random füllen dürfen)
    target = _read_j2_target(sheet, default=6)

    # 2) Bereits manuell belegt zählen (EF oder Grid)
    already = 0
    halves = []
    for (title, ef_top, grid_top, grid_bot, ef_bot) in GROUPS:
        halves.append((title, ef_top, grid_top))
        halves.append((title, ef_bot, grid_bot))
    for (_t, ef_r, grid_r) in halves:
        if _half_has_manual_input(sheet, ef_r, grid_r):
            already += 1

    remaining_random = target - already
    if remaining_random < 0:
        remaining_random = 0

    state = {"remaining_random": remaining_random, "msgs": []}

    doc.lockControllers()
    try:
        # immer neu zeichnen
        delete_all_circles()

        for gi, (title, ef_top, grid_top, grid_bot, ef_bot) in enumerate(GROUPS):
            top_pairs, ok1, _used1 = _pairs_for_half(sheet, title, ef_top, grid_top, state)
            bot_pairs, ok2, _used2 = _pairs_for_half(sheet, title, ef_bot, grid_bot, state)

            # <<< CHANGE: wir brechen NICHT mehr wegen EF <5 ab (ok ist praktisch immer True)
            if not (ok1 and ok2):
                break

            y = _row_top_y(sheet, ef_top)
            fill = ROW_COLORS[gi] if gi < len(ROW_COLORS) else ROW_COLORS[-1]

            for c in range(NUM_COLS):
                idx_tag = f"b{gi}_c{c}"
                t = top_pairs[c] if c < len(top_pairs) else "  "
                b = bot_pairs[c] if c < len(bot_pairs) else "  "

                ul = t[0] if len(t) > 0 else " "
                ur = t[1] if len(t) > 1 else " "
                ll = b[0] if len(b) > 0 else " "
                lr = b[1] if len(b) > 1 else " "
                quad = [ul, ur, ll, lr]

                x = START_X + c * OFFSET_X_BASE
                _draw_circle(doc, dp, x, y, CIRCLE_DIAMETER, fill, quad, idx_tag)

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

    if state["msgs"]:
        _msgbox(doc, "Hinweise", "\n".join(state["msgs"]))

    return True

# =========================
# EXPORTS
# =========================
g_exportedScripts = (
    run_wordspiel_button,
    delete_all_circles,
    rotate_letters_in_circles_and_clear_cells,
    clear_circle_contents_and_cells,
)
