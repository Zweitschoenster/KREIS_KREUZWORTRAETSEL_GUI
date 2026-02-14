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

# Wortspalten (gleicher Sheet)
# Wort steht in COL, Timestamp (Eintrag) steht in COL+1 (existiert schon)
# Used-Wort im Kreis steht in COL+3, Used-Timestamp (30 Tage) in COL+4
WORDLIST_COLS = ["AG", "AN", "AU", "BB"]
RANDOM_DAYS_LOCK = 30
WORDLIST_ROW_START = 4
WORDLIST_ROW_END   = 600

# J2: Zielanzahl automatisch zu füllender Halbkreise (1..6)
DEBUG = False


# =========================
# UNO HELPERS
# =========================
def _get_doc():
    return XSCRIPTCONTEXT.getDocument()

def _get_sheet(doc):
    return doc.Sheets.getByName(SHEET_NAME)

def _get_draw_page(sheet):
    return sheet.DrawPage

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

def _col0_from_letters(col_letters: str) -> int:
    s = (col_letters or "").strip().upper()
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def _cell(sheet, col0, row_1based):
    return sheet.getCellByPosition(col0, row_1based - 1)

def _a1(col0, row0):  # row0 0-based
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
# TEXT NORMALISIERUNG (Umlaute bleiben)
# =========================
def _to_upper_visual(s):
    return (s or "").replace("ß", "ẞ").upper()

def _normalize_keep_umlauts_no_spaces(s: str) -> str:
    x = (s or "")
    x = "".join(ch for ch in x if not ch.isspace())
    x = _to_upper_visual(x)
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜẞ")
    return "".join(ch for ch in x if ch in allowed)


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
# EF / GRID Regeln
# =========================
def _pairs_from_ef_word(sheet, ef_row_1based, label):
    """
    EF:
    - wenn leer: None (nicht benutzt)
    - >8: kürzen (auch in Zelle)
    - <5: MELDUNG, aber EF wird IGNORIERT (Fallback auf Grid/Random)
    """
    msgs = []
    ef_cell = _cell(sheet, 4, ef_row_1based)  # E (merged)
    raw = (ef_cell.getString() or "").strip()
    if not raw:
        return None, msgs  # None = "nicht benutzt"

    norm = _normalize_keep_umlauts_no_spaces(raw)
    n = len(norm)

    # >8 kürzen
    if n > 8:
        msgs.append(f"{label}: >8 Zeichen ('{norm}') → auf 8 gekürzt")
        norm = norm[:8]
        ef_cell.setString(norm)
        n = 8
    else:
        if norm != raw:
            ef_cell.setString(norm)

    # <5 => NICHT gültig für EF, aber KEIN Abbruch!
    if n < 5:
        msgs.append(f"{label}: '{norm}' hat {n} Zeichen (<5) → EF wird ignoriert (Grid/Random)")
        return None, msgs  # <- wichtig: None bedeutet "EF nicht verwenden"

    # 5..8 => in Paare
    pairs = []
    for i in range(0, 8, 2):
        chunk = norm[i:i+2]
        if len(chunk) == 2:
            pairs.append(chunk)
        elif len(chunk) == 1:
            pairs.append(chunk + " ")
        else:
            pairs.append("  ")

    return pairs, msgs


def _pairs_from_grid_row(sheet, row_1based):
    """
    Grid C..F:
    - 0 Zeichen => "  "
    - 1 Zeichen => "<X> " + Hinweis
    - 2 Zeichen => ok
    - >2 => Kürzen auf 2 + Hinweis mit Original und Cut
    """
    msgs = []
    pairs = []
    row0 = row_1based - 1

    for col0 in range(2, 6):  # C..F
        c = sheet.getCellByPosition(col0, row0)
        raw = (c.getString() or "").strip()
        norm = _normalize_keep_umlauts_no_spaces(raw)

        if len(norm) == 0:
            pairs.append("  ")
        elif len(norm) == 1:
            pairs.append(norm + " ")
            msgs.append(f"{_a1(col0,row0)}: 1 Buchstabe → 2. fehlt ('{norm} ')")
            if raw != norm:
                c.setString(norm)
        elif len(norm) == 2:
            pairs.append(norm)
            if raw != norm:
                c.setString(norm)
        else:
            cut = norm[:2]
            pairs.append(cut)
            msgs.append(f"{_a1(col0,row0)}: zu viele Zeichen ('{norm}') → gekürzt auf '{cut}'")
            c.setString(cut)

    return pairs, msgs


def _grid_all_empty(pairs):
    return all(not (p or "").strip() for p in pairs)


def _half_has_manual_input(sheet, ef_row_1based, grid_row_1based):
    ef = (_cell(sheet, 4, ef_row_1based).getString() or "").strip()
    if ef:
        return True
    row0 = grid_row_1based - 1
    for col0 in range(2, 6):
        raw = (sheet.getCellByPosition(col0, row0).getString() or "").strip()
        if _normalize_keep_umlauts_no_spaces(raw):
            return True
    return False


# =========================
# RANDOM aus Wortspalten (5..8), 30 Tage Sperre
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
        for r in range(WORDLIST_ROW_START, WORDLIST_ROW_END + 1):
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
        _msgbox(_get_doc(), "DEBUG", f"Random-Kandidaten: {len(candidates)}")

    if not candidates:
        return None, None, None

    return random.choice(candidates)


# =========================
# SHAPES: löschen / zeichnen
# =========================
def delete_all_circles(*args):
    """
    Löscht alle Shapes auf dem Sheet (außer Controls).
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = sheet.DrawPage

    for i in range(dp.getCount() - 1, -1, -1):
        sh = dp.getByIndex(i)
        try:
            if sh.supportsService("com.sun.star.drawing.ControlShape"):
                continue
        except Exception:
            pass
        try:
            dp.remove(sh)
        except Exception:
            pass

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
    line.Position = Point(int(x1), int(y1))
    line.Size = Size(int(x2 - x1), int(y2 - y1))
    line.LineColor = 0x000000

def _draw_circle(doc, draw_page, x, y, size, fill_color, quad, idx_tag):
    circle = doc.createInstance("com.sun.star.drawing.EllipseShape")
    _mark_shape(circle, f"circle_{idx_tag}")
    draw_page.add(circle)

    circle.Position  = Point(int(x), int(y))
    circle.Size      = Size(int(size), int(size))
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
# Halbkreis -> Paare (EF -> Grid -> Random)
# =========================
def _pairs_for_half(sheet, title, ef_row, grid_row, state):
    """
    Priorität:
    1) EF (nur wenn >=5 Zeichen)
    2) Grid (wenn dort irgendwas steht)
    3) Random (wenn erlaubt)
    """
    # 1) EF (nur wenn gültig >=5, sonst None)
    pairs, m = _pairs_from_ef_word(sheet, ef_row, f"EF{ef_row}")
    if m:
        state["msgs"].extend([f"{title}: {x}" for x in m])
    if pairs is not None:
        return pairs  # EF genutzt

    # 2) Grid
    gpairs, gm = _pairs_from_grid_row(sheet, grid_row)
    if gm:
        state["msgs"].extend([f"{title}: {x}" for x in gm])

    if not _grid_all_empty(gpairs):
        return gpairs  # Grid genutzt

    # 3) Random
    if state["remaining_random"] <= 0:
        return ["  ", "  ", "  ", "  "]  # leer

    w, col_letters, row1 = _pick_random_word(sheet)
    if not w:
        if DEBUG:
            state["msgs"].append(f"{title}: keine Random-Kandidaten gefunden")
        return ["  ", "  ", "  ", "  "]

    _mark_word_used(sheet, col_letters, row1, w)
    state["remaining_random"] -= 1

    pairs = []
    for i in range(0, 8, 2):
        chunk = w[i:i+2]
        if len(chunk) == 2:
            pairs.append(chunk)
        elif len(chunk) == 1:
            pairs.append(chunk + " ")
        else:
            pairs.append("  ")
    return pairs


# =========================
# BUTTON-MAKRO (FIX)
# =========================
def run_wordspiel_button(*args):
    """
    Fix: Erst prüfen/Paare bauen -> dann löschen & zeichnen.
    Dadurch bleiben Kreise sichtbar, wenn EF ungültig ist.
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    target = _read_j2_target(sheet, default=6)

    halves = []
    for (title, ef_top, grid_top, grid_bot, ef_bot) in GROUPS:
        halves.append((title, ef_top, grid_top))
        halves.append((title, ef_bot, grid_bot))

    already = 0
    for (_t, ef_r, grid_r) in halves:
        if _half_has_manual_input(sheet, ef_r, grid_r):
            already += 1

    remaining_random = max(0, target - already)
    state = {"remaining_random": remaining_random, "msgs": []}

    doc.lockControllers()
    try:
        # jetzt immer neu zeichnen
        delete_all_circles()

        for gi, (title, ef_top, grid_top, grid_bot, ef_bot) in enumerate(GROUPS):
            top_pairs = _pairs_for_half(sheet, title, ef_top, grid_top, state)
            bot_pairs = _pairs_for_half(sheet, title, ef_bot, grid_bot, state)

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
)
