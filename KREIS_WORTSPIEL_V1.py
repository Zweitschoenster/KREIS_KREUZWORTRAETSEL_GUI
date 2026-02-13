# -*- coding: utf-8 -*-
# GUI = Graphical User Interface - Deutsch üblich: Grafische Benutzeroberfläche

import uno
import random
from datetime import datetime, timedelta
from com.sun.star.awt import Point, Size

# ============================================================
# 1) KONFIGURATION
# ============================================================

SHEET_NAME = "KREIS_WORTSPIEL"
MARK_DESC  = "KREIS_Wortspiel_GUI"   # Marker zum Löschen (Description)

EF_ROWS   = (3, 6, 8, 11, 13, 16)
GRID_ROWS = (4, 5, 9, 10, 14, 15)

NUM_COLS = 4            # C..F
CIRCLE_DIAMETER = 3900  # 1/100 mm
OFFSET_X_BASE   = 4000
START_X         = 14000

SHOW_MSGBOX = True

ROW_COLORS = [0xCCE5FF, 0xFBE5B6, 0xCCFFCC]
CHAR_HEIGHT = 16

GROUPS = [
    ("Gruppe 1",  3,  4,  5,  6),   # EF3 -> C..F4, EF6 -> C..F5
    ("Gruppe 2",  8,  9, 10, 11),   # EF8 -> C..F9, EF11 -> C..F10
    ("Gruppe 3", 13, 14, 15, 16),   # EF13 -> C..F14, EF16 -> C..F15
]

# Wortlisten (wie du es beschrieben hast)
# Wortspalten: 8=AG, 5=AN, 6=AU, 7=BB
# +1 = Zeitstempel Wortliste (existiert schon)
# +3 = verwendetes Wort im Kreis (same row)
# +4 = Zeitstempel für 30 Tage (same row)
WORDLIST_COLS = {8: "AG", 5: "AN", 6: "AU", 7: "BB"}

# ============================================================
# 2) UNO / CALC HELFER
# ============================================================

U100MM_PER_CM = 1000
def cm(v): return int(round(v * U100MM_PER_CM))

INIT_CENTER_RANGE = "B2:F50"
INIT_HEADER_RANGE = "B2:F2"
INIT_WORD_ROWS    = (3, 6, 8, 11, 13, 16)
INIT_COL_W_C_TO_G = cm(1.7)
INIT_COL_W_B_MAX  = cm(4.5)
INIT_ROW_H_17     = cm(1.7)
INIT_ROW_H_12     = cm(1.2)
INIT_ROW_H_09     = cm(0.9)
INIT_ROWS_H17     = (1, 2, 17, 18)

SHEET_INIT_FLAG = "KREIS_WORTSPIEL_INIT_DONE"

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
    if not items:
        return ""
    items = [x for x in items if x]
    if not items:
        return ""
    if len(items) <= max_lines:
        return f"{title} ({len(items)}):\n" + "\n".join(items)
    head = "\n".join(items[:max_lines])
    rest = len(items) - max_lines
    return f"{title} ({len(items)}):\n{head}\n...\n(+{rest} weitere)"

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

def disable_spellcheck_for_range(sheet, a1_range: str):
    try:
        rng = sheet.getCellRangeByName(a1_range)
        loc = uno.createUnoStruct("com.sun.star.lang.Locale")
        loc.Language = ""
        loc.Country  = ""
        loc.Variant  = ""
        rng.CharLocale = loc
    except Exception:
        pass

def _maybe_msgbox(doc, title, message):
    global SHOW_MSGBOX
    if SHOW_MSGBOX:
        _msgbox(doc, title, message)

# ============================================================
# 3) TEXT / NORMALISIERUNG (Umlaute BEHALTEN)
# ============================================================

def _to_upper_visual(s):
    # ß -> ẞ, Umlaute bleiben
    return (s or "").replace("ß", "ẞ").upper()

def _normalize_for_pairs_keep_umlauts(s: str) -> str:
    """
    Für Kreise/Pairs:
    - Umlaute bleiben EIN Zeichen (ÄÖÜẞ)
    - nur Buchstaben zulassen, Leerzeichen raus
    """
    x = (s or "")
    x = "".join(ch for ch in x if not ch.isspace())
    x = _to_upper_visual(x)
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜẞ")
    x = "".join(ch for ch in x if ch in allowed)
    return x

def _upper_keep_umlauts(s):
    return _to_upper_visual(s)

def _get_ef_word(sheet, ef_row_1based):
    cell = sheet.getCellByPosition(4, ef_row_1based - 1)  # E (EF merged)
    raw = cell.getString() or ""
    vis = _upper_keep_umlauts(raw)
    if raw != vis:
        cell.setString(vis)
    return vis.strip()

def normalize_input_cells(sheet):
    ef_rows = []
    data_rows = []
    for (_title, ef_top, top_row, bot_row, ef_bot) in GROUPS:
        ef_rows.extend([ef_top, ef_bot])
        data_rows.extend([top_row, bot_row])
    for r in set(ef_rows):
        cell = sheet.getCellByPosition(4, r - 1)
        raw = cell.getString() or ""
        up  = _upper_keep_umlauts(raw)
        if raw != up:
            cell.setString(up)
    for r in set(data_rows):
        row0 = r - 1
        for col in range(2, 6):
            cell = sheet.getCellByPosition(col, row0)
            raw = cell.getString() or ""
            up  = _to_upper_visual(raw)
            if raw != up:
                cell.setString(up)

# ============================================================
# 4) PAIRS: EF / GRID / RANDOM
# ============================================================

def _cell(sheet, col0, row_1based):
    return sheet.getCellByPosition(col0, row_1based - 1)

def _col0_from_letters(col_letters: str) -> int:
    s = col_letters.strip().upper()
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

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
    

def _pairs_from_wordspiel_word(word: str, cell_label: str):
    """
    EF-Regel:
    - <5: Fehler + ok=False
    - >8: wird auf 8 gekürzt + Meldung
    - 5..8: ok, KEINE Meldung (auch nicht bei 5..7)
    - Restplätze werden leer, ungerade Länge wird gepaddet (Buchstabe geht nicht verloren)
    """
    msgs = []
    clean = _normalize_for_pairs_keep_umlauts(word)
    n = len(clean)

    ok = True

    if n > 8:
        msgs.append(f"{cell_label}: mehr als 8 Zeichen → auf 8 gekürzt")
        clean = clean[:8]
        n = 8

    if n < 5:
        msgs.append(f"{cell_label}: weniger als 5 Zeichen → nicht gültig")
        ok = False

    pairs = []
    for i in range(0, 8, 2):
        chunk = clean[i:i+2]
        if len(chunk) == 2:
            pairs.append(chunk)
        elif len(chunk) == 1:
            pairs.append(chunk + " ")
        else:
            pairs.append("  ")

    return pairs, msgs, ok

def _pairs_from_grid_row_strict(sheet, row_1based: int):
    """
    Grid-Regel (C..F):
    - 0 Zeichen: ok, keine Meldung
    - 1 Zeichen: ok, wird mit Leerzeichen ergänzt, MELDUNG (1 fehlt)
    - 2 Zeichen: ok, keine Meldung
    - >2 Zeichen: wird auf 2 gekürzt, MELDUNG (zuviel, gekürzt)
    Umwandlungen bleiben wie bisher (Umlaute bleiben 1 Zeichen).
    """
    msgs = []
    pairs = []
    row0 = row_1based - 1

    for col0 in range(2, 6):  # C..F
        addr = _a1(col0, row0)
        cell = sheet.getCellByPosition(col0, row0)
        raw = (cell.getString() or "").strip()

        # deine bisherige Normalisierung (ohne AE/OE/UE-Umwandlung)
        norm = _normalize_for_pairs_keep_umlauts(raw)

        if len(norm) == 0:
            pairs.append("  ")
            continue

        if len(norm) == 1:
            pairs.append(norm + " ")
            msgs.append(f"{addr}: 1 Buchstabe → 2. fehlt")
            # optional: Zelle auf norm setzen (groß/sauber)
            if raw != norm:
                cell.setString(norm)
            continue

        if len(norm) == 2:
            pairs.append(norm)
            if raw != norm:
                cell.setString(norm)
            continue

        # >2: abschneiden
        cut = norm[:2]
        pairs.append(cut)
        msgs.append(f"{addr}: zu viele Zeichen ('{norm}') → gekürzt auf '{cut}'")
        cell.setString(cut)

    return (pairs, msgs, True)

def _grid_row_is_all_empty(pairs):
    for p in pairs:
        if (p or "").strip():
            return False
    return True

def _pick_random_word_from_lists(sheet):
    """
    Random aus 5..8 Spalten verteilt,
    30-Tage Regel über (+4) Timestamp.
    """
    candidates = []
    cutoff = datetime.now() - timedelta(days=30)

    for L, col_letters in WORDLIST_COLS.items():
        col0 = _col0_from_letters(col_letters)
        for r in range(4, 600):
            w = (_cell(sheet, col0, r).getString() or "").strip()
            if not w:
                continue

            # NUR 5..8 zulassen (mit deiner Umlaut-Normalisierung)
            norm = _normalize_for_pairs_keep_umlauts(w)
            if not (5 <= len(norm) <= 8):
                continue

            used_dt = _parse_ts(_cell(sheet, col0 + 4, r).getString())
            if used_dt is None or used_dt < cutoff:
                candidates.append((w, L, col_letters, r))

    if not candidates:
        return (None, None, None, None)

    return random.choice(candidates)

def _mark_word_used(sheet, list_col_letters: str, row_1based: int, used_word: str):
    col0 = _col0_from_letters(list_col_letters)
    _cell(sheet, col0 + 3, row_1based).setString(_to_upper_visual(used_word))
    _cell(sheet, col0 + 4, row_1based).setString(_now_ts())

def _get_pairs_for_half(sheet, doc, title, ef_row_1based, fallback_row_1based, allow_random=False):
    """
    Priorität:
      1) EF (wenn nicht leer) -> EF-Regel 5..8
      2) sonst Grid -> Grid-Regel
      3) sonst Random (5..8) -> schreibt Wort nach EF + used markieren
    Rückgabe: (pairs, messages, ok)
    """
    msgs = []

    # 1) EF
    ef_cell = _cell(sheet, 4, ef_row_1based)  # E (merged)
    raw_ef = (ef_cell.getString() or "").strip()
    if raw_ef:
        vis = _to_upper_visual(raw_ef)
        if vis != raw_ef:
            ef_cell.setString(vis)
        pairs, m, ok = _pairs_from_wordspiel_word(vis, f"EF{ef_row_1based}")
        msgs.extend([f"{title}: {x}" for x in m])
        return (pairs, msgs, ok)

    # 2) Grid
    pairs, m, ok = _pairs_from_grid_row_strict(sheet, fallback_row_1based)
    msgs.extend([f"{title}: {x}" for x in m])
    if not ok:
        return (pairs, msgs, False)

    if not _grid_row_is_all_empty(pairs):
        return (pairs, msgs, True)

    # 3) Random: nur wenn EF leer UND Grid komplett leer
    if not allow_random:
        return (pairs, msgs, True)

    if not _grid_row_is_all_empty(pairs):
        return (pairs, msgs, True)

    w, L, col_letters, row1 = _pick_random_word_from_lists(sheet)
    if not w:
        msgs.append(f"{title}: Keine Random-Wörter verfügbar (5..8) oder alles innerhalb 30 Tage genutzt.")
        return (["  ", "  ", "  ", "  "], msgs, False)

    vis_w = _to_upper_visual(w)

    # WICHTIG: NICHT in EF schreiben (EF bleibt leer)
    _mark_word_used(sheet, col_letters, row1, vis_w)

    pairs2, m2, ok2 = _pairs_from_wordspiel_word(vis_w, f"EF{ef_row_1based} (Random {col_letters}{row1})")
    msgs.extend([f"{title}: {x}" for x in m2])

    return (pairs2, msgs, ok2)


# ============================================================
# 5) CURSOR-SETZEN
# ============================================================

def _select_cell(doc, sheet, col0, row_1based):
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
    for r in EF_ROWS:
        s = (_cell(sheet, 4, r).getString() or "").strip()
        if not s:
            if _select_cell(doc, sheet, 4, r):
                return
    for col0 in (2, 3, 4, 5):  # C,D,E,F
        for r in GRID_ROWS:
            s = (_cell(sheet, col0, r).getString() or "").strip()
            if not s:
                if _select_cell(doc, sheet, col0, r):
                    return

# ============================================================
# 6) ZEICHNEN (SHAPES)
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

def draw_line(doc, draw_page, x1, y1, x2, y2, name_suffix):
    line = doc.createInstance("com.sun.star.drawing.LineShape")
    _mark_shape(line, name_suffix)
    draw_page.add(line)
    line.Position = Point(x1, y1)
    line.Size = Size(x2 - x1, y2 - y1)
    line.LineColor = 0x000000

def draw_circle_with_quadrants(doc, draw_page, x, y, size, fill_color, label_texts, idx_tag):
    circle = doc.createInstance("com.sun.star.drawing.EllipseShape")
    _mark_shape(circle, f"circle_{idx_tag}")
    draw_page.add(circle)

    circle.Position  = Point(x, y)
    circle.Size      = Size(size, size)
    circle.FillColor = fill_color
    circle.LineColor = 0x000000

    cx = x + size // 2
    cy = y + size // 2

    draw_line(doc, draw_page, cx, y,  cx, y + size, f"vline_{idx_tag}")
    draw_line(doc, draw_page, x,  cy, x + size, cy, f"hline_{idx_tag}")

    q1x = x + size // 4
    q3x = x + (3 * size) // 4
    q1y = y + size // 4
    q3y = y + (3 * size) // 4

    centers = [(q1x, q1y), (q3x, q1y), (q1x, q3y), (q3x, q3y)]
    for i, ((tcx, tcy), txt) in enumerate(zip(centers, label_texts), start=1):
        draw_text_center(doc, draw_page, tcx, tcy, txt, f"text_{idx_tag}_{i}")

# ============================================================
# 7) ROW-HEIGHT / POSITION
# ============================================================

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
    radius = int(CIRCLE_DIAMETER // 2)
    base_word = int(INIT_ROW_H_12)
    base_letters = int(INIT_ROW_H_09)
    base_sum = base_word + base_letters
    if base_sum <= 0:
        return
    scale = radius / float(base_sum)
    new_word = int(round(base_word * scale))
    new_letters = radius - new_word
    for (top_word, top_letters), (bot_letters, bot_word) in ROW_PAIRS:
        _set_row_height(sheet, top_word, new_word)
        _set_row_height(sheet, top_letters, new_letters)
        _set_row_height(sheet, bot_letters, new_letters)
        _set_row_height(sheet, bot_word, new_word)

def _row_top_y(sheet, row_1based):
    y = 0
    rows = sheet.Rows
    for i in range(row_1based - 1):
        y += rows.getByIndex(i).Height
    return y

def _circle_top_y_for_group(sheet, group_index):
    ef_top_row = GROUPS[group_index][1]
    return _row_top_y(sheet, ef_top_row)

# ============================================================
# 8) DELETE / INIT
# ============================================================

def delete_all_circles(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    def is_ours(sh):
        try:
            if sh.supportsService("com.sun.star.drawing.ControlShape"):
                return False
        except Exception:
            pass
        nm = getattr(sh, "Name", "") or ""
        desc = getattr(sh, "Description", "") or ""
        return (desc == MARK_DESC or nm.startswith(MARK_DESC + "_"))

    for i in range(dp.getCount() - 1, -1, -1):
        sh = dp.getByIndex(i)
        try:
            if is_ours(sh):
                dp.remove(sh)
        except Exception:
            pass

def _init_sheet_layout(doc):
    sheet = doc.Sheets.getByName(SHEET_NAME)

    disable_spellcheck_for_range(sheet, "E3:F3")
    disable_spellcheck_for_range(sheet, "E6:F6")
    disable_spellcheck_for_range(sheet, "E8:F8")
    disable_spellcheck_for_range(sheet, "E11:F11")
    disable_spellcheck_for_range(sheet, "E13:F13")
    disable_spellcheck_for_range(sheet, "E16:F16")
    disable_spellcheck_for_range(sheet, "C4:F5")
    disable_spellcheck_for_range(sheet, "C9:F10")
    disable_spellcheck_for_range(sheet, "C14:F15")

    sheet.getCellRangeByName("B2").setString("Bezeichnung")
    sheet.getCellRangeByName("C2").setString("Kreis \n-1-")
    sheet.getCellRangeByName("D2").setString("Kreis \n-2-")
    sheet.getCellRangeByName("E2").setString("Kreis \n-3-")
    sheet.getCellRangeByName("F2").setString("Kreis \n-4-")

    hdr = sheet.getCellRangeByName(INIT_HEADER_RANGE)
    _set_center(hdr)
    _set_bold_14(hdr)

    hint = "Begriffe\n(5–8)"
    for r in (3, 6, 8, 11, 13, 16):
        cell = sheet.getCellByPosition(1, r - 1)
        cell.setString(hint)
        try:
            cell.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
            cell.CharHeight = 14
        except Exception:
            pass

    _set_center(sheet.getCellRangeByName(INIT_CENTER_RANGE))

    for r in INIT_WORD_ROWS:
        r0 = r - 1
        ef = sheet.getCellRangeByPosition(4, r0, 5, r0)  # E..F
        try:
            ef.merge(True)
        except Exception:
            pass
        _set_center(ef)
        _set_bold_14(ef)

    cols = sheet.Columns
    for col in range(2, 7):  # C..G
        cols.getByIndex(col).Width = INIT_COL_W_C_TO_G

    col_b = cols.getByIndex(1)
    try:
        col_b.OptimalWidth = True
        if col_b.Width > INIT_COL_W_B_MAX:
            col_b.Width = INIT_COL_W_B_MAX
    except Exception:
        pass

    rows = sheet.Rows
    for r in INIT_ROWS_H17:
        rows.getByIndex(r - 1).Height = INIT_ROW_H_17

    adjust_rows_to_circle_radius(sheet)

    try:
        doc.CurrentController.freezeAtPosition(0, 1)
    except Exception:
        pass

def _ensure_initialized(doc):
    if _get_init_done(doc):
        return
    _init_sheet_layout(doc)
    _set_init_done(doc)

# ============================================================
# 9) UPDATE / CREATE
# ============================================================

def _name_map_from_drawpage(dp):
    m = {}
    for i in range(dp.getCount()):
        sh = dp.getByIndex(i)
        nm = getattr(sh, "Name", "") or ""
        if nm.startswith(MARK_DESC + "_"):
            m[nm] = sh
    return m

def update_texts_only(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)

    all_msgs = []
    missing = []

    doc.lockControllers()
    try:
        _ensure_initialized(doc)
        adjust_rows_to_circle_radius(sheet)
        normalize_input_cells(sheet)

        dp = _get_draw_page(sheet)
        name_map = _name_map_from_drawpage(dp)

        for group_index, (title, ef_top, top_row, bot_row, ef_bot) in enumerate(GROUPS):
            top_pairs, m1, ok1 = _get_pairs_for_half(sheet, doc, title, ef_top, top_row, allow_random=False)
            bot_pairs, m2, ok2 = _get_pairs_for_half(sheet, doc, title, ef_bot, bot_row, allow_random=False)
            all_msgs.extend(m1 + m2)
            if not (ok1 and ok2):
                return False

            for c in range(NUM_COLS):
                idx_tag = f"b{group_index}_c{c}"
                circle = name_map.get(f"{MARK_DESC}_circle_{idx_tag}")
                if circle is None:
                    missing.append(f"{title}: Kreis fehlt – bitte neu zeichnen")
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
                    sh = name_map.get(f"{MARK_DESC}_text_{idx_tag}_{i}")
                    if sh is None:
                        missing.append(f"{title}: Text fehlt – bitte neu zeichnen")
                        continue
                    txt = _to_upper_visual(txt)
                    try:
                        sh.String = txt
                    except Exception:
                        try:
                            sh.Text.setString(txt)
                        except Exception:
                            pass
                    w = sh.Size.Width or 700
                    h = sh.Size.Height or 700
                    sh.Position = Point(int(cx - w/2), int(cy - h/2))

        return not missing

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass
        set_cursor_next_input(doc, sheet)
        if all_msgs or missing:
            parts = []
            if all_msgs:
                parts.append(_format_messages("Hinweise", all_msgs, max_lines=30))
            if missing:
                parts.append(_format_messages("Fehlende Shapes", missing, max_lines=30))
            _maybe_msgbox(doc, "Update", "\n\n".join([p for p in parts if p]))

def create_circle_grid(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    all_msgs = []

    doc.lockControllers()
    try:
        delete_all_circles()
        _ensure_initialized(doc)
        adjust_rows_to_circle_radius(sheet)
        normalize_input_cells(sheet)

        dp = _get_draw_page(sheet)

        for group_index, (title, ef_top, top_row, bot_row, ef_bot) in enumerate(GROUPS):
            top_pairs, m1, ok1 = _get_pairs_for_half(sheet, doc, title, ef_top, top_row, allow_random=False)
            bot_pairs, m2, ok2 = _get_pairs_for_half(sheet, doc, title, ef_bot, bot_row, allow_random=False)
            all_msgs.extend(m1 + m2)
            if not (ok1 and ok2):
                return False

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

                draw_circle_with_quadrants(doc, dp, x, y, CIRCLE_DIAMETER, fill, quad, idx_tag)

        return True

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass
        set_cursor_next_input(doc, sheet)
        if all_msgs:
            _maybe_msgbox(doc, "Hinweise", _format_messages("Hinweise", all_msgs, max_lines=30))

# ============================================================
# 10) SCRAMBLE / REFRESH (mit LEEREN von C3:F16)
# ============================================================

MAX_SOLVED_CIRCLES = 2

def _rotate_right_vals(vals):
    return [vals[2], vals[0], vals[3], vals[1]]

def _rotate_vals(vals, steps_clockwise):
    v = list(vals)
    for _ in range(steps_clockwise % 4):
        v = _rotate_right_vals(v)
    return v

def scramble_all_circles_no_solution(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    missing = []
    forced_solved = 0
    total = 0
    plan = []

    for gi, (title, ef_top, top_row, bot_row, ef_bot) in enumerate(GROUPS):
        top_pairs, _, ok1 = _get_pairs_for_half(sheet, doc, title, ef_top, top_row, allow_random=False)
        bot_pairs, _, ok2 = _get_pairs_for_half(sheet, doc, title, ef_bot, bot_row, allow_random=False)
        if not (ok1 and ok2):
            return

        for c in range(NUM_COLS):
            idx_tag = f"b{gi}_c{c}"
            circle_name = f"{MARK_DESC}_circle_{idx_tag}"
            if circle_name not in name_map:
                missing.append(f"{title}: Kreis fehlt")
                continue

            total += 1
            t = top_pairs[c] if c < len(top_pairs) else ""
            b = bot_pairs[c] if c < len(bot_pairs) else ""

            ul = t[0] if len(t) > 0 else " "
            ur = t[1] if len(t) > 1 else " "
            ll = b[0] if len(b) > 0 else " "
            lr = b[1] if len(b) > 1 else " "
            base = [ul, ur, ll, lr]

            candidates = [k for k in (1, 2, 3) if _rotate_vals(base, k) != base]
            if candidates:
                step = random.choice(candidates)
            else:
                step = 0
                forced_solved += 1
            plan.append((idx_tag, base, step))

    if missing:
        _maybe_msgbox(doc, "Scramble", "Es fehlen Shapes – bitte neu zeichnen.")
        return

    solved_after = 0
    for idx_tag, base, step in plan:
        vals = _rotate_vals(base, step)
        if vals == base:
            solved_after += 1
        for q, val in enumerate(vals, start=1):
            nm = f"{MARK_DESC}_text_{idx_tag}_{q}"
            sh = name_map.get(nm)
            if sh is not None:
                try:
                    sh.String = val
                except Exception:
                    pass

    allowed = min(MAX_SOLVED_CIRCLES, max(total - 1, 0))
    if total > 0 and solved_after > allowed:
        _maybe_msgbox(doc, "Scramble Hinweis",
                f"{solved_after} Kreise stehen in Ausgangsstellung (erlaubt: {allowed}).\n"
                f"Grund: {forced_solved} Kreise rotationssymmetrisch/leer.")

def _clear_manual_input_C3_F16(sheet):
    # NUR hier löschen (dein Wunsch)
    try:
        rng = sheet.getCellRangeByName("C3:F16")
        rng.clearContents(23)  # Strings/Werte/Formeln etc.
    except Exception:
        # fallback: wenigstens Strings
        try:
            sheet.getCellRangeByName("C3:F16").String = ""
        except Exception:
            pass

def refresh_and_scramble(*args):
    """
    EXTRA-BUTTON:
    - löscht NUR HIER C3:F16 (inkl. EF und Grid)
    - dann sorgt der Button-Lauf (run_wordspiel_button) wieder für Inhalte
    - danach Kreise zeichnen/aktualisieren + scramble
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)

    doc.lockControllers()
    try:
        _ensure_initialized(doc)
        _clear_manual_input_C3_F16(sheet)
    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

    # danach ganz normal Button-Logik
    run_wordspiel_button()

# ============================================================
# 11) HAUPT-BUTTON: EF -> GRID -> RANDOM, dann Kreise + Scramble
# ============================================================

def run_wordspiel_button(*args):
    """
    BUTTON-MAKRO:
    - Priorität pro Halbkreis: EF -> Grid -> Random aus Wortlisten (5..8)
    - danach: update (wenn Shapes existieren) sonst create
    - danach scramble
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)

    msgs = []

    doc.lockControllers()
    try:
        _ensure_initialized(doc)
        adjust_rows_to_circle_radius(sheet)
        normalize_input_cells(sheet)

        # sorgt dafür, dass für ALLE Halbkreise Daten existieren (inkl. Random)
        for (title, ef_top, top_row, bot_row, ef_bot) in GROUPS:
            _p1, m1, ok1 = _get_pairs_for_half(sheet, doc, title, ef_top, top_row, allow_random=True)
            _p2, m2, ok2 = _get_pairs_for_half(sheet, doc, title, ef_bot, bot_row, allow_random=True)
            msgs.extend(m1 + m2)
            if not (ok1 and ok2):
                return False
    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

    # Shapes exist?
    dp = _get_draw_page(sheet)
    has_any = False
    try:
        for i in range(dp.getCount()):
            sh = dp.getByIndex(i)
            if getattr(sh, "Description", "") == MARK_DESC:
                has_any = True
                break
    except Exception:
        pass



    # >>> Nur hier EINMAL anzeigen
    if msgs:
        _msgbox(doc, "Hinweise", _format_messages("Hinweise", msgs, max_lines=30))

    return True

# ============================================================
# 12) EXPORTS
# ============================================================

g_exportedScripts = (
    run_wordspiel_button,
    refresh_and_scramble,
    create_circle_grid,
    update_texts_only,
    delete_all_circles,
    scramble_all_circles_no_solution,
)
