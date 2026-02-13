# -*- coding: utf-8 -*-

import uno
import random
from datetime import datetime, timedelta
from com.sun.star.awt import Point, Size

# ============================================================
# KONFIG
# ============================================================

SHEET_NAME = "KREIS_WORTSPIEL"
MARK_DESC  = "KREIS_Wortspiel_GUI"

# Kreise: 3 Gruppen, je Gruppe 4 Kreise (C..F)
NUM_COLS = 4
CIRCLE_DIAMETER = 3900
START_X         = 14000
OFFSET_X_BASE   = 4000

ROW_COLORS = [0xCCE5FF, 0xFBE5B6, 0xCCFFCC]
CHAR_HEIGHT = 16

# Gruppen: (Titel, EF_top, Grid_top_row, Grid_bot_row, EF_bot)
GROUPS = [
    ("Gruppe 1",  3,  4,  5,  6),
    ("Gruppe 2",  8,  9, 10, 11),
    ("Gruppe 3", 13, 14, 15, 16),
]

# Reihenlisten
EF_ROWS   = (3, 6, 8, 11, 13, 16)
GRID_ROWS = (4, 5, 9, 10, 14, 15)

# Wortspalten (1. Zahl nur Info, genutzt wird der Spaltenbuchstabe)
WORDLIST_COLS = {8: "AG", 5: "AN", 6: "AU", 7: "BB"}
RANDOM_DAYS_LOCK = 30

# Layout (minimal – kann erweitert werden)
INIT_HEADER_RANGE = "B2:F2"
INIT_CENTER_RANGE = "B2:F50"
SHEET_INIT_FLAG   = "KREIS_WORTSPIEL_INIT_DONE"

U100MM_PER_CM = 1000
def cm(v): return int(round(v * U100MM_PER_CM))

INIT_COL_W_C_TO_G = cm(1.7)
INIT_COL_W_B_MAX  = cm(4.5)
INIT_ROW_H_17     = cm(1.7)
INIT_ROWS_H17     = (1, 2, 17, 18)

# ============================================================
# UNO / HELPERS
# ============================================================

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

def _format_messages(title: str, items: list, max_lines: int = 35) -> str:
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

def _get_sheet(doc):
    try:
        return doc.Sheets.getByName(SHEET_NAME)
    except Exception:
        _msgbox(doc, "Fehler", f"Sheet '{SHEET_NAME}' nicht gefunden.")
        raise

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

def _cell(sheet, col0, row_1based):
    return sheet.getCellByPosition(col0, row_1based - 1)

def _col0_from_letters(col_letters: str) -> int:
    s = col_letters.strip().upper()
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

# ============================================================
# TEXT NORMALISIERUNG (Umlaute bleiben!)
# ============================================================

def _to_upper_visual(s):
    return (s or "").replace("ß", "ẞ").upper()

def _normalize_keep_umlauts_no_spaces(s: str) -> str:
    x = (s or "")
    x = "".join(ch for ch in x if not ch.isspace())
    x = _to_upper_visual(x)
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜẞ")
    return "".join(ch for ch in x if ch in allowed)

def _read_j2_target(sheet, default=6):
    # J2 = Anzahl Wörter (1-6). Akzeptiert Zahl oder Text.
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

def _half_has_manual_input(sheet, ef_row_1based, grid_row_1based):
    # zählt als "gefüllt", wenn EF nicht leer ODER in C..F der Grid-Zeile irgendwo Buchstaben stehen
    ef = (_cell(sheet, 4, ef_row_1based).getString() or "").strip()
    if ef:
        return True
    row0 = grid_row_1based - 1
    for col0 in range(2, 6):  # C..F
        raw = (sheet.getCellByPosition(col0, row0).getString() or "").strip()
        norm = _normalize_keep_umlauts_no_spaces(raw)
        if norm:
            return True
    return False

def normalize_input_cells(sheet):
    # EF-Zellen und Grid-Zellen auf Großschrift / erlaubte Zeichen trimmen
    # (wichtig: wir verändern hier NICHT auf 2 Zeichen, das macht die Grid-Logik)
    for r in EF_ROWS:
        c = _cell(sheet, 4, r)  # E (merged)
        raw = (c.getString() or "").strip()
        if raw:
            norm = _normalize_keep_umlauts_no_spaces(raw)
            if norm != raw:
                c.setString(norm)

    for r in GRID_ROWS:
        row0 = r - 1
        for col0 in range(2, 6):  # C..F
            c = sheet.getCellByPosition(col0, row0)
            raw = (c.getString() or "")
            if raw.strip():
                norm = _normalize_keep_umlauts_no_spaces(raw)
                if norm != raw:
                    c.setString(norm)

# ============================================================
# EF -> 4 Pairs (8 Slots). Meldung nur bei <5 oder >8
# ============================================================

def _pairs_from_ef_word(word: str, label: str):
    msgs = []
    w = _normalize_keep_umlauts_no_spaces(word)
    n = len(w)

    ok = True
    if n > 8:
        msgs.append(f"{label}: >8 Zeichen → auf 8 gekürzt")
        w = w[:8]
        n = 8

    if n < 5:
        msgs.append(f"{label}: <5 Zeichen → nicht gültig")
        ok = False

    pairs = []
    for i in range(0, 8, 2):
        chunk = w[i:i+2]
        if len(chunk) == 2:
            pairs.append(chunk)
        elif len(chunk) == 1:
            pairs.append(chunk + " ")
        else:
            pairs.append("  ")

    return pairs, msgs, ok

# ============================================================
# GRID: pro Zelle 0/1/2 erlaubt. >2 wird gekürzt.
# Meldung nur bei 1 oder >2. Leer = keine Meldung.
# ============================================================

def _pairs_from_grid_row(sheet, row_1based: int):
    msgs = []
    pairs = []
    row0 = row_1based - 1

    for col0 in range(2, 6):  # C..F
        addr = _a1(col0, row0)
        c = sheet.getCellByPosition(col0, row0)
        raw = (c.getString() or "").strip()

        norm = _normalize_keep_umlauts_no_spaces(raw)

        if len(norm) == 0:
            pairs.append("  ")
            continue

        if len(norm) == 1:
            pairs.append(norm + " ")
            msgs.append(f"{addr}: 1 Buchstabe → 2. fehlt")
            if raw != norm:
                c.setString(norm)
            continue

        if len(norm) == 2:
            pairs.append(norm)
            if raw != norm:
                c.setString(norm)
            continue

        # >2: kürzen
        cut = norm[:2]
        pairs.append(cut)
        msgs.append(f"{addr}: zu viele Zeichen ('{norm}') → gekürzt auf '{cut}'")
        c.setString(cut)

    return pairs, msgs, True

def _grid_all_empty(pairs):
    for p in pairs:
        if (p or "").strip():
            return False
    return True

# ============================================================
# RANDOM aus Wortspalten (nur 5..8), 30-Tage Regel
# ============================================================

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

def _pick_random_word_from_lists(sheet):
    candidates = []
    cutoff = datetime.now() - timedelta(days=RANDOM_DAYS_LOCK)

    for _L, col_letters in WORDLIST_COLS.items():
        col0 = _col0_from_letters(col_letters)

        for r in range(4, 600):
            w = (_cell(sheet, col0, r).getString() or "").strip()
            if not w:
                continue

            norm = _normalize_keep_umlauts_no_spaces(w)
            if not (5 <= len(norm) <= 8):
                continue

            used_dt = _parse_ts(_cell(sheet, col0 + 4, r).getString())
            if used_dt is None or used_dt < cutoff:
                candidates.append((w, col_letters, r))

    if not candidates:
        return (None, None, None)

    return random.choice(candidates)

def _mark_word_used(sheet, list_col_letters: str, row_1based: int, used_word: str):
    col0 = _col0_from_letters(list_col_letters)
    _cell(sheet, col0 + 3, row_1based).setString(_to_upper_visual(used_word))
    _cell(sheet, col0 + 4, row_1based).setString(_now_ts())

# ============================================================
# Paare für Halbkreis holen: EF -> Grid -> Random (wenn EF+Grid leer)
# ============================================================

def _get_pairs_for_half(sheet, title, ef_row_1based, grid_row_1based, allow_random=False, random_budget=None):
    """
    Priorität pro Halbkreis:
    1) EF (wenn nicht leer) -> EF-Regel (<5 oder >8 meldet; >8 kürzt)
    2) sonst Grid (C..F) -> 0/1/2 ok; 1 oder >2 meldet; >2 kürzt
    3) sonst Random -> NUR wenn allow_random=True UND random_budget['n'] > 0 UND EF+Grid leer
       Random wird NICHT in EF geschrieben.
    """
    msgs = []

    # 1) EF
    ef_cell = _cell(sheet, 4, ef_row_1based)
    ef_raw = (ef_cell.getString() or "").strip()
    if ef_raw:
        vis = _to_upper_visual(ef_raw)
        if vis != ef_raw:
            ef_cell.setString(vis)
        pairs, m, ok = _pairs_from_ef_word(vis, f"EF{ef_row_1based}")
        msgs.extend([f"{title}: {x}" for x in m])
        return pairs, msgs, ok

    # 2) GRID
    pairs, m, ok = _pairs_from_grid_row(sheet, grid_row_1based)
    msgs.extend([f"{title}: {x}" for x in m])

    # Wenn im Grid irgendwas steht -> Grid verwenden (kein Random)
    if not _grid_all_empty(pairs):
        return pairs, msgs, True

    # 3) RANDOM: nur wenn erlaubt + Budget vorhanden
    if not allow_random:
        return pairs, msgs, True

    if random_budget is not None:
        try:
            if int(random_budget.get("n", 0)) <= 0:
                return pairs, msgs, True
        except Exception:
            return pairs, msgs, True

    w, col_letters, row1 = _pick_random_word_from_lists(sheet)
    if not w:
        msgs.append(f"{title}: Keine Random-Wörter verfügbar (5–8) oder alles in {RANDOM_DAYS_LOCK} Tagen genutzt.")
        return ["  ", "  ", "  ", "  "], msgs, False

    vis_w = _to_upper_visual(w)

    # NICHT in EF schreiben!
    _mark_word_used(sheet, col_letters, row1, vis_w)

    # Budget reduzieren (damit J2 wirkt)
    if random_budget is not None:
        random_budget["n"] = int(random_budget.get("n", 0)) - 1

    pairs2, m2, ok2 = _pairs_from_ef_word(vis_w, f"EF{ef_row_1based} (Random {col_letters}{row1})")
    msgs.extend([f"{title}: {x}" for x in m2])

    return pairs2, msgs, ok2

# ============================================================
# SHAPES zeichnen / updaten
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

def _name_map_from_drawpage(dp):
    m = {}
    for i in range(dp.getCount()):
        sh = dp.getByIndex(i)
        nm = getattr(sh, "Name", "") or ""
        if nm.startswith(MARK_DESC + "_"):
            m[nm] = sh
    return m

def delete_all_circles(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = sheet.DrawPage

    removed = 0
    skipped = 0

    # rückwärts iterieren
    for i in range(dp.getCount() - 1, -1, -1):
        sh = dp.getByIndex(i)

        # Controls (Buttons etc.) NICHT löschen
        try:
            if sh.supportsService("com.sun.star.drawing.ControlShape"):
                skipped += 1
                continue
        except Exception:
            pass

        try:
            dp.remove(sh)
            removed += 1
        except Exception:
            # wenn ein Shape nicht entfernt werden kann, weiter
            pass

    _msgbox(doc, "Löschen", f"Entfernt: {removed} Shapes\nÜbersprungen (Controls): {skipped}")

# ============================================================
# INIT Layout (minimal)
# ============================================================

def _set_center(rng):
    rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
    rng.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")

def _set_bold_14(rng):
    rng.CharHeight = 14
    rng.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

def _init_sheet_layout(doc, sheet):
    # Header
    sheet.getCellRangeByName("B2").setString("Bezeichnung")
    sheet.getCellRangeByName("C2").setString("Kreis \n-1-")
    sheet.getCellRangeByName("D2").setString("Kreis \n-2-")
    sheet.getCellRangeByName("E2").setString("Kreis \n-3-")
    sheet.getCellRangeByName("F2").setString("Kreis \n-4-")

    hdr = sheet.getCellRangeByName(INIT_HEADER_RANGE)
    _set_center(hdr)
    _set_bold_14(hdr)

    # Hinweis in Spalte B bei EF-Zeilen
    hint = "Begriffe\n(5–8)"
    for r in EF_ROWS:
        sheet.getCellByPosition(1, r - 1).setString(hint)

    _set_center(sheet.getCellRangeByName(INIT_CENTER_RANGE))

    # EF-Zellen (E..F) mergen
    for r in EF_ROWS:
        r0 = r - 1
        ef = sheet.getCellRangeByPosition(4, r0, 5, r0)
        try:
            ef.merge(True)
        except Exception:
            pass
        _set_center(ef)
        _set_bold_14(ef)

    # Spaltenbreiten
    cols = sheet.Columns
    for col in range(2, 7):  # C..G
        try:
            cols.getByIndex(col).Width = INIT_COL_W_C_TO_G
        except Exception:
            pass

    # Spalte B max Breite
    col_b = cols.getByIndex(1)
    try:
        col_b.OptimalWidth = True
        if col_b.Width > INIT_COL_W_B_MAX:
            col_b.Width = INIT_COL_W_B_MAX
    except Exception:
        pass

    # Zeilenhöhen
    rows = sheet.Rows
    for r in INIT_ROWS_H17:
        try:
            rows.getByIndex(r - 1).Height = INIT_ROW_H_17
        except Exception:
            pass

def _ensure_initialized(doc, sheet):
    if _get_init_done(doc):
        return
    _init_sheet_layout(doc, sheet)
    _set_init_done(doc)

# ============================================================
# CREATE / UPDATE (mit allow_random-Option!)
# ============================================================

def create_circle_grid(allow_random=False, random_budget=None, *args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    all_msgs = []
    doc.lockControllers()
    try:
        _ensure_initialized(doc, sheet)
        normalize_input_cells(sheet)

        delete_all_circles()

        for gi, (title, ef_top, grid_top, grid_bot, ef_bot) in enumerate(GROUPS):
            top_pairs, m1, ok1 = _get_pairs_for_half(sheet, title, ef_top, grid_top, allow_random=allow_random, random_budget=random_budget)
            bot_pairs, m2, ok2 = _get_pairs_for_half(sheet, title, ef_bot, grid_bot, allow_random=allow_random, random_budget=random_budget)
            all_msgs.extend(m1 + m2)
            if not (ok1 and ok2):
                return False, all_msgs

            y = _row_top_y(sheet, ef_top)
            fill = ROW_COLORS[gi] if gi < len(ROW_COLORS) else ROW_COLORS[-1]

            for c in range(NUM_COLS):
                x = START_X + c * OFFSET_X_BASE
                idx_tag = f"b{gi}_c{c}"

                t = top_pairs[c] if c < len(top_pairs) else "  "
                b = bot_pairs[c] if c < len(bot_pairs) else "  "

                ul = t[0] if len(t) > 0 else " "
                ur = t[1] if len(t) > 1 else " "
                ll = b[0] if len(b) > 0 else " "
                lr = b[1] if len(b) > 1 else " "
                quad = [ul, ur, ll, lr]

                draw_circle_with_quadrants(doc, dp, x, y, CIRCLE_DIAMETER, fill, quad, idx_tag)

        return True, all_msgs

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

def update_texts_only(allow_random=False, random_budget=None, *args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    all_msgs = []
    missing = []

    doc.lockControllers()
    try:
        _ensure_initialized(doc, sheet)
        normalize_input_cells(sheet)

        name_map = _name_map_from_drawpage(dp)

        for gi, (title, ef_top, grid_top, grid_bot, ef_bot) in enumerate(GROUPS):
            top_pairs, m1, ok1 = _get_pairs_for_half(sheet, title, ef_top, grid_top, allow_random=allow_random)
            bot_pairs, m2, ok2 = _get_pairs_for_half(sheet, title, ef_bot, grid_bot, allow_random=allow_random)
            all_msgs.extend(m1 + m2)
            if not (ok1 and ok2):
                return False, all_msgs

            for c in range(NUM_COLS):
                idx_tag = f"b{gi}_c{c}"
                circle = name_map.get(f"{MARK_DESC}_circle_{idx_tag}")
                if circle is None:
                    missing.append(f"{title}: Kreis fehlt – bitte neu zeichnen")
                    continue

                t = top_pairs[c] if c < len(top_pairs) else "  "
                b = bot_pairs[c] if c < len(bot_pairs) else "  "

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

        if missing:
            all_msgs.append(_format_messages("Fehlende Shapes", missing, max_lines=30))
            return False, all_msgs

        return True, all_msgs

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

# ============================================================
# ROW TOP Y (für Kreis-Position)
# ============================================================

def _row_top_y(sheet, row_1based):
    y = 0
    rows = sheet.Rows
    for i in range(row_1based - 1):
        y += rows.getByIndex(i).Height
    return y

# ============================================================
# SCRAMBLE (optional, keine Meldung außer Shapes fehlen)
# ============================================================

def _rotate_right(vals):
    return [vals[2], vals[0], vals[3], vals[1]]

def scramble_all_circles_no_solution(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    # wenn keine Shapes da -> still raus
    if not name_map:
        return

    for gi in range(len(GROUPS)):
        for c in range(NUM_COLS):
            idx_tag = f"b{gi}_c{c}"
            # 4 Texte lesen
            vals = []
            for q in range(1, 5):
                sh = name_map.get(f"{MARK_DESC}_text_{idx_tag}_{q}")
                if sh is None:
                    vals = None
                    break
                vals.append((sh.String or " ").strip()[:1] or " ")
            if not vals:
                continue

            steps = random.choice([0, 1, 2, 3])
            v = vals
            for _ in range(steps):
                v = _rotate_right(v)

            for q, val in enumerate(v, start=1):
                sh = name_map.get(f"{MARK_DESC}_text_{idx_tag}_{q}")
                if sh is not None:
                    try:
                        sh.String = val
                    except Exception:
                        pass

# ============================================================
# CLEAR INPUT C3:F16 (dein Wunsch)
# ============================================================

def _clear_manual_input_C3_F16(sheet):
    try:
        rng = sheet.getCellRangeByName("C3:F16")
        rng.clearContents(23)
    except Exception:
        try:
            sheet.getCellRangeByName("C3:F16").String = ""
        except Exception:
            pass

def refresh_and_scramble(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)

    doc.lockControllers()
    try:
        _ensure_initialized(doc, sheet)
        _clear_manual_input_C3_F16(sheet)
    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

    run_wordspiel_button()

# ============================================================
# HAUPT-BUTTON
# - EF->Grid->Random (nur wenn EF+Grid leer)
# - Kreise erstellen/aktualisieren MIT allow_random=True
# - Msgbox nur wenn es Meldungen gibt
# ============================================================

def run_wordspiel_button(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    msgs = []

    doc.lockControllers()
    try:
        _ensure_initialized(doc, sheet)
        normalize_input_cells(sheet)
    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

    # Ziel aus J2 (1–6)
    target = _read_j2_target(sheet, default=6)

    # bereits manuell gefüllte Halbkreise zählen (EF oder Grid)
    already = 0
    halves = []
    for (title, ef_top, grid_top, grid_bot, ef_bot) in GROUPS:
        halves.append((title, ef_top, grid_top))
        halves.append((title, ef_bot, grid_bot))

    for (_title, ef_r, grid_r) in halves:
        if _half_has_manual_input(sheet, ef_r, grid_r):
            already += 1

    budget_n = target - already
    if budget_n < 0:
        budget_n = 0

    random_budget = {"n": budget_n}

    # Shapes vorhanden?
    has_any = False
    try:
        for i in range(dp.getCount()):
            sh = dp.getByIndex(i)
            if getattr(sh, "Description", "") == MARK_DESC:
                has_any = True
                break
    except Exception:
        pass

    # >>> Wichtig: Budget per KEYWORD übergeben
    if has_any:
        res = update_texts_only(allow_random=True, random_budget=random_budget)
    else:
        res = create_circle_grid(allow_random=True, random_budget=random_budget)

    # >>> robust: res kann bool oder (ok, msgs) sein
    if isinstance(res, tuple) and len(res) == 2:
        ok, m = res
    else:
        ok, m = bool(res), []

    if m:
        # m kann schon Strings oder Listen enthalten
        if isinstance(m, list):
            msgs.extend([x for x in m if x])
        else:
            msgs.append(str(m))

    if ok:
        scramble_all_circles_no_solution()

    if msgs:
        _msgbox(doc, "Hinweise", _format_messages("Hinweise", msgs, max_lines=35))

    return ok


# ============================================================
# EXPORTS
# ============================================================

g_exportedScripts = (
    run_wordspiel_button,
    refresh_and_scramble,
    create_circle_grid,
    update_texts_only,
    delete_all_circles,
    scramble_all_circles_no_solution,
)
