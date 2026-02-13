# -*- coding: utf-8 -*-
# ============================================================
# 1) CONFIG
# ============================================================
import uno
import random
import math
import time
from datetime import datetime, timedelta

# --- robust fallback: Point/Size always available ---
try:
    from com.sun.star.awt import Point, Size
except Exception:
    def Point(x, y):
        p = uno.createUnoStruct("com.sun.star.awt.Point")
        p.X = int(x)
        p.Y = int(y)
        return p

    def Size(w, h):
        s = uno.createUnoStruct("com.sun.star.awt.Size")
        s.Width = int(w)
        s.Height = int(h)
        return s


# --- GENERAL ---
SHEET_NAME = "KREIS_WORTRAETSEL"
MARK_DESC  = "KREIS_WORTRAETSEL_V1"

# --- TEST / PROTECTION (TESTPHASE: OFF) ---
ENABLE_SHEET_PROTECTION = False
# ENABLE_SHEET_PROTECTION = True # Wenn Testphase beendet ist
PROTECT_PASSWORD = ""  # empty = easy to remove manually

# --- INPUT CELLS ---
CELL_WORDLEN   = "G3"   # word length (5..12)
CELL_WORDCOUNT = "C3"   # how many words to pick (1..6)

DEFAULT_WORDLEN   = 8
DEFAULT_WORDCOUNT = 2
INPUT_COL_D0 = 3   # D

# --- GEOMETRY ---
ROW_H_CM = 1.11
U100MM_PER_CM = 1000

def cm(v):
    """centimeters -> 1/100 mm"""
    return int(round(v * U100MM_PER_CM))

CIRCLE_DIAMETER = cm(2.0 * ROW_H_CM)  # 2 * 1.11 cm = 2.22 cm
H_GAP_CM = 0.20                       # small horizontal gap between circles

ROW_COLORS = (
    0xCCE5FF,  # row 1: blue
    0xFBE5B6,  # row 2: yellow
    0xCCFFCC,  # row 3: green
)

# Circle row starts (1-based):
# Row4/5 = circle row 1, Row7/8 = row 2, Row10/11 = row 3
CIRCLE_ROW_STARTS_1BASED = (4, 7, 10)

# Grid rows that must be set to 1.11cm (including row 12)
GRID_ROWS_1BASED = (4, 5, 6, 7, 8, 9, 10, 11, 12)

# Header rows (we keep them stable and freeze)
HEADER_ROWS_1BASED = (1, 2, 3)

# Left labels in column B (merged blocks)
LABEL_BLOCKS = (
    (4, 5),
    (7, 8),
    (10, 11),
)

# Anchor: RIGHT edge of column D = LEFT edge of column E
ANCHOR_COL0 = 4  # E (0-based)

# Maximum circles per row (for 11/12 letters -> 6 circles)
MAX_CIRCLES = 6

# --- WORDLIST LAYOUT (BLOCKS FROM AG, 5..12 LETTERS) ---
WORDLIST_BASE_COL0 = 32         # AG
WORDLIST_BLOCK_STEP = 7         # 7 columns per length block
WORDLIST_LEN_MIN = 5
WORDLIST_LEN_MAX = 12

WORDLIST_HEADER_ROW_1BASED = 3  # Überschrift steht über der Liste (Liste startet in Zeile 4)
WORDLIST_START_ROW_1BASED  = 4  # first word row
WORDLIST_MAX_ROWS = 300         # scan/format depth for list content

# Row height rule for list area
ROWHEIGHT_FROM_ROW_1BASED = 13  # set row height from row 13 downward
ROWHEIGHT_ROWS_COUNT = 300      # how far down to apply row height (safe buffer)

# ---------- CONFIG FOR WORDLIST LAYOUT ----------
# You told: 8-letter words are in BI (and timestamp is +2 columns => BK).
WORDLIST_ANCHOR_LEN = 8
WORDLIST_ANCHOR_COL_A1 = "BB"     # word column for 8-letter list
WORDLIST_COL_STEP = 7             # move 7 columns per length (adjust if needed)
TIMESTAMP_COL_OFFSET = 2          # timestamp is +2 columns from word column
RECENT_DAYS = 30

# Slots (where we store picked words in Y and map to circle halves)
SLOT_ROWS = [4, 5, 7, 8, 10, 11]     # top/bot for 3 circle rows
STORE_COL_Y0 = 24                    # Y


# ============================================================
# 2) UNO & Cell Helpers (_cell, _set_row_height, _debug_a1, …)
# ============================================================

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
        box_type = uno.Enum("com.sun.star.awt.MessageBoxType", "WARNINGBOX")
        box = toolkit.createMessageBox(parent, box_type, buttons, title, message)
        box.execute()
    except Exception:
        pass

def _debug_a1(sheet, msg):
    try:
        sheet.getCellByPosition(0, 0).setString(str(msg))
    except Exception:
        pass

def _set_center(rng):
    try:
        rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
        rng.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
    except Exception:
        pass

def _set_bold(rng, pt=14):
    try:
        rng.CharHeight = pt
        rng.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
    except Exception:
        pass

def _read_int(sheet, a1, default=0):
    """
    Reads an integer from a single cell (A1 address).
    Prefers numeric value; falls back to string parsing.
    """
    try:
        c = _cell(sheet, a1)  # <-- FIX: get the cell, not the sheet
        v = c.getValue()
        if v is not None and v != 0:
            return int(v)
    except Exception:
        pass

    try:
        c = _cell(sheet, a1)  # safe second attempt
        s = (c.getString() or "").strip()
        return int(s)
    except Exception:
        return default


def _set_row_height(sheet, row_1based, height_u100mm):
    rr = sheet.Rows.getByIndex(row_1based - 1)
    try:
        rr.OptimalHeight = False
    except Exception:
        pass
    rr.Height = int(height_u100mm)

def _set_col_width(sheet, col0, width_u100mm):
    cc = sheet.Columns.getByIndex(col0)
    try:
        cc.OptimalWidth = False
    except Exception:
        pass
    cc.Width = int(width_u100mm)

def _row_top_y(sheet, row_1based):
    y = 0
    rows = sheet.Rows
    for i in range(row_1based - 1):
        y += rows.getByIndex(i).Height
    return y

def _col_left_x(sheet, col0):
    x = 0
    cols = sheet.Columns
    for i in range(col0):
        x += cols.getByIndex(i).Width
    return x

def _lock_range(rng):
    try:
        prot = rng.CellProtection
        prot.IsLocked = True
        rng.CellProtection = prot
    except Exception:
        pass

def _unlock_range(rng):
    try:
        prot = rng.CellProtection
        prot.IsLocked = False
        rng.CellProtection = prot
    except Exception:
        pass

def _cell(sheet, a1):
    """Return a single cell (XCell) from an A1 address."""
    return sheet.getCellRangeByName(a1).getCellByPosition(0, 0)
    
def set_cursor_to_C4(doc, sheet):
    """Move cursor to C4 and focus it."""
    try:
        ctrl = doc.CurrentController
        rng = sheet.getCellRangeByName("C4")
        ctrl.select(rng)  # marks the cell
        # force edit cursor into that cell (more reliable than select alone)
        ctrl.setActiveCell(rng.getCellByPosition(0, 0))
    except Exception:
        pass


def _read_manual_D_overrides(sheet, cap_letters: int):
    """
    Reads manual overrides from D4/D5/D7/D8/D10/D11.
    No length validation:
      - non letters removed by _normalize_crossword
      - too long -> truncated to cap_letters
      - too short -> accepted
    Returns: (manual_map row1->word, rows_to_clear)
    """
    manual = {}
    rows_to_clear = []

    for r in SLOT_ROWS:  # [4,5,7,8,10,11]
        cell = sheet.getCellByPosition(INPUT_COL_D0, r - 1)
        raw = (cell.getString() or "").strip()
        if not raw:
            continue

        rows_to_clear.append(r)

        w = _normalize_crossword(raw)
        if not w:
            # user typed something but nothing usable remains -> treat as "empty override"
            continue

        if cap_letters > 0 and len(w) > cap_letters:
            w = w[:cap_letters]

        manual[r] = w

    return manual, rows_to_clear


def _touch_timestamps_for_used_words(word_to_item: dict, used_words: set, now_str: str):
    """
    Sets timestamp ONLY for words that are actually used in final result.
    word_to_item: mapping word -> candidate item dict from _read_candidates (contains 'tcell').
    """
    for w in used_words:
        it = word_to_item.get(w)
        if it is None:
            continue
        try:
            it["tcell"].setString(now_str)
        except Exception:
            pass


def _clamp_wordlen_to_5_12(sheet, writeback=True) -> int:
    """Reads G3, clamps to 5..12. Optionally writes back to G3."""
    g3 = _cell(sheet, CELL_WORDLEN)
    try:
        raw = int(round(g3.getValue() or 0))
    except Exception:
        raw = DEFAULT_WORDLEN

    L = raw
    if L < WORDLIST_LEN_MIN: L = WORDLIST_LEN_MIN
    if L > WORDLIST_LEN_MAX: L = WORDLIST_LEN_MAX

    if writeback and L != raw:
        try:
            g3.setValue(float(L))     # bleibt eine Zahl in G3
        except Exception:
            g3.setString(str(L))
    return L


def _update_left_labels(sheet, L: int):
    """Updates B4/5, B7/8, B10/11 to 'Begriffe mit L Buchstaben'."""
    txt = f"Begriffe\nmit {L} Buchstaben"
    for (r1, r2) in LABEL_BLOCKS:
        rng = sheet.getCellRangeByPosition(1, r1 - 1, 1, r2 - 1)  # Spalte B
        try:
            rng.merge(True)
        except Exception:
            pass
        sheet.getCellByPosition(1, r1 - 1).setString(txt)
        _set_center(rng)
        _set_bold(rng, pt=14)


# ============================================================
# PART X – WORDLIST BLOCK FORMATTING (AG.., 5..12 letters)
# ============================================================

def format_wordlist_layout(*args):
    """Callable macro: formats the wordlist blocks (AG..)."""
    doc = _get_doc()
    sheet = _get_sheet(doc)
    _format_wordlist_blocks(sheet)
    _debug_a1(sheet, "Wordlist layout formatted (AG.. blocks + row heights from 13).")
    return True
    

def _wordlist_word_col0_for_len_from_AG(L: int) -> int:
    # 5 -> AG, 6 -> AN, 7 -> AU, 8 -> BB, ... 12 -> CD
    return WORDLIST_BASE_COL0 + (L - WORDLIST_LEN_MIN) * WORDLIST_BLOCK_STEP


def _format_wordlist_blocks(sheet):
    hdr_r0 = WORDLIST_HEADER_ROW_1BASED - 1
    start_r0 = WORDLIST_START_ROW_1BASED - 1
    end_r0 = start_r0 + WORDLIST_MAX_ROWS - 1

    def word_col_width_cm(L):
        return max(L * 0.6, 3.6)

    wide_w   = cm(3.8)
    narrow_w = cm(1.0)

    for L in range(WORDLIST_LEN_MIN, WORDLIST_LEN_MAX + 1):  # <-- IMPORTANT +1
        block0 = _wordlist_word_col0_for_len_from_AG(L)

        # widths: [word] + [3x 3.8cm] + [3x 1.0cm]
        _set_col_width(sheet, block0 + 0, cm(word_col_width_cm(L)))
        _set_col_width(sheet, block0 + 1, wide_w)
        _set_col_width(sheet, block0 + 2, wide_w)
        _set_col_width(sheet, block0 + 3, wide_w)
        _set_col_width(sheet, block0 + 4, narrow_w)
        _set_col_width(sheet, block0 + 5, narrow_w)
        _set_col_width(sheet, block0 + 6, narrow_w)

        # header row formatting across the whole block
        try:
            hdr_rng = sheet.getCellRangeByPosition(block0, hdr_r0, block0 + 6, hdr_r0)
            hdr_rng.CharHeight = 10.5
            hdr_rng.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
        except Exception:
            pass

        # AH..AM style within each block: +1..+6
        try:
            rng = sheet.getCellRangeByPosition(block0 + 1, start_r0, block0 + 6, end_r0)
            rng.CharHeight = 12
            rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
            rng.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
        except Exception:
            pass

    # row height from row 13 downward
    rh = cm(0.7)
    for r in range(ROWHEIGHT_FROM_ROW_1BASED, ROWHEIGHT_FROM_ROW_1BASED + ROWHEIGHT_ROWS_COUNT):
        try:
            _set_row_height(sheet, r, rh)
        except Exception:
            break

def sync_wordlist_headers(*args):
    """Callable macro: fix/update all headers 5..12."""
    doc = _get_doc()
    sheet = _get_sheet(doc)
    _sync_wordlist_headers(sheet, overwrite=True)
    _debug_a1(sheet, "Wordlist headers synced (5..12).")
    return True
  
    
# ============================================================
# PART 2A – RANDOM WORD PICK (Wordlist -> Timestamp -> Y -> Circles)
# ============================================================


def _sync_wordlist_headers(sheet, overwrite=True):
    hdr_r0 = WORDLIST_HEADER_ROW_1BASED - 1

    for L in range(WORDLIST_LEN_MIN, WORDLIST_LEN_MAX + 1):  # 5..12 inkl.
        col0 = _wordlist_word_col0_for_len_from_AG(L)
        colA1 = _col_index_to_letters(col0)

        wanted = f"Wörter in {colA1}\n{L} Buchstaben"
        cell = sheet.getCellByPosition(col0, hdr_r0)
        cur = (cell.getString() or "").strip()

        if overwrite or (not cur):
            cell.setString(wanted)

        try:
            cell.CharHeight = 10.5
            cell.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
        except Exception:
            pass


# ---------- helpers Spaltenindex → Buchstaben ----------

def _col_letters_to_index(a1_letters: str) -> int:
    s = (a1_letters or "").strip().upper()
    n = 0
    for ch in s:
        if "A" <= ch <= "Z":
            n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1  # 0-based

def _col_index_to_letters(col0: int) -> str:
    n = col0 + 1
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def _parse_ts_string(ts: str):
    s = (ts or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

def _normalize_crossword(raw: str) -> str:
    """
    Crossword normalization:
    - uppercase
    - ÄÖÜ -> AE/OE/UE
    - ß/ẞ -> SS
    - remove everything not A-Z
    """
    s = (raw or "").strip().upper()
    s = (s.replace("Ä", "AE")
           .replace("Ö", "OE")
           .replace("Ü", "UE")
           .replace("ß", "SS")
           .replace("ẞ", "SS"))
    s = "".join(ch for ch in s if "A" <= ch <= "Z")
    return s

def _desired_len_from_G3(sheet) -> int:
    try:
        v = int(round(_cell(sheet, CELL_WORDLEN).getValue() or 0))
    except Exception:
        v = DEFAULT_WORDLEN
    if v < WORDLIST_LEN_MIN: v = WORDLIST_LEN_MIN
    if v > WORDLIST_LEN_MAX: v = WORDLIST_LEN_MAX
    return v

def _wordcount_from_C3(sheet) -> int:
    try:
        v = int(round(_cell(sheet, CELL_WORDCOUNT).getValue() or 0))
    except Exception:
        v = 1
    if v < 1: v = 1
    if v > 6: v = 6
    return v

def _num_circles_for_len(L: int) -> int:
    # 5-6->3, 7-8->4, 9-10->5, 11-12->6
    n = int(math.ceil(L / 2.0))
    if n < 3: n = 3
    if n > 6: n = 6
    return n

def _wordlist_word_col0_for_len(L: int) -> int:
    anchor0 = _col_letters_to_index(WORDLIST_ANCHOR_COL_A1)  # BB
    return anchor0 + (L - WORDLIST_ANCHOR_LEN) * WORDLIST_COL_STEP


def _read_candidates(sheet, L: int):
    """
    Read all non-empty cells in the word column for length L.
    Candidate is valid iff normalized word length == L.
    Timestamp cell is at (word_col + TIMESTAMP_COL_OFFSET).
    """
    col_word = _wordlist_word_col0_for_len(L)
    col_ts = col_word + TIMESTAMP_COL_OFFSET
    start0 = WORDLIST_START_ROW_1BASED - 1
    end0 = start0 + WORDLIST_MAX_ROWS - 1

    items = []
    for r0 in range(start0, end0 + 1):
        wcell = sheet.getCellByPosition(col_word, r0)
        tcell = sheet.getCellByPosition(col_ts, r0)

        raw = (wcell.getString() or "").strip()
        if not raw:
            continue

        w = _normalize_crossword(raw)
        if len(w) != L:
            continue

        ts_s = (tcell.getString() or "").strip()
        ts_dt = _parse_ts_string(ts_s)

        items.append({
            "row0": r0,
            "word": w,
            "wcell": wcell,
            "tcell": tcell,
            "ts": ts_dt
        })
    return items

def _pick_random_with_30day_rule(items, k: int):
    """
    Prefer items with ts older than RECENT_DAYS or empty.
    If not enough, fall back to all items.
    """
    if k <= 0:
        return [], []

    cutoff = datetime.now() - timedelta(days=RECENT_DAYS)
    eligible = [it for it in items if (it["ts"] is None or it["ts"] < cutoff)]

    pool = eligible if len(eligible) >= k else items
    if len(pool) < k:
        return [], [f"Zu wenige gültige Wörter in der Liste (benötigt {k}, verfügbar {len(pool)})."]
    chosen = random.sample(pool, k)
    return chosen, []


def _write_words_to_circles(sheet, desired_len: int, words_by_slotrow: dict):
    """
    Write letters into the existing circle shapes.
    For each circle row:
      - top half word comes from slot row (4,7,10)
      - bottom half word comes from slot row (5,8,11)
    Each circle shows 2 letters top (UL/UR) and 2 letters bottom (LL/LR).
    """
    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    n_circles = _num_circles_for_len(desired_len)
    cap = 2 * n_circles

    for row_index, (top_r, bot_r) in enumerate([(4,5), (7,8), (10,11)]):
        top_w = _normalize_crossword(words_by_slotrow.get(top_r, ""))
        bot_w = _normalize_crossword(words_by_slotrow.get(bot_r, ""))

        # safety truncate (shouldn't happen if list is correct)
        if len(top_w) > cap: top_w = top_w[:cap]
        if len(bot_w) > cap: bot_w = bot_w[:cap]

        for c in range(n_circles):
            idx_tag = f"r{row_index}_c{c}"

            t = top_w[c*2:(c*2)+2]
            b = bot_w[c*2:(c*2)+2]

            ul = t[0] if len(t) > 0 else " "
            ur = t[1] if len(t) > 1 else " "
            ll = b[0] if len(b) > 0 else " "
            lr = b[1] if len(b) > 1 else " "

            vals = (ul, ur, ll, lr)
            for q, ch in enumerate(vals, start=1):
                nm = f"{MARK_DESC}_text_{idx_tag}_{q}"
                sh = name_map.get(nm)
                if sh is None:
                    continue
                try:
                    sh.String = ch
                except Exception:
                    try:
                        sh.Text.setString(ch)
                    except Exception:
                        pass


def _rotate_cw(vals):
    """
    vals order: [UL, UR, LL, LR]
    90° clockwise => [LL, UL, LR, UR]
    """
    return [vals[2], vals[0], vals[3], vals[1]]

def _rotate_ccw(vals):
    """
    vals order: [UL, UR, LL, LR]
    90° counterclockwise => [UR, LR, UL, LL]
    """
    return [vals[1], vals[3], vals[0], vals[2]]

def _rotate_vals_signed(vals, steps_signed):
    v = list(vals)
    s = int(steps_signed) % 4
    if s == 0:
        return v

    if steps_signed > 0:
        for _ in range(s):
            v = _rotate_cw(v)
    else:
        for _ in range(s):
            v = _rotate_ccw(v)
    return v


def scramble_circles_after_fill(sheet, L):
    """
    Scramble ONLY visible circle letters (text shapes).
    - Randomly rotates left or right (CW/CCW), steps 1..3
    - NEVER keeps a circle unchanged on purpose.
    - Only circles that are rotation-symmetric/empty may remain unchanged.
    - Allowed unchanged (upper bound, not a target):
        L>=8  -> <=2
        L<8   -> <=1
    """
    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    n_circles = _num_circles_for_len(L)
    allowed_unchanged = 2 if L >= 8 else 1

    total = 0
    unchanged = 0
    forced_unchanged = 0

    circles = []
    for row_index in range(3):
        for c in range(n_circles):
            idx_tag = f"r{row_index}_c{c}"

            shapes = []
            base_vals = []
            missing = False
            for q in range(1, 5):
                nm = f"{MARK_DESC}_text_{idx_tag}_{q}"
                sh = name_map.get(nm)
                if sh is None:
                    missing = True
                    break
                shapes.append(sh)
                try:
                    base_vals.append((sh.String or " ")[:1])
                except Exception:
                    try:
                        base_vals.append((sh.Text.getString() or " ")[:1])
                    except Exception:
                        base_vals.append(" ")

            if not missing:
                circles.append((idx_tag, shapes, base_vals))

    for idx_tag, shapes, base_vals in circles:
        total += 1

        candidates = []
        for steps in (1, 2, 3):
            cw_vals  = _rotate_vals_signed(base_vals, +steps)
            ccw_vals = _rotate_vals_signed(base_vals, -steps)
            if cw_vals != base_vals:
                candidates.append(cw_vals)
            if ccw_vals != base_vals:
                candidates.append(ccw_vals)

        if not candidates:
            # cannot be changed (symmetric/empty)
            forced_unchanged += 1
            unchanged += 1
            continue

        # IMPORTANT: never choose "no rotation" -> always change if possible
        new_vals = random.choice(candidates)
        for sh, ch in zip(shapes, new_vals):
            try:
                sh.String = ch
            except Exception:
                try:
                    sh.Text.setString(ch)
                except Exception:
                    pass

    # Only warn if we exceed the allowed max (usually only because of symmetry)
    if unchanged > allowed_unchanged:
        try:
            doc = _get_doc()
            _msgbox(
                doc,
                "Scramble Hinweis",
                # f"{unchanged} Kreise sind unverändert (erlaubt max.: {allowed_unchanged}).\n"
                f"Grund: {forced_unchanged} Kreise sind rotationssymmetrisch/leer."
            )
        except Exception:
            pass


# ---------- MAIN MACRO ----------
def part2a_fill_random_from_wordlist(*args):

    """
    PART 2A + D override (no D length validation):
    1) Clamp G3 to 5..12 (write back) and update left labels.
    2) Ensure circle count matches G3 and clear all quadrant letters.
    3) Pick N words (C3) from wordlist (30-day rule), NO timestamp yet.
    4) Apply manual D overrides (truncate only).
    5) If one half of a row-pair is filled, fill the other half from list (if possible).
    6) Write final words to Y, clear used D cells.
    7) Timestamp ONLY for words actually used.
    8) Write letters into circles + scramble.
    9) Write duration to F2 (10pt).
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)

    t0 = time.perf_counter()

    ok = False
    err_msg = None

    # Keep these for finally (avoid NameError)
    L = None
    N = None
    used_words_count = 0
    manual_count = 0

    doc.lockControllers()
    try:
        # --- normalize inputs early (single source of truth: L) ---
        L = _clamp_wordlen_to_5_12(sheet, writeback=True)
        _update_left_labels(sheet, L)

        # --- make sure the circle grid matches current L and is visually clean ---
        ensure_circle_count_matches_G3(sheet)
        clear_all_circle_quadrant_texts(sheet)

        # --- clear input/output areas (keep formatting) ---
        clear_contents_keep_format(sheet, "C4:C12")
        clear_contents_keep_format(sheet, "Y4:Y12")

        # --- read N and compute caps ---
        N = _wordcount_from_C3(sheet)       # 1..6
        now = _now_str()

        n_circles = _num_circles_for_len(L)
        cap = 2 * n_circles                 # max letters per half-row (2 per circle)

        # --- load candidates from list (once) ---
        items = _read_candidates(sheet, L)
        if not items:
            err_msg = f"Keine gültigen Wörter gefunden für Länge {L}."
            return False

        # Map word -> candidate item for timestamp (first occurrence wins)
        word_to_item = {}
        for it in items:
            w = it["word"]
            if w not in word_to_item:
                word_to_item[w] = it

        # Helper: pick 1 excluding already used words (uses 30-day rule if possible)
        def pick_one_excluding(used_words: set):
            pool = [it for it in items if it["word"] not in used_words]
            if not pool:
                return None
            chosen, errs = _pick_random_with_30day_rule(pool, 1)
            if errs or not chosen:
                return None
            return chosen[0]["word"]

        # --- default fill from list (N words), NO timestamp yet ---
        chosen, errs = _pick_random_with_30day_rule(items, N)
        if errs:
            err_msg = "\n".join(errs)
            return False

        assignments = {}  # row1 -> word
        for i, it in enumerate(chosen):
            if i >= len(SLOT_ROWS):
                break
            r = SLOT_ROWS[i]
            assignments[r] = it["word"]

        # --- apply manual overrides from column D (truncate only) ---
        manual_map, rows_to_clear = _read_manual_D_overrides(sheet, cap)
        manual_count = len(manual_map)
        for r, w in manual_map.items():
            assignments[r] = w

        # --- fill missing half in active row-pairs ---
        used_now = set(w for w in assignments.values() if w)

        for top_r, bot_r in [(4, 5), (7, 8), (10, 11)]:
            top_has = bool(assignments.get(top_r))
            bot_has = bool(assignments.get(bot_r))

            if top_has and not bot_has:
                w = pick_one_excluding(used_now)
                if w:
                    assignments[bot_r] = w
                    used_now.add(w)
            elif bot_has and not top_has:
                w = pick_one_excluding(used_now)
                if w:
                    assignments[top_r] = w
                    used_now.add(w)

        # --- determine actually used words (for timestamps) ---
        used_words = set(w for w in assignments.values() if w)
        used_words_count = len(used_words)

        # --- write final words to Y ---
        for r, w in assignments.items():
            try:
                sheet.getCellByPosition(STORE_COL_Y0, r - 1).setString(w)
            except Exception:
                pass

        # --- clear used D input cells ---
        for r in rows_to_clear:
            try:
                sheet.getCellByPosition(INPUT_COL_D0, r - 1).setString("")
            except Exception:
                pass

        # --- timestamp ONLY for words that are really used AND exist in the list ---
        _touch_timestamps_for_used_words(word_to_item, used_words, now)

        # --- write letters into circles + scramble ---
        _write_words_to_circles(sheet, L, assignments)
        scramble_circles_after_fill(sheet, L)

        ok = True
        return True

    finally:
        try:
            doc.unlockControllers()
        except Exception:
            pass

        # --- move cursor to C4 (after unlock is more reliable) ---
        set_cursor_to_C4(doc, sheet)

        # --- duration output to F2 (always) ---
        dt = time.perf_counter() - t0
        try:
            c = sheet.getCellRangeByName("F2")
            c.setString(f"Dauer: {dt:.1f} sec.".replace(".", ","))
            c.CharHeight = 10
        except Exception:
            pass

        # --- optional debug ---
        try:
            if ok:
                _debug_a1(sheet, f"PART2A OK: Länge={L}, C3={N}, used={used_words_count}, D_overrides={manual_count}")
            else:
                _debug_a1(sheet, f"PART2A ERROR: {err_msg or 'unbekannt'}")
        except Exception:
            pass

        # --- messagebox AFTER unlock ---
        if (not ok) and err_msg:
            try:
                _msgbox(doc, "Wortliste", err_msg)
            except Exception:
                pass

# ============================================================
# 4) INIT: set row heights BEFORE drawing
# ============================================================

def set_left_labels_from_G3(*args):
    doc = _get_doc()
    try:
        sheet = _get_sheet(doc)
    except Exception:
        _msgbox(doc, "Init", f"Sheet '{SHEET_NAME}' not found.")
        return False

    x = _read_int(sheet, CELL_WORDLEN, DEFAULT_WORDLEN)
    txt = f"Begriffe\nmit {x} Buchstaben"

    for (r1, r2) in LABEL_BLOCKS:
        rng = sheet.getCellRangeByPosition(1, r1-1, 1, r2-1)  # column B
        try:
            rng.merge(True)
        except Exception:
            pass
        sheet.getCellByPosition(1, r1 - 1).setString(txt)  # Spalte B, Startzeile

        _set_center(rng)
        _set_bold(rng, pt=14)

    return True

def _apply_sheet_protection(sheet):
    if not ENABLE_SHEET_PROTECTION:
        return

    try:
        big = sheet.getCellRangeByName("A1:AZ2000")
        _lock_range(big)
    except Exception:
        pass

    # Unlock input cells (use ranges)
    try:
        _unlock_range(sheet.getCellRangeByName(CELL_WORDLEN))    # G3
    except Exception:
        pass
    try:
        _unlock_range(sheet.getCellRangeByName(CELL_WORDCOUNT))  # C3
    except Exception:
        pass

    try:
        sheet.protect(PROTECT_PASSWORD)
    except Exception:
        pass


def ensure_initialized(*args):
    """
    Final init (test phase, protection disabled):
    - Rows 1..3: 0.9 cm
    - Rows 4..12: 1.11 cm (GRID_ROWS_1BASED includes 12)
    - Keep B2 and C2 empty
    - B3 label: 14pt bold ("Anzahl Wörter (1-6):")
    - F3 label: 10.5pt bold ("Anzahl der gewünschten ... (5-12):")  (no merge)
    - Inputs:
        C3 = number of words (numeric, clamped 1..6)
        G3 = word length (numeric, clamped 5..12)
    - Column widths: B, D, F = 5.1 cm
    - Extra widths: A, C, E, G, Z = 2.5 cm
    - Z3 label: "Anzahl\nBuchstaben:" bold 10.5 pt, left
    - AA2:AB2 merged label (bold 12 pt, left)
    - AA3 input formatting: centered, 20 pt, bold
    - Left labels in B4/5, B7/8, B10/11: "Begriffe\nmit X Buchstaben" (X from G3)
    - Freeze panes after row 3
    """
    doc = _get_doc()
    try:
        sheet = _get_sheet(doc)
    except Exception:
        _msgbox(doc, "Init", f"Sheet '{SHEET_NAME}' not found.")
        return False

    # --- Keep B2 and C2 empty ---
    _cell(sheet, "B2").setString("")
    _cell(sheet, "C2").setString("")

    # --- Labels (as requested) ---
    b3 = _cell(sheet, "B3")
    b3.setString("Anzahl Wörter (1-6):")
    b3.CharHeight = 14
    b3.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

    f3 = _cell(sheet, "F3")
    f3.setString("Anzahl der gewünschten\nBuchstaben pro Wort (5-12):")
    f3.CharHeight = 10.5
    f3.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")

    # Z3 label: bold 10.5 pt, left aligned
    z3 = _cell(sheet, "Z3")
    z3.setString("Anzahl\nBuchstaben:")
    z3.CharHeight = 10.5
    z3.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
    try:
        z3.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "LEFT")
        z3.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
    except Exception:
        pass

    # AA2:AB2 merged, bold 12pt, left aligned
    try:
        rng = sheet.getCellRangeByName("AA2:AB2")
        rng.merge(True)
        _cell(sheet, "AA2").setString("Eingabe von gültigen Wörter,\ndie aufgenommen werden sollen")
        rng.CharHeight = 12
        rng.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
        rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "LEFT")
        rng.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
    except Exception:
        pass

    # AA3 input cell formatting: centered, 20pt, bold (do not lock in test phase)
    aa3 = _cell(sheet, "AA3")
    try:
        aa3.CharHeight = 20
        aa3.CharWeight = uno.getConstantByName("com.sun.star.awt.FontWeight.BOLD")
        aa3.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
        aa3.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
    except Exception:
        pass

    # (For later protection) keep AA3 editable
    try:
        _unlock_range(sheet.getCellRangeByName("AA3"))
    except Exception:
        pass

    # --- Defaults (C3 = wordcount, G3 = wordlen) ---
    wc = _cell(sheet, "C3")
    if not wc.getValue():
        wc.setValue(DEFAULT_WORDCOUNT)

    wl = _cell(sheet, "G3")
    if not wl.getValue():
        wl.setValue(DEFAULT_WORDLEN)

    # --- Clamp inputs to valid ranges ---
    # C3: 1..6
    n = int(round(wc.getValue() or 0))
    if n < 1: n = 1
    if n > 6: n = 6
    wc.setValue(n)

    # G3: 5..12
    x = int(round(wl.getValue() or 0))
    if x < 5: x = 5
    if x > 12: x = 12
    wl.setValue(x)

    # --- If AA3 is empty: copy content from G3 into AA3 (value only, keep AA3 formatting) ---
    try:
        aa3_cell = _cell(sheet, "AA3")
        aa3_txt  = (aa3_cell.getString() or "").strip()

        # "empty" means: no text and no numeric value
        if not aa3_txt and (aa3_cell.getValue() == 0):
            g3_cell = _cell(sheet, "G3")

            # prefer numeric value from G3
            gv = g3_cell.getValue()
            if gv and gv != 0:
                aa3_cell.setString(str(int(round(gv))))
            else:
                aa3_cell.setString((g3_cell.getString() or "").strip())
    except Exception:
        pass

    # Optional: center input cells
    try:
        _set_center(sheet.getCellRangeByName("C3"))
        _set_center(sheet.getCellRangeByName("G3"))
    except Exception:
        pass

    # --- Row heights ---
    header_h = cm(0.90)  # keep as requested
    for r in HEADER_ROWS_1BASED:
        _set_row_height(sheet, r, header_h)

    grid_h = cm(ROW_H_CM)  # 1.11 cm
    for r in GRID_ROWS_1BASED:  # must include 12
        _set_row_height(sheet, r, grid_h)

    # --- Column widths requested: B, D, F = 5.1 cm ---
    _set_col_width(sheet, 1, cm(5.1))  # B
    _set_col_width(sheet, 3, cm(5.1))  # D
    _set_col_width(sheet, 5, cm(5.1))  # F

    # --- Column widths: A, C, E, G, Z = 2.5 cm ---
    w14 = cm(1.4)
    w25 = cm(2.5)
    w36 = cm(3.6)
    _set_col_width(sheet, 0,  w14)  # A
    _set_col_width(sheet, 2,  w36)  # C
    _set_col_width(sheet, 4,  w25)  # E
    _set_col_width(sheet, 6,  w25)  # G
    _set_col_width(sheet, 25, w25)  # Z

    # IMPORTANT:
    # Do NOT set circle-column widths from E onwards here,
    # otherwise layout columns may be overwritten again.
    # Circles use a fixed step_x in draw_circles_from_G3().

    # --- Left labels "Begriffe mit X Buchstaben" in column B blocks ---
    txt = f"Begriffe\nmit {x} Buchstaben"
    for (r1, r2) in LABEL_BLOCKS:
        rng = sheet.getCellRangeByPosition(1, r1 - 1, 1, r2 - 1)  # B(r1):B(r2)
        try:
            rng.merge(True)
        except Exception:
            pass
        sheet.getCellByPosition(1, r1 - 1).setString(txt)
        _set_center(rng)
        _set_bold(rng, pt=14)

    # Freeze after row 3
    try:
        doc.CurrentController.freezeAtPosition(0, 3)
    except Exception:
        pass
        
    try:
        _format_wordlist_blocks(sheet)
        _sync_wordlist_headers(sheet, overwrite=True)
    except Exception:
        pass


    # TESTPHASE: protection disabled
    # try:
    #     _apply_sheet_protection(sheet)
    # except Exception:
    #     pass

    _debug_a1(sheet, "Init OK: rows set, B/D/F=5.1cm, labels written, inputs C3(1-6)/G3(5-12).")
    return True


# ============================================================
# 5) DRAWING PRIMITIVES (shapes)
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

def _draw_line(doc, draw_page, x1, y1, x2, y2, name_suffix):
    line = doc.createInstance("com.sun.star.drawing.LineShape")
    _mark_shape(line, name_suffix)
    draw_page.add(line)
    line.Position = Point(int(x1), int(y1))
    line.Size = Size(int(x2 - x1), int(y2 - y1))
    try:
        line.LineColor = 0x000000
    except Exception:
        pass
    return line

def _draw_text_center(doc, draw_page, cx, cy, text, name_suffix):
    t = doc.createInstance("com.sun.star.drawing.TextShape")
    _mark_shape(t, name_suffix)
    draw_page.add(t)

    # small fixed text box
    w = cm(0.60)
    h = cm(0.60)
    t.Size = Size(int(w), int(h))
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
        t.CharHeight = 16
        t.TextHorizontalAdjust = 1  # CENTER
        t.TextVerticalAdjust   = 1  # CENTER
    except Exception:
        pass

    return t

def _draw_circle_with_quadrants(doc, draw_page, x, y, size, fill_color, idx_tag):
    circle = doc.createInstance("com.sun.star.drawing.EllipseShape")
    _mark_shape(circle, f"circle_{idx_tag}")
    draw_page.add(circle)

    circle.Position = Point(int(x), int(y))
    circle.Size = Size(int(size), int(size))
    try:
        circle.FillColor = fill_color
        circle.LineColor = 0x000000
        circle.FillStyle = uno.Enum("com.sun.star.drawing.FillStyle", "SOLID")
    except Exception:
        pass

    # Cross lines
    cx = x + size // 2
    cy = y + size // 2
    _draw_line(doc, draw_page, cx, y,  cx, y + size, f"vline_{idx_tag}")
    _draw_line(doc, draw_page, x,  cy, x + size, cy, f"hline_{idx_tag}")

    # Quadrant centers
    q1x = x + size // 4
    q3x = x + (3 * size) // 4
    q1y = y + size // 4
    q3y = y + (3 * size) // 4

    centers = [(q1x, q1y), (q3x, q1y), (q1x, q3y), (q3x, q3y)]

    # Create 4 empty text shapes (later you will fill them)
    for i, (tcx, tcy) in enumerate(centers, start=1):
        _draw_text_center(doc, draw_page, tcx, tcy, " ", f"text_{idx_tag}_{i}")


# ============================================================
# 6) DELETE HELPERS
# ============================================================

def _name_map_from_drawpage(dp):
    """
    Robust: DrawPage count can change while iterating (LO disposes removed shapes).
    Use enumeration if possible, otherwise a safe index loop.
    """
    m = {}

    # 1) Preferred: enumeration (stable even if count changes)
    try:
        enum = dp.createEnumeration()
        while enum.hasMoreElements():
            sh = enum.nextElement()
            nm = getattr(sh, "Name", "") or ""
            if nm.startswith(MARK_DESC + "_"):
                m[nm] = sh
        return m
    except Exception:
        pass

    # 2) Fallback: safe index loop that re-checks count
    i = 0
    while True:
        try:
            cnt = dp.getCount()
        except Exception:
            break
        if i >= cnt:
            break
        try:
            sh = dp.getByIndex(i)
        except Exception:
            # count changed or object disappeared -> stop safely
            break
        nm = getattr(sh, "Name", "") or ""
        if nm.startswith(MARK_DESC + "_"):
            m[nm] = sh
        i += 1

    return m
    

def clear_contents_keep_format(sheet, a1_range: str):
    """Clear values/strings/formulas only (keep formatting/styles)."""
    rng = sheet.getCellRangeByName(a1_range)
    flags = (
        uno.getConstantByName("com.sun.star.sheet.CellFlags.VALUE")   |
        uno.getConstantByName("com.sun.star.sheet.CellFlags.STRING")  |
        uno.getConstantByName("com.sun.star.sheet.CellFlags.FORMULA")
    )
    rng.clearContents(flags)

def clear_all_circle_quadrant_texts(sheet):
    """
    Clears ALL quadrant text-shapes for the puzzle circles:
    3 rows x max 6 circles x 4 quadrants.
    This prevents leftover letters when the number of circles decreases
    or words are shorter than before.
    """
    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    cleared = 0
    MAX_CIRCLES = 6  # because 11/12 letters -> 6 circles

    for row_index in range(3):
        for c in range(MAX_CIRCLES):
            idx_tag = f"r{row_index}_c{c}"
            for q in range(1, 5):
                nm = f"{MARK_DESC}_text_{idx_tag}_{q}"
                sh = name_map.get(nm)
                if sh is None:
                    continue
                try:
                    sh.String = " "   # space = visually empty, avoids some LO quirks with ""
                except Exception:
                    try:
                        sh.Text.setString(" ")
                    except Exception:
                        pass
                cleared += 1

    return cleared


def ensure_circle_count_matches_G3(sheet):
    """
    Ensures the drawn circle grid matches the number implied by G3.
    - If more circles exist than needed: delete the extra circle shapes per row.
    - If fewer circles exist than needed: redraw circles (calls draw_circles_from_G3()).
    Returns: desired_n
    """
    L = _desired_len_from_G3(sheet)
    desired_n = _num_circles_for_len(L)
    MAX_CIRCLES = 6

    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    def circle_exists(r, c):
        return f"{MARK_DESC}_circle_r{r}_c{c}" in name_map

    # Count existing circles per row (0..2)
    row_counts = []
    for r in range(3):
        cnt = 0
        for c in range(MAX_CIRCLES):
            if circle_exists(r, c):
                cnt += 1
        row_counts.append(cnt)

    # If we need more circles than currently exist -> easiest: redraw properly
    if any(cnt < desired_n for cnt in row_counts):
        # draw_circles_from_G3 should delete and recreate the correct amount
        draw_circles_from_G3()
        return desired_n

    # If there are too many circles -> delete the extras (and their quadrant texts/lines)
    # (Assumes naming: circle/vline/hline/text_... as in your current implementation)
    for r in range(3):
        for c in range(desired_n, MAX_CIRCLES):
            idx_tag = f"r{r}_c{c}"

            # remove in safe order: texts, lines, circle
            for q in range(1, 5):
                nm = f"{MARK_DESC}_text_{idx_tag}_{q}"
                sh = name_map.get(nm)
                if sh is not None:
                    try:
                        dp.remove(sh)
                    except Exception:
                        pass

            for nm in (f"{MARK_DESC}_vline_{idx_tag}", f"{MARK_DESC}_hline_{idx_tag}", f"{MARK_DESC}_circle_{idx_tag}"):
                sh = name_map.get(nm)
                if sh is not None:
                    try:
                        dp.remove(sh)
                    except Exception:
                        pass

    return desired_n


def delete_all_circles(*args):
    """
    Robust delete of ALL puzzle shapes (circles/lines/texts), even if older runs used
    different MARK_DESC or naming.
    Keeps form controls/buttons.
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)

    def is_control_shape(sh):
        try:
            return sh.supportsService("com.sun.star.drawing.ControlShape")
        except Exception:
            return False

    def should_delete(sh):
        if is_control_shape(sh):
            return False

        nm = (getattr(sh, "Name", "") or "")
        desc = (getattr(sh, "Description", "") or "")

        # 1) current marker
        if desc == MARK_DESC or nm.startswith(MARK_DESC + "_"):
            return True

        # 2) backwards compatible name patterns (older prefixes)
        up = nm.upper()
        if "KREIS" in up and any(k in nm for k in ("circle_", "vline_", "hline_", "text_")):
            return True

        # 3) any obvious puzzle parts (even if prefix changed)
        if any(k in nm for k in ("_circle_", "_vline_", "_hline_", "_text_")):
            return True

        return False

    removed = 0
    for i in range(dp.getCount() - 1, -1, -1):
        sh = dp.getByIndex(i)
        try:
            if should_delete(sh):
                dp.remove(sh)
                removed += 1
        except Exception:
            pass

    _debug_a1(sheet, f"Deleted puzzle shapes: {removed}")
    return True


def delete_circles_with_content_only(*args):
    """
    Deletes ONLY circles where any quadrant text is not empty.
    Prepared for later when you fill letters.
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)
    dp = _get_draw_page(sheet)
    name_map = _name_map_from_drawpage(dp)

    wl = _read_int(sheet, CELL_WORDLEN, DEFAULT_WORDLEN)
    n_circles = _num_circles_for_len(wl)

    to_remove = set()
    deleted_circles = 0

    for r_index in range(3):
        for c in range(n_circles):
            idx_tag = f"r{r_index}_c{c}"

            has_content = False
            for q in range(1, 5):
                nm = f"{MARK_DESC}_text_{idx_tag}_{q}"
                sh = name_map.get(nm)
                if sh is None:
                    continue
                try:
                    s = (sh.String or "").strip()
                except Exception:
                    s = ""
                if s and s != "":
                    has_content = True
                    break

            if has_content:
                deleted_circles += 1
                to_remove.add(f"{MARK_DESC}_circle_{idx_tag}")
                to_remove.add(f"{MARK_DESC}_vline_{idx_tag}")
                to_remove.add(f"{MARK_DESC}_hline_{idx_tag}")
                for q in range(1, 5):
                    to_remove.add(f"{MARK_DESC}_text_{idx_tag}_{q}")

    removed = 0
    for nm in list(to_remove):
        sh = name_map.get(nm)
        if sh is not None:
            try:
                dp.remove(sh)
                removed += 1
            except Exception:
                pass

    _debug_a1(sheet, f"Deleted circles-with-content: {deleted_circles} | shapes removed: {removed}")
    return True


# ============================================================
# 7) MAIN DRAW: circles anchored to right of column D
# ============================================================

def draw_circles_from_G3(*args):
    """
    Final circle drawing:
    - Calls ensure_initialized() first (row heights must be correct)
    - Reads word length from G3 (numeric)
    - Determines circle count per row:
        5-6 -> 3, 7-8 -> 4, 9-10 -> 5, 11-12 -> 6
      (implemented as ceil(word_len/2) clamped to [3..6])
    - Anchors circles to RIGHT edge of column D (= left edge of column E)
    - Uses FIXED step_x to prevent overlaps and to avoid changing column widths
    """
    doc = _get_doc()
    try:
        sheet = _get_sheet(doc)
    except Exception:
        _msgbox(doc, "Kreis-Worträtsel", f"Sheet '{SHEET_NAME}' not found.")
        return False

    t0 = time.perf_counter()

    
    # Init is mandatory
    if not ensure_initialized():
        return False

    word_len = _clamp_wordlen_to_5_12(sheet, writeback=True)
    _update_left_labels(sheet, word_len)

    # Compute number of circles per row
    n_circles = _num_circles_for_len(word_len)

    if n_circles < 3:
        n_circles = 3
    if n_circles > 6:
        n_circles = 6

    # Fresh draw
    delete_all_circles()
    dp = _get_draw_page(sheet)

    # Anchor X: right edge of column D = left edge of column E
    x0 = _col_left_x(sheet, ANCHOR_COL0)  # ANCHOR_COL0 = 4 (E)

    # FIXED horizontal step (prevents overlap and avoids column-width conflicts)
    step_x = CIRCLE_DIAMETER + cm(H_GAP_CM)

    # Y positions from row starts (4,7,10)
    y_positions = [_row_top_y(sheet, r) for r in CIRCLE_ROW_STARTS_1BASED]

    for r_index in range(3):
        y = y_positions[r_index]
        fill = ROW_COLORS[r_index]
        for c in range(n_circles):
            x = x0 + c * step_x
            idx_tag = f"r{r_index}_c{c}"
            _draw_circle_with_quadrants(doc, dp, x, y, CIRCLE_DIAMETER, fill, idx_tag)

    # Update left labels too (in case G3 changed)
    try:
        txt = f"Begriffe\nmit {word_len} Buchstaben"
        for (r1, r2) in LABEL_BLOCKS:
            rng = sheet.getCellRangeByPosition(1, r1 - 1, 1, r2 - 1)
            try:
                rng.merge(True)
            except Exception:
                pass
            sheet.getCellByPosition(1, r1 - 1).setString(txt)
            _set_center(rng)
            _set_bold(rng, pt=14)
    except Exception:
        pass
        
    clear_contents_keep_format(sheet, "Y4:Y12")

    _debug_a1(sheet, f"Kreise: 3 Reihen x {n_circles} Kreise (G3={word_len}), Anker=D rechts.")
    
    dt = time.perf_counter() - t0
    try:
        c = sheet.getCellRangeByName("J2")
        c.setString(f"Dauer: {dt:.1f} sec.".replace(".", ","))
        c.CharHeight = 10
    except Exception:
        pass

    return True


# ============================================================
# 9) EXPORTED MACROS
# ============================================================

g_exportedScripts = (
    ensure_initialized,
    set_left_labels_from_G3,
    draw_circles_from_G3,
    delete_all_circles,
    delete_circles_with_content_only,
    part2a_fill_random_from_wordlist,
)
