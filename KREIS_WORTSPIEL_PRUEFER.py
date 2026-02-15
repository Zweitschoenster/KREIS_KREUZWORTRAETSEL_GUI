# -*- coding: utf-8 -*-
"""
LibreOffice Calc Python Macro
Macro name : KREIS_WORTSPIEL – Word Checker
File       : kreis_wortspiel_pruefer.py
Sheet      : KREIS_WORTSPIEL

Purpose:
- Validate words for Kreis-Wortspiel
- Dictionary check (LibreOffice)
- Umlaut conversion for crossword logic (ÄÖÜß -> ae/oe/ue/ss)
- Length validation and result listing
- Store words (converted/upper) into length-specific lists + timestamp next to it
- One-time formatting + headers for all word blocks (5..12 + secondary 8 list)
"""

import uno
from datetime import datetime
from com.sun.star.uno import Exception as UnoException

# ============================================================
# CONFIG
# ============================================================

DEBUG = True
INIT_RUNNING = False
INITIALIZED = False

SHEET_NAME = "KREIS_WORTSPIEL"

CELL_HEADER_AA3 = "AA3"
CELL_LENGTH_AB3 = "AB3"
HEADER_ROW = 3
HEADER_START_COL = "AA"
HEADER_EMPTY_STREAK_STOP = 22
TS_HEADER_TEXT = "Aufgenommen am:"



WORDS_COL = "AB"     # AB4...
RESULT_COL = "AC"    # AC4...

START_ROW = 4
LAST_ROW  = 800

MIN_LEN = 5
MAX_LEN = 12

GERMAN_LOCALE = "de-DE"

UMLAUT_MAP = {
    "Ä": "ae", "Ö": "oe", "Ü": "ue", "ß": "ss",
    "ä": "ae", "ö": "oe", "ü": "ue",
}

# Block-/Format-Regeln
BASE_COL_LEN8 = "AG"     # 8er-Liste fix
BLOCK_WIDTH = 7          # Raster-Blockbreite (dein Layout)
WORD_COL_WIDTH_CM = 5.0  # Wortspalte Breite
META_COL_WIDTH_CM = 5.0  # 4 Nachbarspalten Breite
WORD_PT = 16.0           # Wortspalte Schriftgröße
META_PT = 12.0           # Nachbarspalten Schriftgröße
AB_INPUT_PT = 14.0       # AB Eingabe

# ============================================================
# LOGGING / UI
# ============================================================

def log(message: str) -> None:
    if DEBUG:
        print(f"[MACRO] {message}")

def show_messagebox(title: str, message: str) -> None:
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()

    window = doc.CurrentController.Frame.ContainerWindow
    toolkit = window.getToolkit()
    msgbox = toolkit.createMessageBox(window, 1, 1, title, message)
    msgbox.execute()

# ============================================================
# TIME
# ============================================================

def ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")

# ============================================================
# UNO HELPERS
# ============================================================

def get_context():
    return uno.getComponentContext()

def get_desktop(ctx):
    return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

def get_document(desktop):
    doc = desktop.getCurrentComponent()
    if not doc:
        raise RuntimeError("No document is open.")
    return doc

def assert_calc_document(doc) -> None:
    if not hasattr(doc, "Sheets"):
        raise RuntimeError("Current document is not a Calc spreadsheet.")

def get_sheet(doc, sheet_name: str):
    sheets = doc.Sheets
    if not sheets.hasByName(sheet_name):
        raise RuntimeError(f"Sheet not found: {sheet_name}")
    return sheets.getByName(sheet_name)

def get_cell(sheet, a1: str):
    return sheet.getCellRangeByName(a1)
    
def set_cursor_to_cell(doc, sheet, a1: str) -> None:
    """Setzt den Cursor / die aktive Zelle auf a1 (z.B. 'AB4')."""
    try:
        target = sheet.getCellRangeByName(a1)
        controller = doc.CurrentController
        controller.select(target)                 # markiert die Zelle
        controller.setSelection(target)           # extra robust
        controller.setActiveCell(target)          # setzt den Zellcursor
    except Exception:
        # wenn LO gerade 'busy' ist, einfach still ignorieren
        pass    

def commit_any_active_cell_edit(doc, sheet, safe_a1="A1") -> None:
    """
    Beendet eine laufende Zellbearbeitung (falls der Cursor gerade in einer Zelle steht),
    indem eine 'sichere' Zelle selektiert wird. Dadurch wird die Eingabe bestätigt.
    """
    try:
        ctrl = doc.CurrentController
        safe = sheet.getCellRangeByName(safe_a1)
        ctrl.select(safe)  # verlässt Editmodus / bestätigt Eingabe
        # optional: Fokus zurück aufs Sheet (manchmal hilfreich)
        try:
            ctrl.Frame.ContainerWindow.setFocus()
        except Exception:
            pass
    except Exception:
        pass

def commit_edit_by_parking_left_of_last_word(doc, sheet, words_col="AB", start_row=4) -> None:
    """
    Beendet eine laufende Zellbearbeitung (Cursor in Zelle) indem der Controller
    auf eine 'sichere' Zelle links neben dem letzten Wort in words_col springt.
    Standard: links neben AB ist AA.
    Fallback: AA<start_row> bzw. A1.
    """
    try:
        ctrl = doc.CurrentController

        # Letzte belegte Zeile in words_col finden
        last = get_last_row_with_content(sheet, words_col, start_row)

        # Zielzelle: links neben words_col (AB -> AA)
        left_col = shift_col(words_col, -1)  # AB -> AA

        if last < start_row:
            # Falls keine Wörter vorhanden: parke in AA<start_row>
            target_a1 = f"{left_col}{start_row}"
        else:
            target_a1 = f"{left_col}{last}"

        # Selektieren -> beendet Editmodus / committed Eingabe
        ctrl.select(sheet.getCellRangeByName(target_a1))

        # optional: Fokus zurück aufs Sheet (hilft manchmal)
        try:
            ctrl.Frame.ContainerWindow.setFocus()
        except Exception:
            pass

    except Exception:
        # harter Fallback
        try:
            doc.CurrentController.select(sheet.getCellRangeByName("A1"))
        except Exception:
            pass

# ============================================================
# ANZAHL WÖRTER von WORTSPALTE
# ============================================================

def count_non_empty_in_col(sheet, col_letter: str, start_row: int = START_ROW) -> int:
    """Count non-empty cells in a column from start_row down to last used row of that column."""
    last = get_last_row_with_content(sheet, col_letter, start_row)
    if last < start_row:
        return 0

    n = 0
    for r in range(start_row, last + 1):
        if get_cell(sheet, f"{col_letter}{r}").String.strip():
            n += 1
    return n


def write_count_header(cell, label: str, n: int) -> None:
    """
    Writes:
      <label>
      <n>
    and makes <n> bold.
    """
    text = cell.Text
    text.setString("")

    line1 = f"{label}\n"
    num_str = str(n)

    cursor = text.createTextCursor()
    text.insertString(cursor, line1, False)

    start = cursor.getStart()
    text.insertString(cursor, num_str, False)
    end = cursor.getEnd()

    num_cursor = text.createTextCursorByRange(start)
    num_cursor.gotoRange(end, True)
    num_cursor.CharWeight = 150.0  # bold

def update_word_counts_row2(sheet, target_len: int) -> None:
    # (Optional) AB2 Eingaben – wenn du das behalten willst:
    n_inputs = count_non_empty_in_col(sheet, WORDS_COL, START_ROW)
    write_count_header(get_cell(sheet, "AB2"), "Anzahl Eingaben:", n_inputs)

    # Hauptblöcke (Startspalten) für Längen 5..12
    for length in range(MIN_LEN, MAX_LEN + 1):
        list_col = get_list_start_col_for_length(length)
        n_words = count_non_empty_in_col(sheet, list_col, START_ROW)

        if n_words > 0:
            write_count_header(get_cell(sheet, f"{list_col}2"), "Anzahl Wörter:", n_words)
        else:
            get_cell(sheet, f"{list_col}2").String = ""   # leer lassen statt 0

    # Sekundärer 8er-Block (BI)
    sec8 = get_secondary_list_col_for_len8()
    n_words2 = count_non_empty_in_col(sheet, sec8, START_ROW)

    if n_words2 > 0:
        write_count_header(get_cell(sheet, f"{sec8}2"), "Anzahl Wörter:", n_words2)
    else:
        get_cell(sheet, f"{sec8}2").String = ""


# ============================================================
# HEADERS / BASIC FORMATS
# ============================================================ 
def set_headers(sheet) -> None:
    label = "Anzahl Buchstaben \n pro Wort"

    # ------------------------------------------------------------
    # 1) Letzte relevante Header-Spalte finden:
    #    Wir laufen in HEADER_ROW ab HEADER_START_COL nach rechts.
    #    Sobald HEADER_EMPTY_STREAK_STOP Header-Zellen hintereinander leer sind,
    #    brechen wir ab. Die letzte nicht-leere Headerzelle ist unser Ende.
    # ------------------------------------------------------------
    empty_streak = 0
    col_idx0 = col_to_index(HEADER_START_COL) - 1          # 0-basiert
    last_filled_col_idx0 = col_idx0 - 1                    # falls AA3 direkt leer wäre

    while True:
        col_letter = index_to_col(col_idx0 + 1)            # zurück nach "AA", "AB", ...
        txt = (get_cell(sheet, f"{col_letter}{HEADER_ROW}").String or "").strip()

        if txt == "":
            empty_streak += 1
            if empty_streak >= HEADER_EMPTY_STREAK_STOP:
                break
        else:
            empty_streak = 0
            last_filled_col_idx0 = col_idx0

        col_idx0 += 1

    # Wenn wir gar keine gefüllte Headerzelle gefunden haben, nichts tun
    if last_filled_col_idx0 < 0:
        return

    end_col_idx0 = last_filled_col_idx0

    # ------------------------------------------------------------
    # 2) Ab AB (= WORDS_COL) bis zur gefundenen Endspalte laufen
    #    und nur in Spalten mit Inhalt ab START_ROW den Header setzen.
    #    Überschriften, die schon existieren (z.B. "Wörter in AU ..."),
    #    werden NICHT überschrieben.
    # ------------------------------------------------------------
    start_col_idx0 = col_to_index(WORDS_COL) - 1  # AB -> 0-basiert

    # Falls das Ende links von AB liegt: nichts zu tun
    if end_col_idx0 < start_col_idx0:
        return

    for c0 in range(start_col_idx0, end_col_idx0 + 1):
        col_letter = index_to_col(c0 + 1)

        # Prüfen: gibt es ab Zeile 4 (START_ROW) irgendeinen Inhalt in dieser Spalte?
        last = get_last_row_with_content(sheet, col_letter, START_ROW)
        if last < START_ROW:
            continue

        # Header nur setzen, wenn die Headerzelle leer ist (nichts überschreiben!)
        hdr = get_cell(sheet, f"{col_letter}{HEADER_ROW}")
        if (hdr.String or "").strip() == "":
            hdr.String = label
            hdr.CharWeight = 150.0
            hdr.IsTextWrapped = True


def apply_basic_formats(sheet) -> None:
    aa3 = get_cell(sheet, CELL_HEADER_AA3)
    aa3.CharWeight = 150.0
    aa3.IsTextWrapped = True

    ab3 = get_cell(sheet, CELL_LENGTH_AB3)
    ab3.CharWeight = 150.0

    # AC: Ergebnis (links, kein Wrap)
    sheet.getCellRangeByName(f"{RESULT_COL}{START_ROW}:{RESULT_COL}{LAST_ROW}").IsTextWrapped = False
    sheet.getCellRangeByName(f"{RESULT_COL}{START_ROW}:{RESULT_COL}{LAST_ROW}").HoriJustify = 0  # left

    # AB: Eingabeformat ab Zeile 4
    rng_ab = sheet.getCellRangeByName(f"AB{START_ROW}:AB{LAST_ROW}")
    rng_ab.CharHeight = AB_INPUT_PT
    rng_ab.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
    rng_ab.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
    
def set_ts_header(sheet, word_start_col: str) -> None:
    """
    Setzt den Timestamp-Header in Zeile 3 rechts neben der Wortspalte
    (z.B. AG -> AH3), aber überschreibt nichts, wenn schon Text da ist.
    """
    ts_hdr = get_cell(sheet, f"{shift_col(word_start_col, 1)}{HEADER_ROW}")
    """
    Das heißt: nicht überschreiben
    ts_hdr = get_cell(sheet, f"{shift_col(word_start_col, 1)}{HEADER_ROW}")
    if (ts_hdr.String or "").strip() != "":
        return
    """
    # Wenn da irgendwas anderes steht: überschreiben
    current = (ts_hdr.String or "").strip()
    if current == TS_HEADER_TEXT.strip():
        return  # ist schon korrekt

    ts_hdr.String = TS_HEADER_TEXT
    ts_hdr.CharWeight = 150.0
    ts_hdr.IsTextWrapped = True
    ts_hdr.HoriJustify = 2
    ts_hdr.VertJustify = 1
    ts_hdr.CharHeight = 12.0


# ============================================================
# VALIDATION / INPUT
# ============================================================

def ensure_default_length(sheet) -> None:
    ab3 = get_cell(sheet, CELL_LENGTH_AB3)
    if ab3.String.strip() == "" and ab3.Value == 0.0:
        ab3.Value = 8.0
        log("AB3 was empty -> set to 8.")

def read_length_setting(sheet) -> int:
    ab3 = get_cell(sheet, CELL_LENGTH_AB3)

    value = int(ab3.Value) if ab3.Value != 0.0 else 0
    if value == 0:
        t = ab3.String.strip()
        if t.isdigit():
            value = int(t)

    if not (MIN_LEN <= value <= MAX_LEN):
        show_messagebox(
            "Ungültige Zahl",
            f"AB3 muss zwischen {MIN_LEN} und {MAX_LEN} liegen.\n"
            f"Aktuell: {value}\n\nMakro wird abgebrochen."
        )
        raise RuntimeError("Invalid length setting in AB3.")

    return value

# ============================================================
# RANGE / ROW HELPERS
# ============================================================

def get_last_row_with_content(sheet, col_letter: str, start_row: int) -> int:
    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)
    cursor.gotoEndOfUsedArea(True)
    used_end_row = cursor.RangeAddress.EndRow + 1  # 1-based

    if used_end_row < start_row:
        return start_row - 1

    for r in range(used_end_row, start_row - 1, -1):
        if get_cell(sheet, f"{col_letter}{r}").String.strip() != "":
            return r
    return start_row - 1

def find_first_empty_row(sheet, col_letter: str, start_row: int) -> int:
    last = get_last_row_with_content(sheet, col_letter, start_row)
    if last < start_row:
        return start_row
    return last + 1

# ============================================================
# SPELLCHECK (LangID statt Locale-Struct)
# ============================================================

def _locale_to_langid(locale_str: str) -> int:
    lang_map = {
        "de-DE": 1031,
        "de-AT": 3079,
        "de-CH": 2055,
        "de":    1031,
    }
    loc = (locale_str or "de-DE").replace("_", "-").strip()
    parts = loc.split("-")
    if len(parts) >= 2:
        key = f"{parts[0].lower()}-{parts[1].upper()}"
    else:
        key = parts[0].lower() if parts else "de-DE"
    return lang_map.get(key, 1031)

def get_spellchecker(doc):
    try:
        try:
            ctx = doc.getContext()
        except Exception:
            ctx = XSCRIPTCONTEXT.getComponentContext()

        lsm = None
        try:
            lsm = ctx.getValueByName("/singletons/com.sun.star.linguistic2.LinguServiceManager")
        except Exception:
            lsm = None

        if not lsm:
            smgr = ctx.ServiceManager
            lsm = smgr.createInstanceWithContext(
                "com.sun.star.linguistic2.LinguServiceManager", ctx
            )

        return lsm.getSpellChecker() if lsm else None

    except Exception as e:
        log(f"Spellchecker not available: {e}")
        return None

def is_word_in_dictionary(speller, word: str, locale_str: str = "de-DE") -> bool:
    if speller is None or not word:
        return False
    try:
        nLanguage = _locale_to_langid(locale_str)
        return bool(speller.isValid(word, nLanguage, ()))
    except Exception as e:
        log(f"Spellcheck error for '{word.lower()}': {e}")
        return False

def get_nearest_suggestion(speller, word: str, locale_str: str = "de-DE") -> str:
    if speller is None or not word:
        return ""
    try:
        nLanguage = _locale_to_langid(locale_str)
        props = ()
        if speller.isValid(word, nLanguage, props):
            return ""
        alts = speller.spell(word, nLanguage, props)
        if alts is None:
            return ""
        suggestions = alts.getAlternatives() or ()
        return suggestions[0] if suggestions else ""
    except Exception as e:
        log(f"Suggestion error for '{word.lower()}': {e}")
        return ""

# ============================================================
# TEXT / UMLEAUT
# ============================================================

def clear_cell(cell) -> None:
    cell.String = ""

def convert_umlauts_for_crossword(word: str) -> str:
    return "".join(UMLAUT_MAP.get(ch, ch) for ch in word)

# ============================================================
# COLUMN MATH
# ============================================================

def cm_to_100mm(cm: float) -> int:
    return int(cm * 1000)

def col_to_index(col: str) -> int:
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx

def index_to_col(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(ord('A') + r) + s
    return s

def shift_col(col: str, offset: int) -> str:
    return index_to_col(col_to_index(col) + offset)

# ============================================================
# LIST BLOCK MAPPING
# ============================================================

def get_list_start_col_for_length(length: int) -> str:
    """
    8 -> AG (fix)
    5 -> AG + 7
    6 -> AG + 14  (AU)
    7 -> AG + 21
    9 -> AG + 28
    ...
    """
    if length == 8:
        return BASE_COL_LEN8
    if length in (5, 6, 7, 9, 10, 11, 12):
        offset_blocks = (length - 5)  # 5->0, 6->1, 7->2, 9->4, ...
        return shift_col(BASE_COL_LEN8, 7 + offset_blocks * BLOCK_WIDTH)
    raise RuntimeError(f"Unsupported length: {length}")

def get_secondary_list_col_for_len8() -> str:
    # BI bei deinem Raster
    return shift_col(BASE_COL_LEN8, 7 + (8 - 5) * BLOCK_WIDTH)

# ============================================================
# COMPACT LISTS
# ============================================================

def compact_two_columns(sheet, col_a: str, col_b: str, start_row: int) -> None:
    last = get_last_row_with_content(sheet, col_a, start_row)
    if last < start_row:
        return

    items = []
    for r in range(start_row, last + 1):
        a_cell = get_cell(sheet, f"{col_a}{r}")
        b_cell = get_cell(sheet, f"{col_b}{r}")
        a = (a_cell.String or "").strip()
        b = b_cell.String if hasattr(b_cell, "String") else ""
        if a != "":
            items.append((a, b))

    range_a = sheet.getCellRangeByName(f"{col_a}{start_row}:{col_a}{last}")
    range_b = sheet.getCellRangeByName(f"{col_b}{start_row}:{col_b}{last}")
    range_a.clearContents(23)
    range_b.clearContents(23)

    for i, (a, b) in enumerate(items):
        r = start_row + i
        get_cell(sheet, f"{col_a}{r}").String = a
        get_cell(sheet, f"{col_b}{r}").String = b

def build_existing_list_map_for_col(sheet, col_letter: str) -> dict:
    last = get_last_row_with_content(sheet, col_letter, START_ROW)
    m = {}
    if last < START_ROW:
        return m
    for r in range(START_ROW, last + 1):
        v = get_cell(sheet, f"{col_letter}{r}").String.strip()
        if v != "":
            m[v.upper()] = r
    return m

# ============================================================
# BLOCK HEADERS + FORMATTING FOR ALL WORD COLUMNS
# ============================================================

def write_list_header_with_bold_number(cell, n: int, col_label: str) -> None:
    """
    Header:
      Wörter in <COL>
      <N> Buchstaben
    (N fett)
    """
    text = cell.Text
    text.setString("")

    line1 = f"Wörter in {col_label}\n"
    num_str = str(n)
    tail = " Buchstaben"

    cursor = text.createTextCursor()
    text.insertString(cursor, line1, False)

    start = cursor.getStart()
    text.insertString(cursor, num_str, False)
    end = cursor.getEnd()

    num_cursor = text.createTextCursorByRange(start)
    num_cursor.gotoRange(end, True)
    num_cursor.CharWeight = 150.0

    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    text.insertString(cursor, tail, False)

    cell.IsTextWrapped = True
    cell.HoriJustify = 2
    cell.VertJustify = 1
    cell.CharHeight = 14.0

def _set_col_width(sheet, col_letter: str, cm_width: float):
    col = sheet.getColumns().getByName(col_letter)
    col.Width = cm_to_100mm(cm_width)

def _format_col_range(sheet, col_letter: str, pt: float, hori: str, vert: str, width_cm: float):
    _set_col_width(sheet, col_letter, width_cm)
    rng = sheet.getCellRangeByName(f"{col_letter}{START_ROW}:{col_letter}{LAST_ROW}")
    rng.CharHeight = float(pt)
    rng.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", hori)
    rng.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", vert)

def format_all_word_blocks(sheet):
    """
    Formatiert ALLE Wortspalten-Blöcke:
    - Startspalte (Wort): 5cm, 16pt, rechts, vertikal mittig
    - die nächsten 4 Spalten: je 5cm, 12pt, rechts, vertikal mittig
    Zusätzlich: Header in Zeile 3 für alle Längen 5..12 und optional BI (8 sekundär)
    """
    # Headerzeile für den Bereich optisch zentrieren (AA3:?? optional minimal)
    hdr = sheet.getCellRangeByName("AA3:AC3")
    hdr.HoriJustify = 2
    hdr.VertJustify = 1
    hdr.IsTextWrapped = True

    # Alle Längen 5..12
    for length in range(5, 13):
        start_col = get_list_start_col_for_length(length)

        # Header in <start_col>3 setzen (damit AU3 sicher gesetzt ist)
        write_list_header_with_bold_number(get_cell(sheet, f"{start_col}3"), length, start_col)
        set_ts_header(sheet, start_col)

        # Wortspalte
        _format_col_range(sheet, start_col, WORD_PT, "CENTER", "CENTER", WORD_COL_WIDTH_CM)

        # 4 Nachbarspalten (Timestamp + 3 weitere)
        for off in (1, 2, 3, 4):
            c = shift_col(start_col, off)
            _format_col_range(sheet, c, META_PT, "CENTER", "CENTER", META_COL_WIDTH_CM)

    # Sekundärblock für Länge 8 (BI) – optional, aber sauber inkl. Header/Format
    sec8 = get_secondary_list_col_for_len8()
    write_list_header_with_bold_number(get_cell(sheet, f"{sec8}3"), 8, sec8)
    set_ts_header(sheet, sec8)
    
    # Wortspalte
    _format_col_range(sheet, sec8, WORD_PT, "CENTER", "CENTER", WORD_COL_WIDTH_CM)

    # 4 Nachbarspalten (Timestamp + 3 weitere)
    for off in (1, 2, 3, 4):
        c = shift_col(sec8, off)
        _format_col_range(sheet, c, META_PT, "CENTER", "CENTER", META_COL_WIDTH_CM)

def format_count_cells_row2(sheet, cols: list[str]) -> None:
    for c in cols:
        rng = sheet.getCellRangeByName(f"{c}2")
        rng.HoriJustify = 2  # center (oder 3 für rechts, je nach LO-Wert; bei dir nutzt du 2=center)
        rng.VertJustify = 1  # middle
        rng.CharHeight = 12.0
        rng.IsTextWrapped = True

# ============================================================
# CORE PROCESS
# ============================================================

def process_words(sheet, doc, target_len: int) -> None:
    last_row = get_last_row_with_content(sheet, WORDS_COL, START_ROW)
    if last_row < START_ROW:
        show_messagebox("Info", "Keine Wörter in AB ab Zeile 4 gefunden.")
        return

    speller = get_spellchecker(doc)

    # Decide which list block to use based on target length
    list_col = get_list_start_col_for_length(target_len)
    if not list_col:
        raise RuntimeError(f"List column mapping failed for length={target_len}")

    ts_col = shift_col(list_col, 1)  # timestamp is next column

    # Next free row in the chosen list column (append mode)
    next_row = find_first_empty_row(sheet, list_col, START_ROW)

    # Build a map of existing entries in the chosen list column
    existing_map = build_existing_list_map_for_col(sheet, list_col)

    # For length 8: additionally write into secondary 8-block
    list_col2 = None
    ts_col2 = None
    next_row2 = None
    existing_map2 = {}

    if target_len == 8:
        list_col2 = get_secondary_list_col_for_len8()
        if not list_col2:
            raise RuntimeError("Secondary list column mapping failed for length=8")
        ts_col2 = shift_col(list_col2, 1)
        next_row2 = find_first_empty_row(sheet, list_col2, START_ROW)
        existing_map2 = build_existing_list_map_for_col(sheet, list_col2)

    for r in range(START_ROW, last_row + 1):
        word_cell = get_cell(sheet, f"{WORDS_COL}{r}")
        result_cell = get_cell(sheet, f"{RESULT_COL}{r}")

        original = (word_cell.String or "").strip()
        if original == "":
            continue

        # Clear AC first (we write only when needed)
        clear_cell(result_cell)

        # 1) Convert umlauts first (needed for duplicate check and crossword length)
        converted = convert_umlauts_for_crossword(original)
        did_convert = (converted != original)  # <-- PATCH: only true if something changed
        converted_upper = converted.upper()

        # 2) Duplicate check in primary list column
        if converted_upper in existing_map:
            row_found = existing_map[converted_upper]
            result_cell.String = f"Wort in {list_col} {row_found}"
            continue

        # Duplicate check in secondary list (only for len=8)
        if target_len == 8 and converted_upper in existing_map2:
            row_found = existing_map2[converted_upper]
            result_cell.String = f"Wort in {list_col2} {row_found}"
            continue

        # 3) Dictionary check (original word) after duplicate check
        if not is_word_in_dictionary(speller, original, GERMAN_LOCALE):
            sug = get_nearest_suggestion(speller, original, GERMAN_LOCALE)
            if sug:
                result_cell.String = f"Wort nicht im LO-Wörterbuch – Vorschlag: {sug}"
            else:
                result_cell.String = "Wort nicht im LO-Wörterbuch"
            continue

        # 4) Length check on converted word
        conv_len = len(converted)
        if conv_len != target_len:
            # <-- PATCH: suffix only if an umlaut conversion actually occurred
            suffix = " (nach Umlaut-Umwandlung)" if did_convert else ""
            result_cell.String = f"Wort hat {conv_len} Buchstaben{suffix}"
            continue

        # 5) If OK: write into list + timestamp
        stamp = ts()

        # Primary list write
        get_cell(sheet, f"{list_col}{next_row}").String = converted_upper
        get_cell(sheet, f"{ts_col}{next_row}").String = stamp
        existing_map[converted_upper] = next_row
        next_row += 1

        # Secondary list write (only for len=8)
        if target_len == 8:
            get_cell(sheet, f"{list_col2}{next_row2}").String = converted_upper
            get_cell(sheet, f"{ts_col2}{next_row2}").String = stamp
            existing_map2[converted_upper] = next_row2
            next_row2 += 1

    # Compact list + timestamp columns (remove gaps)
    compact_two_columns(sheet, list_col, ts_col, START_ROW)
    if target_len == 8:
        compact_two_columns(sheet, list_col2, ts_col2, START_ROW)
    # -> NEU: Counts aktualisieren
    update_word_counts_row2(sheet, target_len)
    
# ============================================================
# WORKFLOW
# ============================================================

def run_workflow(doc) -> None:
    sheet = get_sheet(doc, SHEET_NAME)

    set_headers(sheet)
    apply_basic_formats(sheet)

    ensure_default_length(sheet)
    target_len = read_length_setting(sheet)

    process_words(sheet, doc, target_len)

# ============================================================
# INIT (ONE-TIME)
# ============================================================

def is_initialized() -> bool:
    return INITIALIZED

def init_kreis_wortspiel(*args):
    global INIT_RUNNING, INITIALIZED

    INIT_RUNNING = True
    try:
        ctx = get_context()
        desktop = get_desktop(ctx)
        doc = get_document(desktop)
        sheet = get_sheet(doc, SHEET_NAME)

        set_headers(sheet)
        apply_basic_formats(sheet)

        # >>> alle Wortspalten-Header + Formatierung (inkl. AU3)
        format_all_word_blocks(sheet)
        
        # format_count_cells_row2(sheet, ["AB", "AG", "AH", "AN", "AO", "AU", "AV"]) diese Zeile wurde ersetzt durch den unteren Block und ist dadurch flexibler geworden
        
        cols = ["AB"]  # AB2 Eingaben
        for length in range(MIN_LEN, MAX_LEN + 1):
            cols.append(get_list_start_col_for_length(length))  # AG/AN/AU/...
        cols.append(get_secondary_list_col_for_len8())          # BI-Block       
        format_count_cells_row2(sheet, cols)
 
        log("format_count_cells_row2 -> Spaltenliste cols = " + ", ".join(cols))
        
        format_count_cells_row2(sheet, cols)

        # Counts auch beim Init aktualisieren (wenn Listen schon gefüllt sind)
        ensure_default_length(sheet)
        target_len = read_length_setting(sheet)   # braucht update_word_counts_row2 wegen AB2 (Eingaben)
        update_word_counts_row2(sheet, target_len)

        INITIALIZED = True
    finally:
        INIT_RUNNING = False

# ============================================================
# CLEAR INPUT & RESULT RANGE (AB4:AC500)
# ============================================================        
        
def clear_input_and_results_range(*args):
    """
    Löscht Inhalte in AB4:AC500 (Eingaben + Ergebnis),
    OHNE Formatierung zu verändern.
    """
    ctx = get_context()
    desktop = get_desktop(ctx)
    doc = get_document(desktop)
    assert_calc_document(doc)
    sheet = get_sheet(doc, SHEET_NAME)

    # Falls gerade eine Zelle im Editmodus ist: erst committen
    commit_edit_by_parking_left_of_last_word(doc, sheet, words_col=WORDS_COL, start_row=START_ROW)

    rng = sheet.getCellRangeByName("AB4:AC500")

    # 23 = löscht Inhalte (Value, String, Formel, Notizen, etc.), lässt Formatierung stehen
    rng.clearContents(23)
    sheet.getCellRangeByName("AB2").clearContents(23)

    # optional: Cursor wieder auf AB4 setzen
    set_cursor_to_cell(doc, sheet, "AB4")
        

# ============================================================
# ENTRY POINT
# ============================================================

def check_words_for_kreis_wordgame(*args):
    if not is_initialized():
        init_kreis_wortspiel()

    try:
        ctx = get_context()
        desktop = get_desktop(ctx)
        doc = get_document(desktop)
        assert_calc_document(doc)
        sheet = get_sheet(doc, SHEET_NAME)

        # >>> WICHTIG: Editmodus verlassen / Cursor aus aktiver Zelle raus
        # commit_any_active_cell_edit(doc, sheet, "AB2")
        commit_edit_by_parking_left_of_last_word(doc, sheet, words_col=WORDS_COL, start_row=START_ROW)

        
        log("Macro started.")
        run_workflow(doc)
        
        # optional am Ende Cursor wohin du willst
        # set_cursor_to_cell(doc, sheet, "AB2")
        
        log("Macro finished.")

    except UnoException as e:
        log(f"UNO Error: {e}")
        show_messagebox("Macro Error (UNO)", str(e))

    except Exception as e:
        log(f"Error: {e}")
        if "Invalid length setting" not in str(e):
            show_messagebox("Macro Error", str(e))

g_exportedScripts = (
    init_kreis_wortspiel,
    check_words_for_kreis_wordgame,
    clear_input_and_results_range,
)
