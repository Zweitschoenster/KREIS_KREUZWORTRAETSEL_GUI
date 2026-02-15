# -*- coding: utf-8 -*-
"""
Kreis_Wortraetsel_Import_V2.py  (single-file, no debug macros)

Task:
- Read candidate words from AA starting at row 4 (AA4..)
- Clean: keep only letters and '-' (hyphen), drop everything else, rebuild word
- Spellcheck (German dictionaries): de-DE / de-CH / de-AT
  using DIRECT service: com.sun.star.linguistic2.SpellChecker
  IMPORTANT: always pass () as PropertyValue sequence (avoids Type-17 issues)
- If NOT in dictionary: write AB "Nicht im LO Wörterbuch; Vorschläge: ..."
  (1–2 suggestions if available), keep AA unchanged.
- If IN dictionary:
  - Convert for crossword: ÄÖÜ -> AE/OE/UE, ß -> SS
  - Uppercase
  - Remove everything not A-Z (hyphen removed here)
  - Length AFTER conversion must be 5..12
    - too short/too long => AB message, keep AA unchanged
  - Determine target word column by length blocks starting at AG (step 7 columns):
      5->AG, 6->AN, 7->AU, 8->BB, 9->BI, 10->BP, 11->BW, 12->CD
  - Timestamp is written DIRECTLY RIGHT of the word column (+1)
  - If word already exists in that column:
      AB: "Wort vorhanden in {COL} {ROW}"
      clear AA (word has a place)
  - Else insert into first empty row in that column, set timestamp at +1 column,
      AB: "Neu aufgenommen in {COL} {ROW}"
      clear AA
- After run: compact each wordlist column (and its timestamp column) upward (no sorting).


Änderungen:
1) AA wird NUR geleert, wenn ein Wort NEU aufgenommen wurde.
   Wenn Wort schon vorhanden oder abgelehnt => bleibt in AA stehen.
2) Spalte AB wird auf 12 cm Breite gesetzt.
"""

import uno
from datetime import datetime
from com.sun.star.lang import Locale
from com.sun.star.beans import PropertyValue


# ============================================================
# 1) CONFIG
# ============================================================

SHEET_NAME = "KREIS_WORTRAETSEL"

# Candidate area
CAND_COL_WORD0   = 26   # AA
CAND_COL_STATUS0 = 27   # AB
CAND_START_ROW_1BASED = 4
CAND_MAX_ROWS = 500

# AB column width
AB_WIDTH_CM = 12.0

# Wordlist blocks
WORDLEN_MIN = 5
WORDLEN_MAX = 12

WORDLIST_BASE_COL0 = 32        # AG (0-based)
WORDLIST_BLOCK_STEP = 7        # 7 columns per length block
WORDLIST_START_ROW_1BASED = 4  # list starts row 4
WORDLIST_MAX_ROWS = 500

# Timestamp columns relative to word column
TS_IMPORT_OFFSET = 1      # direkt rechts neben dem Wort
TS_USED_OFFSET   = 2      # zweite Spalte rechts (für 30-Tage-Logik)
TIMESTAMP_COL_OFFSET = 2

# Optional header row
HEADER_ROW_1BASED = 3
SYNC_HEADERS = True


# Locales to try
LOCALES_DE = (("de", "DE"), ("de", "CH"), ("de", "AT"))


# ============================================================
# 2) UNO HELPERS
# ============================================================

def _get_doc():
    return XSCRIPTCONTEXT.getDocument()

def _get_sheet(doc):
    return doc.Sheets.getByName(SHEET_NAME)

def _msgbox(title, message):
    doc = _get_doc()
    try:
        parent = doc.CurrentController.Frame.ContainerWindow
        toolkit = parent.getToolkit()
        buttons = uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK")
        box_type = uno.Enum("com.sun.star.awt.MessageBoxType", "INFOBOX")
        box = toolkit.createMessageBox(parent, box_type, buttons, title, message)
        box.execute()
    except Exception:
        pass

def _now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M")

def _col_index_to_letters(col0: int) -> str:
    n = col0 + 1
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _cm_to_u100mm(v_cm: float) -> int:
    # 1 cm = 1000 * (1/100 mm)
    return int(round(v_cm * 1000))

def _set_col_width(sheet, col0: int, width_u100mm: int):
    try:
        c = sheet.Columns.getByIndex(col0)
        try:
            c.OptimalWidth = False
        except Exception:
            pass
        c.Width = int(width_u100mm)
    except Exception:
        pass

def _pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

def _goto_cell(doc, sheet_name: str, a1: str):
    # a1 like "C4"
    col = "".join(ch for ch in a1 if ch.isalpha()).upper()
    row = "".join(ch for ch in a1 if ch.isdigit())
    target = f"${sheet_name}.$%s$%s" % (col, row)

    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    dh = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    frame = doc.CurrentController.Frame
    dh.executeDispatch(frame, ".uno:GoToCell", "", 0, (_pv("ToPoint", target),))
    
    
# ============================================================
# 3) WORDLIST VISIBILITY (hide/show AG..CD blocks)
# ============================================================

def _wordlist_visible(sheet, visible: bool):
    # hide/show all wordlist blocks (5..12), each block = 7 columns
    start = WORDLIST_BASE_COL0
    last_start = WORDLIST_BASE_COL0 + (WORDLEN_MAX - WORDLEN_MIN) * WORDLIST_BLOCK_STEP
    end = last_start + (WORDLIST_BLOCK_STEP - 1)  # +6
    for col in range(start, end + 1):
        try:
            sheet.Columns.getByIndex(col).IsVisible = bool(visible)
        except Exception:
            pass

def hide_wordlists(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    _wordlist_visible(sheet, False)
    return True

def show_wordlists(*args):
    doc = _get_doc()
    sheet = _get_sheet(doc)
    _wordlist_visible(sheet, True)
    return True
    
# ============================================================
# 4) SPELLCHECK (DIRECT SERVICE, TYPE-17 SAFE)
# ============================================================

def _get_spellchecker_direct():
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    return smgr.createInstanceWithContext("com.sun.star.linguistic2.SpellChecker", ctx)

def _iter_locales():
    for lang, country in LOCALES_DE:
        loc = Locale()
        loc.Language = lang
        loc.Country = country
        yield loc

def _spell_is_valid_any(sp, word: str) -> bool:
    w = (word or "").strip()
    if not w:
        return False
    variants = (w, w.lower(), w.capitalize())
    for loc in _iter_locales():
        for v in variants:
            try:
                if sp.isValid(v, loc, ()):   # IMPORTANT: () not PropertyValue[]
                    return True
            except Exception:
                pass
    return False

def _spell_suggestions_any(sp, word: str, max_n=2):
    w = (word or "").strip()
    if not w:
        return []

    for loc in _iter_locales():
        try:
            alts = sp.spell(w, loc, ())  # IMPORTANT: ()
            if not alts:
                continue
            try:
                arr = list(alts.getAlternatives() or [])
            except Exception:
                try:
                    arr = list(getattr(alts, "Alternatives", []) or [])
                except Exception:
                    arr = []

            out = []
            seen = set()
            for s in arr:
                s = (s or "").strip()
                if not s:
                    continue
                k = s.lower()
                if k in seen:
                    continue
                seen.add(k)
                out.append(s)
                if len(out) >= max_n:
                    break
            if out:
                return out
        except Exception:
            pass

    return []


# ============================================================
# 5) NORMALIZATION
# ============================================================

def _clean_candidate_keep_hyphen(raw: str) -> str:
    s = (raw or "").strip()
    out = []
    for ch in s:
        if ch.isalpha() or ch == "-":
            out.append(ch)
    return "".join(out)

def _normalize_crossword(raw: str) -> str:
    s = (raw or "").strip().upper()
    s = (s.replace("Ä", "AE")
           .replace("Ö", "OE")
           .replace("Ü", "UE")
           .replace("ß", "SS")
           .replace("ẞ", "SS"))
    s = "".join(ch for ch in s if "A" <= ch <= "Z")
    return s


# ============================================================
# 6) WORDLIST MAPPING / FIND / INSERT / COMPACT
# ============================================================

def _word_col0_for_len(L: int) -> int:
    return WORDLIST_BASE_COL0 + (L - WORDLEN_MIN) * WORDLIST_BLOCK_STEP

def _find_word_in_column(sheet, word_col0: int, target: str):
    start0 = WORDLIST_START_ROW_1BASED - 1
    end0 = start0 + WORDLIST_MAX_ROWS - 1
    t = (target or "").strip().upper()
    for r0 in range(start0, end0 + 1):
        w = (sheet.getCellByPosition(word_col0, r0).getString() or "").strip().upper()
        if not w:
            continue
        if w == t:
            return r0 + 1  # 1-based
    return None

def _append_first_empty(sheet, word_col0: int, w: str, ts_import: str):
    start0 = WORDLIST_START_ROW_1BASED - 1
    end0 = start0 + WORDLIST_MAX_ROWS - 1

    for r0 in range(start0, end0 + 1):
        c = sheet.getCellByPosition(word_col0, r0)
        if not (c.getString() or "").strip():
            c.setString(w)
            sheet.getCellByPosition(word_col0 + TS_IMPORT_OFFSET, r0).setString(ts_import)
            sheet.getCellByPosition(word_col0 + TS_USED_OFFSET, r0).setString("")  # wird erst beim Puzzle gesetzt
            return r0 + 1
    return None


def _compact_word_and_meta(sheet, word_col0: int, col_offsets: tuple):
    """
    Kompaktiert WORD + Meta-Spalten gemeinsam.
    col_offsets z.B. (TS_IMPORT_OFFSET, TS_USED_OFFSET)
    """
    start0 = WORDLIST_START_ROW_1BASED - 1
    end0 = start0 + WORDLIST_MAX_ROWS - 1

    rows = []
    for r0 in range(start0, end0 + 1):
        w = (sheet.getCellByPosition(word_col0, r0).getString() or "").strip()
        if not w:
            continue
        meta = []
        for off in col_offsets:
            meta.append((sheet.getCellByPosition(word_col0 + off, r0).getString() or "").strip())
        rows.append((w, meta))

    write = start0
    for w, meta in rows:
        sheet.getCellByPosition(word_col0, write).setString(w)
        for i, off in enumerate(col_offsets):
            sheet.getCellByPosition(word_col0 + off, write).setString(meta[i])
        write += 1

    for r0 in range(write, end0 + 1):
        sheet.getCellByPosition(word_col0, r0).setString("")
        for off in col_offsets:
            sheet.getCellByPosition(word_col0 + off, r0).setString("")

def _compact_all_wordlists(sheet):
    """
    Kompaktiert ALLE Wortlisten-Spalten (Wort + Timestamp) ab WORDLIST_START_ROW_1BASED.
    Unabhängig von _compact_word_and_ts (falls die fehlt/kaputt ist).
    """
    start0 = WORDLIST_START_ROW_1BASED - 1
    end0 = start0 + WORDLIST_MAX_ROWS - 1

    for L in range(WORDLEN_MIN, WORDLEN_MAX + 1):
        word_col0 = _word_col0_for_len(L)
        ts_col0   = word_col0 + TIMESTAMP_COL_OFFSET

        # 1) sammeln
        pairs = []
        for r0 in range(start0, end0 + 1):
            w = (sheet.getCellByPosition(word_col0, r0).getString() or "").strip()
            if not w:
                continue
            ts = (sheet.getCellByPosition(ts_col0, r0).getString() or "").strip()
            pairs.append((w, ts))

        # 2) kompakt zurückschreiben
        write0 = start0
        for w, ts in pairs:
            sheet.getCellByPosition(word_col0, write0).setString(w)
            sheet.getCellByPosition(ts_col0, write0).setString(ts)
            write0 += 1

        # 3) Rest leeren
        for r0 in range(write0, end0 + 1):
            sheet.getCellByPosition(word_col0, r0).setString("")
            sheet.getCellByPosition(ts_col0, r0).setString("")


def _sync_headers(sheet):
    if not SYNC_HEADERS:
        return
    try:
        r0 = HEADER_ROW_1BASED - 1
        for L in range(WORDLEN_MIN, WORDLEN_MAX + 1):
            col0 = _word_col0_for_len(L)
            colA1 = _col_index_to_letters(col0)
            want = f"Wörter in {colA1}\n{L} Buchstaben"
            cell = sheet.getCellByPosition(col0, r0)
            cur = (cell.getString() or "").strip()
            if cur != want:
                cell.setString(want)
    except Exception:
        pass

def _count_nonempty_in_col(sheet, col0: int, start_row_1based: int, max_rows: int) -> int:
    start0 = start_row_1based - 1
    end0 = start0 + max_rows - 1
    n = 0
    for r0 in range(start0, end0 + 1):
        s = (sheet.getCellByPosition(col0, r0).getString() or "").strip()
        if s:
            n += 1
    return n
        
def _write_count_cell(sheet, col0: int, row_1based: int, n: int):
    r0 = row_1based - 1
    cell = sheet.getCellByPosition(col0, r0)

    cell.setString(f"Anzahl Wörter:\n{n}")

    cell.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
    cell.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
    cell.IsTextWrapped = True
    cell.CharHeight = 11.0

def _update_all_wordlist_counts(sheet):
    """
    Schreibt pro Wortspalte (AG..CD) die Anzahl in Zeile 2.
    Gezählt wird ab Zeile 4 in der jeweiligen Wortspalte.
    """
    for L in range(WORDLEN_MIN, WORDLEN_MAX + 1):
        word_col0 = _word_col0_for_len(L)
        cnt = _count_nonempty_in_col(sheet, word_col0, WORDLIST_START_ROW_1BASED, WORDLIST_MAX_ROWS)
        _write_count_cell(sheet, word_col0, 2, cnt)

def _update_candidate_count_in_AB3(sheet):
    """
    AA (Kandidaten) zählen ab Zeile 4 und in AB3 schreiben.
    """
    cnt = _count_nonempty_in_col(sheet, CAND_COL_WORD0, CAND_START_ROW_1BASED, CAND_MAX_ROWS)
    # AB3 = Spalte AB (27) Zeile 3
    cell = sheet.getCellByPosition(CAND_COL_STATUS0, 2)  # row0=2 => row3
    cell.setString(f"Anzahl der Wörter:\n{cnt}")
    try:
        cell.HoriJustify = _h("CENTER")
        cell.VertJustify = _v("CENTER")
        cell.IsTextWrapped = True
        cell.CharHeight = 14
    except Exception:
        pass

def _write_label_cell(sheet, col0: int, row_1based: int, text: str):
    cell = sheet.getCellByPosition(col0, row_1based - 1)
    cell.setString(text)
    try:
        cell.HoriJustify = uno.Enum("com.sun.star.table.CellHoriJustify", "CENTER")
        cell.VertJustify = uno.Enum("com.sun.star.table.CellVertJustify", "CENTER")
        cell.IsTextWrapped = True
        cell.CharHeight = 11
    except Exception:
        pass

def _update_timestamp_labels_row2(sheet):
    for L in range(WORDLEN_MIN, WORDLEN_MAX + 1):
        wc0 = _word_col0_for_len(L)
        _write_label_cell(sheet, wc0 + TS_IMPORT_OFFSET, 2, "Import\nZeit")
        _write_label_cell(sheet, wc0 + TS_USED_OFFSET,   2, "Kreis\nZeit (30T)")

 # =========================
# KANDIDATENBEREICH LEEREN (bis letzte gefüllte Zeile, Format bleibt)
# =========================      
def clear_candidates_AA_AB(*args):
    """
    Button: löscht nur Inhalte (Text/Wert/Formel), lässt Formatierungen stehen.
    Bereich: AA4:AB<letzte gefüllte Zeile in AA oder AB>
    """
    doc = _get_doc()
    sheet = _get_sheet(doc)

    # Cursor weg von Eingabezelle, damit aktuelle Eingabe übernommen wird
    _goto_cell(doc, SHEET_NAME, "Z4")

    start_row = CAND_START_ROW_1BASED  # = 4
    start0 = start_row - 1
    end0 = start0 + CAND_MAX_ROWS - 1  # Sicherheitslimit (z.B. 1000 Zeilen)

    last0 = None  # 0-based letzte gefüllte Zeile

    # Wir prüfen AA (CAND_COL_WORD0) und AB (CAND_COL_STATUS0)
    for r0 in range(end0, start0 - 1, -1):
        aa = (sheet.getCellByPosition(CAND_COL_WORD0, r0).getString() or "").strip()
        ab = (sheet.getCellByPosition(CAND_COL_STATUS0, r0).getString() or "").strip()
        if aa or ab:
            last0 = r0
            break

    # Nichts gefunden -> nichts zu löschen
    if last0 is None:
        return True

    last_row_1based = last0 + 1
    rng_a1 = f"AA{start_row}:AB{last_row_1based}"

    doc.lockControllers()
    try:
        rng = sheet.getCellRangeByName(rng_a1)
        flags = (1 | 2 | 4 | 16)  # VALUE|DATETIME|STRING|FORMULA
        rng.clearContents(flags)
    finally:
        doc.unlockControllers()

    return True


# ============================================================
# 7) MAIN MACRO
# ============================================================

def import_candidates_from_AA(*args):
    """
    Import:
    - AB width set to 12cm
    - AA cleared ONLY on NEW insert
    - AA kept on: already present / rejected / list full
    """
    doc = _get_doc()
    try:
        sheet = _get_sheet(doc)
    except Exception:
        _msgbox("Import", f"Sheet '{SHEET_NAME}' nicht gefunden.")
        return False
        
    _goto_cell(doc, SHEET_NAME, "Z4")   # Cursor weg von der Eingabezelle -> übernimmt Edit-Inhalt

    # Set AB width
    _set_col_width(sheet, CAND_COL_STATUS0, _cm_to_u100mm(AB_WIDTH_CM))

    try:
        sp = _get_spellchecker_direct()
    except Exception as e:
        _msgbox("Import", f"SpellChecker nicht verfügbar: {e}")
        return False

    _sync_headers(sheet)
    # counts am Anfang (optional)
    _update_all_wordlist_counts(sheet)
    _update_candidate_count_in_AB3(sheet)


    start0 = CAND_START_ROW_1BASED - 1
    end0 = start0 + CAND_MAX_ROWS - 1
    now = _now_str()

    doc.lockControllers()
    try:
        for r0 in range(start0, end0 + 1):
            aa = sheet.getCellByPosition(CAND_COL_WORD0, r0)
            ab = sheet.getCellByPosition(CAND_COL_STATUS0, r0)

            raw = (aa.getString() or "").strip()
            if not raw:
                continue

            ab.setString("")

            # 1) Clean (letters + hyphen only)
            cleaned = _clean_candidate_keep_hyphen(raw)
            if not cleaned:
                ab.setString("Keine Buchstaben übrig.")
                continue

            # 2) Spellcheck
            if not _spell_is_valid_any(sp, cleaned):
                sugg = _spell_suggestions_any(sp, cleaned, max_n=2)
                if sugg:
                    ab.setString("Nicht im LO Wörterbuch; Vorschläge: " + ", ".join(sugg))
                else:
                    ab.setString("Nicht im LO Wörterbuch")
                continue

            # 3) Crossword normalize + length check
            cw = _normalize_crossword(cleaned)
            L = len(cw)

            if L < WORDLEN_MIN:
                ab.setString(f"Wort hat nur {L} Buchstaben.")
                continue

            if L > WORDLEN_MAX:
                ab.setString(f"Wort hat {L} Buchstaben (max. 12).")
                continue

            word_col0 = _word_col0_for_len(L)
            ts_col0 = word_col0 + TIMESTAMP_COL_OFFSET
            colA1 = _col_index_to_letters(word_col0)

            # 4) Already present? -> AA MUST STAY
            found_row1 = _find_word_in_column(sheet, word_col0, cw)
            if found_row1 is not None:
                ab.setString(f"Wort vorhanden in {colA1} {found_row1}")
                # AA bleibt stehen (Änderungswunsch)
                continue

            # 5) Insert new -> AA cleared
            row1 = _append_first_empty(sheet, word_col0, cw, now)
            if row1 is None:
                ab.setString(f"Liste voll in {colA1} (max {WORDLIST_MAX_ROWS}).")
                continue

            ab.setString(f"Neu aufgenommen in {colA1} {row1}")
            aa.setString("")  # NUR hier löschen

        _compact_all_wordlists(sheet)
        
        _update_all_wordlist_counts(sheet)
        _update_timestamp_labels_row2(sheet)
        _update_candidate_count_in_AB3(sheet)

    finally:
        doc.unlockControllers()

    return True


# ============================================================
# 8) EXPORTED MACROS
# ============================================================

g_exportedScripts = (
    import_candidates_from_AA,
    show_wordlists,
    hide_wordlists,
    clear_candidates_AA_AB,
)
