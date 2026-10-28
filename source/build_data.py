"""Build the workbooks for the IN-PERSON edition of the Tatort Excel workshop.

Same planted mess as the other Tatort Excel editions (same seed), plus:
  * a Formulas sheet (every in-room formula, ready to copy)
  * Catch_Up_2/3/4 sheets: the Working sheet as it should look after each sprint,
    so anyone who falls behind can jump back in
  * verified checkpoint numbers for every sprint (printed, and checked by assert)


Ground truth = the episode list used in the R workshop (Tatort_dirty.csv).
We inject *Excel-flavoured* mess into it and log every injection so the
instructor key can account for each planted clue.
"""
import csv, random, re, datetime as dt, sys, copy
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

SRC = sys.argv[1]
OUT = sys.argv[2]          # repo root
random.seed(2026)

rows = list(csv.DictReader(open(SRC, encoding="utf-8")))
COLS = ["Folge", "Titel", "Sender", "Erstausstrahlung", "Ermittler", "Fall",
        "Autor", "Regie", "Besonderheiten"]
for r in rows:
    r["_id"] = int(r["Folge"])

log = []  # (Folge, column, issue, original, planted)
def note(r, col, issue, before, after):
    log.append((r["_id"], col, issue, before, after))

dirty = [dict(r) for r in rows]
by_id = {r["_id"]: r for r in dirty}

def pick(n, pred=lambda r: True, exclude=set()):
    pool = [r for r in dirty if pred(r) and r["_id"] not in exclude]
    return random.sample(pool, n)

used = {c: set() for c in COLS}
def take(col, n, pred=lambda r: True):
    out = pick(n, pred, used[col])
    for r in out:
        used[col].add(r["_id"])
    return out

# ---------- dates: make them real dates first ------------------------------
for r in dirty:
    r["Erstausstrahlung"] = dt.date.fromisoformat(r["Erstausstrahlung"])

# ---------- Titel ----------------------------------------------------------
keep_caps = {"LU", "HAL", "KI"}  # genuinely all-caps titles: must survive!
for r in take("Titel", 30, lambda r: len(r["Titel"]) > 6 and r["Titel"] not in keep_caps):
    b = r["Titel"]; r["Titel"] = b.upper(); note(r, "Titel", "ALL CAPS", b, r["Titel"])
for r in take("Titel", 10, lambda r: len(r["Titel"]) > 6):
    b = r["Titel"]; r["Titel"] = b.lower(); note(r, "Titel", "all lowercase", b, r["Titel"])
for r in take("Titel", 45):
    b = r["Titel"]
    k = random.choice(["lead", "trail", "double", "both"])
    if k == "lead":   r["Titel"] = "  " + b
    if k == "trail":  r["Titel"] = b + "   "
    if k == "both":   r["Titel"] = " " + b + "  "
    if k == "double":
        r["Titel"] = b.replace(" ", "  ", 1) if " " in b else b + "  "
    note(r, "Titel", "extra spaces", b, r["Titel"])
for r in take("Titel", 12):
    b = r["Titel"]; r["Titel"] = b + " "
    note(r, "Titel", "non-breaking space CHAR(160) (TRIM/CLEAN miss it)", b, r["Titel"])

# ---------- Sender ---------------------------------------------------------
aliases = {
    "WDR": ["wdr", "W.D.R.", "Westdeutscher Rundfunk", "Wdr"],
    "NDR": ["ndr", "N.D.R.", "Norddeutscher Rundfunk"],
    "BR": ["br", "Bayerischer Rundfunk", "B.R."],
    "SWR": ["swr", "Südwestrundfunk", "S.W.R."],
    "HR": ["hr", "Hessischer Rundfunk"],
    "MDR": ["mdr", "Mitteldeutscher Rundfunk"],
    "ORF": ["orf", "Österreichischer Rundfunk", "ORF (Österreich)"],
    "Radio Bremen": ["RB", "Radio-Bremen", "radio bremen"],
    "RBB": ["rbb", "Rundfunk Berlin-Brandenburg"],
    "SR": ["sr", "Saarländischer Rundfunk"],
}
for r in take("Sender", 55, lambda r: r["Sender"] in aliases):
    b = r["Sender"]; r["Sender"] = random.choice(aliases[b])
    note(r, "Sender", "alias / inconsistent spelling", b, r["Sender"])
for r in take("Sender", 15):
    b = r["Sender"]; r["Sender"] = b + " "
    note(r, "Sender", "trailing space", b, r["Sender"])

# ---------- Erstausstrahlung -----------------------------------------------
for r in take("Erstausstrahlung", 70):
    b = r["Erstausstrahlung"]; r["Erstausstrahlung"] = b.strftime("%d.%m.%Y")
    note(r, "Erstausstrahlung", "date stored as German text DD.MM.YYYY", b.isoformat(), r["Erstausstrahlung"])
for r in take("Erstausstrahlung", 35):
    b = r["Erstausstrahlung"]; r["Erstausstrahlung"] = b.isoformat()
    note(r, "Erstausstrahlung", "date stored as ISO text YYYY-MM-DD", b.isoformat(), r["Erstausstrahlung"])
for r in take("Erstausstrahlung", 20):
    b = r["Erstausstrahlung"]
    serial = (b - dt.date(1899, 12, 30)).days
    r["Erstausstrahlung"] = ("SERIAL", serial)
    note(r, "Erstausstrahlung", "date shown as serial number (General format)", b.isoformat(), str(serial))
# year typos (outliers)
typos = [(1970, 1907), (1985, 1958), (2019, 2091), (2003, 2030), (1999, 1899)]
for real, fake in typos:
    cand = [r for r in dirty if isinstance(r["Erstausstrahlung"], dt.date)
            and r["Erstausstrahlung"].year == real and r["_id"] not in used["Erstausstrahlung"]]
    r = random.choice(cand); used["Erstausstrahlung"].add(r["_id"])
    b = r["Erstausstrahlung"]
    r["Erstausstrahlung"] = b.replace(year=fake) if fake >= 1900 else b.strftime("%d.%m.") + str(fake)
    note(r, "Erstausstrahlung", "impossible year (typo)", b.isoformat(), str(r["Erstausstrahlung"]))

# ---------- Ermittler ------------------------------------------------------
team_alias = {
    "Batic und Leitmayr": ["Batic & Leitmayr", "Leitmayr und Batic", "Batic/Leitmayr"],
    "Ballauf und Schenk": ["Ballauf & Schenk", "Schenk und Ballauf"],
    "Thiel und Boerne": ["Thiel & Boerne", "Boerne und Thiel", "Thiel/Boerne"],
    "Schimanski und Thanner": ["Schimanski & Thanner"],
    "Odenthal und Kopper": ["Odenthal & Kopper"],
}
for r in take("Ermittler", 30, lambda r: r["Ermittler"] in team_alias):
    b = r["Ermittler"]; r["Ermittler"] = random.choice(team_alias[b])
    note(r, "Ermittler", "alias (& / / or reversed order)", b, r["Ermittler"])
for r in take("Ermittler", 25):
    b = r["Ermittler"]; r["Ermittler"] = random.choice([" ", "  "]) + b + random.choice(["", " "])
    note(r, "Ermittler", "extra spaces", b, r["Ermittler"])

# ---------- Fall / Folge numbers as text -----------------------------------
for r in take("Folge", 60):
    note(r, "Folge", "number stored as text", r["Folge"], "'" + r["Folge"])
    r["Folge"] = ("TEXT", r["Folge"])
for r in take("Folge", 8):
    b = r["Folge"]; r["Folge"] = "Folge " + b
    note(r, "Folge", "prefix 'Folge ' (text)", b, r["Folge"])
for r in take("Folge", 6):
    b = r["Folge"]; r["Folge"] = "#" + b
    note(r, "Folge", "prefix '#' (text)", b, r["Folge"])
for r in take("Fall", 45, lambda r: r["Fall"].isdigit()):
    note(r, "Fall", "number stored as text", r["Fall"], "'" + r["Fall"])
    r["Fall"] = ("TEXT", r["Fall"])
for r, v in zip(take("Fall", 3, lambda r: isinstance(r["Fall"], str) and r["Fall"].isdigit()), ["0", "0", "999"]):
    note(r, "Fall", "impossible value (outlier)", r["Fall"], v); r["Fall"] = v

# ---------- Autor ----------------------------------------------------------
for r in take("Autor", 12, lambda r: ", " in r["Autor"]):
    b = r["Autor"]; r["Autor"] = b.replace(", ", "\n")
    note(r, "Autor", "line break CHAR(10) instead of ', '", b, r["Autor"].replace("\n", "⏎"))
for r, ph in zip(take("Autor", 20), ["n/a"] * 7 + ["unbekannt"] * 6 + ["-"] * 4 + ["?"] * 3):
    note(r, "Autor", f"missing value placeholder '{ph}'", r["Autor"], ph); r["Autor"] = ph

# ---------- Regie ----------------------------------------------------------
for r in take("Regie", 16):
    note(r, "Regie", "missing (blank)", r["Regie"], ""); r["Regie"] = ""
for r in take("Regie", 15):
    b = r["Regie"]; r["Regie"] = b + "  "
    note(r, "Regie", "trailing spaces", b, r["Regie"])

# ---------- Besonderheiten -------------------------------------------------
for r in take("Besonderheiten", 10, lambda r: r["Besonderheiten"] == ""):
    note(r, "Besonderheiten", "placeholder 'NULL' instead of blank", "", "NULL"); r["Besonderheiten"] = "NULL"

# ---------- irrelevant columns ---------------------------------------------
clerks = ["KHK Brandt", "KK Yilmaz", "KHM Novak", "--"]
for r in dirty:
    r["Import_ID"] = f"TT-{random.randint(10000, 99999)}"
    r["Bearbeitet_von"] = random.choice(clerks)
    r["Spalte1"] = None

# ---------- assemble rows + duplicates --------------------------------------
order = sorted(dirty, key=lambda r: r["_id"])
out = list(order)
exact = random.sample(order[5:], 22)
for r in exact:
    i = out.index(r)
    out.insert(i + random.randint(1, 4), r)   # same object -> identical row
    log.append((r["_id"], "(whole row)", "exact duplicate row", "", "row repeated"))
near = random.sample([r for r in order[5:] if r not in exact], 6)
for r in near:
    twin = dict(r)
    if isinstance(twin["Titel"], str):
        twin["Titel"] = twin["Titel"].strip() + " "
    twin["Import_ID"] = f"TT-{random.randint(10000, 99999)}"
    out.insert(out.index(r) + random.randint(1, 3), twin)
    log.append((r["_id"], "(whole row)", "near-duplicate (trailing space in Titel + new Import_ID)", "", "row repeated"))

# ============================================================================
# Workbook styling helpers
# ============================================================================
INK = "1B1B1B"; TAPE = "FFD400"; RED = "B3261E"; PAPER = "F4EFE6"; GREY = "6B6B6B"
hdr_font = Font(name="Calibri", bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor=INK)
tape_fill = PatternFill("solid", fgColor=TAPE)
thin = Side(style="thin", color="BFBFBF")

def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = Alignment(vertical="center")

def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x

def write_value(cell, v):
    if isinstance(v, tuple) and v[0] == "TEXT":
        cell.value = v[1]; cell.number_format = "@"
    elif isinstance(v, tuple) and v[0] == "SERIAL":
        cell.value = v[1]; cell.number_format = "General"
    elif isinstance(v, (dt.date, dt.datetime)):
        cell.value = v; cell.number_format = "yyyy-mm-dd"
    elif isinstance(v, str) and v.isdigit() and len(v) < 6:
        cell.value = int(v)
    else:
        cell.value = v

# ============================================================================
# Participant workbook
# ============================================================================
wb = Workbook()
ws = wb.active
ws.title = "Tatort_Evidence"
ws["A1"] = "TATORT – EPISODENLISTE  |  Export aus Altsystem (Stand 2026)"
ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=INK)
ws["A1"].fill = tape_fill
ws.merge_cells("A1:L1")
ws["A2"] = "ACHTUNG: ungeprüfte Daten – nicht zitieren!  (Caution: unverified data)"
ws["A2"].font = Font(italic=True, color=RED)
ws.merge_cells("A2:L2")
headers = [" Folge", "TITEL", "Sender ", "Erstausstrahlung", "Ermittler", "Fall",
           "Autor", "Regie", "Besonderheiten", "Import_ID", "Bearbeitet_von", "Spalte1"]
HR = 4
for c, h in enumerate(headers, 1):
    ws.cell(row=HR, column=c, value=h)
style_header(ws, HR, len(headers))
keys = COLS + ["Import_ID", "Bearbeitet_von", "Spalte1"]
for i, r in enumerate(out):
    for c, k in enumerate(keys, 1):
        write_value(ws.cell(row=HR + 1 + i, column=c), r[k])
widths(ws, [8, 38, 14, 16, 34, 7, 34, 26, 60, 11, 15, 9])
# alternating fill on some rows: "helpful" manual highlighting from a previous clerk
for i in random.sample(range(len(out)), 25):
    for c in range(1, 10):
        ws.cell(row=HR + 1 + i, column=c).fill = PatternFill("solid", fgColor="FFF2CC")
n_dirty_rows = len(out)

# ---- Known aliases (reference table) ---------------------------------------
ref = wb.create_sheet("Known_Aliases")
ref["A1"] = "FAHNDUNGSLISTE – Known aliases for broadcasters (Sender)"
ref["A1"].font = Font(size=13, bold=True); ref["A1"].fill = tape_fill
ref.merge_cells("A1:D1")
ref["A2"] = "Use this lookup table with XLOOKUP/VLOOKUP to bring every alias back to its official code."
ref["A2"].font = Font(italic=True, color=GREY)
ref.append([])
ref.append(["Alias (as found)", "Sender_Code", "Official name", "Country"])
style_header(ref, 4, 4)
official = {
    "WDR": ("Westdeutscher Rundfunk", "DE"), "NDR": ("Norddeutscher Rundfunk", "DE"),
    "SWR": ("Südwestrundfunk", "DE"), "BR": ("Bayerischer Rundfunk", "DE"),
    "HR": ("Hessischer Rundfunk", "DE"), "MDR": ("Mitteldeutscher Rundfunk", "DE"),
    "ORF": ("Österreichischer Rundfunk", "AT"), "SFB": ("Sender Freies Berlin", "DE"),
    "SR": ("Saarländischer Rundfunk", "DE"), "RBB": ("Rundfunk Berlin-Brandenburg", "DE"),
    "Radio Bremen": ("Radio Bremen", "DE"), "SWF": ("Südwestfunk", "DE"),
    "SDR": ("Süddeutscher Rundfunk", "DE"), "SRF": ("Schweizer Radio und Fernsehen", "CH"),
    "SF": ("Schweizer Fernsehen", "CH"), "DFF": ("Deutscher Fernsehfunk", "DD"),
}
alias_rows = []
for code, (name, ctry) in official.items():
    alias_rows.append((code, code, name, ctry))
    for a in aliases.get(code, []):
        alias_rows.append((a, code, name, ctry))
for a in alias_rows:
    ref.append(list(a))
widths(ref, [28, 14, 32, 9])
ref.freeze_panes = "A5"

# ---- Data dictionary -------------------------------------------------------
dd = wb.create_sheet("Data_Dictionary")
dd["A1"] = "DATA DICTIONARY – what each column should contain once the case is closed"
dd["A1"].font = Font(size=13, bold=True); dd["A1"].fill = tape_fill
dd.merge_cells("A1:E1")
dd.append([])
dd.append(["Column (Spalte)", "English", "Expected type", "Example (clean)", "Notes"])
style_header(dd, 3, 5)
dd_rows = [
    ("Folge", "Episode number", "Whole number, unique", "1", "1 to 1311; every episode appears exactly once"),
    ("Titel", "Title", "Text", "Taxi nach Leipzig", "Keep original capitalisation; some real titles are all caps (LU, HAL, KI)"),
    ("Sender", "Broadcaster (lead)", "Text, from Known_Aliases code list", "NDR", "Co-productions like 'SWR/SRF' can be split into lead + partner"),
    ("Erstausstrahlung", "First broadcast date", "Date", "1970-11-29", "Series began 29 Nov 1970; no dates before that or in the future"),
    ("Ermittler", "Investigator(s)", "Text", "Batic und Leitmayr", "Guest appearances are written in brackets: 'Kressin(Gastauftritt Lutz)'"),
    ("Fall", "Case number for this investigator", "Whole number ≥ 1", "3", "A few entries count cases for two investigators, e.g. '2 (4)'"),
    ("Autor", "Writer(s)", "Text", "Wolfgang Menge", "Multiple writers separated by ', '"),
    ("Regie", "Director", "Text", "Peter Schulze-Rohr", ""),
    ("Besonderheiten", "Notes / trivia", "Text (blank allowed)", "Lief in den USA im Kino", "Bracketed numbers like [1] are leftover footnote markers"),
    ("Import_ID", "Legacy system ID", "—", "", "Not needed for analysis"),
    ("Bearbeitet_von", "Entered by (clerk)", "—", "", "Not needed for analysis"),
    ("Spalte1", "Empty column", "—", "", "Not needed for analysis"),
]
for r_ in dd_rows:
    dd.append(list(r_))
widths(dd, [18, 30, 30, 22, 70])
for row in dd.iter_rows(min_row=4):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")

# ---- Case log template -----------------------------------------------------
cl = wb.create_sheet("Case_Log")
cl["A1"] = "ERMITTLUNGSPROTOKOLL – Case log (chain of custody for every change you make)"
cl["A1"].font = Font(size=13, bold=True); cl["A1"].fill = tape_fill
cl.merge_cells("A1:F1")
cl.append([])
cl.append(["Step", "Clue (what did you notice?)", "Column(s)", "Action taken", "Tool / formula", "How did you verify it?"])
style_header(cl, 3, 6)
cl.append([1, "Example: some titles have extra spaces", "Titel", "Trimmed spaces in a helper column, pasted values back", "=TRIM(B5)", "LEN() before vs after; filter for leading space"])
for c in cl[4]:
    c.font = Font(italic=True, color=GREY)
for i in range(2, 21):
    cl.append([i])
widths(cl, [6, 40, 16, 44, 26, 40])
for row in cl.iter_rows(min_row=4, max_row=23):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = Border(bottom=thin)

ws.freeze_panes = None  # leave it for participants to discover
wb.active = 0
# ============================================================================
# IN-PERSON EDITION: simulate each sprint exactly as Excel would do it,
# write Catch_Up sheets + a Formulas sheet, and verify the checkpoints.
# ============================================================================
from openpyxl.styles import Font as _F
TODAY = dt.date(2026, 10, 27)
BASE = dt.date(1899, 12, 30)
CLEAN_HDR = COLS  # header names after Sprint 2

def cellval(v):
    """What the Excel cell actually holds (mirrors write_value)."""
    if isinstance(v, tuple) and v[0] == "TEXT":   return str(v[1])
    if isinstance(v, tuple) and v[0] == "SERIAL": return v[1]
    if isinstance(v, (dt.date, dt.datetime)):     return v
    if isinstance(v, str) and v.isdigit() and len(v) < 6: return int(v)
    return v

def xl_trim(s):
    return re.sub(" +", " ", str(s)).strip(" ")

# ---- Sprint 2: delete junk rows/cols, Remove Duplicates (all columns) -------
def dupkey(r):
    k = []
    for c in COLS:
        v = cellval(r[c])
        k.append(v.lower() if isinstance(v, str) else v)   # Excel compares text case-insensitively
    return tuple(k)
seen, s2 = set(), []
for r in out:
    k = dupkey(r)
    if k in seen: continue
    seen.add(k); s2.append(r)
CP = {}
CP["rows_before"] = len(out)
CP["dupes_removed"] = len(out) - len(s2)
CP["rows_after_s2"] = len(s2)

# ---- Sprint 3: Titel_Clean, Titel_Changed ----------------------------------
def titel_clean(v):
    return xl_trim(str(v).replace(" ", " ")) if v is not None else ""
s3 = []
for r in s2:
    t = cellval(r["Titel"])
    tc = titel_clean(t)
    s3.append({"Titel_Clean": tc, "Titel_Changed": len(str(t)) != len(tc)})
CP["titles_changed"] = sum(x["Titel_Changed"] for x in s3)

# ---- Sprint 4: Sender_Code via XLOOKUP on Known_Aliases ---------------------
alias_lookup = {}
for row in ref.iter_rows(min_col=1, max_col=2, values_only=True):
    a, code = row
    if a is not None and str(a).lower() not in alias_lookup:
        alias_lookup[str(a).lower()] = code
def sender_code(v):
    s = str(v)
    lead = s.split("/", 1)[0] if "/" in s else s
    return alias_lookup.get(xl_trim(lead).lower(), "CHECK")
s4 = [sender_code(cellval(r["Sender"])) for r in s2]
CP["sender_check"] = s4.count("CHECK")
CP["broadcasters"] = len(set(s4) - {"CHECK"})
CP["wdr"] = s4.count("WDR")

# ---- Sprint 5: Date_Clean, Date_Check --------------------------------------
def xl_date(y, m, d):
    y = int(y); y = y + 1900 if y < 1900 else y
    return dt.date(y, int(m), int(d))
def date_clean(v):
    if isinstance(v, int):  return BASE + dt.timedelta(days=v)
    if isinstance(v, (dt.date, dt.datetime)): return v
    s = str(v)
    if s[2] == ".": return xl_date(s[6:10], s[3:5], s[0:2])
    return xl_date(s[0:4], s[5:7], s[8:10])
s5 = [date_clean(cellval(r["Erstausstrahlung"])) for r in s2]
flags = ["CHECK" if (d < dt.date(1970, 11, 29) or d > TODAY) else "OK" for d in s5]
CP["date_check"] = flags.count("CHECK")
CP["impossible"] = sorted(d.isoformat() for d, f in zip(s5, flags) if f == "CHECK")
okd = [d for d, f in zip(s5, flags) if f == "OK"]
CP["earliest"] = min(okd).isoformat(); CP["latest"] = max(okd).isoformat()

# ---- truth checks: in-room cleaning must agree with the answer key ----------
truth = {r["_id"]: r for r in rows}
def folge_num(v):
    return int(str(v).replace("Folge ", "").replace("#", ""))
for r, code, d, f in zip(s2, s4, s5, flags):
    t = truth[folge_num(cellval(r["Folge"]))]
    assert code == t["Sender"].split("/")[0], (code, t["Sender"])
    if f == "OK":
        assert d.isoformat() == t["Erstausstrahlung"], (d, t["Erstausstrahlung"])
assert CP["sender_check"] == 0 and CP["date_check"] == 5 and CP["earliest"] == "1970-11-29"
# titles: clean title equals truth unless capitalisation was planted (on-your-own)
for r, x in zip(s2, s3):
    t = truth[folge_num(cellval(r["Folge"]))]["Titel"]
    assert x["Titel_Clean"].lower() == xl_trim(t).lower(), (x["Titel_Clean"], t)

# ---- On your own: near-duplicates & Folge -----------------------------------
NEAR_COLS = [c for c in COLS if c != "Titel"]
seen2, s6 = set(), []
for r in s2:
    k = tuple((cellval(r[c]).lower() if isinstance(cellval(r[c]), str) else cellval(r[c])) for c in NEAR_COLS)
    if k in seen2: continue
    seen2.add(k); s6.append(r)
CP["near_removed"] = len(s2) - len(s6)
CP["rows_final"] = len(s6)
assert CP["rows_final"] == len(rows)
assert sorted(folge_num(cellval(r["Folge"])) for r in s6) == list(range(1, len(rows) + 1))

# ---- write Catch_Up sheets ----------------------------------------------------
CU_FILL = PatternFill("solid", fgColor="5E6670")
def catch_up(name, label, extra_cols, extra_vals):
    cu = wb.create_sheet(name)
    cu.sheet_properties.tabColor = "5E6670"
    hdr = CLEAN_HDR + extra_cols
    cu.append(hdr)
    for i, r in enumerate(s2):
        vals = [None] * len(hdr)
        cu.append(vals)
        for c, k in enumerate(COLS, 1):
            write_value(cu.cell(row=i + 2, column=c), r[k])
        for j, v in enumerate(extra_vals[i]):
            cell = cu.cell(row=i + 2, column=len(COLS) + 1 + j, value=v)
            if isinstance(v, dt.date): cell.number_format = "yyyy-mm-dd"
    for c in range(1, len(hdr) + 1):
        cell = cu.cell(row=1, column=c); cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=False)
    t = Table(displayName=name.replace("_", ""), ref=f"A1:{get_column_letter(len(hdr))}{len(s2)+1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium15", showRowStripes=True)
    cu.add_table(t)
    widths(cu, [8, 34, 13, 16, 30, 7, 30, 24, 40] + [26] * len(extra_cols))
    cu.freeze_panes = "B2"
    return cu

catch_up("Catch_Up_2", "after Sprint 2", [], [[] for _ in s2])
catch_up("Catch_Up_3", "after Sprint 3", ["Titel_Clean", "Titel_Changed"],
         [[x["Titel_Clean"], x["Titel_Changed"]] for x in s3])
catch_up("Catch_Up_4", "after Sprint 4", ["Titel_Clean", "Titel_Changed", "Sender_Code"],
         [[x["Titel_Clean"], x["Titel_Changed"], c] for x, c in zip(s3, s4)])

# ---- Formulas sheet -----------------------------------------------------------
FORMULAS = [
    ("3", "Titel_Clean", '=TRIM(SUBSTITUTE([@Titel], UNICHAR(160), " "))',
     "Swaps the invisible non-breaking space for a normal space, then TRIM removes extra spaces."),
    ("3", "Titel_Changed", "=LEN([@Titel])<>LEN([@Titel_Clean])",
     "TRUE if cleaning changed the title. Filter for TRUE to count them."),
    ("4", "Sender_Code", '=XLOOKUP(TRIM(TEXTBEFORE([@Sender], "/", , , , [@Sender])), Known_Aliases!A:A, Known_Aliases!B:B, "CHECK")',
     "Takes the lead broadcaster (before any /), trims it, and looks up its official code. Unknown aliases show CHECK."),
    ("5", "Date_Clean", '=LET(d, [@Erstausstrahlung], IF(ISNUMBER(d), d, IF(MID(d,3,1)=".", DATE(RIGHT(d,4), MID(d,4,2), LEFT(d,2)), DATE(LEFT(d,4), MID(d,6,2), RIGHT(d,2)))))',
     "d is the date cell. Real dates pass through; DD.MM.YYYY and YYYY-MM-DD text are rebuilt with DATE."),
    ("5", "Date_Check", '=IF(OR([@Date_Clean]<DATE(1970,11,29), [@Date_Clean]>TODAY()), "CHECK", "OK")',
     "Tatort began on 29 Nov 1970. Anything earlier, or in the future, is impossible."),
]
fs = wb.create_sheet("Formulas", 1)
fs.sheet_properties.tabColor = TAPE
fs["A1"] = "FORMELN – every formula for today's sprints, ready to copy"
fs["A1"].font = Font(size=13, bold=True); fs["A1"].fill = tape_fill
fs.merge_cells("A1:D1")
fs["A2"] = ("How to copy: click the green formula cell → click in the formula bar → Ctrl/⌘+A, Ctrl/⌘+C → Esc. "
            "Then click the first empty cell under your new column header (row 2) and press Ctrl/⌘+V.")
fs["A2"].font = Font(italic=True, color=GREY); fs["A2"].alignment = Alignment(wrap_text=True)
fs.merge_cells("A2:D2"); fs.row_dimensions[2].height = 32
fs.append([])
fs.append(["Sprint", "New column (type in row 1)", "Formula (copy this)", "What it does"])
style_header(fs, 4, 4)
for i, (sp, col, f, what) in enumerate(FORMULAS, 5):
    fs.cell(row=i, column=1, value=int(sp))
    fs.cell(row=i, column=2, value=col).font = Font(bold=True)
    c = fs.cell(row=i, column=3); c.value = f; c.data_type = "s"     # stored as text, not a live formula
    c.font = Font(name="Consolas", color="1F5C2E"); c.fill = PatternFill("solid", fgColor="E8F3EA")
    fs.cell(row=i, column=4, value=what)
    for cc in range(1, 5):
        fs.cell(row=i, column=cc).alignment = Alignment(wrap_text=True, vertical="top")
r0 = 5 + len(FORMULAS) + 1
fs.cell(row=r0, column=1, value="CHECKPOINT COUNTERS – paste these into any empty cell of the Case_Log sheet (column H is free). If you work in a Catch_Up sheet, change Working to that sheet's name.").font = Font(bold=True)
fs.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=4)
fs.cell(row=r0, column=1).alignment = Alignment(wrap_text=True); fs.row_dimensions[r0].height = 30
COUNTERS = [
    ("3", "Titles changed", "=COUNTIF(Working!K:K, TRUE)", "Counts TRUE in Titel_Changed (column K)."),
    ("4", "Unknown aliases", '=COUNTIF(Working!L:L, "CHECK")', "Counts CHECK in Sender_Code (column L). Should be 0."),
    ("4", "WDR episodes", '=COUNTIF(Working!L:L, "WDR")', "Counts one broadcaster. Try others: NDR, BR, ORF…"),
    ("5", "Impossible dates", '=COUNTIF(Working!N:N, "CHECK")', "Counts CHECK in Date_Check (column N)."),
]
for i, (sp, col, f, what) in enumerate(COUNTERS, r0 + 1):
    fs.cell(row=i, column=1, value=int(sp))
    fs.cell(row=i, column=2, value=col).font = Font(bold=True)
    c = fs.cell(row=i, column=3); c.value = f; c.data_type = "s"
    c.font = Font(name="Consolas", color="1F5C2E"); c.fill = PatternFill("solid", fgColor="E8F3EA")
    fs.cell(row=i, column=4, value=what)
    for cc in range(1, 5):
        fs.cell(row=i, column=cc).alignment = Alignment(wrap_text=True, vertical="top")
widths(fs, [7, 24, 70, 52])

# order: evidence, formulas, aliases, dictionary, log, catch-ups
order_names = ["Tatort_Evidence", "Formulas", "Known_Aliases", "Data_Dictionary", "Case_Log",
               "Catch_Up_2", "Catch_Up_3", "Catch_Up_4"]
wb._sheets = [wb[n] for n in order_names]
ws.sheet_properties.tabColor = INK

# rename case-log example to the in-person workflow
cl["E4"] = "=TRIM(SUBSTITUTE([@Titel], UNICHAR(160), \" \"))"; cl["E4"].data_type = "s"
cl["D4"] = "New helper column Titel_Clean (original kept for comparison)"

import json
print("CHECKPOINTS", json.dumps(CP, indent=1))
json.dump(CP, open(f"{OUT}/documents/instructor/checkpoints.json", "w"), indent=1)

wb.save(f"{OUT}/documents/workshop_materials/Tatort_Evidence.xlsx")

# ---- German-language variants of the evidence workbook ----------------------
# Same data; only the Formulas sheet (and the Case_Log example) use German
# function names. Two separators: ";" (German regional format) and ","
# (German display language with US regional format).
import os
FORMULAS_DE = [
    ('=GLÄTTEN(WECHSELN([@Titel]; UNIZEICHEN(160); " "))',
     "WECHSELN swaps the invisible non-breaking space (UNIZEICHEN(160)) for a normal space, then GLÄTTEN removes extra spaces."),
    ("=LÄNGE([@Titel])<>LÄNGE([@Titel_Clean])",
     "WAHR if cleaning changed the title. Filter for WAHR to count them."),
    ('=XVERWEIS(GLÄTTEN(TEXTVOR([@Sender]; "/"; ; ; ; [@Sender])); Known_Aliases!A:A; Known_Aliases!B:B; "CHECK")',
     "TEXTVOR takes the lead broadcaster (before any /), GLÄTTEN trims it, XVERWEIS looks up its official code. Unknown aliases show CHECK."),
    ('=LET(d; [@Erstausstrahlung]; WENN(ISTZAHL(d); d; WENN(TEIL(d;3;1)="."; DATUM(RECHTS(d;4); TEIL(d;4;2); LINKS(d;2)); DATUM(LINKS(d;4); TEIL(d;6;2); RECHTS(d;2)))))',
     "d is the date cell. Real dates pass through; DD.MM.YYYY and YYYY-MM-DD text are rebuilt with DATUM."),
    ('=WENN(ODER([@Date_Clean]<DATUM(1970;11;29); [@Date_Clean]>HEUTE()); "CHECK"; "OK")',
     "Tatort began on 29 Nov 1970. Anything earlier, or in the future, is impossible."),
]
COUNTERS_DE = [
    ("=ZÄHLENWENN(Working!K:K; WAHR)", "Counts WAHR in Titel_Changed (column K). Paste on the Case_Log tab, cell H4."),
    ('=ZÄHLENWENN(Working!L:L; "CHECK")', "Counts CHECK in Sender_Code (column L). Should be 0. Case_Log H5."),
    ('=ZÄHLENWENN(Working!L:L; "WDR")', "Counts one broadcaster. Case_Log H6."),
    ('=ZÄHLENWENN(Working!N:N; "CHECK")', "Counts CHECK in Date_Check (column N). Case_Log H7."),
]
assert len(FORMULAS_DE) == len(FORMULAS) and len(COUNTERS_DE) == len(COUNTERS)
DE_EXAMPLE = '=GLÄTTEN(WECHSELN([@Titel]; UNIZEICHEN(160); " "))'
os.makedirs(f"{OUT}/documents/german", exist_ok=True)
comma = lambda f: f.replace(";", ",")
fs["A1"] = "FORMELN – every formula for today's sprints, in German Excel"
fs["A2"] = ("Two versions of each formula. Use column C if your Excel uses ; between arguments (German regional format), "
            "column D if it uses , (German Excel with US regional format). Not sure? Start typing =WENN( in a cell: the tooltip shows the separator. "
            "How to copy: click the green cell → click in the formula bar → Strg/⌘+A, Strg/⌘+C → Esc. "
            "Then click the first empty cell under your new column header and press Strg/⌘+V.")
fs.row_dimensions[2].height = 58
fs.unmerge_cells("A1:D1"); fs.merge_cells("A1:E1")
fs.unmerge_cells("A2:D2"); fs.merge_cells("A2:E2")
hdr = ["Sprint", "New column (type in row 1)", "Formula with ;  (German regional format)", "Formula with ,  (US regional format)", "What it does"]
for c, h in enumerate(hdr, 1):
    fs.cell(row=4, column=c, value=h)
style_header(fs, 4, 5)
def put(row, f, what):
    for col, val in ((3, f), (4, comma(f))):
        c = fs.cell(row=row, column=col); c.value = val; c.data_type = "s"
        c.font = Font(name="Consolas", color="1F5C2E"); c.fill = PatternFill("solid", fgColor="E8F3EA")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    d = fs.cell(row=row, column=5, value=what); d.alignment = Alignment(wrap_text=True, vertical="top")
for i, (f, what) in enumerate(FORMULAS_DE, 5):
    put(i, f, what)
fs.unmerge_cells(start_row=r0, start_column=1, end_row=r0, end_column=4)
fs.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=5)
for i, (f, what) in enumerate(COUNTERS_DE, r0 + 1):
    put(i, f, what)
widths(fs, [7, 22, 60, 60, 46])
cl["E4"] = DE_EXAMPLE; cl["E4"].data_type = "s"
wb.save(f"{OUT}/documents/german/Tatort_Evidence_DE.xlsx")
# restore English (the in-memory workbook is not saved again, but keep it tidy)

# ============================================================================
# Instructor workbook: cleaned data + evidence log
# ============================================================================
def clean_row(r):
    erm = r["Ermittler"]
    m = re.match(r"^(.*?)\s*\(Gastauftritt (.*)\)\s*$", erm)
    lead_erm, gast = (m.group(1), m.group(2)) if m else (erm, "")
    sender = r["Sender"]
    lead, *co = sender.split("/")
    fall = r["Fall"]
    fall_num = int(re.match(r"\d+", fall).group()) if re.match(r"\d+", fall) else None
    fall_note = "" if fall.isdigit() else f"original: '{fall}'"
    d = dt.date.fromisoformat(r["Erstausstrahlung"])
    return [int(r["Folge"]), r["Titel"], lead, "/".join(co), official[lead][0], official[lead][1],
            d, d.year, lead_erm, gast, fall_num, fall_note, r["Autor"], r["Regie"],
            re.sub(r"\[\d+\]", "", r["Besonderheiten"]).strip()]

kw = Workbook()
k1 = kw.active; k1.title = "Clean_Data"
clean_hdr = ["Folge", "Titel", "Sender", "Sender_Partner", "Sender_Name", "Land", "Erstausstrahlung",
             "Jahr", "Ermittler", "Gastauftritt", "Fall", "Fall_Hinweis", "Autor", "Regie", "Besonderheiten"]
k1.append(clean_hdr); style_header(k1, 1, len(clean_hdr))
for r in rows:
    k1.append(clean_row(r))
for row in k1.iter_rows(min_row=2, min_col=7, max_col=7):
    for c in row: c.number_format = "yyyy-mm-dd"
widths(k1, [7, 38, 13, 15, 28, 6, 15, 7, 30, 26, 6, 18, 34, 26, 60])
k1.freeze_panes = "C2"
tab = Table(displayName="Tatort_Clean", ref=f"A1:{get_column_letter(len(clean_hdr))}{len(rows)+1}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium1", showRowStripes=True)
k1.add_table(tab)

k2 = kw.create_sheet("Evidence_Log")
k2.append(["Folge", "Column", "Planted issue", "Original value", "Planted value"])
style_header(k2, 1, 5)
for e in sorted(log, key=lambda x: (x[1], x[0])):
    k2.append(list(e))
widths(k2, [8, 18, 52, 50, 50])
k2.freeze_panes = "A2"; k2.auto_filter.ref = f"A1:E{len(log)+1}"

k3 = kw.create_sheet("Summary")
k3.append(["Column", "Planted issue", "Count"]); style_header(k3, 1, 3)
cnt = Counter((e[1], re.sub(r"'.*'", "'…'", e[2]) if "placeholder" in e[2] else e[2]) for e in log)
for (col, iss), n in sorted(cnt.items()):
    k3.append([col, iss, n])
k3.append([]); k3.append(["", "TOTAL planted clues", len(log)])
k3.append(["", "Rows in Tatort_Evidence (excl. header)", n_dirty_rows])
k3.append(["", "Rows after cleaning", len(rows)])
k3.append([])
k3.append(["Pre-existing (not planted) issues kept from the source list:"])
pre = [
    ("Ermittler", "guest appearance inside brackets, e.g. 'Kressin(Gastauftritt Lutz)'", sum("(Gastauftritt" in r["Ermittler"] for r in rows)),
    ("Sender", "co-productions written as 'SWR/SRF'", sum("/" in r["Sender"] for r in rows)),
    ("Fall", "non-numeric case counts, e.g. '2 (4)', '24 /14'", sum(not r["Fall"].isdigit() for r in rows)),
    ("Besonderheiten", "footnote markers like [1]", sum(bool(re.search(r"\[\d+\]", r["Besonderheiten"])) for r in rows)),
    ("Titel", "repeated titles that are NOT duplicates (different episodes)", sum(v for v in Counter(r["Titel"] for r in rows).values() if v > 1)),
    ("Titel", "genuine all-caps titles (LU, HAL, KI) – PROPER() would break them", 3),
]
for p in pre: k3.append(list(p))
widths(k3, [18, 70, 10])
kw.save(f"{OUT}/documents/instructor/Tatort_Excel_SOLVED.xlsx")

# also give the clean data as CSV for convenience
with open(f"{OUT}/documents/instructor/Tatort_clean.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(clean_hdr)
    for r in rows:
        cr = clean_row(r); cr[6] = cr[6].isoformat(); w.writerow(["" if v is None else v for v in cr])

print("dirty rows", n_dirty_rows, "clean rows", len(rows), "planted", len(log))
for (col, iss), n in sorted(cnt.items()):
    print(f"{n:4d}  {col:18s} {iss}")
