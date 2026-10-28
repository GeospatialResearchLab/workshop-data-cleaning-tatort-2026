# Instructor Answer Key: Tatort: Excel, In-Person Edition

All numbers below are computed and checked by `source/build_data.py` (see `checkpoints.json`). They assume students follow the sprints in order and add the helper columns in the order given, so the column letters are the same for everyone: `J` Titel_Clean, `K` Titel_Changed, `L` Sender_Code, `M` Date_Clean, `N` Date_Check.

## In-session checkpoints

| Sprint | Question | Expected | If they get something else |
|---|---|---|---|
| Pre-flight | Own copy + `Working` tab | ✓ | Still in the view-only master → File → Create a Copy |
| 1 · Survey | Clues | Title rows and junk columns; repeated episodes; broadcaster aliases (`wdr`, `W.D.R.`…); text dates in two formats and serial numbers; ALL CAPS titles; `n/a` / `NULL` placeholders | — |
| 2 · Clear the scene | Duplicates removed / rows left | **22 / 1,317** | 0 → junk columns were still there |
| 3 · Hidden characters | Titles changed (`Titel_Changed` = TRUE) | **63** (45 extra spaces, 12 non-breaking spaces, 6 near-duplicate twins with a trailing space) | 57 or 51 → `UNICHAR(160)` missing, or `CHAR(160)` used |
| 4 · Aliases | `CHECK` in Sender_Code | **0** | Many → `TRIM` left out or wrong sheet |
| 4 · Aliases | Broadcasters in the Sender_Code filter | **15** | — |
| 4 · Aliases | WDR episodes | **238** (NDR 198, SWR 169, BR 134, HR 102, MDR 96, ORF 90, SFB 52, Radio Bremen 47, SR 45, RBB 44, SWF 30, SDR 27, SRF 27, SF 12) | — |
| 5 · Dates | Impossible dates (`Date_Check` = CHECK) | **5**: 1907-12-13, 1958-09-15, 2030-10-12, 2091-11-03, 3799-12-05 | 3799 is the text typo `…1899`: Excel's `DATE` adds 1900 to years below 1900 |
| 5 · Dates | First real date after sorting | **29 Nov 1970** (Folge 1, *Taxi nach Leipzig*) | 1907 or 1958 → they read the top row without excluding CHECK |
| 5 · Dates | Text dates converted | 106 (71 German `DD.MM.YYYY`, 35 ISO) + 20 serial numbers | — |

The raw broadcaster column shows about 45 distinct spellings in Excel's filter list (56 exact variants; the filter ignores case).

## On Your Own guide: expected numbers

In the guide's order (Lead B removes 6 rows, so later counts are for 1,311 rows).

| Lead | Expected |
|---|---|
| A · Yellow highlights | about 25 highlighted rows (planted on 25 rows before de-duplication) |
| B · `Folge_Repeat` TRUE | **12** rows (6 pairs) |
| B · Remove Duplicates on `Folge_Clean` only | **6 removed, 1,311 rows**, Folge 1–1311 |
| C · `Fall` | 45 numbers stored as text; 20 genuine notes (`2 (4)`, `24 /14`…); impossible: two `0`, one `999` |
| D · `Autor` | 12 line breaks; **20** placeholders → blanks; 43 hyphenated names that Find & Replace on `-` would break |
| E · `Regie` | **16** blank directors; 15 with trailing spaces. `Besonderheiten`: **10** `NULL` |
| F · Capitalisation | **33** CAPS (30 to fix; *LU*, *HAL*, *KI* are genuine), **10** lowercase |
| G · Guests | **78** (`(Gastauftritt …)`); formula finds **77**: the odd one is `Faber und Bönisch (Gastauftritt: Krusenstern)` |
| G · Team aliases | 16 rows with `&`; reversed order (`Leitmayr und Batic`, `Schenk und Ballauf`, `Boerne und Thiel`); slash variants (`Batic/Leitmayr`, `Thiel/Boerne`); 30 alias rows planted in total, plus 25 with extra spaces |
| H · PivotTable | 15 broadcasters, total **1,311**, WDR 238 |

## Planted problems (same as the other Tatort Excel editions)

1,339 rows, 575 planted issues. The full row-by-row list is in `Tatort_Excel_SOLVED.xlsx` → `Evidence_Log`, with totals in its `Summary` sheet.

## Timing (45 minutes)

| Minutes | Segment |
|---:|---|
| 0–4 | Welcome, how tonight works, the workbook |
| 4–9 | Pre-flight |
| 9–14 | Sprint 1 · Survey (1 brief · 3 work · 1 checkpoint) |
| 14–21 | Sprint 2 · Clear the scene (1.5 · 5 · 0.5) |
| 21–28 | Sprint 3 · Hidden characters (1.5 · 5 · 0.5) |
| 28–35 | Sprint 4 · Aliases (1.5 · 5 · 0.5) |
| 35–42 | Sprint 5 · Dates (1.5 · 5 · 0.5) |
| 42–45 | What's left, habits, exit ticket |

## Not yet tested live

The sprint steps reuse menu paths verified in the live Excel for the web test of the web edition. Two things are new and **should be checked in a dry run** before the first session:

- **File → Create a Copy → Create a copy online** when the file is opened from a *view-only* sharing link
- Pasting a formula copied from the `Formulas` sheet's formula bar into a new Table column (does it paste as a formula, and does the column fill down by itself?)
