# Tatort: Excel — A Forensic Guide to Data Cleaning (In-Person Edition)

> This edition is built for a **45-minute, in-person** session in **Excel for the web**. The room works in five timed sprints: a short briefing, independent work with a countdown on screen, then a shared checkpoint before everyone moves on together. Every formula is in the workbook, ready to copy, and a **catch-up sheet** after each sprint means nobody falls more than one step behind. The rest of the case is documented in the *On Your Own* guide for students to finish afterwards.

**Date:** 2026  
**Hosted by:** Digital Research and Innovation Lab (American University)  
**Instructor:** Jessica Breen, PhD (Program Director, Geospatial Research Support)  
**Tags:** Data Literacy • Excel • Workshop

---

## Abstract

Every messy spreadsheet is a crime scene. In this fast-paced, hands-on workshop you'll investigate one: a corrupted export of every episode of *Tatort*, Germany's longest-running crime drama. In five short sprints you'll secure the evidence, survey the scene for clues, clear out junk and duplicate records, expose invisible characters, unmask broadcasters hiding behind aliases with a reference table, and rescue dates stored in four different disguises. All of it happens in Excel for the web, so everyone sees the same screen. No prior Excel experience beyond the basics is required. You'll leave with a repeatable investigative workflow, a cheat sheet, and a step-by-step guide to closing the rest of the case on your own.

---

## Materials

- **Slides (PDF):** [`Slides.pdf`](documents/Slides.pdf) · [HTML version with sprint timers](documents/Slides.html)
- **Evidence Card – one page for the session (PDF):** [`Evidence_Card.pdf`](documents/Evidence_Card.pdf)
- **On Your Own – finishing the case after the session (PDF):** [`On_Your_Own.pdf`](documents/On_Your_Own.pdf)
- **Data Cleaning Cheat Sheet (PDF):** [`Cheatsheet.pdf`](documents/Cheatsheet.pdf)
- **The evidence (XLSX):** [`Tatort_Evidence.xlsx`](documents/workshop_materials/Tatort_Evidence.xlsx)
- **Everything for participants (ZIP):** [`workshop_materials.zip`](documents/workshop_materials.zip)

Instructor materials (facilitator guide, answer key, solved workbook, cleaned CSV) are in [`documents/instructor/`](documents/instructor/).

---

## Before the Workshop

**Participants**

- Bring a laptop with a current web browser (Chrome, Edge, Firefox or Safari)
- Sign in at office.com with your university Microsoft 365 account and check that you can open **Excel**

**Instructor**

- Upload `Tatort_Evidence.xlsx` to your OneDrive and share a **view-only** link (students make their own copy in the room)
- Add the link or a QR code to the Pre-flight slide, and print one Evidence Card per participant
- Full checklist: [`documents/instructor/Facilitator_Guide.pdf`](documents/instructor/Facilitator_Guide.pdf)

---

## Topics Covered

**In the session**

- Securing raw data and keeping a "chain of custody"
- Surveying a dataset with filters to spot data quality issues
- Removing junk rows and columns, Excel Tables, Remove Duplicates
- Cleaning spaces and non-breaking spaces in helper columns (`TRIM`, `SUBSTITUTE`, `UNICHAR`, `LEN`)
- Standardizing categories with a reference table (`XLOOKUP`, `TEXTBEFORE`)
- Rebuilding text dates with `DATE` and flagging impossible values

**In the On Your Own guide**

- Numbers stored as text, unique keys and near-duplicates (`VALUE`, `COUNTIF`)
- Line breaks and missing-value placeholders (`SUBSTITUTE`, `CHAR(10)`, `LET`)
- Capitalisation checks (`EXACT`), splitting combined columns (`TEXTBEFORE`, `TEXTAFTER`), Find & Replace
- Cross-examining with a PivotTable; saving and documenting a clean CSV

---

## Repository Structure

```
documents/
├── Slides.pdf / Slides.html          Workshop slides (HTML has the sprint timers; press S for notes)
├── Evidence_Card.pdf                 1-page in-session worksheet
├── On_Your_Own.pdf                   3-page guide to the remaining leads, with expected numbers
├── Cheatsheet.pdf                    2-page reference
├── workshop_materials.zip            Evidence + Evidence Card + On Your Own + Cheat Sheet
├── workshop_materials/
│   └── Tatort_Evidence.xlsx Evidence, Formulas, Known_Aliases, Data_Dictionary,
│                                     Case_Log and Catch_Up_2 – Catch_Up_4 sheets
└── instructor/
    ├── Facilitator_Guide.pdf         Setup, timing, troubleshooting, notes for every slide
    ├── answer_key.md                 Every checkpoint and expected number
    ├── checkpoints.json              The same numbers, as computed by build_data.py
    ├── Tatort_Excel_SOLVED.xlsx      Clean data + row-by-row evidence log
    └── Tatort_clean.csv              Clean data as CSV
source/
├── slides.qmd, evidence_card.qmd, on_your_own.qmd, cheatsheet.qmd, facilitator_guide.qmd
├── theme/                            Slide theme, sprint timer, print styles, fonts
└── build_data.py                     Builds the workbook and verifies every checkpoint
```

To rebuild the slides and handouts, run `quarto render <file>.qmd` inside `source/`. The PDFs were printed from the rendered HTML with Chromium. To regenerate the workbook (and re-check every checkpoint number):

```
python3 source/build_data.py path/to/Tatort_dirty.csv .
```

`Tatort_dirty.csv` is in the companion R workshop repository.

---

## Data

The episode list is adapted from the German Wikipedia list of *Tatort* episodes (CC BY-SA 4.0). For this workshop, errors were deliberately planted in it. **Do not use the workshop file as a source of facts about the show.** The workshop materials are licensed under CC BY 4.0; the episode data keeps its CC BY-SA license.

---

## Partners / Series

This workshop is hosted by the **Digital Research and Innovation Lab**.  
Content was developed and delivered by the **Geospatial Research Lab**.  
It is the in-person edition of [*Tatort: Excel*](https://github.com/GeospatialResearchLab/workshop-data-cleaning-excel-tatort-2026-web) (Excel for the web, self-paced) and a companion to [*Spark Joy in Your Data*](https://github.com/GeospatialResearchLab/workshop-data-cleaning-excel-2025) (Excel) and [*AI-Assisted Strategies for Cleaning and Preparing Research Data*](https://github.com/GeospatialResearchLab/workshop-data-cleaning-GenAI-R-2025) (R).

---

## Citation

Breen, J. (2026). *Tatort: Excel — A Forensic Guide to Data Cleaning, In-Person Edition* (workshop materials).  
Geospatial Research Lab, American University Library.  
https://github.com/GeospatialResearchLab/workshop-data-cleaning-excel-tatort-2026-inperson

---

## License

Workshop materials (slides, handouts, exercises and code) are licensed under [Creative Commons Attribution 4.0 International (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/). You're welcome to reuse and adapt them with attribution; please cite as above. The *Tatort* episode data is adapted from Wikipedia and remains under [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/).
