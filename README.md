# Tatort: Excel — A Forensic Guide to Data Cleaning

**Date:** 2026  
**Hosted by:** Digital Research and Innovation Lab (American University)  
**Instructor:** Jessica Breen, PhD (Program Director, Geospatial Research Support)  
**Tags:** Data Literacy • Excel • Data Cleaning • Workshop

---

## Abstract

Every messy spreadsheet is a crime scene. In this hands-on workshop you'll investigate one: a corrupted export of every episode of *Tatort*, Germany's longest-running crime drama. In five timed sprints you'll clear out duplicate rows and hidden characters, then use a reference table to standardize broadcaster names and rebuild dates that were stored as text. Everything happens in Excel for the web, so everyone sees the same screen, and every formula is provided. You only need basic Excel. The steps we don't reach in the room are written up in a guide you can finish on your own.

---

## Materials

- **Workshop page:** [`index.html`](index.html) (participant landing page, published with GitHub Pages)
- **Slides:** [`Slides.pdf`](documents/Slides.pdf) · [HTML version with sprint timers](documents/Slides.html)
- **Evidence Card (handout):** [`Evidence_Card.pdf`](documents/Evidence_Card.pdf)
- **On Your Own guide:** [`On_Your_Own.pdf`](documents/On_Your_Own.pdf)
- **Cheat Sheet:** [`Cheatsheet.pdf`](documents/Cheatsheet.pdf)
- **Workshop dataset:** [`Tatort_Evidence.xlsx`](documents/workshop_materials/Tatort_Evidence.xlsx) · [all participant materials (ZIP)](documents/workshop_materials.zip)
- **German Excel versions:** [`German_language/`](documents/German_language/): the [dataset with German formulas](documents/German_language/Tatort_Evidence_DE.xlsx) (the Formulas sheet shows each formula with `;` and `,` separators side by side), the [On Your Own guide](documents/German_language/On_Your_Own_DE.pdf) and the [Cheat Sheet](documents/German_language/Cheatsheet_DE.pdf) with German formulas
- **Facilitator Guide:** [`Facilitator_Guide.pdf`](documents/instructor/Facilitator_Guide.pdf) · [answer key](documents/instructor/answer_key.md)

---

## Topics Covered

- Securing raw data and keeping a chain of custody
- Surveying a dataset with filters to spot data quality issues
- Removing junk rows and columns; Excel Tables; Remove Duplicates
- Cleaning spaces and non-breaking spaces in helper columns (`TRIM`, `SUBSTITUTE`, `UNICHAR`, `LEN`)
- Standardizing categories with a reference table (`XLOOKUP`, `TEXTBEFORE`)
- Rebuilding text dates with `DATE` and flagging impossible values

---

## Partners / Series

This workshop is hosted by the **Digital Research and Innovation Lab**.  
Content was developed and delivered by the **Geospatial Research Lab**.  
It is a companion to [*Tatort: Excel* (Excel for the web, self-paced)](https://github.com/GeospatialResearchLab/workshop-data-cleaning-excel-tatort-2026-web), [*Spark Joy in Your Data*](https://github.com/GeospatialResearchLab/workshop-data-cleaning-excel-2025) and [*AI-Assisted Strategies for Cleaning and Preparing Research Data*](https://github.com/GeospatialResearchLab/workshop-data-cleaning-GenAI-R-2025).

---

## Citation

Breen, J. (2026). *Tatort: Excel — A Forensic Guide to Data Cleaning* (workshop materials).  
Geospatial Research Lab, American University Library.  
https://github.com/GeospatialResearchLab/workshop-data-cleaning-tatort-2026

---

## License

Workshop materials are licensed under [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). The *Tatort* episode data is adapted from Wikipedia and remains under [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/).
