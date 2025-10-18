# Innovation Vitals Tool – Google Apps Script

This repository contains the Google Apps Script code developed to extend the functionality of the Innovation Vitals digital tool in Looker Studio.

## Purpose
The script automates the interaction between **Looker Studio** and **Google Sheets**, enabling users to:
- Add or remove indicators to an indicator selection associated with specific projects from the indicators database directly in Looker Studio;
- Export selected indicators (with sub-indicators, scales, and source articles) as CSV files.

## Main Functions
- `doGet(e)` — Handles user requests triggered from Looker Studio (add, remove, clear, export).
- `exportFull_(project)` — Generates downloadable CSV files with selected indicators.
- `openSheet_(id, name)` — Connects to the appropriate Google Sheets files.
- `htmlMessage_(msg)` — Displays confirmation messages.
