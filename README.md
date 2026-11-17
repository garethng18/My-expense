# My expense

Daily Expense Tracker — a Google Apps Script web app backed by Google Sheets.

## Live app

https://script.google.com/macros/s/AKfycbxMmF_E3Q1eZ2IROSiPAmQUsqPdlDaT6lm13CXroHwy2uizp-qpZiwM5nvK94UE-gyO/exec

## Project files

| File | Role |
|------|------|
| `Code.gs` | Server-side Apps Script (dashboard data, add/update expenses) |
| `Index.html` | Web app UI (budget overview, logging form, recent entries, 3-month summary) |

## Setup (Google Apps Script)

1. Create a Google Sheet with sheets named **Expenses** and **Budget**.
2. Open **Extensions → Apps Script** and paste `Code.gs` and `Index.html`.
3. Deploy as a **Web app** (execute as you; access as appropriate for your use).
4. On the **Budget** sheet, set the target month in `A2` and budget amount in `B2`.

### Expenses columns

Date | Category | Amount | Items | Significant | Record timestamp
