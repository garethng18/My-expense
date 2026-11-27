# My expense

Daily Expense Tracker — a Google Apps Script web app backed by Google Sheets.

## Live app

https://script.google.com/macros/s/AKfycbxMmF_E3Q1eZ2IROSiPAmQUsqPdlDaT6lm13CXroHwy2uizp-qpZiwM5nvK94UE-gyO/exec

## Project files

| File | Role |
|------|------|
| `Code.gs` | Server-side Apps Script (dashboard data, add/update expenses, Fin dashboard) |
| `Index.html` | Daily expense UI |
| `Fin.html` | Monthly Fin grid. Open it from the Daily page, or add `?page=fin` to the web app URL. |

## Setup (Google Apps Script)

1. Create a Google Sheet with sheets named **Expenses** and **Budget**.
2. Open **Extensions → Apps Script** and paste `Code.gs` and `Index.html`.
3. Deploy as a **Web app** (execute as you; access as appropriate for your use).
4. On the **Budget** sheet, one row per month: Month (`yyyy-MM`), Target Budget, Actual Expense, Extra Expense. The monthly tiles read that row.

### Expenses columns

Date | Category | Amount | Items | Significant | Record timestamp

### Day templates

The first dashboard load creates a **Templates** sheet. It stores three day setups. Each setup holds expenses you can add or remove. Applying a template appends every saved expense to the Expenses sheet for the date you pick. Redeploy the web app after updating `Code.gs` and `Index.html`.
