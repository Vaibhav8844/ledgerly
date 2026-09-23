# Ledgerly Portfolio — V2 (Google Sheets + Apps Script)

Ledgerly is a responsive personal portfolio dashboard for Indian stocks and mutual funds. This version removes the Express/SQLite backend and uses **Google Sheets as the data store** and **Google Apps Script as the API/backend**.

## Architecture

```text
React + Vite dashboard
        ↓ HTTPS
Google Apps Script Web App
        ↓
Google Sheets
```

## What is included
- Responsive React/Vite frontend for desktop and mobile
- Google Sheets database template/schema
- Google Apps Script API (`apps-script/Code.gs`)
- Manual BUY/SELL transaction entry
- Independent purchase lots
- FIFO sell validation and lot depletion
- Realized P&L calculation from FIFO lots
- Portfolio value, invested amount, unrealized P&L and return
- Holdings, transactions and analytics views
- Historical portfolio snapshots
- Mutual-fund CSV/TSV import using Scheme Name, AMC, Category, Sub-category, Folio No., Source, Units, Invested Value, Current Value, Returns and XIRR
- Manual quote endpoint for future market-price integrations
- Empty-start state: no sample financial data
- No browser localStorage as the source of truth

## Setup

### 1. Create the Google Sheet
Create a new Google Spreadsheet, for example `Ledgerly Portfolio DB`.

Open **Extensions → Apps Script** and paste the contents of `apps-script/Code.gs` into `Code.gs`.

Run `setupLedgerlySheet()` once and approve the requested Google permissions. It creates:

- `Transactions`
- `Quotes`
- `Snapshots`
- `Lots`
- `Settings`

### 2. Deploy Apps Script as a Web App
In Apps Script:

**Deploy → New deployment → Web app**

Use an access setting that allows your dashboard to call the web app. For a personal deployment, restrict access as much as your Google account/environment permits while still allowing the frontend to reach it.

Copy the `/exec` deployment URL.

### 3. Configure the React app
Copy `client/.env.example` to `client/.env` and set:

```env
VITE_API_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
```

### 4. Run the frontend
Requirements: Node.js 20+

```bash
npm install
npm run dev
```

The dashboard will use Apps Script/Sheets instead of a local server/database.

## API actions

The Apps Script web app supports:

- `GET ?action=portfolio`
- `GET ?action=transactions`
- `GET ?action=quotes`
- `GET ?action=lots`
- `POST {"action":"transaction", ...}`
- `GET ?action=deleteTransaction&id=TXN-...`
- `POST {"action":"quote", "symbol":"RELIANCE", "price":1500}`
- `POST {"action":"snapshot"}`
- `POST {"action":"importMutualFunds","rows":[...]}`

### Import mutual funds

Use the **Import MF** button in the dashboard and select an Excel workbook (`.xlsx`), CSV, or tab-separated file. Excel imports read the first worksheet. The file must contain this header row:

```text
Scheme Name,AMC,Category,Sub-category,Folio No.,Source,Units,Invested Value,Current Value,Returns,XIRR
```

The import replaces the previous `MutualFunds` sheet contents. Invested and current values are added to the portfolio totals; mutual-fund P&L is calculated as current value minus invested value.

## Lot model
Every BUY is a separate lot with its own date, quantity and purchase price. SELL transactions consume available lots in FIFO order. The transaction ledger is the source of truth; `Lots` is a derived/rebuilt view.

## Important note about market prices
V2 does not include a live market data provider yet. If a quote exists in the `Quotes` sheet, the dashboard uses it for current value. Otherwise it uses invested cost as a conservative fallback so the dashboard never invents a market price.

## Next integrations
1. Live NSE/BSE price feed
2. MF NAV feed
3. Groww/email/CAS transaction import
4. XIRR
5. Tax-lot reporting
6. Dividends, splits, bonuses and other corporate actions
7. Optional authentication/hardened deployment

## CORS / deployment note
Google Apps Script web apps are subject to Google's web-app redirect and browser cross-origin behavior. If a separately hosted React site reports a CORS error when calling the Apps Script URL, the most reliable deployment is to host the frontend through a Google/Apps Script-compatible same-origin layer or add a small proxy. The portfolio logic and Sheets backend do not need to change.
