# OM MARKETING — Receipt & Party Portal

Mobile-first invoice search, party ledger, and payment collection app hosted on GitHub Pages.

## What's New & Architecture

### ⚡ Instant Client-Side Search (Single-Fetch + IndexedDB)
- **Zero-Lag Search**: Instead of waiting 10–15 seconds for Google Apps Script on every keystroke, the web app loads all invoice data once into high-speed browser storage (`IndexedDB` + in-memory cache).
- **Offline / Instant Ready**: App opens in under 100ms. Subsequent visits load in 0ms from local IndexedDB cache.
- **🔄 Live Sheet Sync**: A "Sync" button in the top bar pulls fresh updates from the Google Sheet (`?action=fetchAll`) whenever needed.

### 👤 Grouped Party View (No Duplicate Invoice Clutter)
- **Unique Party Cards**: Searching by customer name now groups invoices by unique Party. Instead of seeing the same shop repeated 30–50 times, you see **one clean card** per party.
- **Party Summary Cards**:
  - Clean Party Name, Shop Code (`#ID`), Beat/Area, and Agent.
  - Total Invoices count & Total Billed Amount.
  - Live Outstanding Balance badge (e.g. `₹11,908 Pending (6 unpaid)` or `All Settled`).
- **Drill-down to Party Ledger**:
  - Tap any party card to open **All Invoices** for that specific customer.
  - Filter by `All`, `⚠️ Pending Only`, or `✅ Paid Only`.
  - Search within that party's invoices.
  - Tap any invoice to view details or record a payment.

### 🔢 Direct Invoice Lookup
- Dedicated tab to search directly by invoice number (e.g., partial match `5503`, `FY26`, `17063`).

### 💰 Split Payment & Live Updates
- Record split **Cash + Bank** payments, receipt number, and remarks.
- Quick 1-tap fill buttons (`⚡ Fill Full in Cash` / `⚡ Fill Full in Bank`).
- Live remaining balance preview.
- When submitted:
  - Updates the invoice in Google Sheets and logs to `PAYMENT_LOG`.
  - Immediately updates the local IndexedDB and memory cache, recalculating party outstanding balance in real time without needing a page refresh!

---

## Google Apps Script Setup (`Code.gs`)

1. Open your Google Sheet: `ALL INVOICE PARTY (OM MARKETING) - MARCH-SEPT`.
2. Go to **Extensions** → **Apps Script**.
3. Replace the script code with [Code.gs](./Code.gs).
4. Click **Deploy** → **Manage deployments** → **Edit (pencil icon)**.
5. Under **Version**, select **New version**, then click **Deploy**.
6. Ensure the Web App URL matches `API_URL` in [index.html](./index.html).
