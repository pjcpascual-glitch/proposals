# ATI Offer & Tender Tracker — Dashboard v2.1

A modern, high-performance web dashboard for engineering proposals, tender pipeline analytics, follow-up radar, and win-loss tracking. Built for ATI Consultores & Konsulta.

---

## 🌟 Key Features

1. **Zero-Click Executive Viewing (Option 1)**:
   - Executives and bosses can simply **double-click [`Offer Dashboard.html`](Offer%20Dashboard.html)** in any browser.
   - All 550+ proposals, KPIs, 6 interactive charts, tables, and leaderboards load **immediately** with zero prompts, zero permissions, and no server running.

2. **Direct Browser-to-Excel Sync (No Server Needed)**:
   - When administrators click **"Link Excel"**, the browser securely connects directly to [`REG_16_Tender List_v1.xlsx`](REG_16_Tender%20List_v1.xlsx) via the Chrome/Edge **File System Access API**.
   - Handles are remembered across sessions in **IndexedDB** (`OfferTrackerV2_DB`).

3. **Two-Way Proposal Creation & Editing**:
   - **+ New Proposal Modal**: Auto-generates the next sequential ID (`26071`), auto-converts foreign currencies to EUR (`USD/1.15`, `PHP/70`, `CFA/655.957`), and appends the row directly into Excel sheet `REG_16_Tender List_v1` using native formulas (e.g. `=I557/1.15`).
   - **Edit Proposal & Status Modal**: Advance proposals (`Issued` ➔ `Accepted` / `Rejected`), update decision dates, add follow-up notes, or edit values and scope directly from the UI.

4. **Follow-Up Radar (Action Center)**:
   - Identifies and prioritizes active tenders (`Issued` & `Preliminary`) whose expiration/decision dates are overdue or approaching within 7 days.

5. **User Success Leaderboard**:
   - Performance ranking of all engineers and responsibles (`AMF`, `DCS`, `PPC`, `GAS`, etc.) with win rate percentages, total won EUR, and pipeline under management.

---

## 📂 File Structure

| File | Purpose |
| :--- | :--- |
| **`Offer Dashboard.html`** | Main dashboard application (Self-contained, includes embedded pre-baked data snapshot). |
| **`REG_16_Tender List_v1.xlsx`** | Master Excel tender database. |
| **`sync_excel.py`** | Automated Python CLI helper script for validation, JSON cloud backup, and snapshot baking. |
| **`server.py`** | Optional local development server with CORS and REST API endpoints. |
| **`Launch_Offer_Tracker.vbs`** | Silent Windows desktop launcher (launches server in background and opens browser). |
| **`Launch_Offer_Tracker.bat`** | Standard batch script launcher. |
| **`offer_tracker_backup.json`** | Structured JSON database for cloud backup & merge synchronization. |

---

## 🚀 How to Use

### For Executives / Bosses (View-Only / Read Mode)
* Simply **double-click `Offer Dashboard.html`**.
* The app opens immediately in Chrome or Edge with all proposals, KPIs, and graphs populated.

### For Administrators (Adding & Editing Proposals)
1. Open `Offer Dashboard.html`.
2. Click **"Link Excel"** in the top bar and select `REG_16_Tender List_v1.xlsx` once.
3. Use **"+ New Proposal"** to add proposals, or click **"Edit"** on any proposal row to update status or details.
4. Changes are written directly to the `.xlsx` file on disk.

### Keeping the Executive Snapshot Updated
Whenever you update the Excel file, run:
```bash
py sync_excel.py --bake
```
*This extracts the latest Excel rows and bakes them into `Offer Dashboard.html` in under 1 second.*

Or inside `Offer Dashboard.html`, open **Settings** (gear icon) and click **"Export Self-Contained HTML for Bosses"**.

---

## 🛠️ CLI Utilities (`sync_excel.py`)

```bash
# View summary statistics & overdue proposals in the terminal
py sync_excel.py --stats

# Re-bake the latest Excel data into Offer Dashboard.html
py sync_excel.py --bake

# Bidirectional sync between Excel and JSON backup
py sync_excel.py --sync

# Export Excel data to offer_tracker_backup.json
py sync_excel.py --export-json

# Run local development server on port 8000
py sync_excel.py --serve
```

