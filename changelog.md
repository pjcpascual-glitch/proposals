# Changelog

All notable changes to the **ATI Offer & Tender Tracker** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to Semantic Versioning.

---

## [2.2.0] - 2026-11-11

### Added
- **Browser-Native HTML Auto-Bake Engine**:
  - Implemented `autoBakeHtmlSnapshot()` inside `ADM_APP_ATI_000 Proposals.html`: When linked via **"Link Workspace Folder"**, any proposal creation, status edit, or Excel sync automatically rewrites the embedded `<script id="prebaked-data">` on disk, ensuring viewers in Incognito and executives without linked files always see the latest data.
- **Dynamic HTML File Name Resolution**:
  - Updated `sync_excel.py` and `server.py` to dynamically detect and support `ADM_APP_ATI_000 Proposals.html` as well as legacy names.
- **Revenue-Weighted Performance Mode (Chart 2)**:
  - Added an interactive toggle between **`Won EUR (€)`** and **`Win Rate (%)`** on Chart 2 (*Top Users by Won Revenue / Rate of Success*).
  - Default view ranks team members by **Total Won Revenue (€ EUR)** to properly credit the primary commercial drivers lifting the firm (e.g. PPC with €890k and DCS with €320k in 2026).
  - Each bar dynamically displays both the won value and the conversion efficiency (e.g., `€890.3k (19%)`).
- **Interactive Leaderboard Column Sorting**:
  - Clickable headers on the **User Performance & Success Rates Leaderboard** allowing instant ascending/descending sorting by *Responsible*, *Total Offers*, *Awarded (Won)*, *Ongoing*, *Rejected / Lost*, *Won Value (€ EUR)*, *Ongoing Pipeline (€)*, or *Win Rate (% of Sent)*.
  - Set default leaderboard sorting to **`Won Value (€ EUR)`** to reflect commercial impact.
- **Low-Sample Qualification Badges**:
  - Users with fewer than 3 proposals (such as 1-of-1 single wins) are flagged with an `(< 3 offers)` badge and prioritized below active team members when sorting by percentage.
- **Developer Attribution & Support Footnote**:
  - Added formal attribution in the dashboard footer and Sync Drawer:
    > *“This tool has been coded by ppascual@aticonsult.es. Any bugs or query needs to be routed to him.”*
  - Includes a direct `mailto:ppascual@aticonsult.es` link.
- **Knowledge Documentation**:
  - Created `changelog.md` to track version history and feature rollouts.
  - Created `preference.md` to capture engineering learnings, statistical formulas, and design solutions that work.

### Changed
- **Wide Container Layout (Horizontal Scroll Elimination)**:
  - Expanded dashboard max-width from `max-w-7xl` (1280px) to **`max-w-[1750px] 2xl:max-w-[1850px]`** across the header, main container, and footer.
  - Adjusted proposal table description min-widths from `min-w-[220px]` to `min-w-[170px]`.
  - Result: All 10 table columns (from Tender ID through Actions) are fully visible without requiring horizontal scrolling on standard and widescreen displays.
- **Fair Win Rate Formula (Option A)**:
  - Replaced the flawed ratio `Won / (Won + Rejected)` with:
    $$\text{Win Rate \%} = \frac{\text{Awarded (Won)}}{\text{Total Proposals Sent}} \times 100\%$$
  - Pending proposals (`Issued` / `Preliminary`) and cancelled tenders now properly count toward the total effort until converted into wins.
  - Fixed PSP anomaly where cancelled tenders caused an artificial 100% win rate; PSP now accurately reflects 45.5% (5 won of 11 sent).
- **KPI Card #5 Updated**:
  - Re-labeled from *"Won vs Decided Offers"* to **`Win Rate (Conversion): Won vs Total Proposals Sent`**.
- **Python CLI Tool (`sync_excel.py`)**:
  - Updated `py sync_excel.py --stats` to calculate win rates against total proposals sent.
  - Added robust numeric parsing (`parseFloat(item.eur) || 0`) to handle non-numeric Excel cells like `'N/A'`.

---

## [2.1.0] - 2026-11-05

### Added
- **Direct Browser File System Access API**:
  - Browser directly connects and writes back to `REG_16_Tender List_v1.xlsx` via native Chromium File System API.
  - Handle persistence across browser sessions using IndexedDB (`OfferTrackerV2_DB`).
- **Zero-Click Executive Snapshot Mode**:
  - Self-contained HTML bundle with pre-baked JSON data embedded inside `<script id="prebaked-data" type="application/json">`.
  - Executives can double-click `Offer Dashboard.html` to view all KPIs and charts with zero setup, zero server, and no file pickers.
- **One-Click Re-Bake CLI**:
  - Added `py sync_excel.py --bake` to synchronize the latest Excel data into `Offer Dashboard.html` in under one second.
- **Follow-Up Radar (Action Center)**:
  - Tab 3 prioritized list of active proposals (`Issued` and `Preliminary`) with overdue alerts and `< 7 days` remaining warnings.
- **Quick Edit Modal**:
  - Instant in-browser status updating (`Issued` ➔ `Accepted` / `Rejected`), comments editing, and decision tracking.

---

## [2.0.0] - 2026-10-20

### Added
- Complete rewrite of the legacy spreadsheet tracker into a responsive Single-Page Application (SPA).
- Modern UI using Tailwind CSS CDN, Chart.js 4, and Inter typography.
- 5 high-level KPI summary cards.
- 6 interactive visual analytics charts (Status Donut, Success Rate Bar, Follow-Up Radar, Sector Pipeline, Top Clients, Country Breakdown).
- Proposals spreadsheet registry with search, filter, and pagination.
- + New Proposal modal with automatic ID incrementing and currency conversion to EUR.

---

## [1.0.0] - 2017 to 2026

### Initial Release
- Legacy master Excel workbook `REG_16_Tender List_v1.xlsx` tracking historical engineering tenders and proposals.

