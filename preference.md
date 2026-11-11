# Engineering Preferences & Key Learnings

This document records architectural principles, analytical rules, and coding patterns that have been proven to work for the **ATI Offer & Tender Tracker**. It serves as a persistent guide for future enhancements and bug fixes.

---

## 1. Pipeline Analytics & Statistical Fairness

### ⚠️ The "Law of Small Numbers" Fallacy
* **Problem**: In raw win-rate calculations ($\frac{\text{Won}}{\text{Total}}$), someone who drafts **1 single small proposal** that happens to win (e.g. a €2,800 repeat order) achieves a **100.0% win rate**. Under unweighted ranking, they rank #1 above senior team members who manage dozens of complex tenders (e.g. AMF with 100 wins out of 252 proposals, or PPC with 57 wins).
* **Proven Solution**:
  1. **Minimum Sample Qualification**: Always require a minimum threshold ($\ge 3$ or $\ge 5$ proposals in scope) before an engineer qualifies for ranking in win-rate charts.
  2. **Sample Badging**: Mark low-sample users with an `(< 3 offers)` badge on leaderboards rather than presenting their 100% at face value.
  3. **Tie-Breaking Order**: Always break percentage ties using:
     1. Number of Awarded Deals (`won`)
     2. Total Won Value in EUR (`wonEur`)

### 📐 Denominator Selection: Option A (Total Sent vs Decided)
* **Problem**: Calculating win rate as $\frac{\text{Won}}{\text{Won} + \text{Lost}}$ ignores active/in-review proposals (`Issued`, `Preliminary`) and `Cancelled` proposals. An engineer with 1 win and 4 proposals still awaiting client decision was showing as $\frac{1}{1 + 0} = 100\%$, which misrepresents reality.
* **Proven Solution (Option A)**:
  $$\text{Win Rate \%} = \frac{\text{Awarded (Won)}}{\text{Total Proposals Sent}} \times 100\%$$
  * All proposals released into the market count against the conversion until they are officially awarded.
  * Inactive/cancelled bids are accounted for, preventing artificial 100% scores (e.g. PSP had 5 won and 6 cancelled; under Option A this correctly shows as 45.5% rather than 100%).

---

## 2. Commercial Ranking: Revenue Impact vs. Conversion Efficiency

### 💰 Why Commercial Leaders Must Be Ranked by Won Value
* In B2B engineering consultancy and EPC contracting, the team members "lifting the firm" are those bringing in millions of euros in closed contracts, even if their competitive tender conversion rate is 15%–25%.
* **Example (2026 Pipeline)**:
  * **PPC**: Won €890k across 7 major contracts + manages €2.97M in active pipeline (19% count win rate).
  * **DCS**: Won €320k across 3 contracts + manages €845k in active pipeline (21% count win rate).
  * Together, PPC and DCS generated **over 96% of the company's won revenue in 2026**.
  * Ranking AMC (1 deal, €22.5k) as #1 above PPC and DCS based solely on count percentage obscured the firm's primary commercial drivers.
* **Proven Solution**:
  1. **Chart 2 Toggle (`Won EUR (€)` vs `Win Rate (%)`)**:
     * Default Chart 2 to **`Won EUR (€)`** so executive viewers immediately see who brings in the revenue.
     * Always annotate the bars with both metrics: `€890.3k (19%)`.
  2. **Default Leaderboard Sorting by Won Value (`wonEur`)**:
     * Set default leaderboard sort to `Won Value (€ EUR)`. PPC & DCS naturally head the 2026 leaderboard, and AMF (€24.1M) & PPC (€5.78M) head the all-time board.
     * Keep column headers clickable for interactive re-sorting.

---

## 3. UI Layout & Responsive Sizing

### 🖥️ Widescreen Container Optimization
* **Problem**: Standard Tailwind `max-w-7xl` caps dashboard width at `1280px` (`80rem`). With side padding (`px-4 sm:px-6 lg:px-8`), the content box is reduced to `1216px`. A 10-column spreadsheet table requires ~`1150px`–`1250px`. Consequently, on standard 1080p desktop monitors (especially with Windows scaling at 125% or 150%), the table overflowed horizontally, hiding the **Actions** column or clipping **Tender ID**.
* **Proven Solution**:
  * Set the primary container to **`max-w-[1750px] 2xl:max-w-[1850px] mx-auto`** across `<header>`, `<main>`, and `<footer>`.
  * Reduce description column constraints from fixed `min-w-[220px]` to flexible `min-w-[170px] max-w-sm break-words`.
  * Result: Tables expand across wide viewports without horizontal scrollbars, keeping Tender ID and Action buttons simultaneously in view.

---

## 4. Safe Data Handling & Data Types

### 🔢 Excel Non-Numeric Cell Sanitation
* **Problem**: Excel data imported into JSON frequently includes non-numeric string representations (e.g. `'N/A'`, `'-'`, `''`, `null`). Using raw JavaScript arithmetic (`item.eur || 0`) leads to string concatenation bugs:
  ```javascript
  // Bug: 'N/A' + 50000 === 'N/A50000'
  userStats[u].wonEur += (item.eur || 0);
  ```
* **Proven Solution**:
  * Always coerce financial values with `parseFloat`:
    ```javascript
    const eur = parseFloat(item.eur) || 0;
    userStats[u].wonEur += eur;
    ```
  * In Python (`sync_excel.py`), always wrap numeric extraction in a `try...except ValueError` block:
    ```python
    try:
        eur = float(d.get('eur') or 0)
    except (ValueError, TypeError):
        eur = 0.0
    ```

---

## 5. Architecture & Zero-Server Deployment

### 📦 The Dual-Mode Synchronization Architecture
* **Mode 1: Zero-Click Executive Viewing (Self-Contained Snapshot)**:
  * Executives and management will not run local Python servers or link file pickers.
  * Storing the latest database snapshot inside `<script id="prebaked-data" type="application/json">` allows `Offer Dashboard.html` to be opened via simple double-click on any machine.
* **Mode 2: Chromium File System Access API with IndexedDB**:
  * For administrators editing data directly in the browser without installing a backend server.
  * Saving the file handle in `indexedDB` (`OfferTrackerV2_DB`) allows subsequent sessions to write directly to `REG_16_Tender List_v1.xlsx` with zero backend dependencies.
* **CLI Baking**:
  * `py sync_excel.py --bake` replaces the embedded JSON in `Offer Dashboard.html` in under 1 second whenever the Excel workbook is modified externally.

---

## 8. Git Commit & Versioning Conventions

### 🏷️ Version-Prefixed Commit Message Standard
* **Format**: All Git commits and pushes must strictly follow the user's versioned standard:
  ```text
  v<Version>: <Concise Description of Changes>
  ```
  * **Examples**:
    * `v2.2.0: Fair Win-Rate Engine, Revenue-Weighted Rankings & Layout Expansion`
    * `v0.4.33: Common Project Path Parser & Standardized Path Autofill Engine`
* **Rule**: Never use conventional commit prefixes (e.g. `feat:`, `chore:`, `fix:`) — always use the `vX.Y.Z: Title` format.
* **Branches**: Support both `master` and `main` branches as required by the user's repository settings.


### 🔒 Why Incognito Loads Pre-Baked Snapshots (and How to Keep It Fresh)
* **The Browser Privacy Sandbox**:
  * In Chrome and Edge, **Incognito windows strictly isolate storage**: IndexedDB handles, local permissions, and cookies are deliberately discarded when an Incognito session begins.
  * When you open the HTML app in Incognito, the browser cannot access the administrator's previously granted file handle to `REG_16_Tender List_v1.xlsx`.
  * The app gracefully falls back to the embedded `<script id="prebaked-data" type="application/json">` block.
* **Why Normal Browser Sessions Don't Automatically Rewrite the HTML File by Default**:
  * When clicking **"Link Excel"**, Chrome's File System API grants permission to read/write **only the `.xlsx` file**.
  * Chrome strictly forbids a web page from writing to any other file on disk (including the `.html` file itself) without explicit folder permission.
  * Therefore, edits made in your browser session were saved into `REG_16_Tender List_v1.xlsx`, while the static text inside the `.html` file remained unchanged until re-baked.
* **The Solution (Auto-Bake via Folder Link)**:
  * Linking the workspace folder via **"Link Workspace Folder"** (`showDirectoryPicker`) grants permission to both the Excel file and the HTML file.
  * The app now contains an **`autoBakeHtmlSnapshot()`** engine: Whenever an administrator adds a proposal, updates a status, or clicks "Sync", the browser automatically rewrites `<script id="prebaked-data">` inside `ADM_APP_ATI_000 Proposals.html` on disk.
  * Alternatively, running `py sync_excel.py --bake` re-bakes the latest Excel data into `ADM_APP_ATI_000 Proposals.html` in under 1 second.

* **Primary Code Owner**: `ppascual@aticonsult.es`
* **Routing Policy**: Any questions, feature enhancements, or bug reports regarding calculation formulas, sync scripts, or layout adjustments should be directed to the code owner.
* **UI Placement**: Ensure the contact footnote remains permanently visible in the dashboard `<footer>` and inside the **Cloud Sync & Helper Hub** drawer across all release updates.

