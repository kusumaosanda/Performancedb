# Dashboard Setup Guide
Free · Public · 24/7 · No Mac needed

---

## PART 1 — Set up Google Sheets (your data source)

### Step 1 — Create the spreadsheet

1. Go to **sheets.google.com** → click **"Blank spreadsheet"**
2. Rename it (click "Untitled spreadsheet" at top) → e.g. **Dashboard Data**
3. You need **3 tabs**. Rename the default tab and add two more:

| Tab name (exact) | What goes here |
|---|---|
| `HOME` | Project identity info |
| `KPIS` | KPI numbers |
| `PROJECTS` | Project list |

To rename a tab: right-click the tab at the bottom → Rename

---

### Step 2 — Fill in the HOME tab

Columns: **Type** | **Label** | **Value**

| Type | Label | Value |
|------|-------|-------|
| meta | dashboard_title | My Dashboard |
| meta | dashboard_sub | Management Portal |
| meta | avatar | 🏢 |
| meta | name | YOUR PROJECT NAME |
| meta | subtitle | Division / Department |
| meta | status | Active |
| meta | description | Brief description of this project or entity. |
| meta | tags | Tag1, Tag2, Tag3 |
| field | Project Code | PRJ-001 |
| field | Manager / PIC | Your Name |
| field | Location | City, Country |
| field | Start Date | 01 Jan 2025 |
| field | End Date | 31 Dec 2025 |
| field | Client | Client Name |
| field | Contract Value | RM 5,000,000 |
| field | Phase | Construction |

**Rules:**
- `Type = meta` → special config fields (dashboard title, status, etc.)
- `Type = field` → shows up in the info grid on the dashboard
- Add as many `field` rows as you want
- `status` options: `Active` · `On Hold` · `Complete` · `Inactive`

---

### Step 3 — Fill in the KPIS tab

Columns: **Label** | **Actual** | **Target** | **Prefix** | **Suffix** | **Trend** | **Delta** | **Icon**

| Label | Actual | Target | Prefix | Suffix | Trend | Delta | Icon |
|-------|--------|--------|--------|--------|-------|-------|------|
| Overall Progress | 67 | 100 | | % | up | 5% | 📈 |
| Budget Used | 1250000 | 2000000 | RM | | flat | | 💰 |
| Tasks Completed | 34 | 50 | | | up | 8 | ✅ |
| Issues Open | 7 | 0 | | | down | 3 | ⚠️ |

**Rules:**
- `Actual` and `Target` must be numbers only (no RM, no %, no commas)
- `Trend`: `up` · `down` · `flat` (controls the arrow colour)
- `Prefix` / `Suffix`: text shown before/after the number (e.g. `RM` or `%`)
- `Icon`: any emoji

---

### Step 4 — Fill in the PROJECTS tab

Columns: **Name** | **Status** | **Priority** | **PIC** | **Due Date** | **Progress** | **Category**

| Name | Status | Priority | PIC | Due Date | Progress | Category |
|------|--------|----------|-----|----------|----------|----------|
| Site Preparation | Completed | High | Ahmad | 2025-03-31 | 100 | Civil |
| Foundation Works | In Progress | High | Syafiq | 2025-08-15 | 65 | Structural |
| MEP Rough-In | Not Started | Medium | Lim CW | 2025-10-01 | 0 | M&E |

**Status options (exact spelling):**
`Not Started` · `In Progress` · `Completed` · `On Hold` · `Cancelled`

**Priority options:** `High` · `Medium` · `Low`

**Due Date format:** `YYYY-MM-DD` (e.g. `2025-12-31`) — overdue items will show a ⚠ warning

**Progress:** number 0–100 (no % sign)

---

### Step 5 — Publish the sheet to the web

1. In your Google Sheet: **File → Share → Publish to web**
2. Set "Link" to **Entire Document**, format to **Web page**
3. Click **Publish** → confirm
4. Close the dialog (you don't need the published URL)

> This makes your data publicly readable. The sheet is still only editable by you.

---

### Step 6 — Get your Sheet ID

Look at your Google Sheet URL:

```
https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit
                                        ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
                                        this is your Sheet ID
```

Copy everything between `/d/` and `/edit`.

---

### Step 7 — Paste the Sheet ID into index.html

Open `index.html` in any text editor (Notepad, TextEdit, VS Code).

Find this line near the top of the `<script>` section:

```javascript
const GOOGLE_SHEET_ID = 'YOUR_SHEET_ID_HERE';
```

Replace `YOUR_SHEET_ID_HERE` with your actual Sheet ID:

```javascript
const GOOGLE_SHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms';
```

Save the file.

---

## PART 2 — Deploy to GitHub Pages (free public hosting)

### Step 8 — Upload to GitHub

1. Go to **github.com** → log in → **New repository**
2. Name it `dashboard` → set **Public** → **Create repository**
3. Click **"uploading an existing file"**
4. Upload `index.html` (the data folder is no longer needed)
5. Click **Commit changes**

### Step 9 — Enable GitHub Pages

1. In your repo → **Settings** → **Pages** (left sidebar)
2. Branch: `main` → folder: `/ (root)` → **Save**
3. Wait ~1 minute → your URL appears at the top

✅ Your dashboard is now live at: `https://YOUR-USERNAME.github.io/dashboard/`

---

## Updating data (from anywhere, no Mac needed)

Just open your Google Sheet and edit it — the dashboard pulls fresh data every time someone opens it.

- Edit KPIs → update the `KPIS` tab
- Add/update projects → update the `PROJECTS` tab
- Change project info → update the `HOME` tab

No need to redeploy. Changes show up immediately.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Dashboard shows "Sheet not public" | Go to File → Share → Publish to web and publish again |
| Data not showing | Check tab names match exactly: `HOME`, `KPIS`, `PROJECTS` |
| KPI numbers wrong | Make sure Actual/Target columns have numbers only (no commas or symbols) |
| Project dates not showing | Use format `YYYY-MM-DD` (e.g. `2025-12-31`) |
