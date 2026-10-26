# Prompt: EPM to CLH feature dependency dashboard and GitHub hosting files

Copy everything below the line into a new session and attach the latest **Client Directory** workbook (.xlsx).

---

## Role and goal

You are building an executive-ready, Elsevier-branded HTML dashboard for a Customer Readiness Manager. It shows why building one feature doesn't move EPM databases to CLH: a database can migrate only when every feature it relies on exists in CLH, and most databases need several.

Deliver two things:

1. **One self-contained HTML file** (`index.html`) with the data, fonts and icon built in. It must open by double-clicking and also work when hosted on GitHub Pages.
2. **A small set of GitHub hosting files** that go with it.

Apply the Elsevier brand skill (colors, Tiempos Text and National 2 fonts, logo, accessibility). Write in plain, professional language for an executive audience. Use sentence case for headings. Never use the word "customer" as a counting unit; count **databases** only.

## Input

The attached Client Directory workbook, tab **"Main - Customer Directory"**. Find columns by header name, not position (columns move between versions). Read in read-only mode and stop after about 50 consecutive blank rows (the sheet has ~1M formatted empty rows).

Columns used: Account, ECR ID, Database Name, CMX ID, Contract Value, Phase, Migration Status, CLH Customer Account Status (Active or Canceled), Client Type, Number of Active Learners, Number of Active Admins, Country, SSO Eligibility, Functionality: Mosby Skills, Functionality: Automated User Upload (HR Download), Functionality: SCORM Uploader (SCO12 file type), Multi-database (Org Picker), Linked Tests, LinkedContent (Tests), Self-Registration Link on Login Page, Self-Enroll: eLearning/Skills Turned On, CK Nursing Subscription, CK Nursing Link on Skills Instance, Change Type.

## Scope rules

1. **Unit:** one record per CMX ID (one row per database). Skip blank or placeholder CMX IDs (NULL, N/A, NA, New Customer, TTP Cut Over, TTP Only). If a CMX ID repeats, keep the first.
2. **Workbook filters:** if the main tab has Excel AutoFilter criteria set, apply exactly those (compare trimmed values; respect "blanks" being allowed).
3. **Otherwise use these fallback rules,** which mirror the filters used previously:
   - Exclude Phase: Cancel/Remove, Cancelled, Internal, NA, Data Moved to CLH, DirectCourse.
   - Exclude Migration Status: Complete, NA, Net New, TTP Cutover, Handoff to AM.
4. **Also exclude:**
   - Databases marked Canceled/Cancelled under CLH Customer Account Status or Change Type.
   - Client types Workforce Development, Internal MCS Site, Connect Partner, EPM Connect, LMS Connect and Clinical Skills Connect.
5. Blank contract value, learners or admins count as 0. Always state how many databases have a recorded contract value.

## The six features (gates)

A feature is needed when its column starts with "Yes".

| Key | Label | Source column(s) |
|---|---|---|
| orgpicker | Multi-database (Org Picker) | Multi-database (Org Picker). Do not use "Multiple Databases (Old Def)". |
| hr | HR Download (automated user management) | Functionality: Automated User Upload (HR Download) |
| selfreg | Self-registration / self-enrollment | Self-Registration Link on Login Page OR Self-Enroll: eLearning/Skills Turned On |
| scorm | SCORM Uploader (SCO12) | Functionality: SCORM Uploader (SCO12 file type) |
| linked | LinkedContent / Linked Tests | LinkedContent (Tests) OR Linked Tests |
| ck | CK Nursing (subscription / Skills link) | CK Nursing Subscription OR CK Nursing Link on Skills Instance |

**Dependency:** Org Picker depends on HR Download. Whenever a build order is calculated, Org Picker is built together with HR Download in a single step.

**Build order:** a greedy order that, at each step, adds the feature (plus its dependencies) that releases the most additional databases.

For each database, also keep Country, SSO Eligibility and whether it is flagged for Mosby Skills.

## Page layout (in this order)

**Header:** title "Why building one feature doesn't move EPM databases to *CLH*" (CLH in Vital Orange italic), subtitle "A database can migrate only when **every** feature it relies on exists in CLH. Most databases need several.", and the Elsevier wordmark on the right. Under the header, a small right-aligned line "Page version: [date time]", then a disclaimer box (sand background, orange left bar):

> **Disclaimer:** The data may change as information is uploaded upon request and with monthly database reconciliation.

**1. Executive summary** (collapsible, open by default; all figures cover the full scope)
- **Why it matters:** contract value and active learners held by databases blocked by at least one feature, the learner share, and the contract-value coverage caveat.
- Three statements:
  - "[X] databases depend on at least one of the six features CLH does not have yet."
  - "[X] databases depend on two or more of these features."
  - "Most common need: **Multi-database (Org Picker)**. Org Picker tool requires SSO. As currently scoped, SSO on CLH requires manual user management of facilities and departments. Manual facility/department assignment is not operationally feasible at scale. Needed by [X] databases."
- **What this means:** no single feature clears the backlog. Say what the first build step releases (count and %), when half and 90% are reached, and that all are released only when all six are built.
- **Decisions needed:**
  1. Confirm the build order (Org Picker depends on SSO with automated user management, so the first step must deliver both; state what it releases).
  2. Define "customer" (until agreed we report databases only; see section 8).
- Closing line: "How to use this document: Section 2 shows an overview; section 3 lets you test what happens as features are built; sections 4 to 6 let you select features and explore the databases."

**2. At a glance** (full scope, no filters)
- Boxes that visibly add up: **[total] still to migrate = [blocked] blocked by at least one missing feature** (with "of which [N] are blocked by two or more" inside that box) **+ [not blocked] not blocked by any of the six features**. The not-blocked box has a dark left bar.
- Directly below, a full-width note with the same dark left bar, explaining **why the not-blocked databases are not blocked by functionality**. Calculate it from Phase: how many are already in Phases 2 to 4 (with per-phase counts), and how the rest split (Phase 5, REVIEW, ANZ TBD, blank). Use this framing: databases outside Phases 2 to 4 were recently decoupled from Connect, were originally treated as multi-database, and need the Customer Evaluation Form completed before they can be migrated. Make it editable text.
- A totals table with columns "Still to migrate", "= Blocked by at least one feature", "+ Not blocked by the six", and "Of the blocked: two or more" (shaded, last). Rows: Databases, Contract value (with recorded coverage), Active learners, Active admins.
- "Databases able to migrate as each feature is built": a real HTML table with fixed column widths and headings aligned to their data. Columns: Build step, Built at this step, Newly able to migrate, Progress (stacked bar: grey for earlier steps, orange for this step, with a key in the heading), Total able to migrate, Share of blocked, Still blocked. Do not include a "Today (nothing built)" row. Add a note on the Org Picker + HR Download row saying they are built together.

**3. Build simulator**
- A card for each feature (click to toggle built or not built; built cards turn orange).
- Four stat cards: Features built (x / 6 and % of the gap closed), Databases able to migrate (of the blocked), Contract value released, Active learners released.
- A stacked bar "Features still missing per database": Ready / 1 / 2 / 3 or more, with legend counts.
- Buttons: Reset, and "Show suggested order, step by step" (applies the dependency-aware order one step per click).
- If Org Picker is built without HR Download, show a note about the dependency.

**4. Features needed**
- One card per feature with the count of databases needing it. Clicking selects it for section 5.
- Selecting features shows only databases that need **every** selected feature (no any/all toggle). A total box under the cards shows the matching database count, recorded contract value, active learners and admins. Control: Clear selection.

**5. Databases by selected feature**
- Hidden until at least one feature is selected.
- Searchable list with columns Account, Database name, CMX ID, ECR ID, Client type, Migration status, Contract value, Learners, Admins, and Selected features (chips for the selected features only; other features stay hidden until selected).
- Export CSV.

**6. Databases sorted by features needed**
- Tabs: **All**, **Migrations in progress (Phases 2 to 4)**, **Not yet in migration**, each showing its count.
- Rolled up by ECR ID: groups are collapsed by default and click to expand. Every account name containing "VISN" forms one group, "VISN (VA networks)", whatever its ECR or CMX ID.
- Columns: Account, Database name, Databases (CMX ID), ECR ID, Client type, Contract value, Learners, Admins, Features needed, plus one column per feature. Group rows show sums and per-feature counts.
- Default sort: Multi-database count, highest first. Every column header re-sorts on click.
- Quick filters by number of features needed, a "Group by ECR ID" toggle, "Hide databases that need none of the six" (on by default), search, and Export CSV (all rows in the current tab; include a "Databases with this ECR ID" column).

**7. How we counted** (editable list)
- Data pull. Built from the Client Directory, which is refreshed monthly from a finance file.
- Counted by database. Each database is counted once.
- No customer count yet. We report databases until "customer" has an agreed definition. See section 8.
- Data excluded: completed migrations, Net New, TTP Cutover, cancellations, Connect, Internal and DirectCourse.
- Six key features to unlock migrations: Multi-database (Org Picker), HR Download, Self-registration / self-enrollment, SCORM Uploader, Linked Tests and CK Nursing.
- A "Terms used" glossary: EPM, CLH, Database (CMX ID), ECR ID, Multi-database (Org Picker), SSO, HR Download, SCORM Uploader (SCO12), and the excluded account types.

**8. Why we count databases, not customers** (editable): the explanation of why "customer" is hard to define for EPM, with these examples: Baptist Health Corbin, Mount Sinai, VA (with its ECR-ID_a/b/c placeholder IDs), Common Spirit, NEAH and Medstar, Kaiser Permanente, plus the note that EPM and CLH customer counts will differ because databases are combined on CLH (Fortis Hospitals).

**9. Databases outside this view: Connect, ANZ and DirectCourse** (reference only, not counted anywhere above): a three-column table (Connect | ANZ | DirectCourse) with Databases, Contract value (with recorded coverage), Active learners, Active admins, Would be blocked by the six features, and Of which by feature. Connect = Connect client types (LMS Connect, Clinical Skills Connect, EPM Connect, Connect Partner) that are still to migrate under the same scope rules. ANZ = databases still to migrate with Phase "ANZ TBD" or country Australia / New Zealand (excluded from all dashboard counts; Connect excludes ANZ so columns do not overlap). DirectCourse = databases with Phase "DirectCourse". One record per CMX ID. Calculate both in `build_data.py` and store them in the data block as `outside`.

## Behaviour and quality requirements

- **Edit mode:** an "Edit text" button (bottom right) makes static headings, hints, notes and lists editable. "Save edited copy" downloads an updated HTML file. Numbers and tables are generated from the data and are not editable.
- All numbers come from the embedded data; nothing is hard-coded except the editable notes.
- Fonts and the favicon are embedded as base64. The data is embedded as a script block starting with `/* Generated by scripts/build_data.py. Do not edit by hand. */` followed by `window.DASHBOARD_DATA = {...};`, so the build script can replace it in place.
- Accessibility: WCAG AA contrast, no orange text under 24px, aria-pressed on toggles, and tabs with role="tab".
- `<meta name="robots" content="noindex, nofollow">`.
- **Before delivering, verify:**
  - Totals add up: total = blocked + not blocked.
  - The "two or more" count is a subset of blocked and is never shown as if it adds to the total.
  - Feature counts match the workbook.
  - The page script runs without errors.
  - The file has no external references.

## GitHub hosting files (one folder, no subfolders)

| File | Content |
|---|---|
| `index.html` | The self-contained dashboard |
| `build_data.py` | Reads the workbook, applies the rules above, and replaces only the data block inside `index.html` (same folder). Usage: `python build_data.py "path/to/Client Directory.xlsx"` |
| `requirements.txt` | `openpyxl>=3.1` |
| `.nojekyll` | Empty file |
| `robots.txt` | `User-agent: *` then `Disallow: /` |
| `.gitignore` | Exclude `*.xlsx`, `*.xlsm`, `*.xls`, `*.csv`, `~$*`, Python caches, OS and editor files |
| `README.md` | What the dashboard shows, new-repository setup, data refresh, editing text, data-sensitivity warning, contact |

**New-repository setup steps for the README:**

1. Create a private repository.
2. Choose **Add file → Upload files** (or **Create new file** and paste), add the files, and commit to `main`.
3. Go to **Settings → Pages → Deploy from a branch**, choose `main` and `/ (root)`, and save.
4. Wait 1 to 2 minutes, then open the site address.
5. Confirm the "Page version" line shows the latest upload.

**Data-sensitivity warning:** the file contains account names, IDs, contract values and learner counts. On most GitHub plans a Pages site is public even if the repository is private. Use private Pages or get data-owner approval before publishing.

## Final response

Keep it short:

- **Headline figures:** total, blocked, two or more, not blocked, most common need, first-step release, contract value and learners at blocked databases.
- **What changed since the last workbook,** if one is provided: databases added or removed and flags changed.
- **Data issues to confirm:** blank fields, inconsistent country values, and notes that no longer match the data.
- **Which files to upload.**

Then present `index.html` and the GitHub files.

**Contact:** Tessa Richards, Customer Readiness Manager (t.richards@elsevier.com)
