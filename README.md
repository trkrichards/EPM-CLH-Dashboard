# EPM to CLH: features needed by database

`index.html` is the complete dashboard in one file. The data, fonts and icon are built in, so it works on GitHub Pages and also opens by double-clicking it on your computer. Section 9 gives reference totals for Connect, ANZ and DirectCourse databases, which are not included in the main counts.

## Files

**Required**
- `index.html`: the dashboard
- `.nojekyll`: tells GitHub Pages to serve the file as-is (empty file)

**Recommended**
- `robots.txt`: asks search engines not to index the site
- `.gitignore`: keeps Excel workbooks out of the repository

**Optional, for refreshing the data**
- `build_data.py`: updates the data inside `index.html` from a new Client Directory workbook
- `requirements.txt`: the Python package it needs
- `README.md`: this file

## Set up a new repository

1. In GitHub, choose **New repository**, set it to **Private**, and create it.
2. Choose **Add file → Upload files**, drag in the files (not the folder), and click **Commit changes**.
3. Go to **Settings → Pages**, set **Source** to **Deploy from a branch**, choose **main** and **/ (root)**, and click **Save**.
4. After 1 to 2 minutes, refresh **Settings → Pages**. The site address appears at the top. The "Page version" line above the disclaimer confirms you're seeing the latest file.

**Before you publish:** the dashboard contains account names, IDs, contract values and learner counts. On most GitHub plans a Pages site is public even if the repository is private. Confirm with your data owner, or use private Pages.

## Refreshing the data

```
pip install -r requirements.txt
python build_data.py "path/to/Client Directory.xlsx"
```

This updates `index.html` in place. Upload the updated `index.html` to GitHub.

## Features counted

Seven features: Multi-database (Org Picker), HR Download, Self-registration / self-enrollment (including Skills and eLearning self-assign counts of 1 or more), SCORM Uploader, LinkedContent / Linked Tests, CK Nursing and Skills Editor (Skills Editor Count of 1 or more).

## Scope

Databases still to migrate, one per CMX ID. Excluded: completed migrations, Net New, TTP Cutover, cancellations, Connect, Workforce Development, Internal, DirectCourse and ANZ (Phase "ANZ TBD" or Australia / New Zealand). If the workbook's main tab has Excel filters set, those filters are applied first.

## Editing text

Open the dashboard, click **Edit text** (bottom right), make changes, click **Save edited copy**, and use the downloaded file as the new `index.html`.

## Contact

Tessa Richards, Customer Readiness Manager (t.richards@elsevier.com)
