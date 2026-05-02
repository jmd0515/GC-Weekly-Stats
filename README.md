# GC-Weekly-Stats
Weekly Performance Reports to Share with Managers

Live: https://jmd0515.github.io/GC-Weekly-Stats/

## Automation

Every Saturday at 9:00 AM ET, GitHub Actions ([.github/workflows/weekly-report.yml](.github/workflows/weekly-report.yml)) rebuilds the dashboards:

1. Logs into reports.salondata.com with Playwright and exports `Employee_Stats.csv` and `Employee_Return_Stats.csv` (week ending most recent Friday).
2. Converts those CSVs to xlsx ([convert_downloads.py](convert_downloads.py)).
3. Runs [generate_data.py](generate_data.py) using those two files plus the committed `All_Salons.xlsx`.
4. Commits the regenerated HTML files. GitHub Pages republishes within a minute.

### Required repo secrets
Settings → Secrets and variables → Actions:

| Secret | Purpose |
|---|---|
| `SALONDATA_USERNAME` | reports.salondata.com login |
| `SALONDATA_PASSWORD` | reports.salondata.com login |

### Manual run
Actions tab → **Weekly Report** → **Run workflow**.

## Updating All_Salons.xlsx

`All_Salons.xlsx` is the system-wide weekly dataset (93 salons × 19 metrics, accumulating week over week). It comes from a Power BI Customer Experience report and is currently **updated manually**:

1. In Power BI, set start/end date to the most recent Friday.
2. Hover the table visual → `⋯` → **Export data** → save the xlsx.
3. Replace `All_Salons.xlsx` in this repo and commit.

The next scheduled (or manual) workflow run will pick it up.

> Automation of this step is parked in [pbi_export.js](pbi_export.js) — Microsoft Entra auth makes headless Power BI access tricky. To revisit later.

## Local manual run (fallback)
The original double-click flow still works: drop the three `.xlsx` files in this folder and run `run.bat`.
