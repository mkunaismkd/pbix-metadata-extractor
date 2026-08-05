# PBIX Metadata Extractor

Extracts full metadata from a Power BI Desktop file (`.pbix`) and exports it
to a single, formatted Excel workbook — no Power BI Desktop installation
required.

Built with [PBIXRay](https://github.com/Hugoberry/pbixray), which parses the
`.pbix` file's internal VertiPaq data model directly.

## What gets extracted

| Sheet | Contents |
|---|---|
| Overview | File size, counts, and Power BI Desktop build metadata |
| Tables | All table names in the model |
| Columns (Schema) | Every column, its table, and data type |
| Measures (DAX) | All DAX measures with full expressions, display folder, description |
| Calculated Columns | DAX-calculated columns and their expressions |
| Calculated Tables | DAX-calculated tables and their expressions |
| Relationships | Table relationships, cardinality, active/inactive, cross-filter direction |
| Power Query (M) | M code for every query/table's data source transformation |
| Parameters | M parameters (if any) |
| Perspectives | Perspective definitions (if any) |
| Row-Level Security | RLS roles and filter expressions (if any) |

Sheets that are entirely empty for a given file (e.g. no parameters defined)
are automatically skipped.

## Install

```bash
pip install -r requirements.txt
```

## Usage

Single file:
```bash
python extract_pbix_metadata.py "MyReport.pbix"
# -> MyReport_metadata.xlsx written alongside it
```

Custom output path:
```bash
python extract_pbix_metadata.py "MyReport.pbix" -o "output/MyReport_docs.xlsx"
```

Batch mode — process every `.pbix` in a folder:
```bash
python extract_pbix_metadata.py ./reports/
```

## Notes

- Works entirely offline on the file's binary contents — the report never
  needs to be opened in Power BI Desktop, and no data connections are made.
- DAX/M expression columns are wrapped and column-widened for readability
  in Excel.
- If a `.pbix` uses a newer model feature PBIXRay doesn't yet support, that
  one sheet is skipped with a warning printed to the console — the rest of
  the export still completes.

## Web app (Streamlit)

`streamlit_app.py` wraps the extractor in a drag-and-drop upload UI with a
sheet-by-sheet preview and a download button — no CLI needed.

Run locally:
```bash
streamlit run streamlit_app.py
```

### Deploy for free — Streamlit Community Cloud

GitHub Pages can't run Python, so this repo is deployed instead through
[Streamlit Community Cloud](https://share.streamlit.io), which redeploys
automatically on every push to `main`:

1. Push this repo to GitHub (include `streamlit_app.py` and `requirements.txt`).
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Pick this repo/branch, set **Main file path** to `streamlit_app.py`.
4. Deploy. You'll get a public URL like `https://<your-app>.streamlit.app`.

Optionally link that URL from your GitHub repo's description or a GitHub
Pages landing page, so the repo still "looks like" it's hosted on GitHub
Pages even though Streamlit Cloud is doing the actual serving.

## Possible extensions

- Add a GitHub Action that runs this on every `.pbix` pushed to a repo and
  commits the metadata Excel as build documentation.
- Add a simple auth gate or rate limit if you expect public traffic and want
  to control PBIXRay compute usage on the free tier.
