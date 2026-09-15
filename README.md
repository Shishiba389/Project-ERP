# Project ERP

A static browser workspace for creating Matas EPR packaging-card files.

## Workflow

1. Select the Matas assortment workbook. The app uses columns A (EAN), D (Brand), and E (Produktnavn).
2. Select Master Data. Products are matched using `BAR CODE`; only `Discontinued` products are excluded.
3. Select a blank Kopia template and its output worksheet.
4. Generate all brands automatically or choose the brands to export.

Each downloaded file is a copy of the selected template. The app only writes columns A, B, and D in the selected output sheet. Processing happens in the browser; source workbooks are not uploaded.

## Local preview

Open `index.html` with a local web server, or publish this repository with GitHub Pages once it is ready to be public.
