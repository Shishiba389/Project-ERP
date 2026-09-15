# Project ERP — Product Packaging Export Workflow

## Purpose

Create brand-specific Kopia packaging workbooks from a Matas assortment file. The application runs in the browser. Source files are read locally and never uploaded.

## Input files

| File | Required data | Role |
|---|---|---|
| Matas assortment | A: EAN, D: Brand, E: Produktnavn | Defines products to process and their brand grouping. |
| Master Data | `BAR CODE`, `STATUS` in `artikkel` sheet | Validates every Matas EAN and identifies discontinued products. |
| Blank Kopia template | Selected output worksheet with prepared rows | Source workbook for each generated file. |

## Core rules

1. Match `Matas assortment.A (EAN)` to `Master Data.artikkel.BAR CODE`.
2. Normalize Status before evaluating it: trim spaces and ignore letter case.
3. Exclude only products where Master Data Status is `Discontinued`.
4. Keep all other statuses, including `ACTIVE`, `Non-Active`, `Limited`, `Upcoming`, `N/A`, and `Others`.
5. Group remaining products by `Matas assortment.D (Brand)`.
6. Use `Matas assortment.E (Produktnavn)` as the product name in output.

## Processing flow

```text
Select source files
      ↓
Read Matas assortment (EAN / Brand / Product name)
      ↓
Read Master Data (BAR CODE / STATUS)
      ↓
Match each EAN
      ↓
 ┌───────────────────────────────────────────┐
 │ EAN missing from Master Data?              │
 │ Yes → stop and show unmatched EAN list     │
 │ No  → continue                             │
 └───────────────────────────────────────────┘
      ↓
Is STATUS = Discontinued?
      ↓
Yes → exclude and report count
No  → keep product
      ↓
Group kept products by Brand
      ↓
Choose generation mode
      ↓
Create one Kopia workbook per selected Brand
      ↓
Download output files
```

## Generation modes

### Automatic mode

Generate one output file for every Brand that has at least one eligible product.

### User-select mode

Show all eligible Brands and their product counts. The user chooses one or more Brands to generate.

## Kopia output contract

Each output file is a direct copy of the selected blank Kopia template.

Filename:

```text
Kopia_Mass_Product_Packaging_EN_<Brand>.xlsx
```

Only these cells in the selected Kopia output worksheet may change:

| Kopia column | Value |
|---|---|
| A — `L.P. PRODUCT` | Sequential product index, starting at 1. |
| B — `EAN SZT` | Matas EAN. |
| D — `PRODUCT'S NAME` | Matas Produktnavn. |

All other workbook content must remain unchanged: sheet names/order, formulas, styles, validation, dropdown lists, column widths, row heights, hidden areas, named ranges, print settings, and workbook structure.

## Required UI states for mockup

1. **Start** — select the three input files.
2. **Reading** — file parsing/progress state.
3. **Validation success** — total received, matched, excluded, eligible, and Brand count.
4. **Validation issue** — unmatched EAN list with clear recovery action.
5. **Mode selection** — Automatic or Select Brands.
6. **Brand selection** — eligible Brand list, counts, selected state.
7. **Template selection** — Kopia worksheet picker.
8. **Generation in progress** — one or more Brand files being prepared.
9. **Success** — downloaded file count and per-Brand result.
10. **Error** — invalid file, missing worksheet, insufficient template rows, or failed export.

## Example result from current sample files

```text
Matas products: 188
Matched to Master Data: 188
Discontinued excluded: 0
Eligible: 188
Output brands: 5

coxir:   28 products
Derma:B:  8 products
Goodal:  10 products
LAKA:    91 products
PERIPERA: 51 products
```
