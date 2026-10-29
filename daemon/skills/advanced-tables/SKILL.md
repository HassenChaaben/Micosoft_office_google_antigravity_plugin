---
name: advanced-tables
description: Complex Word tables - merged cells, multi-row headers, repeat header rows, shading, widths, totals, matrices (RACI, P×I) - built atomically and edited without breaking structure.
hosts: [word]
always_on: false
triggers:
  doc_types: [minutes]
  keywords: [tabella complessa, complex table, celle unite, merge, merged, unisci celle, intestazione ripetuta, repeat header, matrice, matrix, raci, ram, griglia, grid, ombreggiatura, shading, totali, totals, colonne, columns, righe, rows, aggiungi righe, add rows]
---
# Advanced tables

## Office mental model
- A Word table is a grid with merged cells: horizontal merge = colspan (gridSpan), vertical
  merge = rowspan (vMerge). Addresses `T3r2c1` refer to the top-left real cell of a merge.
- Professional tables: header row(s) bold with shading, repeat_header on for tables that
  can cross pages, numbers right-aligned with Italian format, text left, widths in % summing
  to 100, a table style (e.g. "Grid Table 4 - Accent 1", "Table Grid") instead of
  hand-drawn borders, totals row bold with top border.
- Tables are for DATA. Do not use them to lay out paragraphs; do not use tabs/spaces to fake
  columns.
- Multi-row header: first header row has group cells with colspan (e.g. "Probabilità"
  over "Val. qual.vo | Val. num."), second has sub-columns; header_rows = 2.
- Matrices: RACI (WP rows × roles columns, values R/A/C/I, exactly one A per row),
  P×I heat map (5×5, shading green/yellow/red by P×I score), traceability matrices.
- Wide tables (> 7 columns) go into a landscape section (sections-and-page-layout).
- Editing existing tables: never rebuild a filled table to change one cell.
  Fill with word_fill_fields, add rows with word_table_clone_row, read with word_read_table.

## Procedure
1. Decide columns, header structure, merges, alignment and widths BEFORE inserting.
2. Build the whole table in ONE word_insert_content block (with its caption).
3. For existing tables: word_read_table, then word_fill_fields / word_table_clone_row;
   replace a whole generated table with word_update_block only if the agent created it.
4. Long tables: repeat_header true; avoid rowspan across page boundaries.
5. office_verify; check widths and that totals equal the sum of the rows.

## Tool recipes
Two-row header with merges and totals:
```json
word_insert_content({"location": "after:Analisi qualitativa", "blocks": [
  {"type": "caption", "label": "Tabella", "text": "Analisi qualitativa delle minacce"},
  {"type": "table", "header_rows": 2, "repeat_header": true, "style": "Table Grid",
   "columns": [{"width_pct": 10}, {"width_pct": 30}, {"width_pct": 15}, {"width_pct": 15}, {"width_pct": 15}, {"width_pct": 15}],
   "rows": [
    [{"text": "Cod.", "rowspan": 2, "bold": true, "shading": "#D9E2F3", "valign": "center"},
     {"text": "Rischio", "rowspan": 2, "bold": true, "shading": "#D9E2F3", "valign": "center"},
     {"text": "Probabilità", "colspan": 2, "bold": true, "shading": "#D9E2F3", "align": "center"},
     {"text": "Impatto", "colspan": 2, "bold": true, "shading": "#D9E2F3", "align": "center"}],
    [{"text": "Qual.", "bold": true, "shading": "#D9E2F3"}, {"text": "Num.", "bold": true, "shading": "#D9E2F3"},
     {"text": "Qual.", "bold": true, "shading": "#D9E2F3"}, {"text": "Num.", "bold": true, "shading": "#D9E2F3"}],
    [{"text": "M01"}, {"text": "Errori migrazione dati"}, {"text": "Alto"}, {"text": "4", "align": "center"}, {"text": "Moderato"}, {"text": "3", "align": "center"}]]}]})
```
RACI matrix (one A per row):
```json
{"type": "table", "header_rows": 1, "repeat_header": true,
 "columns": [{"width_pct": 34}, {"width_pct": 22}, {"width_pct": 22}, {"width_pct": 22}],
 "rows": [[{"text": "WP", "bold": true}, {"text": "PM", "bold": true, "align": "center"}, {"text": "Team IT", "bold": true, "align": "center"}, {"text": "Sponsor", "bold": true, "align": "center"}],
          [{"text": "1.1 Analisi requisiti"}, {"text": "A", "align": "center"}, {"text": "R", "align": "center"}, {"text": "C", "align": "center"}]]}
```
Add 3 rows to an existing grid and fill one cell:
```json
word_table_clone_row({"table": "T4", "row": 6, "count": 3})
word_fill_fields({"fields": {"T4r7c0": "WP 1.4", "T4r7c1": "R"}})
```

## Table design rules
- Header text short (≤ 3 words), units in header ("Costo (€)") not in every cell.
- Alignment: text left, numbers right, codes/flags/R-A-C-I centered, dates centered.
- Zebra shading or light borders only; header shading one accent color (e.g. #D9E2F3).
- P×I heat map shading: score 1-4 green (#C6EFCE), 5-9 yellow (#FFEB9C), 10-25 red (#FFC7CE).
- Totals: computed by you, bold, last row; put "Totale" in the first cell.

## Pitfalls
- Widths that do not sum to 100% or tiny columns causing one-letter-per-line wrapping.
- Rowspan/colspan arithmetic errors: every row must cover the same number of grid columns.
- Repeat header missing on multi-page tables; header shading not applied to all header rows.
- Inserting a new table to "update" an existing one (duplicates); rebuilding user tables.
- Numbers as "75000" or left-aligned; totals that don't add up.
- RACI with zero or two "A" in a row.

## Verification checklist
- [ ] Caption + header rows + repeat header; widths sensible; style applied.
- [ ] Merge geometry correct (word_read_table shows expected grid).
- [ ] Numbers right-aligned, Italian format, totals verified.
- [ ] Wide tables in landscape; office_verify clean.
