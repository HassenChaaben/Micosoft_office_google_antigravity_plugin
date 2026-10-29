---
name: template-form-filling
description: Fill an existing Word template/form (label-value tables, record grids, checkboxes, content controls) without breaking its formatting or structure.
hosts: [word]
always_on: false
triggers:
  doc_types: [form_template]
  keywords: [compila, compilare, riempi, riempire, completa, inserisci i dati, modulo, scheda, template, form, fill, fill in, complete the form, campi, fields, registro, register]
---
# Template form filling

## Office mental model
- Professional Italian templates are mostly TABLES used as forms, styled with DIRECT
  formatting (font, size, bold, shading per cell). The map warns "DIRECT FORMATTING is used".
  The value cell already has the right run properties (often an empty run or a
  placeholder run): word_fill_fields writes INTO that run so formatting survives.
- Three table shapes in the map:
  - `form`: label cell -> value cell (`Cliente › Azienda [T0 r6 c1]`). Band rows
    ("Cliente", "Sponsor") are section titles, not fields.
  - `record_grid`: header row(s) + N empty record rows
    (`Minacce › records #1-#6 × [Cod., Categoria, WBS, ...] [T1 rows r2-r7]`). Each record is
    one row; each column one field. Two-level headers become paths like
    `Probabilità accadimento › Val. num.`.
  - `mixed`: a form part followed by grid parts in the same table (Project charter T1:
    Deliverable grid, Milestone grid, Rischi, Project manager form).
- Merged cells: `merges=13` means vertical/horizontal merges; the field path already
  points at the real (top-left) cell. Never address cells inside a merge.
- Content controls (tagged) and legacy checkboxes ("SI/NO", "☐") exist in some templates:
  fill controls by tag, toggle checkboxes with word_set_checkbox, never type "X".
- Placeholders that must be REPLACED, not appended to: "Nome Cognome", "XX", "(X)",
  "Minore di X", "gg/mm/aaaa", "Titolo del progetto", "aggiornato il 31/03/21".

## Procedure
1. From the map list every field path and grid; note `empty_records` per grid and
   `... and N more` truncation (then word_read_table each truncated table).
2. Gather data: user message, attached text, memory_search (project code, client, sponsor,
   PM, budget, dates, WBS codes), and other documents of the same project. For PM templates
   also load the rules of pm-documentation-suite (consistency across 36 documents).
3. agent_plan: one step per table (fill form part, fill grid, clone rows, checkboxes),
   then verify.
4. Count records vs empty rows. If records > empty_records, call word_table_clone_row on
   the LAST empty record row of that grid with count = missing rows BEFORE filling (clone
   copies borders, shading and run formatting). Re-read the table: row indexes after the
   grid shift by `count`.
5. Fill each table with ONE word_fill_fields call. Prefer field paths from the map; use
   `T{t}r{r}c{c}` only for cells the map does not name. Fill grids row by row with
   `<grid path> #<n> › <column>` keys, or with cell coordinates read from the map ranges.
6. Values: short, factual, consistent. Italian formats:
   - money `€ 75.000,00` (thousand dot, decimal comma, space after €); percentages `12,5%`;
   - dates `15/03/2026` (gg/mm/aaaa); durations `20 gg` / `4 settimane`;
   - names "Cognome Nome" when the label says so (`Cognome Nome` -> "Rossi Mario");
   - codes keep the project pattern (`PRJ-2026-014`, `R01`, `WP 1.2.3`, `RQ-005`).
   Unknown or not provided -> "[da definire]" — never invent people, amounts or dates.
   Derived values (EMV, risk score, totals) must be CALCULATED from filled inputs.
7. Leave structure alone: no new tables, no deleted rows (unused record rows stay empty
   or get "—" only if the user asks), no re-styling, no moving labels.
8. office_verify; then report: fields filled, rows cloned, items "[da definire]".

## Tool recipes
Form part (paths exactly as in the map):
```json
word_fill_fields({"fields": {
  "Codice del documento": "PRJ-2026-014-PC-01",
  "Anagrafica progetto › Codice": "PRJ-2026-014",
  "Anagrafica progetto › Titolo": "Migrazione CRM in cloud",
  "Cliente › Azienda": "Alfa S.p.A.",
  "Cliente › Cognome Nome": "Bianchi Laura",
  "Budget › Budget complessivo": "€ 75.000,00",
  "Budget › Note esplicative sul budget": "[da definire]"}})
```
Record grid, 8 threats but only 6 empty rows (T1 rows r2-r7): clone first, then fill:
```json
word_table_clone_row({"table": "T1", "row": 7, "count": 2})
word_read_table({"table": "T1"})
word_fill_fields({"fields": {
  "Minacce › #1 › Cod.": "M01", "Minacce › #1 › Categoria": "Tecnico",
  "Minacce › #1 › WBS": "1.2.1", "Minacce › #1 › Causa": "Dati legacy non normalizzati",
  "Minacce › #1 › Evento": "Errori nella migrazione dati",
  "Minacce › #1 › Effetto": "Ritardo go-live di 2 settimane",
  "Minacce › #1 › Periodo accadimento": "03/2026", "Minacce › #1 › Nickname": "Dati sporchi",
  "T1r9c0": "M08", "T1r9c1": "Organizzativo"}})
```
Content control and checkbox:
```json
word_fill_fields({"fields": {"ga:data_approvazione": "30/04/2026"}})
word_set_checkbox({"anchor": "Progetto completato (SI/NO)", "checked": true})
```
Edit a value that is already filled (keeps run formatting): same word_fill_fields call with
the new value. Fix a typo inside a long paragraph: word_replace_text.

## Pitfalls
- word_insert_content / word_update_block inside a form table: it rebuilds runs and loses
  the template look. Use word_fill_fields.
- Writing into the LABEL cell (c0) instead of the value cell (c1/c2); writing into band rows.
- Cloning the header row or a totals row instead of the last empty record row.
- Forgetting that clone shifts all following row indexes (and following grids in mixed tables).
- Appending to a placeholder ("Nome Cognome Rossi Mario") instead of replacing it.
- Inconsistent codes between grids of the same document (R01 in registro, R1 in analisi).
- Inventing plausible data to "look complete". Mark "[da definire]" and say so.
- Using English formats in Italian forms (75,000.00 / 03/15/2026).
- Filling "Val. num." with words or "Val. qual.vo" with numbers; check the scale legend.

## Verification checklist
- [ ] Every field path in the map is filled or deliberately "[da definire]".
- [ ] Records > empty rows handled with word_table_clone_row; no record lost.
- [ ] No placeholder text left ("Nome Cognome", "XX", "(X)", "gg/mm/aaaa").
- [ ] Fonts, shading, borders, merges unchanged (word_get_formatting on 1-2 filled cells).
- [ ] Calculated fields recomputed (totals, EMV = P × impatto, risk score = P × I).
- [ ] Italian number/date formats everywhere; codes consistent with other project documents.
- [ ] office_verify clean; report lists open "[da definire]" items.
