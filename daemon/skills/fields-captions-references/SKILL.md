---
name: fields-captions-references
description: Table of contents, captions (Tabella/Figura), cross-references and field updates - everything that must stay correct when the document changes.
hosts: [word]
always_on: false
triggers:
  doc_types: [report]
  keywords: [indice, sommario, toc, table of contents, indice delle figure, didascalia, didascalie, caption, captions, riferimento incrociato, riferimenti incrociati, cross-reference, cross reference, campo, campi calcolati, field, fields, aggiorna campi, update fields, numerazione figure, numerazione tabelle, vedi tabella, vedi figura]
---
# Fields, captions and cross-references

## Office mental model
- Fields are live formulas: TOC (from heading outline levels), SEQ (caption numbers:
  "Tabella {SEQ Tabella}"), REF/PAGEREF (cross-references to bookmarks/headings/captions),
  PAGE/NUMPAGES, DATE, STYLEREF. They show cached results until UPDATED.
- TOC lists Heading 1..n (levels 1-3 for business documents). It is wrong after every
  edit until word_update_fields runs. Place it after the title page, before the summary.
- Captions: label + automatic number + text, style "Caption". Italian labels "Tabella",
  "Figura", "Grafico", "Diagramma"; English "Table", "Figure". Convention: table captions
  ABOVE the table, figure/chart captions BELOW the figure. Optionally chapter-numbered
  ("Tabella 2.3") when headings use linked numbering.
- Cross-references point to the caption ("vedi Tabella 3"), heading number/text
  ("cfr. paragrafo 4.2") or page ("a pagina 12"). They renumber automatically.
- An "Indice delle figure/tabelle" is a TOC built from captions: create via toc block only
  if supported; otherwise mention as open point.

## Procedure
1. Map: is there a TOC? stale? Captions as plain text ("Tabella 1 – ...")? Typed refs
   ("vedi tabella sotto")? office_verify reports stale TOC.
2. agent_plan: TOC position, list of objects needing captions, list of text references to
   convert, update + verify.
3. Insert TOC once (word_insert_toc or toc block in a skeleton insert).
4. Caption each table/figure/chart with word_insert_caption on the object anchor
   (T3, chart/shape anchor) — or as caption block right inside the same insert as the table.
5. Replace manual references: find them (word_find "vedi tabella", "Figura 2"), then insert
   word_insert_cross_reference at the location and remove the typed text.
6. word_update_fields at the end (updates TOC, SEQ, REF, PAGE); then office_verify.

## Tool recipes
```json
word_insert_toc({"location": "after:p3", "levels": 3})
word_insert_caption({"anchor": "T2", "label": "Tabella", "text": "Riepilogo dei costi per WP"})
word_insert_caption({"anchor": "T4", "label": "Tabella", "text": "Matrice RACI"})
word_find({"query": "vedi tabella"})
word_insert_cross_reference({"location": "after:p41", "target_anchor": "Tabella 2", "kind": "label_and_number"})
word_insert_cross_reference({"location": "after:p88", "target_anchor": "Analisi dei rischi", "kind": "heading_number"})
word_insert_cross_reference({"location": "after:p88", "target_anchor": "Analisi dei rischi", "kind": "page"})
word_update_fields({})
```
In a new chapter, caption inside the atomic insert:
```json
word_insert_content({"location": "end", "blocks": [
  {"type": "caption", "label": "Tabella", "text": "Milestone principali"},
  {"type": "table", "columns": [{"width_pct": 60}, {"width_pct": 40}], "header_rows": 1, "repeat_header": true,
   "rows": [[{"text": "Milestone", "bold": true}, {"text": "Data", "bold": true}], [{"text": "Go-live"}, {"text": "30/09/2026"}]]}]})
```

## Caption and reference conventions
| Object | Label (IT / EN) | Position | Reference text |
|---|---|---|---|
| Table | Tabella / Table | above | "come mostrato in Tabella 3" |
| Chart | Figura or Grafico / Figure | below | "(cfr. Figura 2)" |
| Diagram, SmartArt | Figura / Figure | below | "la Figura 4 illustra la WBS" |
| Equation | Equazione / Equation | right | "dall'Equazione 1" |
| Chapter | — | — | "nel paragrafo 3.2" (heading_number) |
- Use one label per object type across the whole document; never mix Figura and Grafico.
- Source line for external data goes in the caption text: "... – Fonte: ISTAT, 2025".

## Pitfalls
- Typing "Tabella 1" by hand: breaks after the first insert/delete.
- Captions mixing labels (Tab./Tabella/Table) or positions (some above, some below).
- Inserting the TOC before headings are styled -> empty TOC; never updating it.
- Cross-reference to a heading that has no automatic number (kind heading_number gives "").
- Fields in form templates: don't add TOCs/captions to fill-in forms.
- Leaving "Errore. L'origine riferimento non è stata trovata." after deleting a target —
  office_verify flags it; fix by re-pointing or deleting the reference.

## Verification checklist
- [ ] TOC present (if the document needs one), placed after the title page, updated.
- [ ] Every table/figure/chart captioned with consistent label and position.
- [ ] No typed numbers or "below/above" references; all refs are fields.
- [ ] word_update_fields executed last; no broken-reference errors; office_verify clean.
