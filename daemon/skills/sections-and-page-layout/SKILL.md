---
name: sections-and-page-layout
description: Sections, section breaks, landscape pages, headers/footers, page numbering and different first page using native Word layout.
hosts: [word]
always_on: false
triggers:
  doc_types: [report]
  keywords: [sezione, sezioni, section, sections, interruzione di sezione, section break, orizzontale, landscape, verticale, portrait, intestazione, header, piè di pagina, pie di pagina, footer, numero di pagina, numeri di pagina, page number, page numbers, prima pagina diversa, different first page, margini, margins, layout, impaginazione, pagina x di y]
---
# Sections and page layout

## Office mental model
- A SECTION owns page setup: orientation, size, margins, columns, header/footer set,
  page-number format and restart. Change any of these mid-document = new section.
- Section break types: next_page (new page; orientation changes), continuous (same page;
  column changes). A page_break block is NOT a section break.
- Headers/footers are per section and per kind (first page, primary, even). New sections
  are "linked to previous" by default — so set header/footer on the section where the
  change starts and expect later ones to inherit.
- Page numbers are PAGE / NUMPAGES fields ("Pagina 3 di 12"), never typed digits.
- Typical professional layout:
  | Part | Section | Orientation | Header/Footer |
  |---|---|---|---|
  | Title page + front matter | 1 | portrait | different first page, none on cover |
  | Body | 2 | portrait | title + code / Pagina X di Y |
  | Wide table, Gantt, matrix | 3 | landscape | same as body |
  | Remaining body + appendices | 4 | portrait | same as body |
- Margins standard 2-2.5 cm; landscape sections keep the same header text so the
  document reads as one.

## Procedure
1. Map: `Section n: portrait|landscape` lines tell the current structure. Count sections.
2. Decide where layout changes: around wide tables (> 6-7 columns or grids like
   baseline/Gantt/RACI), diagrams, A3-like content.
3. agent_plan; insert breaks from the END of the document backwards so earlier anchors
   do not shift unexpectedly (or re-read map after each break).
4. Landscape block = next_page break with orientation landscape BEFORE the wide content
   and next_page break with orientation portrait AFTER it.
5. Headers/footers: set on the first body section; different_first_page true on the cover
   section; page_numbers true in the footer.
6. Re-read with office_verify: section count, orientation, header/footer presence.

## Tool recipes
Landscape island around table T5 (p40 = last paragraph before T5; locations accept only after:/replace:/start/end/cursor):
```json
word_insert_section_break({"location": "after:p40", "type": "next_page", "orientation": "landscape"})
word_insert_section_break({"location": "after:T5", "type": "next_page", "orientation": "portrait"})
```
Header and footer:
```json
word_set_header_footer({"section": 1, "kind": "footer", "text": "", "page_numbers": false, "different_first_page": true})
word_set_header_footer({"section": 2, "kind": "header", "text": "Alfa S.p.A. – Studio di fattibilità – PRJ-2026-014", "page_numbers": false, "different_first_page": false})
word_set_header_footer({"section": 2, "kind": "footer", "text": "Riservato – uso interno", "page_numbers": true, "different_first_page": false})
```
Separate cover from body after the TOC:
```json
word_insert_section_break({"location": "after:p6", "type": "next_page", "orientation": "portrait"})
```

## Layout decisions
- Wide-table test: columns × minimum readable width (≈ 1,8 cm) > 16 cm -> landscape.
- Margins: business reports 2,5 cm; forms keep template margins; landscape same margins.
- Page numbering restart: body starts at 1 only when the front matter uses roman numerals
  or no numbers; otherwise continuous numbering is simpler and expected in Italian reports.
- Columns (two-column newsletters, glossaries) need continuous section breaks around them.
- Headers: left = organisation/document title, right = code/version; footer = classification
  + "Pagina X di Y". Keep header text short (one line).

## Pitfalls
- Using page_break when orientation must change (whole document flips).
- Setting the header on section 3 and being surprised it changed sections 4+ (linked).
- Typing page numbers or "Pagina 1 di 10" as text.
- A landscape section containing only part of the wide table (break inside the table).
- Empty pages created by a page_break followed by a next_page break.
- Changing margins in a form template (tables overflow): don't.

## Verification checklist
- [ ] Each orientation change has its own next_page section; no blank pages.
- [ ] Cover without header/footer; body with header text + Pagina X di Y.
- [ ] Wide tables fit the landscape page width; headers consistent across sections.
- [ ] office_verify run; TOC page numbers refreshed with word_update_fields.
