---
name: long-report-architecture
description: Build long structured documents (relazioni, report, studi di fattibilità, tesi, manuali) with title page, TOC, numbered headings, captions, cross-references, landscape sections and appendices.
hosts: [word]
always_on: false
triggers:
  doc_types: [report, blank]
  keywords: [relazione, report, rapporto, studio di fattibilità, feasibility study, tesi, thesis, manuale, manual, whitepaper, documento tecnico, technical document, capitoli, chapters, indice, sommario, table of contents, toc, executive summary, sintesi, appendice, appendix, allegati, scrivi, write a, redigi, draft]
---
# Long report architecture

## Office mental model
A professional long document is a STRUCTURE first, text second:
1. Title page (section 1, different first page, no page number): Title style, Subtitle
   style, organisation, author, version, date, document code, classification.
2. Front matter: revision table (Versione | Data | Autore | Modifiche), optional
   distribution list, TOC (field) — then a next-page section break.
3. Executive summary / Sintesi (½-1 page: context, key findings, recommendations, numbers).
4. Body chapters with built-in Heading 1-3 ONLY, numbered 1 / 1.1 / 1.1.1 via a multilevel
   list LINKED to the heading styles (never typed numbers).
5. Figures/tables with captions (SEQ fields "Figura n", "Tabella n") and cross-references
   ("vedi Tabella 3") as REF fields; wide tables in a landscape section.
6. Conclusions / recommendations, glossary, bibliography, appendices (Appendice A, B... as
   separate heading level or own numbering).
7. Headers/footers: document title + code in header, "Pagina X di Y" in footer, starting
   after the title page.
Word builds the TOC, numbering and captions from styles and fields: that is what makes a
200-page document maintainable. Body text uses Normal (or a "Body Text" style) with space
after, not empty paragraphs; page breaks are explicit page_break or section breaks.

## Procedure
1. Analyze map: blank vs existing report. Existing: keep its styles and numbering scheme;
   append/restructure only what is asked. Check language.
2. Content: gather facts (user input, memory_search, research-for-content for external
   facts with sources). Draft an outline: 4-8 chapters, 2-5 sections each, planned figures
   and tables. agent_plan with one step per chapter batch + layout steps.
3. Skeleton in ONE word_insert_content call: title, subtitle, metadata paragraphs,
   page_break, toc block, page_break, Executive summary heading + text.
4. Apply heading numbering once: word_link_numbering_to_headings scheme "1.1.1"
   (unnumbered headings like "Sintesi"/"Appendici" can use a non-numbered style or
   be outside the numbered range; keep it consistent).
5. Write chapters in batches (1-2 chapters per word_insert_content, location "end"):
   headings + paragraphs + bullets + tables + caption blocks. Real content, specific numbers,
   no filler. Tables: header row, repeat_header, widths in %.
6. Visuals: native charts (native-charts), SmartArt/diagrams (smartart-and-diagrams),
   each with a caption below (figures) or above (tables, Italian/EU convention).
7. Layout: section break after front matter; landscape section around wide tables
   (sections-and-page-layout); header/footer with page numbers, different first page.
8. Cross-references in text to tables/figures/chapters; word_update_fields; refresh TOC.
9. document-quality-check pass + office_verify; report outline and open points.

## Tool recipes
Skeleton:
```json
word_insert_content({"location": "start", "blocks": [
  {"type": "title", "text": "Studio di fattibilità – Migrazione CRM in cloud"},
  {"type": "subtitle", "text": "Alfa S.p.A. – Direzione Sistemi Informativi"},
  {"type": "paragraph", "runs": [{"text": "Versione 1.0 – 15/03/2026 – Codice PRJ-2026-014-SF", "size": 10}], "align": "center"},
  {"type": "page_break"},
  {"type": "toc", "levels": 3},
  {"type": "page_break"},
  {"type": "heading", "level": 1, "text": "Sintesi"},
  {"type": "paragraph", "runs": [{"text": "Il progetto riduce i costi operativi del 18% (€ 42.000,00/anno) ..."}]}]})
word_link_numbering_to_headings({"scheme": "1.1.1"})
```
Chapter batch with table + caption:
```json
word_insert_content({"location": "end", "blocks": [
  {"type": "heading", "level": 1, "text": "Analisi dei costi"},
  {"type": "paragraph", "runs": [{"text": "La Tabella seguente riassume i costi per fase."}]},
  {"type": "caption", "label": "Tabella", "text": "Costi per fase di progetto"},
  {"type": "table", "columns": [{"width_pct": 40}, {"width_pct": 30}, {"width_pct": 30}],
   "header_rows": 1, "repeat_header": true, "style": "Grid Table 4 - Accent 1",
   "rows": [[{"text": "Fase", "bold": true}, {"text": "Costo (€)", "bold": true, "align": "right"}, {"text": "Durata", "bold": true}],
            [{"text": "Analisi"}, {"text": "12.000,00", "align": "right"}, {"text": "4 settimane"}]]},
  {"type": "heading", "level": 2, "text": "Ipotesi di calcolo"},
  {"type": "bullets", "items": ["Tariffa media € 450,00/gg", {"text": "Licenze", "children": ["SaaS 50 utenti", "Ambiente di test"]}]}]})
```
Cross-reference, landscape annex, header/footer, refresh:
```json
word_insert_cross_reference({"location": "after:p57", "target_anchor": "Tabella 2", "kind": "label_and_number"})
word_insert_section_break({"location": "after:p212", "type": "next_page", "orientation": "landscape"})
word_set_header_footer({"section": 2, "kind": "footer", "text": "PRJ-2026-014 – Studio di fattibilità", "page_numbers": true, "different_first_page": false})
word_update_fields({})
```

## Pitfalls
- Bold Normal paragraphs as "headings" (invisible to TOC/navigation pane).
- Typed numbering "1.2 Obiettivi" plus automatic numbering -> "1.2 1.2 Obiettivi".
- TOC inserted but never updated (shows "No table of contents entries found" or old pages).
- Heading level skips (H1 -> H3); more than 3 levels in a business report.
- One giant insert of 30 pages: split into chapter batches so errors are local.
- Empty paragraphs for spacing; manual page breaks before every heading (use
  "page break before" in Heading 1 style via word_modify_style if needed).
- Captions as plain bold text; "see table below" instead of a cross-reference.
- Executive summary without numbers or decisions; generic AI filler text.

## Verification checklist
- [ ] Title page, TOC, executive summary, chapters, conclusions, appendices present.
- [ ] Only Heading 1-3 for structure; numbering linked to headings; no heading skips.
- [ ] Every table/figure has a caption and is referenced in the text.
- [ ] Wide tables in landscape section; headers/footers + page numbers correct per section.
- [ ] word_update_fields run; TOC current; office_verify clean.
