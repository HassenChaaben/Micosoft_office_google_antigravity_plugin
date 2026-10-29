---
name: styles-and-numbering
description: Use built-in styles, modify style definitions, and link multilevel numbering to headings instead of direct formatting or typed numbers.
hosts: [word]
always_on: false
triggers:
  doc_types: [report, contract]
  keywords: [stile, stili, style, styles, titolo, titoli, heading, headings, numerazione, numbering, multilevel, elenco numerato, capitoli numerati, 1.1.1, formattazione, formatting, font, carattere, interlinea, spaziatura, spacing, uniforma, consistent, coerente]
---
# Styles and numbering

## Office mental model
- Paragraph styles carry the whole look (font, size, color, spacing, keep-with-next,
  outline level). Character styles for inline emphasis. Table styles for tables.
- Built-in styles have meaning: Title, Subtitle, Heading 1-9 (outline levels -> TOC,
  navigation pane, cross-references), Normal, List Paragraph, Caption, Quote, TOC 1-3.
  Italian Word shows "Titolo 1"; the anchor/style name "Heading 1" maps to the same
  built-in style. Prefer built-ins over custom "MyHeading" styles.
- Direct formatting (bold/size on runs) overrides the style. In REPORTS, convert direct
  formatting to styles. In FORM TEMPLATES (map says DIRECT FORMATTING is used), leave it.
- Multilevel numbering ("1 / 1.1 / 1.1.1" or "1 / a / i") must be ONE list definition
  LINKED to Heading 1..3 — then every heading numbers itself, cross-references show
  "paragrafo 3.2", and inserting a chapter renumbers everything.
- Bullets/numbered lists in body text: word_insert_content bullets/numbered blocks use
  List Bullet/List Number; nesting via children = list levels, not tabs.
- Changing the style definition once (word_modify_style) restyles the whole document:
  the expert way to "make all titles blue 14pt".

## Procedure
1. Analyze map: styles in use, direct-formatting warning, headings detected (or fake
   headings: short bold Normal paragraphs).
2. word_get_formatting on 2-3 representative paragraphs to learn the current look.
3. agent_plan: (a) map fake headings to levels, (b) apply styles, (c) modify style
   definitions to the house look, (d) link numbering, (e) remove typed numbers, (f) verify.
4. Apply Heading styles with word_apply_style per anchor (batch by walking the outline).
5. Typed numbers ("2.1 Obiettivi") -> word_set_paragraph_text with the bare title, BEFORE
   linking numbering, to avoid doubles.
6. word_link_numbering_to_headings with the requested scheme.
7. word_modify_style for Normal, Heading 1-3, Caption, Title to match the house style
   (font family, size, color, space before/after). Keep body 10-12 pt, H1 14-18 pt,
   space after Normal 6 pt, line spacing via style not empty lines.
8. office_verify (heading skips, mixed formatting); report the style changes.

## Tool recipes
```json
word_apply_style({"anchor": "p14", "style": "Heading 1"})
word_apply_style({"anchor": "p15", "style": "Heading 2"})
word_set_paragraph_text({"anchor": "p15", "text": "Obiettivi del progetto"})
word_link_numbering_to_headings({"scheme": "1.1.1"})
word_modify_style({"style": "Heading 1", "font": "Calibri Light", "size_pt": 16, "color": "#1F3864", "bold": true, "space_before": 18, "space_after": 6})
word_modify_style({"style": "Normal", "font": "Calibri", "size_pt": 11, "space_before": 0, "space_after": 6})
word_modify_style({"style": "Caption", "size_pt": 9, "color": "#404040", "bold": false})
```
Nested numbered list in body (legal-style clauses use headings instead):
```json
word_insert_content({"location": "after:Requisiti", "blocks": [
  {"type": "numbered", "items": ["Autenticazione SSO", {"text": "Integrazioni", "children": ["ERP", "Posta elettronica"]}, "Reportistica"]}]})
```
Contracts: articles as Heading 1 ("Articolo 1 – Oggetto"), clauses as Heading 2 numbered
1.1, 1.2 with scheme "1.1.1"; cross-reference clauses by heading, never by typed number.

## House style defaults (when the user has none)
| Style | Font / size | Spacing | Other |
|---|---|---|---|
| Title | Calibri Light 26 | after 12 | cover page only |
| Heading 1 | Calibri Light 16, accent color | before 18 / after 6 | numbered 1; keep with next |
| Heading 2 | Calibri Light 13 | before 12 / after 4 | numbered 1.1 |
| Heading 3 | Calibri 11 bold | before 8 / after 2 | numbered 1.1.1 |
| Normal | Calibri 11 | after 6, line 1,15 | justified only if the user asks |
| Caption | Calibri 9 italic or grey | after 6 | label + number automatic |
Apply this with word_modify_style once at the start of a new report, then only styles.

## Pitfalls
- Formatting each heading by hand (bold + size) instead of the style.
- Creating a new list per chapter -> numbering restarts at 1 in each chapter.
- Mixing typed and automatic numbers; "1.1.1" scheme on documents using "Articolo n".
- Applying Heading styles inside form tables or to table cells.
- Changing Normal's font in a form template (reflows the entire form).
- Using Heading 4+ for emphasis; skipping levels.
- Color/size changes on runs instead of word_modify_style (inconsistent, not maintainable).

## Verification checklist
- [ ] All structural headings use Heading 1-3; no fake bold headings left.
- [ ] Numbering is automatic, continuous and linked; no duplicated typed numbers.
- [ ] Style definitions match the requested look; no stray direct formatting in reports.
- [ ] TOC updated after style changes (word_update_fields); office_verify clean.
