---
name: document-quality-check
description: Final quality gate before delivering a Word document - structure, fields, placeholders, consistency, formatting, accessibility - using office_verify plus expert checks.
hosts: [word]
always_on: false
triggers:
  doc_types: []
  keywords: [controlla, controllo, verifica, verify, check, qualità, quality, rivedi tutto, finale, final, consegna, deliver, pronto, ready, errori, errors, coerenza, consistency, refuso, refusi, typo, qa]
---
# Document quality check

## Office mental model
Senior consultants never hand over a document without a QA pass. office_verify gives the
machine view (re-snapshot diff + issues: empty fields, stale TOC, heading skips, leftover
placeholders). You add the expert view:
| Area | Check |
|---|---|
| Completeness | every template field filled or "[da definire]"; all requested sections exist |
| Placeholders | "Nome Cognome", "XX", "(X)", "gg/mm/aaaa", "Lorem", "[inserire]", "TBD" |
| Structure | Heading 1-3 only, no skips, numbering linked, TOC current |
| Fields | captions numbered, cross-references resolve, page numbers fields |
| Numbers | totals = sum of rows; P×I, EMV, EVM formulas correct; % sum to 100 |
| Consistency | project code, names, dates, amounts identical across sections/documents |
| Formats | Italian: € 1.234,56, gg/mm/aaaa, 12,5%; one date style; one currency style |
| Formatting | no stray direct formatting in reports; template formatting intact in forms |
| Layout | no blank pages, wide tables in landscape, headers/footers per section |
| Language | one language, correct terminology, no AI filler ("In conclusione, è importante...") |
| Accessibility | header rows marked, captions/alt text on charts/diagrams, no color-only meaning |

## Procedure
1. office_verify first; list issues.
2. Targeted reads: word_find for placeholder patterns and suspicious formats; word_read_table
   for tables with numbers; word_read_section for summary/conclusions vs body.
3. memory_search project master data and compare.
4. agent_plan the fixes ordered: content errors > numbers > consistency > structure > cosmetics.
5. Fix with the narrowest tool (word_fill_fields, word_replace_text, word_apply_style,
   word_update_fields). If the document is not yours to rewrite (review context), use
   tracked changes/comments instead (review-and-track-changes).
6. office_verify again until clean (or remaining issues explained).
7. Report: PASS/FAIL per area (short table), what you fixed, what needs the user.

## Tool recipes
```json
office_verify({})
word_find({"query": "Nome Cognome"})
word_find({"query": "XX"})
word_find({"query": "gg/mm/aaaa"})
word_find({"query": "[da definire]"})
word_read_table({"table": "T5"})
word_fill_fields({"fields": {"T5r10c6": "€ 12.300,00"}})
word_apply_style({"anchor": "p31", "style": "Heading 2"})
word_update_fields({})
office_verify({})
```
Report format:
```
QA: 7/9 PASS
- Completezza: FAIL -> 3 campi [da definire] (Sponsor › Ruolo, Budget › Note, Milestone #4 › Data)
- Numeri: PASS (EMV ricalcolati, totale € 18.400,00)
- Coerenza: FIXED codice progetto "PRJ-14" -> "PRJ-2026-014" (2 occorrenze)
```

## Consistency probes (quick wins)
- Search each master value from memory (codice progetto, budget, sponsor surname) with
  word_find; every occurrence must be identical.
- Search typical wrong formats: "€75", ".00", "/2026 " vs "-2026", "EUR", double spaces.
- Compare executive summary numbers with body tables; compare captions with references.

## Pitfalls
- Declaring "done" without office_verify, or ignoring its warnings.
- Fixing cosmetics while wrong numbers remain.
- "Fixing" a template's intentional direct formatting.
- Filling unknown data to make the check pass. "[da definire]" is a valid, honest state.
- Over-long QA reports; the user wants the verdict and open points.

## Verification checklist
- [ ] office_verify clean or every remaining issue explained.
- [ ] No placeholders; numbers recomputed; consistency vs memory confirmed.
- [ ] Structure/fields/layout checks done; TOC and fields updated last.
- [ ] Short PASS/FAIL report with open points delivered.
