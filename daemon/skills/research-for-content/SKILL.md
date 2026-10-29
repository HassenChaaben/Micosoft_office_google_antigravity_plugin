---
name: research-for-content
description: Use the web only to get FACTS (data, norms, prices, definitions, market figures) for document content, with sources - never for how to format Word.
hosts: [word]
always_on: false
triggers:
  doc_types: []
  keywords: [ricerca, cerca, search, research, fonti, fonte, sources, source, dati di mercato, market data, statistiche, statistics, normativa, regulation, legge, law, gdpr, iso, benchmark, prezzi, prices, aggiornato, latest, attuale, current, cita, cite, bibliografia, bibliography]
---
# Research for content

## Office mental model
- Two kinds of knowledge:
  - HOW to build the document (styles, numbering, sections, fields, charts, SmartArt,
    PM methodology): already in the skills. Do NOT search the web for it.
  - WHAT to write (facts that change or that you cannot know: market size, ISTAT/Eurostat
    figures, current regulations, standard numbers, vendor prices, company facts):
    research, with source and date.
- A professional document cites sources: inline "(Fonte: ISTAT, 2025)", table/figure
  source line under the caption, and a "Bibliografia"/"Fonti" section at the end
  (author/organisation, title, year, URL, accessed date "consultato il 12/03/2026").
- Prefer primary/authoritative sources: official statistics (ISTAT, Eurostat, OECD),
  regulators and law texts (Gazzetta Ufficiale, EUR-Lex), standards bodies (ISO, UNI,
  PMI), company reports/investor relations, peer-reviewed or major consultancies. Avoid
  content farms and unsourced blogs. Note when figures are estimates.
- Data that the user's organisation owns (budgets, names, dates) is NOT on the web: ask or
  mark "[da definire]".

## Procedure
1. List the factual gaps in the planned content (agent_plan step "research: ...").
2. memory_search first (facts may already be known for this project/user).
3. Research only those gaps; capture value + unit + date + source URL for each fact.
   Cross-check important numbers with a second source; prefer the most recent.
4. memory_save useful facts (scope "document" for project facts, "global" for reusable
   references) with their source.
5. Write content with word_insert_content: facts in sentences/tables in the document's
   language and number format (convert units/currency formats), citations inline,
   bibliography section at the end (or update it with word_update_block).
6. Charts from researched data: native-charts + source line in caption.
7. office_verify; report which facts came from which sources and anything unverified.

## Tool recipes
```json
memory_search({"query": "dimensione mercato CRM Italia fonte"})
memory_save({"text": "Mercato software CRM Italia 2025: [valore] (Fonte: [organizzazione], [anno], URL) - consultato 12/03/2026", "scope": "document"})
word_insert_content({"location": "after:Contesto di mercato", "blocks": [
  {"type": "paragraph", "runs": [{"text": "Secondo l'ultima rilevazione disponibile, "}, {"text": "[dato]", "bold": true}, {"text": " (Fonte: [organizzazione], [anno])."}]},
  {"type": "caption", "label": "Tabella", "text": "Indicatori di mercato – Fonte: [organizzazione], [anno]"}]})
word_insert_content({"location": "end", "blocks": [
  {"type": "heading", "level": 1, "text": "Fonti"},
  {"type": "numbered", "items": ["[Organizzazione], \"[Titolo]\", [anno]. [URL] (consultato il 12/03/2026)."]}]})
```

## Fact handling rules
| Fact type | Good sources | How to write it |
|---|---|---|
| Statistics (population, GDP, sector size) | ISTAT, Eurostat, OECD, Banca d'Italia | value + unit + year: "4,2 mld € (2024)" |
| Laws and regulations | Normattiva, Gazzetta Ufficiale, EUR-Lex, Garante Privacy | full reference: "Reg. (UE) 2016/679 (GDPR), art. 32" |
| Standards and methods | ISO, UNI, PMI (PMBOK), AXELOS | name + edition: "ISO 21500:2021" |
| Prices and tariffs | vendor price lists, public tenders (ANAC), Consip | "listino pubblico, consultato il gg/mm/aaaa" |
| Company facts | official site, bilancio, investor relations, Registro Imprese | cite page and year |
- Round sensibly (2 significant decimals max in prose); keep the unit next to the number.
- Ranges and estimates: "tra 30 e 40 mln € (stima)". Contradicting sources: report both.
- Quotations: short, in quotes, with source; otherwise paraphrase.
- If research fails: write "[dato da verificare]" and list it in the final report.

## Pitfalls
- Searching "how to insert a TOC in Word" — use the skills/tools instead.
- Figures without date/source; mixing years; copying long passages verbatim.
- Presenting estimates as facts; inventing a source to support a number.
- Leaving US/UK number formats from the source in an Italian document.
- Researching the user's internal data (it is not public) instead of asking.

## Verification checklist
- [ ] Every external fact has source + year (+ access date in bibliography).
- [ ] Key numbers cross-checked; estimates labelled as such.
- [ ] Formats converted to the document language; citations consistent.
- [ ] Bibliography/Fonti section present when sources are used; office_verify clean.
