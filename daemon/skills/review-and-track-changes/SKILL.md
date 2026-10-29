---
name: review-and-track-changes
description: Professional review of a document with Track Changes and comments - redlining contracts, copy-editing, reviewer feedback - so every edit stays visible and reversible.
hosts: [word]
always_on: false
triggers:
  doc_types: [contract]
  keywords: [revisione, revisiona, rivedi, review, revisioni, track changes, revisioni tracciate, rilevamento modifiche, traccia modifiche, redline, commento, commenti, comment, comments, correggi, proofread, copy edit, bozza, feedback, osservazioni, contratto, contract, clausola, clause]
---
# Review and track changes

## Office mental model
- Track Changes records every insertion/deletion/format change with author and date; the
  owner accepts or rejects them. A review WITHOUT tracking is a silent rewrite: unacceptable
  for contracts, shared drafts or anything the user did not ask to rewrite.
- Comments (balloons) carry reasoning, questions and suggestions without changing text.
  Anchor them to the smallest meaningful range (paragraph, cell, clause heading).
- Review levels:
  1. Copy-edit: typos, grammar, punctuation, consistency (terms, capitalisation, numbers,
     dates, Italian vs English formats) -> tracked edits directly.
  2. Substantive: unclear obligations, missing data, contradictions between sections,
     wrong calculations -> comment explaining + proposed wording (tracked) when safe.
  3. Legal/commercial risk (contracts): uncapped liability, unilateral termination,
     vague deliverables/SLAs, missing penalties, payment terms, governing law -> comment with
     risk level (Alto/Medio/Basso) and alternative clause; do not change legal substance
     without the user asking.
- Consistency is a review target: defined terms ("il Fornitore") used identically,
  amounts in words and figures matching ("€ 75.000,00 (settantacinquemila/00)"), article
  cross-references correct.

## Procedure
1. Analyze map + word_read_section of the parts to review; memory_search for the user's
   preferences (tone, terminology, which party they represent).
2. word_track_changes on BEFORE the first edit; keep it on (tell the user).
3. agent_plan by section: copy-edit pass, consistency pass, risk/content pass.
4. Make edits with the smallest operation: word_replace_text for words/phrases,
   word_set_paragraph_text for a rewritten sentence/paragraph, word_update_block for a block
   the user asked to rewrite. Avoid deleting and re-inserting whole sections (noisy redline).
5. Add word_add_comment for every non-trivial change or question (why + suggestion).
6. office_verify; final report: N tracked edits, N comments, top 3 issues by severity.
   Never accept/reject the user's existing revisions unless asked.

## Tool recipes
```json
word_track_changes({"on": true})
word_replace_text({"find": "entro 30 giorni lavorativi", "replace": "entro 30 (trenta) giorni di calendario"})
word_set_paragraph_text({"anchor": "p47", "text": "Il Fornitore garantisce la disponibilità del servizio pari al 99,5% su base mensile, misurata secondo l'Allegato B."})
word_add_comment({"anchor": "Articolo 9 – Limitazione di responsabilità", "text": "Rischio ALTO: la responsabilità del Fornitore non ha un massimale. Proposta: limitare al 100% del corrispettivo annuo, salvo dolo o colpa grave."})
word_add_comment({"anchor": "T2r3c2", "text": "Il totale (€ 74.500,00) non coincide con il budget dell'Art. 5 (€ 75.000,00). Verificare."})
```

## Review heuristics (contracts and formal documents)
| Clause / area | Typical issue | Comment template |
|---|---|---|
| Oggetto / Deliverable | vague scope ("supporto", "attività varie") | "Specificare deliverable e criteri di accettazione (rif. Allegato)." |
| Corrispettivi e pagamenti | no milestones, no invoicing terms | "Legare i pagamenti alle milestone; termini 30/60 gg d.f.f.m." |
| SLA / penali | SLA without measurement or cap on penalties | "Definire metrica, finestra di misura e tetto penali (es. 10%)." |
| Responsabilità | unlimited liability | "Massimale pari al corrispettivo annuo, salvo dolo/colpa grave." |
| Recesso / risoluzione | unilateral, no notice | "Preavviso di 60 gg e rimborso costi sostenuti." |
| Riservatezza / dati | no GDPR roles | "Nominare responsabile del trattamento ex art. 28 GDPR." |
| Foro / legge | missing | "Indicare legge italiana e foro competente." |
Severity: ALTO (financial/legal exposure), MEDIO (ambiguity likely to cause disputes),
BASSO (wording/style). Group the final report by severity.

## Pitfalls
- Editing with tracking off; turning it off at the end without telling the user.
- Rewriting whole paragraphs for a one-word fix (unreadable redline).
- Comments that only say "unclear" without a proposal; comments on the whole document.
- Changing the meaning of legal clauses silently; "improving" defined terms.
- Accepting the author's previous revisions.
- Using English typographic rules in Italian text (quotes, decimal separators).

## Verification checklist
- [ ] Track Changes on for all edits; existing revisions untouched.
- [ ] Every substantive change has a comment with reason; risks rated.
- [ ] Terminology, numbers and cross-references consistent.
- [ ] office_verify run; summary of edits/comments given to the user.
