---
name: pm-documentation-suite
description: Italian project-management document suite (36 templates from Frontespizio to Chiusura progetto) - purpose of each document and cross-document consistency rules.
hosts: [word]
always_on: false
triggers:
  doc_types: [form_template]
  keywords: [project management, project charter, business case, stakeholder, registro dei rischi, risk register, rischi, risk, wbs, dizionario della wbs, baseline, matrice probabilità, probabilità, impatto, emv, earned value, cpi, spi, eac, raci, obs, milestone, approvvigionamenti, requisiti, tracciabilità, questione, richiesta di modifica, change request, chiusura progetto, lesson learned, avanzamento, prestazioni, piano dei tempi, piano dei costi, pmbok]
---
# PM documentation suite (Italian, PMBOK-style)

## Office mental model
The suite is 36 Word templates (tables + direct formatting) that describe ONE project over
its life cycle. Each is a form: fill it with template-form-filling rules. The expertise here
is WHAT goes in each document and keeping all of them CONSISTENT.

Avvio (initiating)
- 0 Frontespizio: titolo del progetto, autori (Nome Cognome) — cover of the dossier.
- 1 Business case: iniziativa, iniziatore, impulsi scatenanti (domanda di mercato, requisito
  di legge, bisogno interno, cliente esterno...), prodotto/servizi, benefici, costi stimati.
- 2 Project charter: anagrafica (codice, titolo), cliente, sponsor, descrizione e
  giustificazione, budget complessivo, obiettivi (tipo tempi/costi/qualità/ambito, cod,
  criterio), deliverable principali, milestone principali, rischi principali, PM e autorità.
- 3 Stakeholder register: id, ruolo, organizzazione, requisiti di comunicazione, aspettative,
  potere/influenza/interesse/impatto, valutazione globale, strategia (gestire da vicino,
  soddisfare, informare, monitorare).
Pianificazione (planning)
- 4 Piano di Project Management: standard (PMBOK/ISO 21500), fasi e processi con input,
  metodi, strumenti, output; elenco dei piani sussidiari con versione.
- 5 Piano di gestione dei rischi / 7 Piano di gestione degli approvvigionamenti:
  processo -> azioni, stakeholder coinvolti, metodi e approcci, template usati.
- 6 Matrice probabilità-impatto: scale P (Improbabile, Basso, Medio, Alto) and I (Lieve,
  Tollerabile, Moderato, Grave) with numeric values (X), matrix P×I, range del rischio ->
  azione (es. < 4 accettare/monitorare, 4-9 mitigare, > 9 escalation/evitare).
- 8 WBS e dizionario della WBS: codice WP, titolo, responsabile, descrizione, stima durata,
  stima costi. WBS diagram: see smartart-and-diagrams (hierarchy).
- 9 Baseline dei tempi: WBS cod, ES/EF/LS/LF, free/total slack, critical (Sì/No).
- 10 Baseline dei costi: costi e fondi incrementali/cumulativi per periodo, curva a S
  (native line chart of cumulative cost: see native-charts).
- 11 Piano delle comunicazioni: informazione, SH emittente/destinatario, tempistica,
  modalità, tipologia.
- 12 Documento dei requisiti / 29 aggiornato: ID, descrizione, SH richiedente, criterio di
  accettazione. 13 Matrice di tracciabilità: ID, versione, stato, classe (SIC, PER, TEC,
  PROG, PROD), priorità, SH responsabile, WBS.
- 14 Descrizione dell'ambito: obiettivi (ambito/tempi/costi/qualità), confini (in/out of
  scope), deliverable + requisiti + criteri + assunti + vincoli, WBS di massima.
- 15 Piano della qualità: regole QA (codice regola, processo, responsabile, auditor,
  metriche) e QC (deliverable, ispettore, metriche).
- 16 Elenco attività e milestone (WBE/WP, codice attività/milestone).
- 17 Stima durate, risorse e costi (durata, vincoli data, risorse umane/materiali/
  strumentali, quantità, costo).
- 18 Piano risorse umane: OBS (livelli), RAM con approccio RACI per WP, acquisizione.
- 19 Reticolo di progetto: legami (attività condizionante -> condizionata, FS/SS/FF/SF,
  anticipo/ritardo); diagram in smartart-and-diagrams (network, critical path).
- 20 Piano dei tempi: ES, EF, LS, LF, TF = LS − ES, FF, criticità (TF = 0).
- 21 Distribuzione costi su WBS e OBS: costi per livello WBS, per OBS, evoluzione nel tempo.
- 22 Piano degli approvvigionamenti: cosa acquistare, quantità, data attesa, fornitore, WBE.
- 23 Registro dei rischi: minacce e opportunità (cod, categoria, WBS, causa-evento-effetto),
  analisi qualitativa (P, I, fattore = P×I, strategia, priorità), quantitativa (EMV, ETV),
  piano di risposta, piano di contingenza, riserve (contingenza, gestione).
Esecuzione e controllo (executing & monitoring)
- 24 Richiesta di modifica, 25 Registro delle modifiche (codice, data, richiedente, esito).
- 26 Scheda questione, 27 Registro delle questioni (issue: originatore, responsabile,
  azione, stato, data chiusura).
- 28 Stato di avanzamento del lavoro: baseline (BAC, PV), effettivi (AC, EV, % avanzamento
  fisico), nuove previsioni.
- 30 Esiti della qualità (audit QA, ispezioni QC, azioni correttive).
- 31 Accettazione dei deliverable (ispezione, criteri, esito).
- 32 Stato dei rischi (in essere/accaduta/chiusa, EMV aggiornato, contingenza usata/rilasciata).
- 33 Report sulle prestazioni: deliverable completati/non, modifiche, criticità, EVM per WP
  (PV, AC, EV, SV, CV, SPI, CPI) e previsioni a finire (BAC, EAC1/2/3, ETC, VAC).
- 34 Baseline dei costi aggiornata (data di aggiornamento nel titolo).
Chiusura
- 35 Chiusura progetto: completato/approvato (SI/NO), firme sponsor e PM, grado di
  raggiungimento obiettivi (tempi, costi, qualità), deliverable, modifiche, lesson learned.

## Procedure
1. Identify which of the 36 documents is open (title in T0 or first paragraph).
2. memory_search for the project master data; if missing, ask once or use what the user
   gave and memory_save it (scope "document") so every later document reuses it:
   codice progetto, titolo, cliente, sponsor, PM, budget, data inizio/fine, WBS codes,
   stakeholder list (id + Cognome Nome + ruolo), risk codes, requirement IDs.
3. agent_plan per table; apply template-form-filling (clone rows, Italian formats, "[da definire]").
4. Apply the consistency rules and formulas below; compute, never guess, derived numbers.
5. office_verify; report numbers you computed and any inconsistency found with memory.

## Cross-document consistency rules
- Same codice progetto and titolo everywhere; document code pattern
  `<codice progetto>-<sigla doc>-<versione>` (e.g. PRJ-2026-014-RR-01) unless the user has one.
- Budget: charter budget = sum of baseline dei costi = WBS cost distribution total
  (+ riserve di contingenza; riserva di gestione is outside the cost baseline). BAC in 28/33
  = baseline cost of the WP/project.
- Dates: milestone dates in charter = baseline dei tempi = piano dei tempi; all within
  project start/end; gg/mm/aaaa.
- Stakeholder names and roles identical in charter, stakeholder register, communication plan,
  RACI, approvals and signatures ("Cognome Nome").
- WBS codes (1, 1.1, 1.1.1) identical in WBS dictionary, activity list, schedule, cost
  distribution, risk register (WBS column), requirements traceability, change requests.
- Risk codes (M01.. minacce, O01.. opportunità) identical across registro, analisi,
  risposta, contingenza, stato dei rischi.
- Requirement IDs identical in 12, 13, 14, 29, 31.

## Formulas (show them in notes when useful)
- Fattore di rischio = P × I on the 1-5 (or template's) numeric scale; priority from the
  ranges of document 6. Keep qualitative label coherent with the number.
- EMV = P (probability 0-1 or %) × impatto economico (€); threats negative/opportunities
  positive in the totals; contingency reserve ≈ Σ EMV of accepted/mitigated threats.
- ETV = P × impatto sui tempi (giorni).
- EVM: SV = EV − PV; CV = EV − AC; SPI = EV / PV; CPI = EV / AC;
  EAC1 = AC + (BAC − EV) (senza trend); EAC2 = BAC / CPI (trend CPI);
  EAC3 = AC + (BAC − EV) / (CPI × SPI); ETC = EAC − AC; VAC = BAC − EAC.
  CPI/SPI with 2 decimals (0,92); < 1 = over budget / behind schedule.
- Schedule: EF = ES + durata; LS = LF − durata; TF = LS − ES; critical path = TF 0.

## Tool recipes
```json
memory_save({"text": "PRJ-2026-014 'Migrazione CRM in cloud'; cliente Alfa S.p.A.; sponsor Rossi Mario; PM Verdi Anna; budget € 75.000,00; 01/02/2026-30/09/2026", "scope": "document"})
word_fill_fields({"fields": {
  "Analisi qualitativa dei rischi identificati › Minacce › #1 › Cod.": "M01",
  "Analisi qualitativa dei rischi identificati › Minacce › #1 › Probabilità accadimento › Val. num.": "4",
  "Analisi qualitativa dei rischi identificati › Minacce › #1 › Impatto › Val. num.": "3",
  "Analisi qualitativa dei rischi identificati › Minacce › #1 › Fattore di rischio › Val. num.": "12",
  "Analisi qualitativa dei rischi identificati › Minacce › #1 › Strategia": "Mitigare"}})
word_insert_chart({"type": "line", "title": "Curva a S dei costi cumulati",
  "categories": ["Gen", "Feb", "Mar", "Apr"],
  "series": [{"name": "Costi cumulati (€)", "values": [8000, 21000, 39000, 52000]}],
  "axes": {"x": "Periodo", "y": "€"}, "location": "after:Grafico della curva dei costi (curva “ad S”)"})
word_add_comment({"anchor": "T5r4c6", "text": "EMV = 0,40 × € 12.000,00 = € 4.800,00"})
```

## Pitfalls
- Different budget/date/name in two documents of the same project (most common reviewer complaint).
- Fattore di rischio not equal to P × I; EMV computed with % as 40 instead of 0,40.
- Mixing WBS levels (1.2 vs 1.2.0) or renumbering WBS between documents.
- Confusing riserva di contingenza (inside baseline) with riserva di gestione (outside).
- Translating Italian PM terms to English inside an Italian template.

## Verification checklist
- [ ] Project master data identical to memory (code, title, client, sponsor, PM, budget, dates).
- [ ] WBS, risk, requirement and stakeholder codes consistent with other documents.
- [ ] All derived numbers recomputed (P×I, EMV, ETV, SV/CV/SPI/CPI/EAC/ETC/VAC, TF).
- [ ] Charts/diagrams the template asks for are native (curva a S, WBS, reticolo).
- [ ] template-form-filling checklist satisfied; office_verify clean.
