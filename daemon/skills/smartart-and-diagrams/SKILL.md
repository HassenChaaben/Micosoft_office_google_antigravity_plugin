---
name: smartart-and-diagrams
description: Native SmartArt and shape diagrams - WBS, organigramma/OBS, processi, reticolo di progetto with critical path, Gantt, flowcharts - chosen and built like a PM/Office expert.
hosts: [word]
always_on: false
triggers:
  doc_types: []
  keywords: [diagramma, diagram, smartart, wbs, work breakdown, organigramma, org chart, obs, gerarchia, hierarchy, processo, process, flusso, flowchart, diagramma di flusso, reticolo, network diagram, pert, cpm, percorso critico, critical path, gantt, cronoprogramma, timeline, linea temporale, ciclo, cycle, piramide, pyramid, matrice swot, swot]
---
# SmartArt and diagrams

## Office mental model
Two native engines, pick by STRUCTURE of the information:
| Need | Tool | Layout/kind |
|---|---|---|
| Hierarchy ≤ ~25 nodes, ≤ 4 levels (WBS summary, OBS, organigramma) | word_insert_smartart | hierarchy / org_chart |
| Linear steps (processo, fasi) ≤ 7 | word_insert_smartart | process |
| Repeating loop (PDCA, ciclo di vita) | word_insert_smartart | cycle |
| Grouped list, pyramid, 2×2 (SWOT) | word_insert_smartart | list / pyramid / matrix |
| Few milestones on a line | word_insert_smartart | timeline |
| Nodes with arbitrary edges (reticolo, decision flow) | word_insert_diagram | network / flowchart |
| Large WBS with codes (> 25 nodes) or WBS with cost/owner per box | word_insert_diagram | wbs |
| Schedule bars over time | word_insert_diagram | gantt |
SmartArt = automatic layout, text-driven, easy to edit, but no arbitrary connectors.
Shape diagrams = grouped shapes + connectors on a canvas, full control (critical path color,
labels on arrows). Both are native and editable; never insert a picture of a diagram.

Domain rules:
- WBS: level 0 = project, level 1 = phases or deliverables (100% rule: children cover all
  the parent scope, no overlap), work packages at the leaves, codes 1, 1.1, 1.1.1 shown in
  the box text ("1.2 Migrazione dati"). Deliverable-oriented nouns, not verbs.
  Codes must equal those in the WBS dictionary / schedule / costs.
- OBS/organigramma: roles + names ("Project Manager – Verdi Anna").
- Reticolo (AON/PDM): node = activity (code, title, durata, ES/EF/LS/LF optional),
  edge = dependency (FS default; label SS/FF/SF and lag "+2g"). Start and End nodes.
  Critical path = activities with total float 0: highlight nodes AND edges (red / bold);
  compute it with forward/backward pass before drawing.
- Gantt: one bar per activity with start/finish (gg/mm/aaaa), milestones as diamonds
  (durata 0), critical tasks highlighted, grouped by WBS; time scale weeks or months.
  Wide -> landscape section.
- Flowchart: start/end terminators, process boxes, decision diamonds with "Sì/No" labels,
  one entry/exit, top-to-bottom.

## Procedure
1. Get the data: from the document tables (word_read_table on WBS dictionary, reticolo
   legami, piano dei tempi) or memory; validate codes and dependencies.
2. For reticolo/Gantt compute ES/EF/LS/LF and float (EF = ES + d; LS = LF − d;
   TF = LS − ES) and mark critical activities.
3. Choose engine with the table above; plan location (landscape for wide/large diagrams).
4. Insert with ONE call; caption "Figura n – ..." below; reference it in text.
5. office_verify; report node count, critical path (e.g. A → C → F → H, 46 gg).

## Tool recipes
WBS as hierarchy SmartArt (small):
```json
word_insert_smartart({"layout": "hierarchy", "location": "after:WBS", "nodes": [
  {"text": "0 CRM Cloud", "level": 1},
  {"text": "1 Analisi", "level": 2}, {"text": "1.1 Requisiti", "level": 3}, {"text": "1.2 Architettura", "level": 3},
  {"text": "2 Realizzazione", "level": 2}, {"text": "2.1 Migrazione dati", "level": 3}, {"text": "2.2 Integrazioni", "level": 3},
  {"text": "3 Rilascio", "level": 2}, {"text": "3.1 Formazione", "level": 3}, {"text": "3.2 Go-live", "level": 3}]})
```
Reticolo with critical path:
```json
word_insert_diagram({"kind": "network", "location": "after:Legami", "nodes": [
  {"id": "S", "text": "Inizio"},
  {"id": "A", "text": "A – Requisiti\n10 gg\nES 0 EF 10", "critical": true},
  {"id": "B", "text": "B – Licenze\n5 gg\nTF 7"},
  {"id": "C", "text": "C – Migrazione\n15 gg\nES 10 EF 25", "critical": true},
  {"id": "E", "text": "Fine"}],
  "edges": [{"from": "S", "to": "A"}, {"from": "A", "to": "B"}, {"from": "A", "to": "C", "critical": true},
            {"from": "B", "to": "E", "type": "FS"}, {"from": "C", "to": "E", "critical": true}]})
```
Gantt (put it in a landscape section first):
```json
word_insert_diagram({"kind": "gantt", "location": "end", "nodes": [
  {"id": "1.1", "text": "Requisiti", "start": "02/02/2026", "finish": "13/02/2026", "critical": true},
  {"id": "1.2", "text": "Licenze", "start": "16/02/2026", "finish": "20/02/2026"},
  {"id": "M1", "text": "Go-live", "start": "30/09/2026", "finish": "30/09/2026", "milestone": true}],
  "edges": [{"from": "1.1", "to": "1.2"}]})
```
Process / cycle:
```json
word_insert_smartart({"layout": "process", "location": "end", "nodes": [
  {"text": "Richiesta di modifica", "level": 1}, {"text": "Valutazione impatti", "level": 1},
  {"text": "Approvazione CCB", "level": 1}, {"text": "Aggiornamento baseline", "level": 1}]})
```

## Pitfalls
- SmartArt hierarchy for a 60-node WBS (unreadable): use word_insert_diagram kind wbs or
  split by level-1 branch (one diagram per phase).
- WBS codes in the diagram different from the WBS dictionary; verbs instead of deliverables.
- Reticolo without computed float -> "critical path" guessed. Always calculate.
- Arrows of the critical path not highlighted, only nodes (or vice versa).
- Gantt in portrait with 40 columns of weeks; Gantt dates in US format.
- Pictures/ASCII art instead of native objects; diagrams without caption.

## Verification checklist
- [ ] Engine matches structure (SmartArt vs shapes); size readable on the page.
- [ ] Codes/names/durations equal the source tables; 100% rule for WBS.
- [ ] Critical path computed (TF = 0) and highlighted on nodes and edges.
- [ ] Caption + text reference; landscape for wide diagrams; office_verify clean.
