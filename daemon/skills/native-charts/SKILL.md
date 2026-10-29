---
name: native-charts
description: Insert native, editable Word charts (column, bar, line, pie, area, scatter, combo) chosen by the data story, with captions and Italian number formats.
hosts: [word]
always_on: false
triggers:
  doc_types: []
  keywords: [grafico, grafici, chart, charts, istogramma, histogram, torta, pie, linee, line chart, barre, bar chart, colonne, curva a s, s-curve, andamento, trend, confronto, comparison, distribuzione, distribution, dashboard, kpi, visualizza dati, plot]
---
# Native charts

## Office mental model
- word_insert_chart creates a NATIVE chart (embedded workbook) the user can edit with
  "Modifica dati". Never paste images of charts or ASCII charts.
- Choose the chart from the data STORY, not from taste:
  | Story | Chart | Notes |
  |---|---|---|
  | Compare categories (costs per WP, vendors) | bar (long labels) / column (≤ 8 short) | sort descending, one color |
  | Trend over time (monthly costs, KPIs) | line | time on X, ≤ 4 series |
  | Cumulative over time (curva a S, PV/EV/AC) | line | cumulative values; EVM: 3 series PV, EV, AC |
  | Part of a whole (budget split) | pie only if ≤ 5 slices summing to 100% | else bar |
  | Composition over time | area (stacked) or stacked column | |
  | Relationship two measures (P vs impatto) | scatter | label points |
  | Two units (costs € + % avanzamento) | combo (column + line, secondary axis) | |
- Every chart: clear title (the message, e.g. "I costi di integrazione pesano il 42%"),
  axis titles with units ("€", "giorni", "%"), caption "Grafico n"/"Figura n" below,
  referenced in the text. Data in the chart must match the table/text exactly.
- Italian documents: category names in Italian, months "Gen, Feb, Mar", units "€".
  Values passed as numbers (no thousand separators in JSON); Word formats them.

## Procedure
1. Identify the dataset (from a table via word_read_table, from the user, or researched
   facts with source). Check totals and units.
2. Pick type by the table above; limit series/categories; decide order.
3. agent_plan if more than one chart; insert each chart after its introducing paragraph.
4. word_insert_chart with categories, series, title, axes, location.
5. word_insert_caption (label "Figura" or "Grafico") and a cross-reference in the text.
6. office_verify; state the data source in the caption or a note when external.

## Tool recipes
Cost comparison (caption anchor = the ga:<id> returned by word_insert_chart):
```json
word_insert_chart({"type": "bar", "title": "Costo stimato per work package",
  "categories": ["1.1 Analisi", "1.2 Migrazione dati", "1.3 Integrazioni", "1.4 Formazione"],
  "series": [{"name": "Costo (€)", "values": [12000, 21000, 31500, 10500]}],
  "axes": {"x": "€", "y": "Work package"}, "location": "after:p34"})
word_insert_caption({"anchor": "ga:chart_costi_wp", "label": "Figura", "text": "Costo stimato per work package"})
```
Earned value S-curves:
```json
word_insert_chart({"type": "line", "title": "Andamento PV, EV, AC al 31/05/2026",
  "categories": ["Feb", "Mar", "Apr", "Mag"],
  "series": [{"name": "PV", "values": [10000, 25000, 42000, 58000]},
             {"name": "EV", "values": [9000, 22000, 37000, 51000]},
             {"name": "AC", "values": [9500, 24500, 41000, 56000]}],
  "axes": {"x": "Mese", "y": "€ cumulati"}, "location": "end"})
```
Combo costs + progress:
```json
word_insert_chart({"type": "combo", "title": "Costi mensili e avanzamento fisico",
  "categories": ["Feb", "Mar", "Apr"],
  "series": [{"name": "Costi (€)", "values": [9500, 15000, 16500]}, {"name": "Avanzamento (%)", "values": [15, 38, 64]}],
  "axes": {"x": "Mese", "y": "€ / %"}, "location": "end"})
```

## Chart styling rules
- One message per chart; sort bars by value unless the order is natural (time, WBS code).
- Max 6 categories in a pie, 12 bars, 4 lines; otherwise aggregate into "Altro".
- Same series = same color across all charts in the document (e.g. PV blue, EV green, AC orange).
- Start value axes at zero for bars/columns; lines may zoom but say so.
- Chart near its introducing paragraph; width = text width; in landscape sections for wide time axes.

## Pitfalls
- Pie charts with many slices or values not summing to a whole; 3D charts.
- Line chart for unordered categories; bar chart for time series with 24 points.
- Values typed as strings with "." thousands -> wrong numbers. Pass plain numbers.
- Chart data different from the table in the document.
- No caption/axis units; legend for a single series (redundant).
- Inserting the chart in a form table cell of a template not designed for it.

## Verification checklist
- [ ] Chart type matches the data story; ≤ 5-6 series, readable labels.
- [ ] Title states the message; axis titles with units; caption + text reference.
- [ ] Numbers consistent with source table/text; source cited if external.
- [ ] office_verify clean.
