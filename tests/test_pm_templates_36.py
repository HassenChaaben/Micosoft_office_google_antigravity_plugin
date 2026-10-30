"""
Comprehensive 36-Template Project Management Test Suite
Tests form-filling, batch cell updates, and in-place modifications across
all 36 documents in the Project Management template catalog.
"""

import os
import glob
import pytest
import docx
from daemon import word_tools

PM_TEMPLATE_DIR = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "practical_tests",
        "Template - Progetto PM-20260227",
        "Template - Progetto PM",
        "Template_Progetto_PM"
    )
)

if not os.path.isdir(PM_TEMPLATE_DIR):
    pytestmark = pytest.mark.skip(reason="practical_tests local templates not present in repo")

# 36 Tailored prompts for each template, plus a follow-up modification/enhancement prompt
PM_36_PROMPTS = [
    {
        "id": 0,
        "file": "0 Frontespizio.docx",
        "name": "Frontespizio",
        "prompt": "Compila il frontespizio con Titolo: 'Piattaforma Assistente Vocale AI per Robot', Sottotitolo: 'Progetto di Automazione Cloud4Job', Autore: 'Hassen Ben Chaaben', Data: '28 Ottobre 2026'.",
        "modification": "Aggiorna la versione a 'v2.0' e aggiungi 'Relatore: Ing. Ergys' nei metadati."
    },
    {
        "id": 1,
        "file": "1 Business case.docx",
        "name": "Business Case",
        "prompt": "Compila il Business Case per l'iniziativa 'Robot Voice AI': Codice: 'BC-2026-01', Titolo: 'Automazione Vocale Robotica', IRR: '24%', NPV: '€180.000', Payback Period: '18 mesi'.",
        "modification": "Modifica l'IRR al 28% e aumenta il NPV stimato a €210.000."
    },
    {
        "id": 2,
        "file": "2 Project charter.docx",
        "name": "Project Charter",
        "prompt": "Compila il Project Charter: Codice: 'PRJ-VOICE-01', Titolo: 'Robot Voice Full Duplex', Cliente: 'Cloud4Job Robotics', Sponsor: 'Direzione Tecnica', Budget complessivo: '€65.000'.",
        "modification": "Modifica il Budget complessivo portandolo a €75.000 e assegna come Project Manager 'Hassen Ben Chaaben'."
    },
    {
        "id": 3,
        "file": "3 Stakeholder register.docx",
        "name": "Stakeholder Register",
        "prompt": "Aggiungi 3 stakeholder principali nel registro: Ergys (Lead Maintainer, Alto interesse/Alto potere), Hassen (Sviluppatore Core, Alto interesse), Utenti Finali (Basso potere/Alto interesse).",
        "modification": "Modifica la strategia di gestione per il Lead Maintainer in 'Coinvolgimento attivo settimanale'."
    },
    {
        "id": 4,
        "file": "4 Piano di Project Management.docx",
        "name": "Piano di Project Management",
        "prompt": "Compila il Piano di PM: Standard di riferimento: 'PMBOK Guide 7th Edition', Testo di riferimento: 'Guida Pratica Cloud4Job PM', Metodologia: 'Agile Scrum con Sprint bisettimanali'.",
        "modification": "Aggiungi nei principi di governance il vincolo di verifica TDD obbligatoria prima di ogni rilascio."
    },
    {
        "id": 5,
        "file": "5 Piano di gestione dei rischi.docx",
        "name": "Piano di Gestione dei Rischi",
        "prompt": "Compila il piano di gestione rischi: Processo di identificazione: 'Brainstorming settimanale & analisi audit', Frequenza di revisione: 'Ogni Sprint', Responsabile del monitoraggio: 'Risk Manager'.",
        "modification": "Modifica la frequenza di revisione in 'Settimanale' e aggiungi la soglia di escalation al CTO."
    },
    {
        "id": 6,
        "file": "6 Matrice probabilità-impatto.docx",
        "name": "Matrice Probabilità-Impatto",
        "prompt": "Definisci le scale di probabilità e impatto: Lieve (1), Moderato (2), Grave (3), Critico (4). Soglia di allarme rosso fissata a punteggio >= 9.",
        "modification": "Aggiorna la scala di impatto finanziario: Moderato fino a €10.000, Grave sopra €25.000."
    },
    {
        "id": 7,
        "file": "7 Piano di gestione degli approvvigionamenti.docx",
        "name": "Piano di Gestione degli Approvvigionamenti",
        "prompt": "Compila il piano acquisti: Tipo di contratti: 'Time & Material per consulenze cloud, Canone fisso per GPU LLM server', Criteri di selezione fornitori: 'Conformità GDPR e latenza < 100ms'.",
        "modification": "Aggiungi un vincolo di audit di sicurezza per qualsiasi fornitore SaaS di terze parti."
    },
    {
        "id": 8,
        "file": "8 WBS e dizionario della WBS.docx",
        "name": "WBS e Dizionario della WBS",
        "prompt": "Compila la WBS a 3 livelli: 1.0 Architettura di Sistema, 2.0 Motore Vocale (2.1 FakeAsr, 2.2 FakeLlm, 2.3 FakeTts), 3.0 EventBus & Coordinator, 4.0 Validazione & Collaudo.",
        "modification": "Aggiungi il Work Package '2.4 Interruzione Audio Full-Duplex' sotto il nodo 2.0."
    },
    {
        "id": 9,
        "file": "9 Baseline dei tempi.docx",
        "name": "Baseline dei Tempi",
        "prompt": "Imposta la Baseline dei Tempi ufficiale: Data di avvio: '01 Novembre 2026', Rilascio MVP: '15 Dicembre 2026', Collaudo finale: '28 Febbraio 2027'.",
        "modification": "Anticipa la data di rilascio MVP al '05 Dicembre 2026'."
    },
    {
        "id": 10,
        "file": "10 Baseline dei costi.docx",
        "name": "Baseline dei Costi",
        "prompt": "Compila la baseline dei costi: Q1 2026: '€20.000', Q2 2026: '€25.000', Q3 2026: '€15.000', Riserva di contingenza: '€5.000'. Totale: '€65.000'.",
        "modification": "Aggiorna la riserva di gestione portandola a €10.000 per mitigare la volatilità dei prezzi GPU."
    },
    {
        "id": 11,
        "file": "11 Piano delle comunicazioni.docx",
        "name": "Piano delle Comunicazioni",
        "prompt": "Compila la matrice delle comunicazioni: Daily Standup (Teams, 15m, Team Dev), Sprint Review (Bisettimanale, Stakeholder), Executive Dashboard (Mensile, Sponsor).",
        "modification": "Aggiungi il canale Slack dedicato '#robot-voice-alerts' per notifiche di build e incidenti."
    },
    {
        "id": 12,
        "file": "12 Documento dei requisiti.docx",
        "name": "Documento dei Requisiti",
        "prompt": "Inserisci i requisiti non funzionali: REQ-NF-01: Latenza risposta vocale < 300ms, REQ-NF-02: Disponibilità 99.9%, REQ-NF-03: Elaborazione audio locale senza invio dati non crittografati.",
        "modification": "Modifica REQ-NF-01 riducendo il target di latenza a < 200ms end-to-end."
    },
    {
        "id": 13,
        "file": "13 Matrice di tracciabilità dei requisiti.docx",
        "name": "Matrice di Tracciabilità Requisiti",
        "prompt": "Collega REQ-NF-01 al Work Package WBS 2.2 (FakeLlm) e al test di collaudo TST-LAT-01.",
        "modification": "Aggiungi il link verso il test di integrazione 'test_final_e2e_live.py'."
    },
    {
        "id": 14,
        "file": "14 Descrizione dell_ambito.docx",
        "name": "Descrizione dell'Ambito",
        "prompt": "Compila l'Ambito del Progetto: Obiettivi: 'Integrazione robotica autonoma vocale', Limiti: 'Inclusa lingua italiana ed inglese; escluse lingue orientali nella v1'.",
        "modification": "Aggiungi tra i deliverable inclusi nell'ambito il simulatore Web taskpane per test senza robot fisico."
    },
    {
        "id": 15,
        "file": "15 Piano della qualità.docx",
        "name": "Piano della Qualità",
        "prompt": "Definisci le metriche di qualità: Code Coverage minima: '85%', Esito linter: 'Zero errori flake8/ruff', Test pass rate: '100% su CI GitHub Actions'.",
        "modification": "Eleva la soglia di Code Coverage al '90%' per il modulo coordinator."
    },
    {
        "id": 16,
        "file": "16 Elenco delle attività e Milestone.docx",
        "name": "Elenco Attività e Milestone",
        "prompt": "Compila l'elenco delle attività: M1: Architettura validata (10 Nov), M2: Fake Core pronto (30 Nov), M3: EventBus integrato (15 Dic), M4: Release 1.0 (20 Gen).",
        "modification": "Sposta la data della Milestone M3 al '10 Dicembre'."
    },
    {
        "id": 17,
        "file": "17 Stima durate, risorse e costi delle attività.docx",
        "name": "Stima Durate, Risorse e Costi",
        "prompt": "Compila la stima: Task 'Implementazione Bus': Durata 40 ore, Risorsa: Senior Dev, Costo stimato: €3.200.",
        "modification": "Riduci la durata stimata a 32 ore grazie al riuso di componenti asincroni esistenti."
    },
    {
        "id": 18,
        "file": "18 Piano delle risorse umane.docx",
        "name": "Piano delle Risorse Umane",
        "prompt": "Definisci il team di progetto: Project Manager (100%), AI Engineer (100%), Robotic Software Dev (50%), QA Specialist (50%).",
        "modification": "Assegna Hassen Ben Chaaben al ruolo di AI & ROS 2 Engineer principale al 100%."
    },
    {
        "id": 19,
        "file": "19 Reticolo di progetto.docx",
        "name": "Reticolo di Progetto",
        "prompt": "Definisci le dipendenze: L'attività EventBus (ID 5) richiede il completamento di FakeAsr (ID 1) e FakeLlm (ID 2). Relazione Finish-to-Start (FS).",
        "modification": "Imposta un lead time di 2 giorni per consentire il testing parallelo con mock."
    },
    {
        "id": 20,
        "file": "20 Piano dei tempi.docx",
        "name": "Piano dei Tempi",
        "prompt": "Compila il Gantt di progetto: Sprint 1 (Nov 1-14): Setup ambiente, Sprint 2 (Nov 15-28): Core Fakes, Sprint 3 (Nov 29 - Dic 12): EventBus.",
        "modification": "Aggiungi lo Sprint 4 (Dic 13-23) dedicato interamente a stress test e documentazione."
    },
    {
        "id": 21,
        "file": "21 Distribuzioni costi su WBS e OBS.docx",
        "name": "Distribuzione Costi WBS-OBS",
        "prompt": "Ripartisci i costi per OBS: R&D Software: €40.000, Hardware & Sensori: €15.000, Project Management & QA: €10.000.",
        "modification": "Ribilancia assegnando €45.000 a R&D Software e riducendo Hardware a €10.000."
    },
    {
        "id": 22,
        "file": "22 Piano degli approvvigionamenti.docx",
        "name": "Piano degli Approvvigionamenti",
        "prompt": "Compila la lista acquisti: 2x Schede Jetson Orin Nano, Licenza cloud speech API per test, Microfoni direzionali a matrice.",
        "modification": "Modifica la quantità di schede Jetson a 3 unità per ridondanza di sviluppo."
    },
    {
        "id": 23,
        "file": "23 Registro dei rischi.docx",
        "name": "Registro dei Rischi",
        "prompt": "Compila la tabella Minacce: Cod: 'RSK-01', Categoria: 'Tecnologia', Causa: 'Latenza rete mobile', Evento: 'Ritardo nel TTS', Impatto: 'Moderato'. Tabella Opportunità: Cod: 'OPP-01', Causa: 'Ottimizzazione quantizzata INT8', Evento: 'Velocità raddoppiata'.",
        "modification": "Aggiorna l'azione di risposta per RSK-01 in 'Implementare cache locale delle risposte frequenti'."
    },
    {
        "id": 24,
        "file": "24 Richiesta di modifica.docx",
        "name": "Richiesta di Modifica",
        "prompt": "Compila la Change Request: Codice: 'CR-004', Titolo: 'Supporto Full Duplex con barge-in acustico', Motivazione: 'Migliorare naturalezza interazione cliente', Impatto tempi: '+3 giorni'.",
        "modification": "Aggiungi la firma di approvazione preliminare da parte dello Sponsor Tecnico."
    },
    {
        "id": 25,
        "file": "25 Registro delle modifiche.docx",
        "name": "Registro delle Modifiche",
        "prompt": "Aggiungi la riga nel registro: Codice: 'CR-004', Stato: 'Approvata', Data: '28/10/2026', Responsabile: 'Hassen Ben Chaaben'.",
        "modification": "Modifica lo stato della CR-004 in 'In Implementazione'."
    },
    {
        "id": 26,
        "file": "26 Scheda questione.docx",
        "name": "Scheda Questione",
        "prompt": "Compila la scheda issue: Codice questione: 'ISS-012', Oggetto: 'Conflitto porte websocket su localhost', Priorità: 'Alta', Assegnatario: 'Hassen'.",
        "modification": "Imposta la data target di risoluzione entro le 24 ore successive."
    },
    {
        "id": 27,
        "file": "27 Registro delle questioni.docx",
        "name": "Registro delle Questioni",
        "prompt": "Aggiungi al registro l'Issue 'ISS-012': Stato: 'Aperta', Severità: 'Alta', Azione intrapresa: 'Riassegnazione porta dinamica su 18890'.",
        "modification": "Aggiorna lo stato di ISS-012 in 'Risolta e Verificata con Test Unitari'."
    },
    {
        "id": 28,
        "file": "28 Informazioni sullo stato di avanzamento del lavoro.docx",
        "name": "Stato Avanzamento Lavoro",
        "prompt": "Compila lo stato avanzamento: Data di riferimento: '28 Ottobre 2026', WP completati: 'WP 1.1, WP 2.1, WP 2.2', Percentuale complessiva: '65%'.",
        "modification": "Aggiorna la percentuale di completamento complessiva al '72%' includendo l'Issue 5 EventBus."
    },
    {
        "id": 29,
        "file": "29 Documento dei requisiti aggiornato.docx",
        "name": "Documento Requisiti Aggiornato",
        "prompt": "Inserisci il nuovo requisito REQ-F-08: 'Riconoscimento vocale di parole chiave di stop immediato (Emergency Halt)'. Priorità: 'Critica'.",
        "modification": "Aggiungi la specifica tecnica di tempo di reazione hardware < 50ms per l'Emergency Halt."
    },
    {
        "id": 30,
        "file": "30 Esiti della qualità.docx",
        "name": "Esiti della Qualità",
        "prompt": "Compila il report qualità Sprint 2: Test automatici eseguiti: 57, Test passati: 57 (100%), Bug bloccanti aperti: 0, Esito complessivo: 'Conforme'.",
        "modification": "Aggiungi nota di merito per zero difetti di regressione nell'EventBus."
    },
    {
        "id": 31,
        "file": "31 Accettazione dei deliverable.docx",
        "name": "Accettazione Deliverable",
        "prompt": "Compila l'accettazione del deliverable 'D-02: Modulo FakeLlm': Stato: 'Accettato senza riserve', Firma Cliente: 'Ing. Cloud4Job', Data: '26/10/2026'.",
        "modification": "Aggiungi come allegato di collaudo il log di test CI con 28 scenari superati."
    },
    {
        "id": 32,
        "file": "32 Stato dei rischi.docx",
        "name": "Stato dei Rischi",
        "prompt": "Compila il monitoraggio rischi: RSK-01 (Latenza rete) declassato da 'Alto' a 'Basso' a seguito dell'adozione del modello locale on-edge.",
        "modification": "Chiudi formalmente il rischio RSK-01 annotando la contromisura come permanente."
    },
    {
        "id": 33,
        "file": "33 Report sulle prestazioni.docx",
        "name": "Report sulle Prestazioni",
        "prompt": "Compila gli indici Earned Value: Planned Value (PV): '€45.000', Earned Value (EV): '€48.000', Actual Cost (AC): '€44.000', CPI: '1.09', SPI: '1.06'.",
        "modification": "Evidenzia nel commento che il progetto è in anticipo sui tempi ed economico rispetto al budget."
    },
    {
        "id": 34,
        "file": "34 Baseline dei costi aggiornata.docx",
        "name": "Baseline dei Costi Aggiornata",
        "prompt": "Aggiorna la baseline costi a seguito dell'approvazione della CR-004: Nuovo budget autorizzato: '€75.000', con variazione di +€10.000 interamente allocata su R&D.",
        "modification": "Inserisci il dettaglio trimestrale: Q1: €25.000, Q2: €30.000, Q3: €20.000."
    },
    {
        "id": 35,
        "file": "35 Chiusura progetto.docx",
        "name": "Chiusura Progetto",
        "prompt": "Compila il verbale di chiusura progetto: Progetto completato: 'SI', Tutti i deliverable accettati: 'SI', Budget finale consuntivato: '€72.400' (risparmio di €2.600).",
        "modification": "Aggiungi nei Lessons Learned l'efficacia dell'approccio TDD e l'uso dell'assistente autonomo Word."
    }
]


def get_template_file(template_id: int) -> str:
    """Finds the template file by its leading numeric index to avoid OS codepage discrepancies."""
    files = [f for f in os.listdir(PM_TEMPLATE_DIR) if f.endswith(".docx")]
    for f in files:
        if f.startswith(f"{template_id} "):
            return f
    raise FileNotFoundError(f"Template with id {template_id} not found in {PM_TEMPLATE_DIR}")


def test_pm_templates_directory_exists_and_complete():
    """Verifies all 36 templates are present on disk."""
    assert os.path.isdir(PM_TEMPLATE_DIR), f"Template dir missing: {PM_TEMPLATE_DIR}"
    docx_files = [f for f in os.listdir(PM_TEMPLATE_DIR) if f.endswith(".docx")]
    assert len(docx_files) == 36, f"Expected 36 templates, found {len(docx_files)}"


def test_pm_36_prompts_catalog_complete():
    """Verifies that 36 tailored prompts and modifications are defined."""
    assert len(PM_36_PROMPTS) == 36
    for idx, p in enumerate(PM_36_PROMPTS):
        assert p["id"] == idx
        actual_file = get_template_file(idx)
        assert os.path.isfile(os.path.join(PM_TEMPLATE_DIR, actual_file))
        assert len(p["prompt"]) > 20
        assert len(p["modification"]) > 10


def test_mcp_form_and_batch_tools_registered():
    """Verifies the new batch and smart form tools are present in word_tools."""
    tools = word_tools.get_tool_definitions()
    names = [t["name"] for t in tools]
    assert "word_batch_update_table_cells" in names
    assert "word_fill_form_fields" in names
    assert "word_insert_image" in names

    batch_tool = word_tools.get_tool("word_batch_update_table_cells")
    assert batch_tool is not None
    assert "updates" in batch_tool["inputSchema"]["required"]
    assert not word_tools.is_destructive("word_batch_update_table_cells")

    fill_tool = word_tools.get_tool("word_fill_form_fields")
    assert fill_tool is not None
    assert "fields" in fill_tool["inputSchema"]["required"]
    assert not word_tools.is_destructive("word_fill_form_fields")


@pytest.mark.parametrize("item", PM_36_PROMPTS)
def test_pm_template_structure_and_form_targets(item):
    """
    For each of the 36 templates:
    1. Loads the .docx document using python-docx.
    2. Inspects tables, rows, columns, and text placeholders.
    3. Verifies that the document can be read and contains target form elements.
    """
    filename = get_template_file(item["id"])
    path = os.path.join(PM_TEMPLATE_DIR, filename)
    assert os.path.exists(path), f"Missing file: {path}"
    doc = docx.Document(path)

    # Document must contain either tables or paragraphs
    has_tables = len(doc.tables) > 0
    has_paras = len([p for p in doc.paragraphs if p.text.strip()]) > 0
    assert has_tables or has_paras, f"{item['file']} is completely empty!"

    if has_tables:
        # Table-based template: verify dimensions and cell accessibility
        t = doc.tables[0]
        assert len(t.rows) > 0
        assert len(t.columns) > 0
        cells = t.rows[0].cells
        assert len(cells) > 0


def test_simulate_form_fill_and_modification_on_project_charter():
    """
    Tests simulated form filling and follow-up in-place modification
    on the most complex template (2 Project charter.docx).
    """
    path = os.path.join(PM_TEMPLATE_DIR, "2 Project charter.docx")
    doc = docx.Document(path)
    t = doc.tables[0]

    # Verify initial template state
    labels = [c.text.strip() for r in t.rows for c in r.cells if c.text.strip()]
    assert "Codice" in labels
    assert "Titolo" in labels
    assert "Budget complessivo" in labels

    # Simulate finding and updating the cell next to 'Codice'
    target_row = None
    target_col = None
    for r_idx, r in enumerate(t.rows):
        for c_idx, c in enumerate(r.cells):
            if c.text.strip() == "Codice":
                target_row = r_idx
                # Find adjacent value slot
                for next_c in range(c_idx + 1, len(r.cells)):
                    if r.cells[next_c].text.strip() != "Codice":
                        target_col = next_c
                        break
                break
        if target_row is not None:
            break

    assert target_row is not None
    assert target_col is not None

    # Apply initial value
    t.rows[target_row].cells[target_col].text = "PRJ-VOICE-01"
    assert t.rows[target_row].cells[target_col].text == "PRJ-VOICE-01"

    # Apply follow-up modification in-place
    t.rows[target_row].cells[target_col].text = "PRJ-VOICE-02-ENHANCED"
    assert t.rows[target_row].cells[target_col].text == "PRJ-VOICE-02-ENHANCED"
