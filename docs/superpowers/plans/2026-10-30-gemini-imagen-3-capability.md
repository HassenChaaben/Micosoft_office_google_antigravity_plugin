# Piano di Implementazione: Generazione Immagini AI con Gemini Imagen 3

## 1. Obiettivo e Vincoli
L'agente Microsoft Office funziona nativamente con l'abbonamento Google Gemini (tramite Antigravity CLI `agy` e credenziali Windows `gemini:antigravity`), senza richiedere alcuna chiave API per le normali attività sui documenti (scrittura, formattazione, tabelle, SmartArt, grafici).

La funzionalità di **Generazione Immagini AI con Gemini Imagen 3** (`imagen-3.0-generate-002`) è una capacità aggiuntiva e opzionale:
- **Disattivata per impostazione predefinita** se l'utente non ha inserito una chiave API Gemini.
- Se l'utente richiede un'immagine generata dall'AI senza chiave API configurata, l'agente restituisce un messaggio chiaro che invita ad aggiungere la chiave API nelle impostazioni del pannello laterale.
- **Attiva immediatamente** non appena l'utente inserisce la propria chiave API Gemini nelle impostazioni del pannello laterale (`/api/auth/api_key`).
- L'immagine generata viene salvata nella cache locale e inserita nativamente nel documento Word con didascalia automatica tramite `word_insert_image`.

---

## 2. Architettura dei File Coinvolti

| File | Responsabilità |
|---|---|
| `daemon/office/imagen_client.py` (Nuovo) | Client REST per Google Imagen 3 (`imagen-3.0-generate-002:predict`). Gestisce autenticazione con API key, richiesta JSON, decodifica base64 e salvataggio file PNG temporaneo. |
| `daemon/office/word_com_tools.py` | Registra lo strumento `word_generate_image` in `EXTRA_TOOLS`. Verifica la presenza della chiave API tramite `auth_bridge.get_api_key()`, richiama il client Imagen 3 e inserisce l'immagine nel documento Word. |
| `daemon/word_tools.py` | Registra lo schema dello strumento `word_generate_image` per il server MCP e per l'agente. |
| `daemon/mcp_server.py` | Espone `word_generate_image` nel server Model Context Protocol. |
| `daemon/agent_session.py` / `daemon/antigravity_backend.py` | Aggiorna le regole di sistema dell'agente documentando la disponibilità di `word_generate_image` e la dipendenza dalla chiave API. |
| `web/taskpane.js` | Mostra l'indicatore di stato di Gemini Imagen 3 nelle impostazioni (Attivo / Inattivo). |
| `tests/test_imagen_client.py` (Nuovo) | Test unitari con mock per verificare: assenza chiave API, chiamata API con successo, gestione errori di rete e inserimento nel documento. |

---

## 3. Dettaglio delle Fasi di Implementazione (TDD)

### Task 1: Client Imagen 3 (`daemon/office/imagen_client.py`) e Test Unitari
- **Red**: Creare `tests/test_imagen_client.py` con test per:
  1. `generate_imagen3_image` fallisce in modo pulito se `api_key` non è presente.
  2. `generate_imagen3_image` invia la richiesta REST corretta e salva l'immagine PNG decodificando `bytesBase64Encoded`.
  3. Gestione di errori HTTP (400, 403, timeout).
- **Green**: Implementare `daemon/office/imagen_client.py` usando `urllib.request` standard (zero dipendenze esterne, compatibile al 100% con PyInstaller standalone).
- **Verify**: `pytest tests/test_imagen_client.py` verde.

### Task 2: Strumento COM `word_generate_image` in `word_com_tools.py` e `mcp_server.py`
- Registrare `word_generate_image(doc, prompt, caption, location, aspect_ratio, width_pt)`:
  - Legge `auth_bridge.get_api_key()`.
  - Se assente, restituisce: `{"ok": False, "error": "La generazione di immagini AI con Imagen 3 richiede una chiave API Gemini. Aggiungi la tua chiave API nelle impostazioni del pannello laterale per attivare questa funzionalità."}`.
  - Se presente, genera l'immagine con `generate_imagen3_image` e richiama `word_insert_image(doc, file_path, location, caption=caption, width_pt=width_pt)`.
- Registrare lo schema in `EXTRA_TOOLS` e nel router `mcp_server.py`.
- **Verify**: Test unitario in `tests/test_imagen_client.py`.

### Task 3: Aggiornamento Prompt di Sistema e Schema Agente
- In `daemon/agent_session.py` e `daemon/antigravity_backend.py`: aggiungere la regola per `word_generate_image`.
- In `daemon/word_tools.py`: registrare lo schema JSON per `word_generate_image`.

### Task 4: UI Indicatore di Stato in `web/taskpane.js`
- All'apertura della finestra delle impostazioni o al salvataggio della chiave API:
  - Se `has_key === true` -> Mostra badge verde `Gemini Imagen 3: Attivo`.
  - Se `has_key === false` -> Mostra badge neutro `Gemini Imagen 3: Inattivo (Inserisci chiave API)`.

### Task 5: Verifica Completa e Commit
- Esecuzione di tutta la suite di test: `pytest tests/test_imagen_client.py` e test di non regressione.
- Commit in italiano secondo le convenzioni di progetto:
  `feat(imagen): generazione immagini AI con Gemini Imagen 3 attivabile con chiave API`
- Rebuild con `scripts/build.ps1` per aggiornare l'eseguibile standalone e l'installer.
