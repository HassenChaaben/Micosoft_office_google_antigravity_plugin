/**
 * Gemini Antigravity Office — Autonomous Word Agent Executor
 * Runs inside Microsoft Word Taskpane (WebView2).
 * Translates AI MCP tool invocations into native Word Office.js actions:
 * - Native tables with Word styles (Grid Table, Accent)
 * - Built-in Word paragraph and list styles (Heading 1, 2, List Bullet)
 * - Inline formatting, search/replace, comments, and selection focus
 * - User approval dialogs for destructive edits
 */

(function () {
  /**
   * Resolves Word built-in styles across languages (French, Italian, English)
   * to avoid InvalidArgument errors in localized Word environments.
   */
  function safeApplyParagraphStyle(para, styleName) {
    if (!styleName || styleName === 'Normal') return;
    const s = String(styleName).toLowerCase().replace(/[\s\-_]/g, '');
    let enumStyle = null;
    if (typeof Word !== 'undefined' && Word.Style) {
      if (s.includes('heading1') || s.includes('titre1') || s.includes('titolo1')) enumStyle = Word.Style.heading1;
      else if (s.includes('heading2') || s.includes('titre2') || s.includes('titolo2')) enumStyle = Word.Style.heading2;
      else if (s.includes('heading3') || s.includes('titre3') || s.includes('titolo3')) enumStyle = Word.Style.heading3;
      else if (s.includes('subtitle') || s.includes('soustitre') || s.includes('sottotitolo')) enumStyle = Word.Style.subtitle;
      else if (s.includes('title') || s.includes('titre') || s.includes('titolo')) enumStyle = Word.Style.title;
      else if (s.includes('quote') || s.includes('citation')) enumStyle = Word.Style.quote;
    }
    if (enumStyle) {
      try { para.styleBuiltIn = enumStyle; return; } catch (_) {}
    }
    try { para.style = styleName; } catch (_) {}
  }

  class WordAgent {
    constructor() {
      this.onApprovalRequired = null; // function(callId, tool, args): Promise<boolean>
      this.executionLogs = []; // Trace history of all tool executions
    }

    /**
     * Executes a tool invocation requested by the AI agent via the MCP server.
     * @param {string} toolName
     * @param {object} args
     * @param {string} callId
     * @returns {Promise<object>} Result object { ok: boolean, ... }
     */
    /**
     * Perception: read the whole document as OOXML (one call, fast), plus selection,
     * headers/footers and identity, so the daemon can analyse it before the model acts.
     */
    async captureDocumentContext() {
      const ctx = { document_ooxml: '', selection: '', headers_footers: [], doc_url: '', doc_title: '' };
      try {
        ctx.doc_url = (Office.context.document && Office.context.document.url) || '';
        ctx.doc_title = ctx.doc_url ? ctx.doc_url.split(/[\\/]/).pop() : '';
        await Word.run(async (context) => {
          const ooxml = context.document.body.getOoxml();
          const sel = context.document.getSelection();
          sel.load('text');
          const sections = context.document.sections;
          sections.load('items');
          await context.sync();
          ctx.document_ooxml = ooxml.value;
          ctx.selection = (sel.text || '').trim().slice(0, 2000);

          const parts = [];
          sections.items.slice(0, 5).forEach((section, i) => {
            ['Primary', 'FirstPage'].forEach((kind) => {
              const h = section.getHeader(kind); h.load('text');
              const f = section.getFooter(kind); f.load('text');
              parts.push({ i, kind, h, f });
            });
          });
          await context.sync();
          parts.forEach(({ i, kind, h, f }) => {
            if (h.text && h.text.trim()) ctx.headers_footers.push(`section ${i + 1} header (${kind}): ${h.text.trim().slice(0, 200)}`);
            if (f.text && f.text.trim()) ctx.headers_footers.push(`section ${i + 1} footer (${kind}): ${f.text.trim().slice(0, 200)}`);
          });
        });
      } catch (e) {
        console.warn('[WordAgent] captureDocumentContext failed:', e);
      }
      return ctx;
    }

    /**
     * Write resolved form cells in place. Replacing the text of the cell's first paragraph keeps
     * the cell's own run/paragraph formatting (fonts, bold, size, shading stay as in the template).
     * One batched sync; if Word rejects the batch, retry cell by cell to report the bad ones.
     */
    async fillCells(cells, contentControls) {
      const L = Word.InsertLocation;
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const failed = [];
        const ops = [];
        for (const c of cells) {
          const tbl = tables.items[c.table];
          const label = `T${c.table}r${c.row}c${c.cell}`;
          if (!tbl) { failed.push(`${label}: table not found`); continue; }
          ops.push({ label, run: () => tbl.getCell(c.row, c.cell).body.paragraphs.getFirst().insertText(String(c.text), L.replace) });
        }
        for (const [tag, value] of Object.entries(contentControls)) {
          ops.push({
            label: `content control '${tag}'`,
            run: () => {
              const cc = context.document.body.contentControls.getByTag(tag).getFirstOrNullObject();
              if (cc.isNullObject) throw new Error('not found');
              cc.insertText(String(value), L.replace);
            },
          });
        }

        let written = 0;
        try {
          ops.forEach((op) => op.run());
          await context.sync();
          written = ops.length;
        } catch (batchError) {
          for (const op of ops) {
            try {
              op.run();
              await context.sync();
              written++;
            } catch (e) {
              failed.push(`${op.label}: ${e.message || e}`);
            }
          }
        }
        return { ok: written > 0 || ops.length === 0, written, failed };
      });
    }

    /**
     * Insert a daemon-built OOXML package in ONE native operation.
     * location: cursor | start | end | replace_selection | after:p<N> | after:<heading text>
     *           | after:ga:<id> | replace:ga:<id>
     */
    async insertOoxmlAt(ooxml, location, blockId) {
      if (!ooxml) return { ok: false, error: "Missing 'ooxml'." };
      const L = Word.InsertLocation;
      return await Word.run(async (context) => {
        const body = context.document.body;
        let target = null;
        let where = null;
        const loc = String(location || 'cursor').trim();
        const m = loc.match(/^(after|replace):(.*)$/i);

        if (loc === 'start' || loc === 'end') {
          target = body; where = loc === 'start' ? L.start : L.end;
        } else if (loc === 'cursor' || loc === 'replace_selection') {
          const sel = context.document.getSelection();
          sel.load('text');
          await context.sync();
          target = sel;
          where = loc === 'replace_selection' || !(sel.text || '').trim() ? L.replace : L.after;
        } else if (m && m[2].startsWith('ga:')) {
          const cc = body.contentControls.getByTag(m[2]).getFirstOrNullObject();
          await context.sync();
          if (cc.isNullObject) return { ok: false, error: `Anchor ${m[2]} not found (it may have been deleted).` };
          target = cc.getRange('Whole');
          where = m[1].toLowerCase() === 'replace' ? L.replace : L.after;
        } else if (m && m[1].toLowerCase() === 'after') {
          body.paragraphs.load('items/text');
          await context.sync();
          const items = body.paragraphs.items;
          const ref = m[2].trim();
          let para = null;
          const pm = ref.match(/^p(\d+)$/i);
          if (pm) para = items[parseInt(pm[1], 10)] || null;
          else {
            const needle = ref.toLowerCase();
            para = items.find((p) => (p.text || '').trim().toLowerCase().startsWith(needle)) || null;
          }
          if (!para) return { ok: false, error: `Anchor '${ref}' not found in the document.` };
          target = para.getRange(); where = L.after;
        } else {
          return { ok: false, error: `Unsupported location '${loc}'.` };
        }

        target.insertOoxml(ooxml, where);
        await context.sync();
        return { ok: true, inserted_at: loc, anchor: blockId ? `ga:${blockId}` : null };
      });
    }

    async executeTool(toolName, args, callId) {
      console.log(`[WordAgent] Executing tool: ${toolName}`, args);
      const startTime = Date.now();
      let outcome = { ok: false };
      try {
        if (typeof Word === 'undefined' || !Word.run) {
          outcome = { ok: false, error: 'Office.js Word API is not available in this environment.' };
        } else {
          // Destructive tools that delete whole tables or document content require user approval
          const destructiveTools = [
            'word_delete_paragraph',
            'word_delete_table',
            'word_clear_document'
          ];
          if (destructiveTools.includes(toolName) && typeof this.onApprovalRequired === 'function') {
            const approved = await this.onApprovalRequired(callId, toolName, args);
            if (!approved) {
              outcome = { ok: false, error: 'User rejected the destructive document edit.' };
              this.logExecution(callId, toolName, args, outcome, Date.now() - startTime);
              return outcome;
            }
          }

          switch (toolName) {
            case '__capture_document':
              outcome = { ok: true, ...(await this.captureDocumentContext()) };
              break;
            case 'word_fill_cells':
              outcome = await this.fillCells(args.cells || [], args.content_controls || {});
              break;
            case 'word_insert_ooxml':
              outcome = await this.insertOoxmlAt(args.ooxml, args.location || 'cursor', args.block_id);
              break;
            case 'word_get_document_outline':
              outcome = await this.getDocumentOutline();
              break;
            case 'word_read_paragraphs':
              outcome = await this.readParagraphs(args.start_index, args.end_index);
              break;
            case 'word_search':
              outcome = await this.search(args.search_text, args.match_case);
              break;
            case 'word_insert_paragraph':
              outcome = await this.insertParagraph(args.text, args.style, args.location);
              break;
            case 'word_insert_list':
              outcome = await this.insertList(args.items, args.list_type, args.location);
              break;
            case 'word_insert_table':
              outcome = await this.insertTable(args.headers, args.rows, args.table_style, args.location, args.columns);
              break;
            case 'word_insert_image':
              outcome = await this.insertImage(
                args.source || args.file_path || args.url,
                args.caption,
                args.location,
                args.width_px,
                args.height_px,
                args
              );
              break;
            case 'word_insert_diagram':
              outcome = await this.insertImage(
                args.kind || args.type || 'diagram',
                args.caption || args.title || 'Diagramma di processo',
                args.location || 'end',
                args.width_px || 600,
                args.height_px || 280,
                args
              );
              break;
            case 'word_insert_chart':
              outcome = await this.insertImage(
                'chart',
                args.caption || args.title || 'Grafico di confronto',
                args.location || 'end',
                args.width_px || 600,
                args.height_px || 280,
                args
              );
              break;
            case 'word_format_table':
              outcome = await this.formatTable(args.table_index, args.style, args.header_bold, args.autofit);
              break;
            case 'word_set_table_cell':
              outcome = await this.setTableCell(args.table_index, args.row_index, args.col_index, args.text, args.bold);
              break;
            case 'word_add_table_row':
              outcome = await this.addTableRow(args.table_index, args.values);
              break;
            case 'word_delete_table':
              outcome = await this.deleteTable(args.table_index);
              break;
            case 'word_delete_column':
              outcome = await this.deleteColumn(args.table_index, args.col_index !== undefined ? args.col_index : args.column_index);
              break;
            case 'word_delete_row':
              outcome = await this.deleteRow(args.table_index, args.row_index);
              break;
            case 'word_read_table':
              outcome = await this.readTable(args.table_index);
              break;
            case 'word_clear_document':
              outcome = await this.clearDocument();
              break;
            case 'word_format_text':
              outcome = await this.formatText(args);
              break;
            case 'word_apply_style':
              outcome = await this.applyStyle(args.paragraph_index, args.style);
              break;
            case 'word_add_comment':
              outcome = await this.addComment(args.paragraph_index, args.comment, args.search_text);
              break;
            case 'word_insert_page_break':
              outcome = await this.insertPageBreak(args.location);
              break;
            case 'word_select':
              outcome = await this.selectParagraph(args.paragraph_index);
              break;
            case 'word_set_paragraph_text':
              outcome = await this.setParagraphText(args.paragraph_index, args.new_text !== undefined ? args.new_text : args.text);
              break;
            case 'word_replace_text':
              outcome = await this.replaceText(args.search_text, args.replacement_text, args.match_case);
              break;
            case 'word_delete_paragraph':
              outcome = await this.deleteParagraph(args.paragraph_index);
              break;
            case 'word_batch_update_table_cells':
              outcome = await this.batchUpdateTableCells(args.table_index, args.updates);
              break;
            case 'word_fill_form_fields':
              outcome = await this.fillFormFields(args.fields, args.table_index);
              break;
            default:
              outcome = { ok: false, error: `Tool not supported by Word agent: ${toolName}` };
          }
        }
      } catch (err) {
        console.error(`[WordAgent] Error executing ${toolName}:`, err);
        outcome = { ok: false, error: err.message || String(err) };
      }

      this.logExecution(callId, toolName, args, outcome, Date.now() - startTime);
      return outcome;
    }

    logExecution(callId, tool, args, result, durationMs) {
      const entry = {
        callId,
        tool,
        args,
        result,
        durationMs,
        time: new Date().toLocaleTimeString()
      };
      this.executionLogs.push(entry);
      if (typeof window !== 'undefined' && typeof window.onWordAgentLog === 'function') {
        window.onWordAgentLog(entry);
      }
    }

    // -------------------------------------------------------------------------
    // Word Office.js Tool Implementations
    // -------------------------------------------------------------------------

    async getDocumentOutline() {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        const tables = body.tables;
        paragraphs.load(['text', 'styleBuiltIn', 'style']);
        tables.load(['rowCount', 'values']);

        let selectedText = '';
        try {
          const sel = context.document.getSelection();
          sel.load('text');
          await context.sync();
          selectedText = (sel.text || '').trim();
        } catch (_) {}

        await context.sync();

        const paraList = [];
        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          const raw = (p.text || '').trim();
          if (raw.length > 0 || paragraphs.items.length === 1) {
            paraList.push({
              index: i,
              style: p.styleBuiltIn || p.style || 'Normal',
              preview: raw.length > 250 ? raw.substring(0, 247) + '...' : raw
            });
          }
        }

        const tableList = [];
        for (let t = 0; t < tables.items.length; t++) {
          const tbl = tables.items[t];
          const firstRow = tbl.values && tbl.values.length > 0 ? tbl.values[0] : [];
          tableList.push({
            table_index: t,
            rows: tbl.rowCount,
            cols: firstRow.length,
            headers: firstRow
          });
        }

        let picCount = 0;
        try {
          const pics = body.inlinePictures;
          pics.load('items');
          await context.sync();
          picCount = pics.items.length;
        } catch (_) {}

        return {
          ok: true,
          selected_text: selectedText,
          total_paragraphs: paragraphs.items.length,
          active_paragraphs: paraList.length,
          paragraphs: paraList.slice(0, 35),
          tables: tableList,
          pictures_count: picCount
        };
      });
    }

    async insertImage(source, caption, location = 'end', widthPx, heightPx, options = {}) {
      if (!source) {
        return { ok: false, error: "Missing required parameter 'source'. Provide an image URL, base64 string, or keyword ('diagram', 'chart', 'quantum circuit', etc.)." };
      }

      let base64Data = '';
      if (source.startsWith('data:image/')) {
        base64Data = source.split(',')[1] || '';
      } else if (source.startsWith('http://') || source.startsWith('https://')) {
        try {
          const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
          const timeoutId = controller ? setTimeout(() => controller.abort(), 3500) : null;
          const resp = await fetch(source, controller ? { signal: controller.signal } : {});
          if (timeoutId) clearTimeout(timeoutId);
          if (resp && resp.ok) {
            const blob = await resp.blob();
            if (blob && blob.type && blob.type.startsWith('image/')) {
              base64Data = await new Promise((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => {
                  const res = reader.result;
                  resolve(typeof res === 'string' ? res.split(',')[1] : '');
                };
                reader.onerror = () => resolve('');
                reader.readAsDataURL(blob);
              });
            }
          }
        } catch (_) {
          base64Data = '';
        }
        // Fallback: If external fetch was blocked by CORS, returned 403 HTML, or timed out,
        // dynamically generate a relevant graphic instead of returning broken images to Word!
        if (!base64Data) {
          base64Data = this.generateDynamicGraphic(source, caption, options);
        }
      } else {
        base64Data = this.generateDynamicGraphic(source, caption, options);
      }

      if (!base64Data) {
        return { ok: false, error: 'Could not resolve image data.' };
      }

      return await Word.run(async (context) => {
        const body = context.document.body;
        let targetRange = body;
        let insertLoc = Word.InsertLocation.end;

        if (location === 'start') {
          insertLoc = Word.InsertLocation.start;
        } else if (location && location.startsWith('after:')) {
          const locTarget = location.substring(6).trim();
          const idx = parseInt(locTarget, 10);
          body.paragraphs.load('items');
          await context.sync();
          if (!isNaN(idx)) {
            if (body.paragraphs.items[idx]) {
              targetRange = body.paragraphs.items[idx];
              insertLoc = Word.InsertLocation.after;
            }
          } else if (typeof body.search === 'function') {
            const matches = body.search(locTarget, { matchCase: false });
            matches.load('items');
            await context.sync();
            if (matches.items && matches.items.length > 0) {
              targetRange = matches.items[0];
              insertLoc = Word.InsertLocation.after;
            }
          }
        } else if (location && location.startsWith('before:')) {
          const locTarget = location.substring(7).trim();
          const idx = parseInt(locTarget, 10);
          body.paragraphs.load('items');
          await context.sync();
          if (!isNaN(idx)) {
            if (body.paragraphs.items[idx]) {
              targetRange = body.paragraphs.items[idx];
              insertLoc = Word.InsertLocation.before;
            }
          } else if (typeof body.search === 'function') {
            const matches = body.search(locTarget, { matchCase: false });
            matches.load('items');
            await context.sync();
            if (matches.items && matches.items.length > 0) {
              targetRange = matches.items[0];
              insertLoc = Word.InsertLocation.before;
            }
          }
        }

        const pic = targetRange.insertInlinePictureFromBase64(base64Data, insertLoc);
        if (widthPx && typeof widthPx === 'number') pic.width = widthPx;
        if (heightPx && typeof heightPx === 'number') pic.height = heightPx;
        if (typeof pic.select === 'function') pic.select();
        await context.sync();

        if (caption) {
          let capPara = null;
          if (pic && typeof pic.insertParagraph === 'function') {
            try {
              capPara = pic.insertParagraph(caption, Word.InsertLocation.after);
            } catch (_) {}
          }
          if (!capPara && targetRange && typeof targetRange.insertParagraph === 'function' && targetRange !== body) {
            try {
              capPara = targetRange.insertParagraph(caption, Word.InsertLocation.after);
            } catch (_) {}
          }
          if (!capPara && body && typeof body.insertParagraph === 'function') {
            capPara = body.insertParagraph(caption, Word.InsertLocation.end);
          }
          if (capPara) {
            safeApplyParagraphStyle(capPara, 'Caption');
          }
          await context.sync();
        }

        return {
          ok: true,
          image_inserted: true,
          source_type: source.startsWith('http') ? 'dynamic_fallback' : (source.startsWith('data:') ? 'base64' : 'dynamic'),
          caption: caption || null
        };
      });
    }

    generateDynamicGraphic(source, caption, options = {}) {
      if (typeof document === 'undefined') {
        return 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
      }
      const canvas = document.createElement('canvas');
      const w = 720;
      const h = 340;
      const scale = 2; // High-DPI 2x Retina rendering for crisp diagrams in Word
      canvas.width = w * scale;
      canvas.height = h * scale;
      const ctx = canvas.getContext('2d');
      if (!ctx) {
        return 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
      }
      ctx.scale(scale, scale);

      const title = options.title || caption || '';
      const textToScan = `${source || ''} ${caption || ''} ${title} ${options.kind || ''} ${options.type || ''}`.toLowerCase();

      if (
        textToScan.includes('quantum') ||
        textToScan.includes('circuito') ||
        textToScan.includes('circuit') ||
        textToScan.includes('bell') ||
        textToScan.includes('qubit') ||
        textToScan.includes('hadamard') ||
        textToScan.includes('porta logica') ||
        textToScan.includes('algoritm')
      ) {
        this.drawQuantumCircuit(ctx, w, h, title, caption, options);
      } else if (
        textToScan.includes('process') ||
        textToScan.includes('workflow') ||
        textToScan.includes('flusso') ||
        textToScan.includes('pipeline') ||
        textToScan.includes('fase') ||
        textToScan.includes('fasi') ||
        textToScan.includes('step') ||
        textToScan.includes('passaggi')
      ) {
        this.drawProcessWorkflow(ctx, w, h, title, caption, options);
      } else if (
        textToScan.includes('chart') ||
        textToScan.includes('grafico') ||
        textToScan.includes('metric') ||
        textToScan.includes('confronto') ||
        textToScan.includes('performance') ||
        textToScan.includes('budget') ||
        textToScan.includes('tabella') ||
        textToScan.includes('istogramma') ||
        textToScan.includes('trend')
      ) {
        this.drawMetricsChart(ctx, w, h, title, caption, options);
      } else if (
        textToScan.includes('arch') ||
        textToScan.includes('system') ||
        textToScan.includes('rete') ||
        textToScan.includes('network') ||
        textToScan.includes('tier') ||
        textToScan.includes('layer') ||
        textToScan.includes('component')
      ) {
        this.drawSystemArchitecture(ctx, w, h, title, caption, options);
      } else {
        this.drawUniversalDiagram(ctx, w, h, title, caption, options);
      }

      return canvas.toDataURL('image/png').split(',')[1];
    }

    drawRoundedRect(ctx, x, y, width, height, radius) {
      ctx.beginPath();
      ctx.moveTo(x + radius, y);
      ctx.lineTo(x + width - radius, y);
      ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
      ctx.lineTo(x + width, y + height - radius);
      ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
      ctx.lineTo(x + radius, y + height);
      ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
      ctx.lineTo(x, y + radius);
      ctx.quadraticCurveTo(x, y, x + radius, y);
      ctx.closePath();
    }

    drawQuantumCircuit(ctx, w, h, title, caption, options) {
      const bgGrad = ctx.createLinearGradient(0, 0, w, h);
      bgGrad.addColorStop(0, '#0B1120');
      bgGrad.addColorStop(1, '#0F172A');
      ctx.fillStyle = bgGrad;
      this.drawRoundedRect(ctx, 0, 0, w, h, 12);
      ctx.fill();

      ctx.strokeStyle = '#1E293B';
      ctx.lineWidth = 2;
      this.drawRoundedRect(ctx, 1, 1, w - 2, h - 2, 12);
      ctx.stroke();

      ctx.fillStyle = 'rgba(6, 182, 212, 0.15)';
      this.drawRoundedRect(ctx, 30, 24, 210, 26, 6);
      ctx.fill();
      ctx.fillStyle = '#38BDF8';
      ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText('CIRCUITO LOGICO QUANTISTICO', 40, 41);

      ctx.fillStyle = '#F8FAFC';
      ctx.font = 'bold 18px -apple-system, Segoe UI, Roboto, sans-serif';
      const cleanTitle = (options.title || caption || 'Circuito per la Generazione dello Stato di Bell |Φ⁺⟩').replace(/^(Figura|Figure|Fig\.)\s*\d+[:\- ]*/i, '');
      ctx.fillText(cleanTitle.length > 55 ? cleanTitle.substring(0, 52) + '...' : cleanTitle, 30, 75);

      const wire0Y = 145;
      const wire1Y = 225;
      const wireStartX = 110;
      const wireEndX = 480;

      ctx.fillStyle = '#38BDF8';
      ctx.font = 'bold 14px "Courier New", monospace';
      ctx.fillText('|q₀⟩ = |0⟩', 30, wire0Y + 5);
      ctx.fillText('|q₁⟩ = |0⟩', 30, wire1Y + 5);

      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      ctx.moveTo(wireStartX, wire0Y);
      ctx.lineTo(wireEndX, wire0Y);
      ctx.moveTo(wireStartX, wire1Y);
      ctx.lineTo(wireEndX, wire1Y);
      ctx.stroke();

      ctx.strokeStyle = 'rgba(56, 189, 248, 0.3)';
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(wireStartX, wire0Y);
      ctx.lineTo(wireEndX, wire0Y);
      ctx.moveTo(wireStartX, wire1Y);
      ctx.lineTo(wireEndX, wire1Y);
      ctx.stroke();

      const hGateX = 180;
      const hGateW = 46;
      const hGateH = 46;
      const hGrad = ctx.createLinearGradient(hGateX, wire0Y - 23, hGateX + hGateW, wire0Y + 23);
      hGrad.addColorStop(0, '#0284C7');
      hGrad.addColorStop(1, '#0EA5E9');
      ctx.fillStyle = hGrad;
      this.drawRoundedRect(ctx, hGateX, wire0Y - 23, hGateW, hGateH, 6);
      ctx.fill();
      ctx.strokeStyle = '#38BDF8';
      ctx.lineWidth = 1.5;
      this.drawRoundedRect(ctx, hGateX, wire0Y - 23, hGateW, hGateH, 6);
      ctx.stroke();

      ctx.fillStyle = '#FFFFFF';
      ctx.font = 'bold 20px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('H', hGateX + hGateW / 2, wire0Y + 7);
      ctx.font = '10px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillStyle = '#94A3B8';
      ctx.fillText('Hadamard', hGateX + hGateW / 2, wire0Y + 36);

      const cnotX = 300;
      ctx.fillStyle = '#06B6D4';
      ctx.beginPath();
      ctx.arc(cnotX, wire0Y, 6.5, 0, Math.PI * 2);
      ctx.fill();

      ctx.strokeStyle = '#06B6D4';
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      ctx.moveTo(cnotX, wire0Y);
      ctx.lineTo(cnotX, wire1Y);
      ctx.stroke();

      ctx.strokeStyle = '#06B6D4';
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      ctx.arc(cnotX, wire1Y, 15, 0, Math.PI * 2);
      ctx.stroke();
      ctx.beginPath();
      ctx.moveTo(cnotX - 15, wire1Y);
      ctx.lineTo(cnotX + 15, wire1Y);
      ctx.moveTo(cnotX, wire1Y - 15);
      ctx.lineTo(cnotX, wire1Y + 15);
      ctx.stroke();

      ctx.font = '10px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillStyle = '#94A3B8';
      ctx.fillText('CNOT', cnotX, wire1Y + 32);

      const meterX = 410;
      [wire0Y, wire1Y].forEach((y) => {
        ctx.fillStyle = '#1E293B';
        this.drawRoundedRect(ctx, meterX, y - 18, 44, 36, 6);
        ctx.fill();
        ctx.strokeStyle = '#475569';
        ctx.lineWidth = 1.5;
        this.drawRoundedRect(ctx, meterX, y - 18, 44, 36, 6);
        ctx.stroke();

        ctx.fillStyle = '#CBD5E1';
        ctx.font = '12px "Courier New", monospace';
        ctx.fillText('◷ M', meterX + 22, y + 4);
      });

      const outBoxX = 505;
      const outBoxY = 115;
      const outBoxW = 185;
      const outBoxH = 150;
      ctx.fillStyle = 'rgba(15, 33, 55, 0.85)';
      this.drawRoundedRect(ctx, outBoxX, outBoxY, outBoxW, outBoxH, 10);
      ctx.fill();
      ctx.strokeStyle = '#0284C7';
      ctx.lineWidth = 1.5;
      this.drawRoundedRect(ctx, outBoxX, outBoxY, outBoxW, outBoxH, 10);
      ctx.stroke();

      ctx.textAlign = 'left';
      ctx.fillStyle = '#38BDF8';
      ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('STATO DI BELL RISULTANTE', outBoxX + 14, outBoxY + 24);

      ctx.fillStyle = '#FFFFFF';
      ctx.font = 'bold 14px "Courier New", monospace';
      ctx.fillText('|Φ⁺⟩ = (|00⟩ + |11⟩) / √2', outBoxX + 14, outBoxY + 54);

      ctx.fillStyle = '#94A3B8';
      ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('• Entanglement: 100%', outBoxX + 14, outBoxY + 84);
      ctx.fillText('• Base: {|00⟩, |01⟩, |10⟩, |11⟩}', outBoxX + 14, outBoxY + 106);
      ctx.fillText('• Misure correlate', outBoxX + 14, outBoxY + 128);

      ctx.fillStyle = '#1E293B';
      this.drawRoundedRect(ctx, 30, 290, w - 60, 32, 6);
      ctx.fill();

      ctx.fillStyle = '#94A3B8';
      ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('Qubit: 2   •   Profondità circuito: 2 cicli   •   Porte logiche: Hadamard + CNOT   •   Coerenza: Conservata', 45, 310);
    }

    drawProcessWorkflow(ctx, w, h, title, caption, options) {
      const bgGrad = ctx.createLinearGradient(0, 0, w, h);
      bgGrad.addColorStop(0, '#0F172A');
      bgGrad.addColorStop(1, '#1E293B');
      ctx.fillStyle = bgGrad;
      this.drawRoundedRect(ctx, 0, 0, w, h, 12);
      ctx.fill();

      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 2;
      this.drawRoundedRect(ctx, 1, 1, w - 2, h - 2, 12);
      ctx.stroke();

      ctx.fillStyle = 'rgba(16, 185, 129, 0.15)';
      this.drawRoundedRect(ctx, 30, 24, 180, 26, 6);
      ctx.fill();
      ctx.fillStyle = '#34D399';
      ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText('WORKFLOW & PROCESSO', 40, 41);

      ctx.fillStyle = '#F8FAFC';
      ctx.font = 'bold 18px -apple-system, Segoe UI, Roboto, sans-serif';
      const cleanTitle = (options.title || caption || 'Workflow Operativo & Fasi di Avanzamento').replace(/^(Figura|Figure|Fig\.)\s*\d+[:\- ]*/i, '');
      ctx.fillText(cleanTitle.length > 55 ? cleanTitle.substring(0, 52) + '...' : cleanTitle, 30, 75);

      const defaultSteps = [
        { num: '01', title: 'Pianificazione', desc: 'Requisiti & Scope' },
        { num: '02', title: 'Sviluppo', desc: 'Implementazione core' },
        { num: '03', title: 'Validazione', desc: 'Collaudo & Audit QA' },
        { num: '04', title: 'Rilascio', desc: 'Deploy & Monitoraggio' }
      ];

      const customSteps = Array.isArray(options.steps) && options.steps.length >= 2 ? options.steps.slice(0, 4).map((s, i) => {
        const text = typeof s === 'string' ? s : (s.title || s.name || `Fase ${i + 1}`);
        return { num: `0${i + 1}`, title: text, desc: (s.desc || s.subtitle || 'Stato operativo') };
      }) : defaultSteps;

      const numCards = customSteps.length;
      const totalWidth = w - 60;
      const cardGap = 16;
      const cardWidth = (totalWidth - (numCards - 1) * cardGap) / numCards;
      const cardHeight = 160;
      const cardY = 110;

      customSteps.forEach((step, i) => {
        const cardX = 30 + i * (cardWidth + cardGap);
        const cardGrad = ctx.createLinearGradient(cardX, cardY, cardX, cardY + cardHeight);
        cardGrad.addColorStop(0, '#1E293B');
        cardGrad.addColorStop(1, '#0F172A');
        ctx.fillStyle = cardGrad;
        this.drawRoundedRect(ctx, cardX, cardY, cardWidth, cardHeight, 10);
        ctx.fill();

        ctx.strokeStyle = i === 0 ? '#38BDF8' : '#334155';
        ctx.lineWidth = 1.5;
        this.drawRoundedRect(ctx, cardX, cardY, cardWidth, cardHeight, 10);
        ctx.stroke();

        ctx.fillStyle = i === 0 ? '#0284C7' : '#334155';
        this.drawRoundedRect(ctx, cardX + 14, cardY + 14, 32, 22, 6);
        ctx.fill();
        ctx.fillStyle = '#FFFFFF';
        ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(step.num, cardX + 30, cardY + 29);

        ctx.fillStyle = '#10B981';
        ctx.font = 'bold 12px sans-serif';
        ctx.fillText('✓', cardX + cardWidth - 22, cardY + 29);

        ctx.textAlign = 'left';
        ctx.fillStyle = '#F8FAFC';
        ctx.font = 'bold 13px -apple-system, Segoe UI, Roboto, sans-serif';
        const displayTitle = step.title.length > 18 ? step.title.substring(0, 16) + '..' : step.title;
        ctx.fillText(displayTitle, cardX + 14, cardY + 68);

        ctx.fillStyle = '#94A3B8';
        ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
        const displayDesc = step.desc.length > 22 ? step.desc.substring(0, 20) + '..' : step.desc;
        ctx.fillText(displayDesc, cardX + 14, cardY + 92);

        ctx.fillStyle = '#334155';
        this.drawRoundedRect(ctx, cardX + 14, cardY + 125, cardWidth - 28, 6, 3);
        ctx.fill();
        ctx.fillStyle = i === 0 ? '#38BDF8' : '#10B981';
        const progressPct = i === 0 ? 0.9 : 1.0;
        this.drawRoundedRect(ctx, cardX + 14, cardY + 125, (cardWidth - 28) * progressPct, 6, 3);
        ctx.fill();

        if (i < numCards - 1) {
          ctx.fillStyle = '#64748B';
          ctx.font = 'bold 16px sans-serif';
          ctx.textAlign = 'center';
          ctx.fillText('➔', cardX + cardWidth + cardGap / 2, cardY + cardHeight / 2);
        }
      });

      ctx.fillStyle = '#1E293B';
      this.drawRoundedRect(ctx, 30, 290, w - 60, 32, 6);
      ctx.fill();
      ctx.textAlign = 'left';
      ctx.fillStyle = '#94A3B8';
      ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('Avanzamento Globale: 4 Fasi Configurate   •   Stato: In Corso   •   Milestone Attiva: Fase 1', 45, 310);
    }

    drawMetricsChart(ctx, w, h, title, caption, options) {
      const bgGrad = ctx.createLinearGradient(0, 0, w, h);
      bgGrad.addColorStop(0, '#0A192F');
      bgGrad.addColorStop(1, '#112240');
      ctx.fillStyle = bgGrad;
      this.drawRoundedRect(ctx, 0, 0, w, h, 12);
      ctx.fill();

      ctx.strokeStyle = '#233554';
      ctx.lineWidth = 2;
      this.drawRoundedRect(ctx, 1, 1, w - 2, h - 2, 12);
      ctx.stroke();

      ctx.fillStyle = 'rgba(99, 102, 241, 0.15)';
      this.drawRoundedRect(ctx, 30, 24, 210, 26, 6);
      ctx.fill();
      ctx.fillStyle = '#818CF8';
      ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText('ANALISI COMPARATIVA & METRICHE', 40, 41);

      ctx.fillStyle = '#F8FAFC';
      ctx.font = 'bold 18px -apple-system, Segoe UI, Roboto, sans-serif';
      const cleanTitle = (options.title || caption || 'Distribuzione Valori e Indicatori di Performance').replace(/^(Figura|Figure|Fig\.)\s*\d+[:\- ]*/i, '');
      ctx.fillText(cleanTitle.length > 55 ? cleanTitle.substring(0, 52) + '...' : cleanTitle, 30, 75);

      ctx.fillStyle = 'rgba(16, 185, 129, 0.15)';
      this.drawRoundedRect(ctx, w - 175, 24, 145, 26, 6);
      ctx.fill();
      ctx.fillStyle = '#34D399';
      ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('▲ +32.4% Crescita', w - 160, 41);

      const chartBottom = 260;
      const chartTop = 110;
      const chartHeight = chartBottom - chartTop;
      const gridLevels = [0.25, 0.5, 0.75, 1.0];
      gridLevels.forEach((lvl) => {
        const y = chartBottom - chartHeight * lvl;
        ctx.strokeStyle = 'rgba(255, 255, 255, 0.07)';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(70, y);
        ctx.lineTo(w - 40, y);
        ctx.stroke();
      });

      const defaultBars = [
        { label: 'Q1', val: '€ 45k', height: 75 },
        { label: 'Q2', val: '€ 78k', height: 110 },
        { label: 'Q3', val: '€ 115k', height: 140 },
        { label: 'Q4', val: '€ 160k', height: 175 },
        { label: 'Q5 Target', val: '€ 210k', height: 195 }
      ];

      const numBars = defaultBars.length;
      const chartWidth = w - 140;
      const barWidth = 64;
      const barSpacing = (chartWidth - numBars * barWidth) / (numBars - 1);

      defaultBars.forEach((bar, i) => {
        const x = 70 + i * (barWidth + barSpacing);
        const y = chartBottom - bar.height;

        const barGrad = ctx.createLinearGradient(x, y, x, chartBottom);
        if (i === numBars - 1) {
          barGrad.addColorStop(0, '#38BDF8');
          barGrad.addColorStop(1, '#0284C7');
        } else {
          barGrad.addColorStop(0, '#6366F1');
          barGrad.addColorStop(1, '#4338CA');
        }
        ctx.fillStyle = barGrad;
        this.drawRoundedRect(ctx, x, y, barWidth, bar.height, 6);
        ctx.fill();

        ctx.fillStyle = '#FFFFFF';
        ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(bar.val, x + barWidth / 2, y - 8);

        ctx.fillStyle = '#94A3B8';
        ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
        ctx.fillText(bar.label, x + barWidth / 2, chartBottom + 20);
      });

      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(50, chartBottom);
      ctx.lineTo(w - 30, chartBottom);
      ctx.stroke();

      ctx.fillStyle = '#112240';
      this.drawRoundedRect(ctx, 30, 295, w - 60, 30, 6);
      ctx.fill();
      ctx.textAlign = 'left';
      ctx.fillStyle = '#94A3B8';
      ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('Fonte dati: Proiezioni di budget e avanzamento deliverable consolidate', 45, 314);
    }

    drawSystemArchitecture(ctx, w, h, title, caption, options) {
      const bgGrad = ctx.createLinearGradient(0, 0, w, h);
      bgGrad.addColorStop(0, '#0F172A');
      bgGrad.addColorStop(1, '#1E293B');
      ctx.fillStyle = bgGrad;
      this.drawRoundedRect(ctx, 0, 0, w, h, 12);
      ctx.fill();

      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 2;
      this.drawRoundedRect(ctx, 1, 1, w - 2, h - 2, 12);
      ctx.stroke();

      ctx.fillStyle = 'rgba(14, 165, 233, 0.15)';
      this.drawRoundedRect(ctx, 30, 24, 210, 26, 6);
      ctx.fill();
      ctx.fillStyle = '#38BDF8';
      ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText('ARCHITETTURA DI SISTEMA', 40, 41);

      ctx.fillStyle = '#F8FAFC';
      ctx.font = 'bold 18px -apple-system, Segoe UI, Roboto, sans-serif';
      const cleanTitle = (options.title || caption || 'Architettura e Componenti Modulari').replace(/^(Figura|Figure|Fig\.)\s*\d+[:\- ]*/i, '');
      ctx.fillText(cleanTitle.length > 55 ? cleanTitle.substring(0, 52) + '...' : cleanTitle, 30, 75);

      const layers = [
        { name: 'Layer 1: Presentazione & Client', desc: 'Interfacce Utente, Office Taskpane, Webhooks', color: '#0284C7' },
        { name: 'Layer 2: Core Orchestration Bus', desc: 'Gestione Eventi, State Machine, Protocolli MCP', color: '#6366F1' },
        { name: 'Layer 3: Persistenza & Servizi Dati', desc: 'Storage Relazionale, Cache In-Memory, Sync Gateway', color: '#10B981' }
      ];

      const layerY = [105, 165, 225];
      const layerH = 46;
      const layerW = w - 60;

      layers.forEach((layer, i) => {
        const y = layerY[i];
        ctx.fillStyle = '#1E293B';
        this.drawRoundedRect(ctx, 30, y, layerW, layerH, 8);
        ctx.fill();

        ctx.strokeStyle = layer.color;
        ctx.lineWidth = 1.5;
        this.drawRoundedRect(ctx, 30, y, layerW, layerH, 8);
        ctx.stroke();

        ctx.fillStyle = '#FFFFFF';
        ctx.font = 'bold 13px -apple-system, Segoe UI, Roboto, sans-serif';
        ctx.textAlign = 'left';
        ctx.fillText(layer.name, 45, y + 22);

        ctx.fillStyle = '#94A3B8';
        ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
        ctx.fillText(layer.desc, 45, y + 38);

        ctx.fillStyle = layer.color;
        this.drawRoundedRect(ctx, w - 120, y + 12, 75, 22, 4);
        ctx.fill();
        ctx.fillStyle = '#FFFFFF';
        ctx.font = 'bold 10px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('ATTIVO', w - 82, y + 26);
      });

      [151, 211].forEach((y) => {
        ctx.fillStyle = '#64748B';
        ctx.font = '13px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('⇅', w / 2, y + 7);
      });

      ctx.fillStyle = '#1E293B';
      this.drawRoundedRect(ctx, 30, 290, w - 60, 32, 6);
      ctx.fill();
      ctx.textAlign = 'left';
      ctx.fillStyle = '#94A3B8';
      ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('Protocollo: Zero Trust   •   Ridondanza: Attiva   •   Isolamento componenti: Garantito', 45, 310);
    }

    drawUniversalDiagram(ctx, w, h, title, caption, options) {
      const bgGrad = ctx.createLinearGradient(0, 0, w, h);
      bgGrad.addColorStop(0, '#0F172A');
      bgGrad.addColorStop(1, '#1E293B');
      ctx.fillStyle = bgGrad;
      this.drawRoundedRect(ctx, 0, 0, w, h, 12);
      ctx.fill();

      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 2;
      this.drawRoundedRect(ctx, 1, 1, w - 2, h - 2, 12);
      ctx.stroke();

      ctx.fillStyle = 'rgba(56, 189, 248, 0.15)';
      this.drawRoundedRect(ctx, 30, 24, 180, 26, 6);
      ctx.fill();
      ctx.fillStyle = '#38BDF8';
      ctx.font = 'bold 11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText('DIAGRAMMA ESPLICATIVO', 40, 41);

      ctx.fillStyle = '#F8FAFC';
      ctx.font = 'bold 18px -apple-system, Segoe UI, Roboto, sans-serif';
      const cleanTitle = (options.title || caption || 'Schema Funzionale & Struttura del Documento').replace(/^(Figura|Figure|Fig\.)\s*\d+[:\- ]*/i, '');
      ctx.fillText(cleanTitle.length > 55 ? cleanTitle.substring(0, 52) + '...' : cleanTitle, 30, 75);

      const cards = [
        { title: 'Requisiti & Struttura', icon: '📋', desc: 'Definizione chiara delle componenti di base e vincoli.' },
        { title: 'Elaborazione Dati', icon: '⚡', desc: 'Trasformazione logica e allineamento agli standard stabiliti.' },
        { title: 'Risultati & Sintesi', icon: '🎯', desc: 'Validazione metrica e integrazione formale nel documento.' }
      ];

      const cardW = (w - 60 - 32) / 3;
      const cardY = 110;
      const cardH = 160;

      cards.forEach((c, i) => {
        const x = 30 + i * (cardW + 16);
        ctx.fillStyle = '#1E293B';
        this.drawRoundedRect(ctx, x, cardY, cardW, cardH, 8);
        ctx.fill();

        ctx.strokeStyle = '#334155';
        ctx.lineWidth = 1.5;
        this.drawRoundedRect(ctx, x, cardY, cardW, cardH, 8);
        ctx.stroke();

        ctx.font = '24px sans-serif';
        ctx.textAlign = 'left';
        ctx.fillText(c.icon, x + 16, cardY + 38);

        ctx.fillStyle = '#F8FAFC';
        ctx.font = 'bold 13px -apple-system, Segoe UI, Roboto, sans-serif';
        ctx.fillText(c.title, x + 16, cardY + 75);

        ctx.fillStyle = '#94A3B8';
        ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
        const words = c.desc.split(' ');
        let line = '';
        let ly = cardY + 98;
        words.forEach((word) => {
          const testLine = line + word + ' ';
          if (testLine.length > 25) {
            ctx.fillText(line, x + 16, ly);
            line = word + ' ';
            ly += 16;
          } else {
            line = testLine;
          }
        });
        if (line) ctx.fillText(line, x + 16, ly);
      });

      ctx.fillStyle = '#1E293B';
      this.drawRoundedRect(ctx, 30, 290, w - 60, 32, 6);
      ctx.fill();
      ctx.textAlign = 'left';
      ctx.fillStyle = '#94A3B8';
      ctx.font = '11px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillText('Verifica di conformità completata   •   Parametri di sistema verificati', 45, 310);
    }

    generatePresetGraphic(type) {
      return this.generateDynamicGraphic(type, '', {});
    }

    async readParagraphs(startIndex, endIndex) {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load(['text', 'styleBuiltIn']);
        await context.sync();

        const end = typeof endIndex === 'number' ? Math.min(endIndex, paragraphs.items.length - 1) : startIndex;
        const items = [];
        for (let i = startIndex; i <= end; i++) {
          if (paragraphs.items[i]) {
            items.push({
              index: i,
              style: paragraphs.items[i].styleBuiltIn || 'Normal',
              text: paragraphs.items[i].text
            });
          }
        }
        return { ok: true, paragraphs: items };
      });
    }

    async search(searchText, matchCase = false) {
      return await Word.run(async (context) => {
        const searchResults = context.document.body.search(searchText, { matchCase: !!matchCase });
        searchResults.load(['text']);
        await context.sync();

        return {
          ok: true,
          count: searchResults.items.length,
          matches: searchResults.items.slice(0, 10).map((r) => r.text)
        };
      });
    }

    async insertParagraph(text, style = 'Normal', location = 'end') {
      return await Word.run(async (context) => {
        const body = context.document.body;
        let insertLoc = Word.InsertLocation.end;
        let targetRange = body;

        if (location === 'start') {
          insertLoc = Word.InsertLocation.start;
        } else if (location && location.startsWith('after:')) {
          const idx = parseInt(location.split(':')[1], 10);
          body.paragraphs.load('items');
          await context.sync();
          if (body.paragraphs.items[idx]) {
            targetRange = body.paragraphs.items[idx];
            insertLoc = Word.InsertLocation.after;
          }
        } else if (location && location.startsWith('before:')) {
          const idx = parseInt(location.split(':')[1], 10);
          body.paragraphs.load('items');
          await context.sync();
          if (body.paragraphs.items[idx]) {
            targetRange = body.paragraphs.items[idx];
            insertLoc = Word.InsertLocation.before;
          }
        }

        // 1. Insert text and commit immediately
        const newPara = targetRange.insertParagraph(text || '', insertLoc);
        newPara.select();
        await context.sync();

        // 2. Apply style in safe secondary step so it never fails insertion
        if (style && style !== 'Normal') {
          try {
            safeApplyParagraphStyle(newPara, style);
            await context.sync();
          } catch (_) {}
        }

        return { ok: true, inserted_text: text, style: style };
      });
    }

    async insertList(items, listType = 'bullet', location = 'end') {
      return await Word.run(async (context) => {
        const body = context.document.body;
        const normItems = Array.isArray(items) ? items : [String(items || '')];
        for (const itemText of normItems) {
          const p = body.insertParagraph(itemText, Word.InsertLocation.end);
          try {
            p.styleBuiltIn = listType === 'number' ? 'List Number' : 'List Bullet';
          } catch (_) {
            try { p.style = listType === 'number' ? 'List Number' : 'List Bullet'; } catch (_) {}
          }
        }
        await context.sync();
        return { ok: true, items_count: normItems.length, list_type: listType };
      });
    }

    async insertTable(headers, rows, tableStyle = 'Grid Table 4 - Accent 1', location = 'end', columns = null) {
      let finalHeaders = Array.isArray(headers) && headers.length > 0 ? headers.map(String) : null;
      if (!finalHeaders) {
        let colCount = typeof columns === 'number' && columns > 0 ? columns : 0;
        if (!colCount && Array.isArray(rows) && rows.length > 0 && Array.isArray(rows[0])) {
          colCount = rows[0].length;
        }
        if (!colCount) {
          throw new Error("Missing required parameter 'headers'. Example: headers: ['product_name', 'quantity', 'price']");
        }
        finalHeaders = Array.from({ length: colCount }, (_, i) => `Colonna ${i + 1}`);
      }

      let normRows;
      if (typeof rows === 'number' && rows > 0) {
        const dataCount = rows > 1 ? rows - 1 : rows;
        normRows = Array.from({ length: dataCount }, (_, r) =>
          finalHeaders.map((_, c) => `Elemento ${r + 1}`)
        );
      } else if (Array.isArray(rows) && rows.length > 0) {
        normRows = rows;
      } else {
        normRows = [
          finalHeaders.map((h, i) => `Dato 1`),
          finalHeaders.map((h, i) => `Dato 2`)
        ];
      }

      return await Word.run(async (context) => {
        const body = context.document.body;
        const totalRows = normRows.length + 1;
        const totalCols = finalHeaders.length;

        // Build 2D string matrix
        const matrix = [finalHeaders];
        for (const r of normRows) {
          const rowData = [];
          for (let c = 0; c < totalCols; c++) {
            rowData.push(r && r[c] !== undefined && r[c] !== null ? String(r[c]) : '');
          }
          matrix.push(rowData);
        }

        let insertLoc = Word.InsertLocation.end;
        if (location === 'start') insertLoc = Word.InsertLocation.start;

        // 1. Insert native Word Table and commit immediately
        const table = body.insertTable(totalRows, totalCols, insertLoc, matrix);
        await context.sync(); // Table is 100% created!

        // 2. Safe style application
        if (tableStyle) {
          try {
            table.style = tableStyle;
            await context.sync();
          } catch (_) {
            try {
              table.styleBuiltIn = 'Grid Table 1 Light';
              await context.sync();
            } catch (_) {
              try {
                table.style = 'Table Grid';
                await context.sync();
              } catch (_) {}
            }
          }
        }

        // 3. Bold the header row safely via Range
        try {
          const headerRow = table.rows.getFirst();
          const headerRange = headerRow.getRange();
          headerRange.font.bold = true;
          await context.sync();
        } catch (_) {}

        // 4. Select table so cursor moves there
        try {
          table.select();
          await context.sync();
        } catch (_) {}

        return {
          ok: true,
          table_created: true,
          rows: totalRows,
          columns: totalCols,
          headers: finalHeaders
        };
      });
    }

    async formatTable(tableIndex, style, headerBold = true, autofit = true) {
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const table = tables.items[tableIndex];
        if (!table) {
          return { ok: false, error: `Table at index ${tableIndex} not found (total tables: ${tables.items.length}).` };
        }

        if (style) {
          try { table.style = style; await context.sync(); } catch (_) {
            try { table.styleBuiltIn = style; await context.sync(); } catch (_) {}
          }
        }

        if (headerBold) {
          try {
            const headerRow = table.rows.getFirst();
            const headerRange = headerRow.getRange();
            headerRange.font.bold = true;
            await context.sync();
          } catch (_) {}
        }

        try {
          table.select();
          await context.sync();
        } catch (_) {}

        return { ok: true, formatted: true, table_index: tableIndex };
      });
    }

    async deleteTable(tableIndex) {
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const idx = typeof tableIndex === 'number' ? tableIndex : parseInt(tableIndex, 10);
        if (isNaN(idx) || idx < 0 || idx >= tables.items.length) {
          return { ok: false, error: `Table index ${tableIndex} not found. Total tables: ${tables.items.length}.` };
        }

        const table = tables.items[idx];
        table.delete();
        await context.sync();
        return { ok: true, deleted_table_index: idx };
      });
    }

    async deleteColumn(tableIndex = 0, colIndex = 0) {
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const tIdx = typeof tableIndex === 'number' ? tableIndex : parseInt(tableIndex, 10);
        if (isNaN(tIdx) || tIdx < 0 || tIdx >= tables.items.length) {
          return { ok: false, error: `Table index ${tableIndex} not found. Total tables: ${tables.items.length}.` };
        }

        const table = tables.items[tIdx];
        const firstRow = table.rows.getFirst();
        firstRow.load('cellCount');
        await context.sync();

        const cIdx = typeof colIndex === 'number' ? colIndex : parseInt(colIndex, 10);
        if (isNaN(cIdx) || cIdx < 0 || cIdx >= firstRow.cellCount) {
          return { ok: false, error: `Column index ${colIndex} out of bounds (table has ${firstRow.cellCount} columns).` };
        }

        if (firstRow.cellCount <= 1) {
          return { ok: false, error: 'Cannot delete the last remaining column of a table. Use word_delete_table instead.' };
        }

        table.deleteColumns(cIdx, 1);
        await context.sync();
        return { ok: true, deleted_column_index: cIdx, table_index: tIdx, remaining_columns: firstRow.cellCount - 1 };
      });
    }

    async deleteRow(tableIndex = 0, rowIndex = 0) {
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const tIdx = typeof tableIndex === 'number' ? tableIndex : parseInt(tableIndex, 10);
        if (isNaN(tIdx) || tIdx < 0 || tIdx >= tables.items.length) {
          return { ok: false, error: `Table index ${tableIndex} not found. Total tables: ${tables.items.length}.` };
        }

        const table = tables.items[tIdx];
        table.load('rowCount');
        await context.sync();

        const rIdx = typeof rowIndex === 'number' ? rowIndex : parseInt(rowIndex, 10);
        if (isNaN(rIdx) || rIdx < 0 || rIdx >= table.rowCount) {
          return { ok: false, error: `Row index ${rowIndex} out of bounds (table has ${table.rowCount} rows).` };
        }

        if (table.rowCount <= 1) {
          return { ok: false, error: 'Cannot delete the last remaining row of a table. Use word_delete_table instead.' };
        }

        table.deleteRows(rIdx, 1);
        await context.sync();
        return { ok: true, deleted_row_index: rIdx, table_index: tIdx, remaining_rows: table.rowCount - 1 };
      });
    }

    async readTable(tableIndex) {
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const idx = typeof tableIndex === 'number' ? tableIndex : parseInt(tableIndex, 10);
        if (isNaN(idx) || idx < 0 || idx >= tables.items.length) {
          return { ok: false, error: `Table index ${tableIndex} not found. Total tables: ${tables.items.length}.` };
        }

        const table = tables.items[idx];
        table.load(['rowCount', 'values']);
        await context.sync();

        return {
          ok: true,
          table_index: idx,
          rows: table.rowCount,
          values: table.values
        };
      });
    }

    async clearDocument() {
      return await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();
        await context.sync();
        return { ok: true, cleared: true };
      });
    }

    async setTableCell(tableIndex, rowIndex, colIndex, text, bold) {
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const table = tables.items[tableIndex];
        if (!table) return { ok: false, error: `Table ${tableIndex} not found.` };

        const cell = table.getCell(rowIndex, colIndex);
        cell.value = text;
        if (typeof bold === 'boolean') {
          cell.font.bold = bold;
        }
        cell.select();
        await context.sync();
        return { ok: true, updated_cell: [rowIndex, colIndex] };
      });
    }

    async addTableRow(tableIndex, values) {
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const table = tables.items[tableIndex];
        if (!table) return { ok: false, error: `Table ${tableIndex} not found.` };

        const row = table.addRows(Word.InsertLocation.end, 1, [values.map(String)]);
        row.select();
        await context.sync();
        return { ok: true, added_row: values };
      });
    }

    async batchUpdateTableCells(tableIndex = 0, updates = []) {
      if (!Array.isArray(updates) || updates.length === 0) {
        return { ok: false, error: "Missing or empty 'updates' array." };
      }
      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load('items');
        await context.sync();

        const idx = typeof tableIndex === 'number' ? tableIndex : (parseInt(tableIndex, 10) || 0);
        const table = tables.items[idx];
        if (!table) return { ok: false, error: `Table index ${idx} not found. Total tables: ${tables.items.length}.` };

        let successCount = 0;
        const details = [];
        for (const item of updates) {
          try {
            const r = parseInt(item.row, 10);
            const c = parseInt(item.col, 10);
            const cell = table.getCell(r, c);
            cell.value = String(item.text !== undefined ? item.text : '');
            if (typeof item.bold === 'boolean') {
              cell.font.bold = item.bold;
            }
            successCount++;
            details.push({ row: r, col: c, ok: true });
          } catch (cellErr) {
            details.push({ row: item.row, col: item.col, ok: false, error: cellErr.message || String(cellErr) });
          }
        }
        await context.sync();
        return { ok: true, table_index: idx, updated_count: successCount, details };
      });
    }

    async fillFormFields(fields, targetTableIndex) {
      if (!fields || typeof fields !== 'object') {
        return { ok: false, error: "Missing or invalid 'fields' parameter. Provide an object mapping labels to values." };
      }

      return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load(['items', 'rowCount']);
        await context.sync();

        const matchedFields = [];
        const unmatchedFields = [];
        const fieldEntries = Object.entries(fields);

        const targetTables = (typeof targetTableIndex === 'number' && tables.items[targetTableIndex])
          ? [tables.items[targetTableIndex]]
          : tables.items;

        for (const tbl of targetTables) {
          tbl.load(['rowCount', 'values']);
        }
        await context.sync();

        for (const [rawLabel, rawValue] of fieldEntries) {
          const label = String(rawLabel).trim().toLowerCase();
          const valStr = String(rawValue !== undefined && rawValue !== null ? rawValue : '');
          let fieldFound = false;

          // 1. Search in tables (exact match pass first, then boundary pass)
          for (let pass = 0; pass < 2 && !fieldFound; pass++) {
            for (let t = 0; t < targetTables.length && !fieldFound; t++) {
              const tbl = targetTables[t];
              const matrix = tbl.values || [];

              for (let r = 0; r < matrix.length && !fieldFound; r++) {
                const row = matrix[r] || [];
                for (let c = 0; c < row.length && !fieldFound; c++) {
                  const cellText = String(row[c] || '').trim().toLowerCase();
                  if (!cellText) continue;

                  const isMatch = pass === 0
                    ? (cellText === label)
                    : (cellText === label || cellText.split(/[\s:_]+/).includes(label));

                  if (isMatch) {
                    // Horizontal form: find the next cell to the right that doesn't equal the label
                    let targetCol = -1;
                    for (let right = c + 1; right < row.length; right++) {
                      const nextText = String(row[right] || '').trim().toLowerCase();
                      if (nextText !== cellText) {
                        targetCol = right;
                        break;
                      }
                    }

                    if (targetCol !== -1) {
                      try {
                        const targetCell = tbl.getCell(r, targetCol);
                        targetCell.value = valStr;
                        if (tbl.values && tbl.values[r]) {
                          tbl.values[r][targetCol] = valStr;
                        }
                        fieldFound = true;
                        matchedFields.push({ label: rawLabel, location: `Table ${t} [row ${r}, col ${targetCol}]`, value: valStr });
                        break;
                      } catch (_) {}
                    }

                    // Vertical / stacked form fallback: check cell directly below
                    if (!fieldFound && r + 1 < matrix.length) {
                      try {
                        const belowCell = tbl.getCell(r + 1, c);
                        belowCell.value = valStr;
                        if (tbl.values && tbl.values[r + 1]) {
                          tbl.values[r + 1][c] = valStr;
                        }
                        fieldFound = true;
                        matchedFields.push({ label: rawLabel, location: `Table ${t} [row ${r + 1}, col ${c}]`, value: valStr });
                        break;
                      } catch (_) {}
                    }
                  }
                }
              }
            }
          }

          // 2. If not found in tables, check for document body placeholders like [LABEL], <LABEL>, LABEL:
          if (!fieldFound) {
            try {
              const placeholderVariants = [`[${rawLabel}]`, `<${rawLabel}>`, `${rawLabel}:`];
              for (const ph of placeholderVariants) {
                const res = context.document.body.search(ph, { matchCase: false });
                res.load('items');
                await context.sync();
                if (res.items && res.items.length > 0) {
                  res.items[0].insertText(valStr, Word.InsertLocation.replace);
                  fieldFound = true;
                  matchedFields.push({ label: rawLabel, location: 'Document text placeholder', value: valStr });
                  break;
                }
              }
            } catch (_) {}
          }

          if (!fieldFound) {
            unmatchedFields.push(rawLabel);
          }
        }

        await context.sync();

        return {
          ok: true,
          matched_count: matchedFields.length,
          unmatched_count: unmatchedFields.length,
          matched_fields: matchedFields,
          unmatched_fields: unmatchedFields
        };
      });
    }

    async formatText(args) {
      return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const p = paragraphs.items[args.paragraph_index];
        if (!p) return { ok: false, error: `Paragraph ${args.paragraph_index} not found.` };

        let targetRange = p;
        if (args.search_text) {
          const res = p.search(args.search_text);
          res.load('items');
          await context.sync();
          if (res.items.length > 0) {
            targetRange = res.items[0];
          }
        }

        if (typeof args.bold === 'boolean') targetRange.font.bold = args.bold;
        if (typeof args.italic === 'boolean') targetRange.font.italic = args.italic;
        if (typeof args.underline === 'boolean') targetRange.font.underline = args.underline ? 'Single' : 'None';
        if (args.color) targetRange.font.color = args.color;
        if (args.highlight) targetRange.font.highlightColor = args.highlight;
        if (args.size_pt) targetRange.font.size = args.size_pt;

        // Alignment: 'center' | 'left' | 'right' | 'justified'
        if (args.alignment) {
          const a = String(args.alignment).toLowerCase().trim();
          if (a.includes('center') || a.includes('centr')) p.alignment = 'Centered';
          else if (a.includes('right') || a.includes('destra')) p.alignment = 'Right';
          else if (a.includes('left') || a.includes('sinistra')) p.alignment = 'Left';
          else if (a.includes('justif') || a.includes('giustific')) p.alignment = 'Justified';
        }

        // Direct style application within formatText if specified
        if (args.style) {
          safeApplyParagraphStyle(p, args.style);
          const s = String(args.style).toLowerCase();
          if (s.includes('subtitle') || s.includes('sottotitolo') || s.includes('soustitre')) {
            if (!args.size_pt) p.font.size = 14;
            if (args.bold === undefined) p.font.bold = false;
          }
        }

        targetRange.select();
        await context.sync();
        return { ok: true, formatted_paragraph: args.paragraph_index };
      });
    }

    async applyStyle(paragraphIndex, style) {
      return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const p = paragraphs.items[paragraphIndex];
        if (!p) return { ok: false, error: `Paragraph ${paragraphIndex} not found.` };

        safeApplyParagraphStyle(p, style);

        // Subtitle normalization: if Subtitle was applied, reset any title-level font blowups
        const s = String(style || '').toLowerCase();
        if (s.includes('subtitle') || s.includes('sottotitolo') || s.includes('soustitre')) {
          try {
            p.font.size = 14;
            p.font.bold = false;
          } catch (_) {}
        }

        p.select();
        await context.sync();
        return { ok: true, applied_style: style };
      });
    }

    async addComment(paragraphIndex, commentText, searchText) {
      return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const p = paragraphs.items[paragraphIndex];
        if (!p) return { ok: false, error: `Paragraph ${paragraphIndex} not found.` };

        let target = p;
        if (searchText) {
          const res = p.search(searchText);
          res.load('items');
          await context.sync();
          if (res.items.length > 0) target = res.items[0];
        }

        target.insertComment(commentText);
        await context.sync();
        return { ok: true, comment_added: true };
      });
    }

    async insertPageBreak(location = 'end') {
      return await Word.run(async (context) => {
        const body = context.document.body;
        body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
        await context.sync();
        return { ok: true, page_break_inserted: true };
      });
    }

    async selectParagraph(paragraphIndex) {
      return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const p = paragraphs.items[paragraphIndex];
        if (!p) return { ok: false, error: `Paragraph ${paragraphIndex} not found.` };

        p.select();
        await context.sync();
        return { ok: true, selected_index: paragraphIndex };
      });
    }

    async setParagraphText(paragraphIndex, newText) {
      return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const p = paragraphs.items[paragraphIndex];
        if (!p) return { ok: false, error: `Paragraph ${paragraphIndex} not found.` };

        p.insertText(newText, Word.InsertLocation.replace);
        p.select();
        await context.sync();
        return { ok: true, rewritten_paragraph: paragraphIndex };
      });
    }

    async replaceText(searchText, replacementText, matchCase = false) {
      return await Word.run(async (context) => {
        const res = context.document.body.search(searchText, { matchCase: !!matchCase });
        res.load('items');
        await context.sync();

        const count = res.items.length;
        for (const item of res.items) {
          item.insertText(replacementText, Word.InsertLocation.replace);
        }
        await context.sync();
        return { ok: true, replacements_made: count };
      });
    }

    async deleteParagraph(paragraphIndex) {
      return await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();

        const p = paragraphs.items[paragraphIndex];
        if (!p) return { ok: false, error: `Paragraph ${paragraphIndex} not found.` };

        p.delete();
        await context.sync();
        return { ok: true, deleted_index: paragraphIndex };
      });
    }
  }

  // Expose globally to taskpane window
  window.wordAgent = new WordAgent();
})();
