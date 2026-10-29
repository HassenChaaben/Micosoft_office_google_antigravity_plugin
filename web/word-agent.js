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
              outcome = await this.insertTable(args.headers, args.rows, args.table_style, args.location);
              break;
            case 'word_insert_image':
              outcome = await this.insertImage(args.source, args.caption, args.location, args.width_px, args.height_px);
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

    async insertImage(source, caption, location = 'end', widthPx, heightPx) {
      if (!source) {
        return { ok: false, error: "Missing required parameter 'source'. Provide an image URL, base64 string, or keyword ('logo', 'chart', 'diagram', 'banner')." };
      }

      let base64Data = '';
      if (source.startsWith('data:image/')) {
        base64Data = source.split(',')[1] || '';
      } else if (source.startsWith('http://') || source.startsWith('https://')) {
        try {
          const resp = await fetch(source);
          const blob = await resp.blob();
          base64Data = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => {
              const res = reader.result;
              resolve(typeof res === 'string' ? res.split(',')[1] : '');
            };
            reader.onerror = reject;
            reader.readAsDataURL(blob);
          });
        } catch (fetchErr) {
          return { ok: false, error: `Failed to fetch image from URL: ${fetchErr.message || fetchErr}` };
        }
      } else {
        base64Data = this.generatePresetGraphic(source);
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

        const pic = targetRange.insertInlinePictureFromBase64(base64Data, insertLoc);
        if (widthPx && typeof widthPx === 'number') pic.width = widthPx;
        if (heightPx && typeof heightPx === 'number') pic.height = heightPx;
        pic.select();
        await context.sync();

        if (caption) {
          const capPara = body.insertParagraph(caption, Word.InsertLocation.end);
          safeApplyParagraphStyle(capPara, 'Caption');
          await context.sync();
        }

        return {
          ok: true,
          image_inserted: true,
          source_type: source.startsWith('http') ? 'url' : (source.startsWith('data:') ? 'base64' : 'preset'),
          caption: caption || null
        };
      });
    }

    generatePresetGraphic(type) {
      if (typeof document === 'undefined') {
        return 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
      }
      const canvas = document.createElement('canvas');
      canvas.width = 600;
      canvas.height = 300;
      const ctx = canvas.getContext('2d');
      if (!ctx) return 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

      const t = String(type).toLowerCase();
      const grad = ctx.createLinearGradient(0, 0, 600, 300);
      if (t.includes('chart') || t.includes('metric')) {
        grad.addColorStop(0, '#1E3A8A');
        grad.addColorStop(1, '#3B82F6');
      } else if (t.includes('logo')) {
        grad.addColorStop(0, '#0F172A');
        grad.addColorStop(1, '#334155');
      } else {
        grad.addColorStop(0, '#4338CA');
        grad.addColorStop(1, '#6366F1');
      }
      ctx.fillStyle = grad;
      ctx.fillRect(0, 0, 600, 300);

      ctx.fillStyle = '#FFFFFF';
      ctx.font = 'bold 28px -apple-system, Segoe UI, Roboto, sans-serif';
      const label = t.includes('chart') ? '📊 Performance Metrics' : (t.includes('logo') ? '🏢 Enterprise System' : '✨ Gemini Antigravity');
      ctx.fillText(label, 40, 70);

      ctx.font = '16px -apple-system, Segoe UI, Roboto, sans-serif';
      ctx.fillStyle = '#E0E7FF';
      ctx.fillText('Generated natively inside Microsoft Word', 40, 110);

      if (t.includes('chart')) {
        const barHeights = [90, 140, 110, 170, 150];
        const labels = ['Q1', 'Q2', 'Q3', 'Q4', 'Q5'];
        for (let i = 0; i < barHeights.length; i++) {
          const x = 50 + i * 105;
          const y = 260 - barHeights[i];
          ctx.fillStyle = 'rgba(255, 255, 255, 0.85)';
          ctx.fillRect(x, y, 70, barHeights[i]);
          ctx.fillStyle = '#E0E7FF';
          ctx.font = '14px sans-serif';
          ctx.fillText(labels[i], x + 25, 280);
        }
      } else {
        ctx.strokeStyle = 'rgba(255, 255, 255, 0.2)';
        ctx.lineWidth = 2;
        ctx.strokeRect(40, 140, 520, 120);
        ctx.font = '15px sans-serif';
        ctx.fillStyle = '#FFFFFF';
        ctx.fillText('• Automated Office Orchestration Active', 60, 180);
        ctx.fillText('• High Precision Document Engineering', 60, 215);
      }

      return canvas.toDataURL('image/png').split(',')[1];
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

    async insertTable(headers, rows, tableStyle = 'Grid Table 4 - Accent 1', location = 'end') {
      if (!headers || !Array.isArray(headers) || headers.length === 0) {
        return {
          ok: false,
          error: "Missing required parameter 'headers'. Example: headers: ['product_name', 'quantity', 'price']"
        };
      }

      const normHeaders = headers.map(String);
      const normRows = Array.isArray(rows) && rows.length > 0 ? rows : [
        normHeaders.map((h, i) => `Sample ${h} 1`),
        normHeaders.map((h, i) => `Sample ${h} 2`)
      ];

      return await Word.run(async (context) => {
        const body = context.document.body;
        const totalRows = normRows.length + 1;
        const totalCols = normHeaders.length;

        // Build 2D string matrix
        const matrix = [normHeaders];
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
          headers: normHeaders
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
