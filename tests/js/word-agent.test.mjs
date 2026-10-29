import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';

// Load word-agent.js in a mock browser environment
const scriptCode = fs.readFileSync(path.resolve('web/word-agent.js'), 'utf-8');

function createMockWordContext() {
  function makePara(text, style = 'Normal') {
    const p = {
      text,
      styleBuiltIn: style,
      alignment: 'Left',
      font: {},
      select: () => {},
      insertText: (t, loc) => { p.text = t; return p; },
      delete: function() {
        const idx = paragraphs.indexOf(this);
        if (idx >= 0) paragraphs.splice(idx, 1);
      },
      search: (q) => ({
        load: () => {},
        items: [{ text: q, insertText: () => {}, font: {} }]
      })
    };
    return p;
  }

  const paragraphs = [
    makePara('Heading 1', 'Heading 1'),
    makePara('First paragraph content.', 'Normal')
  ];
  const tables = [];

  const body = {
    paragraphs: {
      load: () => {},
      items: paragraphs
    },
    tables: {
      load: () => {},
      items: tables
    },
    inlinePictures: {
      load: () => {},
      items: []
    },
    insertInlinePictureFromBase64(b64, loc) {
      const pic = { width: 0, height: 0, select: () => {} };
      body.inlinePictures.items.push(pic);
      return pic;
    },
    insertTable(rows, cols, loc, matrix) {
      const tbl = {
        rowCount: rows,
        values: matrix,
        styleBuiltIn: '',
        select: () => {},
        load: () => {},
        delete: function() {
          const idx = tables.indexOf(this);
          if (idx >= 0) tables.splice(idx, 1);
        },
        getCell: function(r, c) {
          if (!this.values[r]) this.values[r] = [];
          return {
            set value(v) { tbl.values[r][c] = v; },
            get value() { return tbl.values[r][c]; },
            font: {},
            select: () => {}
          };
        },
        addRows: function(loc, count, vals) {
          if (vals && vals.length > 0) {
            tbl.values.push(...vals);
            tbl.rowCount += vals.length;
          }
          return { select: () => {} };
        },
        rows: {
          getFirst: () => ({ font: {}, getRange: () => ({ font: {} }), shadingColor: '' })
        }
      };
      tables.push(tbl);
      return tbl;
    },
    insertParagraph(text, loc) {
      const p = makePara(text, 'Normal');
      paragraphs.push(p);
      return p;
    },
    search(query) {
      return {
        load: () => {},
        items: [{ text: query, insertText: () => {} }]
      };
    }
  };

  const Word = {
    Style: {
      heading1: 'Heading 1',
      heading2: 'Heading 2',
      heading3: 'Heading 3',
      title: 'Title',
      subtitle: 'Subtitle',
      normal: 'Normal'
    },
    Alignment: { centered: 'Centered', left: 'Left', right: 'Right', justified: 'Justified' },
    InsertLocation: { end: 'End', start: 'Start', replace: 'Replace', after: 'After', before: 'Before' },
    BreakType: { page: 'Page' },
    run: async (fn) => {
      const ctx = {
        document: {
          body,
          getSelection: () => ({ text: 'my selection', load: () => {} })
        },
        sync: async () => {}
      };
      return await fn(ctx);
    }
  };

  return { Word, body, paragraphs, tables };
}

test('wordAgent.getDocumentOutline returns paragraphs and tables', async () => {
  const { Word } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  const res = await sandbox.window.wordAgent.executeTool('word_get_document_outline', {});
  assert.equal(res.ok, true);
  assert.equal(res.total_paragraphs, 2);
  assert.equal(res.paragraphs[0].preview, 'Heading 1');
});

test('wordAgent.insertTable inserts native table with exact columns', async () => {
  const { Word, tables } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  const headers = ['product_name', 'quantity', 'price'];
  const rows = [
    ['Standard Item A', '150', '$19.99'],
    ['Premium Item B', '45', '$49.50']
  ];

  const res = await sandbox.window.wordAgent.executeTool('word_insert_table', {
    headers,
    rows,
    table_style: 'Grid Table 4 - Accent 1'
  });

  assert.equal(res.ok, true);
  assert.equal(res.rows, 3);
  assert.equal(res.columns, 3);
  assert.deepEqual(JSON.parse(JSON.stringify(res.headers)), headers);
  assert.equal(tables.length, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(tables[0].values[0])), headers);
  assert.deepEqual(JSON.parse(JSON.stringify(tables[0].values[1])), rows[0]);
});

test('wordAgent requires approval for destructive actions', async () => {
  const { Word } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  let approvalPrompted = false;
  sandbox.window.wordAgent.onApprovalRequired = async (callId, tool, args) => {
    approvalPrompted = true;
    return false; // User rejects
  };

  const res = await sandbox.window.wordAgent.executeTool('word_delete_paragraph', {
    paragraph_index: 0
  });

  assert.equal(approvalPrompted, true);
  assert.equal(res.ok, false);
  assert.match(res.error, /rejected/i);
});

test('wordAgent.insertTable rejects missing or empty headers without creating phantom table', async () => {
  const { Word, tables } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  // Calling with empty object should return clear error without polluting document
  const res = await sandbox.window.wordAgent.executeTool('word_insert_table', {});
  assert.equal(res.ok, false);
  assert.match(res.error, /Missing required parameter 'headers'/);
  assert.equal(tables.length, 0); // No phantom table created!
});

test('wordAgent.readTable and deleteTable operate precisely on native tables', async () => {
  const { Word, tables } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  // Insert a table first
  const ins = await sandbox.window.wordAgent.executeTool('word_insert_table', {
    headers: ['A', 'B'],
    rows: [['1', '2']]
  });
  assert.equal(ins.ok, true);
  assert.equal(tables.length, 1);

  // Read table
  const read = await sandbox.window.wordAgent.executeTool('word_read_table', { table_index: 0 });
  assert.equal(read.ok, true);
  assert.equal(read.table_index, 0);
  assert.deepEqual(JSON.parse(JSON.stringify(read.values[0])), ['A', 'B']);

  // Delete table (requires approval)
  sandbox.window.wordAgent.onApprovalRequired = async () => true;
  const del = await sandbox.window.wordAgent.executeTool('word_delete_table', { table_index: 0 });
  assert.equal(del.ok, true);
  assert.equal(del.deleted_table_index, 0);
  assert.equal(tables.length, 0);
});

test('wordAgent.insertParagraph gracefully maps international styles', async () => {
  const { Word, paragraphs } = createMockWordContext();
  Word.Style = { heading1: 'Heading1', normal: 'Normal' };
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  const res = await sandbox.window.wordAgent.executeTool('word_insert_paragraph', {
    text: 'My Document Title',
    style: 'Titolo 1'
  });

  assert.equal(res.ok, true);
  assert.equal(paragraphs.length, 3);
  assert.equal(paragraphs[2].text, 'My Document Title');
});

test('wordAgent.setParagraphText edits paragraph in-place smoothly without duplicating', async () => {
  const { Word, paragraphs } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  let approvalPrompted = false;
  sandbox.window.wordAgent.onApprovalRequired = async () => {
    approvalPrompted = true;
    return true;
  };

  const initialCount = paragraphs.length;
  const res = await sandbox.window.wordAgent.executeTool('word_set_paragraph_text', {
    paragraph_index: 1,
    text: 'Updated first paragraph text without duplication.'
  });

  assert.equal(res.ok, true);
  assert.equal(approvalPrompted, false); // In-place edits are smooth, non-blocking
  assert.equal(paragraphs.length, initialCount); // No duplicate paragraph created!
  assert.equal(paragraphs[1].text, 'Updated first paragraph text without duplication.');
});

test('wordAgent.insertImage inserts native inline image with optional caption', async () => {
  const { Word, body, paragraphs } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  const res = await sandbox.window.wordAgent.executeTool('word_insert_image', {
    source: 'chart',
    caption: 'Figure 1: Performance Metrics'
  });

  assert.equal(res.ok, true);
  assert.equal(res.image_inserted, true);
  assert.equal(body.inlinePictures.items.length, 1);
  assert.equal(paragraphs[paragraphs.length - 1].text, 'Figure 1: Performance Metrics');
});

test('wordAgent.batchUpdateTableCells updates multiple table cells atomically', async () => {
  const { Word } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  // Insert a table
  await sandbox.window.wordAgent.executeTool('word_insert_table', {
    headers: ['Field', 'Value'],
    rows: [['Codice', ''], ['Titolo', '']]
  });

  const res = await sandbox.window.wordAgent.executeTool('word_batch_update_table_cells', {
    table_index: 0,
    updates: [
      { row: 1, col: 1, text: 'PRJ-2026-01', bold: true },
      { row: 2, col: 1, text: 'Autonomous Voice AI', bold: false }
    ]
  });

  assert.equal(res.ok, true);
  assert.equal(res.updated_count, 2);

  const read = await sandbox.window.wordAgent.executeTool('word_read_table', { table_index: 0 });
  assert.equal(read.values[1][1], 'PRJ-2026-01');
  assert.equal(read.values[2][1], 'Autonomous Voice AI');
});

test('wordAgent.fillFormFields matches labels and fills adjacent cells for templates', async () => {
  const { Word } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  // Simulate PM Project Charter table with merged cells
  await sandbox.window.wordAgent.executeTool('word_insert_table', {
    headers: ['Anagrafica', '', '', ''],
    rows: [
      ['Codice', 'Codice', '', ''],
      ['Titolo', 'Titolo', 'Titolo', ''],
      ['Budget complessivo', 'Budget complessivo', '', '']
    ]
  });

  // Initial Form Filling
  const res = await sandbox.window.wordAgent.executeTool('word_fill_form_fields', {
    table_index: 0,
    fields: {
      'Codice': 'PRJ-VOICE-01',
      'Titolo': 'AI Voice Customer Care',
      'Budget complessivo': '€50.000'
    }
  });

  assert.equal(res.ok, true);
  assert.equal(res.matched_count, 3);

  // Verify initial values
  let read = await sandbox.window.wordAgent.executeTool('word_read_table', { table_index: 0 });
  assert.equal(read.values[1][2], 'PRJ-VOICE-01');
  assert.equal(read.values[2][3], 'AI Voice Customer Care');
  assert.equal(read.values[3][2], '€50.000');

  // Modification / Enhancement: User asks to change the budget and title
  const modRes = await sandbox.window.wordAgent.executeTool('word_fill_form_fields', {
    table_index: 0,
    fields: {
      'Budget complessivo': '€75.000',
      'Titolo': 'Enterprise AI Voice Agent v2'
    }
  });

  assert.equal(modRes.ok, true);
  assert.equal(modRes.matched_count, 2);

  // Verify in-place modification
  read = await sandbox.window.wordAgent.executeTool('word_read_table', { table_index: 0 });
  assert.equal(read.values[3][2], '€75.000');
  assert.equal(read.values[2][3], 'Enterprise AI Voice Agent v2');
});

test('wordAgent.applyStyle applies Subtitle without colliding with Title', async () => {
  const { Word, paragraphs } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  // Apply Subtitle to paragraph 1
  const res = await sandbox.window.wordAgent.executeTool('word_apply_style', {
    paragraph_index: 1,
    style: 'Subtitle'
  });

  assert.equal(res.ok, true);
  assert.equal(paragraphs[1].styleBuiltIn, 'Subtitle'); // MUST NOT be 'Title'!
});

test('wordAgent.formatText supports centering alignment and font styling', async () => {
  const { Word, paragraphs } = createMockWordContext();
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);

  const res = await sandbox.window.wordAgent.executeTool('word_format_text', {
    paragraph_index: 1,
    alignment: 'center',
    size_pt: 14,
    bold: false
  });

  assert.equal(res.ok, true);
  assert.equal(paragraphs[1].alignment, 'Centered');
  assert.equal(paragraphs[1].font.size, 14);
  assert.equal(paragraphs[1].font.bold, false);
});



