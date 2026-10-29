import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';

const scriptCode = fs.readFileSync(path.resolve('web/word-agent.js'), 'utf-8');

function setup({ badCell = null } = {}) {
  const writes = [];
  let syncs = 0;
  let pendingBad = false;
  const tables = [0, 1].map((t) => ({
    getCell: (r, c) => {
      if (badCell && badCell[0] === t && badCell[1] === r && badCell[2] === c) pendingBad = true;
      return {
        body: { paragraphs: { getFirst: () => ({ insertText: (text, loc) => writes.push({ t, r, c, text, loc }) }) } },
      };
    },
  }));
  const ccs = { cliente: { isNullObject: false, insertText: (text, loc) => writes.push({ cc: 'cliente', text, loc }) } };
  const body = {
    tables: { load: () => {}, items: tables },
    contentControls: { getByTag: (tag) => ({ getFirstOrNullObject: () => ccs[tag] || { isNullObject: true } }) },
  };
  const Word = {
    InsertLocation: { replace: 'Replace' },
    run: async (fn) => fn({
      document: { body },
      sync: async () => { syncs++; if (pendingBad) { pendingBad = false; throw new Error('InvalidArgument'); } },
    }),
  };
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);
  return { agent: sandbox.window.wordAgent, writes, syncs: () => syncs };
}

test('writes every cell in place with one batched sync', async () => {
  const s = setup();
  const res = await s.agent.executeTool('word_fill_cells', {
    cells: [{ table: 0, row: 6, cell: 1, text: 'Cloud4Job' }, { table: 1, row: 2, cell: 3, text: 'Ritardo' }],
    content_controls: { cliente: 'ACME' },
  });
  assert.equal(res.ok, true);
  assert.equal(res.written, 3);
  assert.deepEqual(s.writes.map((w) => w.loc), ['Replace', 'Replace', 'Replace']);
  assert.ok(s.syncs() <= 2);
});

test('a bad cell does not block the others and is reported', async () => {
  const s = setup({ badCell: [1, 9, 9] });
  const res = await s.agent.executeTool('word_fill_cells', {
    cells: [{ table: 0, row: 6, cell: 1, text: 'ok' }, { table: 1, row: 9, cell: 9, text: 'bad' }],
  });
  assert.equal(res.ok, true);
  assert.equal(res.written, 1);
  assert.equal(res.failed.length, 1);
  assert.match(res.failed[0], /T1r9c9/);
});

test('missing table index fails clearly', async () => {
  const s = setup();
  const res = await s.agent.executeTool('word_fill_cells', { cells: [{ table: 7, row: 0, cell: 0, text: 'x' }] });
  assert.equal(res.written, 0);
  assert.match(res.failed[0], /T7/);
});
