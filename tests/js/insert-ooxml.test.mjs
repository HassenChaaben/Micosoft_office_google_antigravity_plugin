import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';

const scriptCode = fs.readFileSync(path.resolve('web/word-agent.js'), 'utf-8');

function setup({ selectionText = '' } = {}) {
  const calls = [];
  let syncs = 0;
  const target = (name) => ({
    insertOoxml: (ooxml, loc) => { calls.push({ on: name, loc, ooxml }); return {}; },
    getRange: (kind) => target(`${name}.range(${kind || ''})`),
  });
  const paragraphs = ['Intro', 'Rischi principali', 'Body'].map((text, i) => ({ text, ...target(`p${i}`) }));
  const cc = { isNullObject: false, ...target('cc:ga:abc') };
  const body = {
    ...target('body'),
    paragraphs: { load: () => {}, items: paragraphs },
    contentControls: { getByTag: (tag) => ({ getFirstOrNullObject: () => (tag === 'ga:abc' ? cc : { isNullObject: true }) }) },
  };
  const selection = { text: selectionText, load: () => {}, ...target('selection') };
  const Word = {
    InsertLocation: { end: 'End', start: 'Start', replace: 'Replace', after: 'After', before: 'Before' },
    run: async (fn) => fn({ document: { body, getSelection: () => selection }, sync: async () => { syncs++; } }),
  };
  const sandbox = { window: {}, console, Word };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);
  return { agent: sandbox.window.wordAgent, calls, syncs: () => syncs };
}

const run = (agent, location) => agent.executeTool('word_insert_ooxml', { ooxml: '<pkg:package/>', location, block_id: 'x1' });

test('end inserts once at body end with a single sync', async () => {
  const { agent, calls, syncs } = setup();
  const res = await run(agent, 'end');
  assert.equal(res.ok, true);
  assert.equal(calls.length, 1);
  assert.deepEqual([calls[0].on, calls[0].loc], ['body', 'End']);
  assert.ok(syncs() <= 2);
});

test('cursor with empty selection replaces the caret, with text inserts after it', async () => {
  let s = setup();
  await run(s.agent, 'cursor');
  assert.deepEqual([s.calls[0].on, s.calls[0].loc], ['selection', 'Replace']);
  s = setup({ selectionText: 'chosen words' });
  await run(s.agent, 'cursor');
  assert.deepEqual([s.calls[0].on, s.calls[0].loc], ['selection', 'After']);
});

test('after:p1 and after:<heading text> insert after that paragraph', async () => {
  let s = setup();
  await run(s.agent, 'after:p1');
  assert.deepEqual([s.calls[0].on, s.calls[0].loc], ['p1.range()', 'After']);
  s = setup();
  await run(s.agent, 'after:rischi principali');
  assert.deepEqual([s.calls[0].on, s.calls[0].loc], ['p1.range()', 'After']);
});

test('replace:ga:<id> replaces the whole agent block', async () => {
  const s = setup();
  const res = await run(s.agent, 'replace:ga:abc');
  assert.equal(res.ok, true);
  assert.deepEqual([s.calls[0].on, s.calls[0].loc], ['cc:ga:abc.range(Whole)', 'Replace']);
});

test('unknown anchor fails clearly without inserting', async () => {
  const s = setup();
  const res = await run(s.agent, 'after:ga:missing');
  assert.equal(res.ok, false);
  assert.match(res.error, /not found/i);
  assert.equal(s.calls.length, 0);
});
