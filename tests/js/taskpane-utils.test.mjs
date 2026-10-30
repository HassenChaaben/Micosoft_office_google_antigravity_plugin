import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';

const scriptCode = fs.readFileSync(path.resolve('web/taskpane.js'), 'utf-8');

function loadUtils() {
  const sandbox = { window: {}, document: { addEventListener: () => {} }, console };
  vm.createContext(sandbox);
  vm.runInContext(scriptCode, sandbox);
  return sandbox.window.taskpaneUtils;
}

const { escapeHtml, renderMarkdown, relativeTime, humanToolLabel, formatFamilyLabel, FAMILY_LABELS } = loadUtils();

test('escapeHtml escapes tags, ampersands and quotes', () => {
  assert.equal(escapeHtml(`<b a="1">'&'</b>`), '&lt;b a=&quot;1&quot;&gt;&#39;&amp;&#39;&lt;/b&gt;');
  assert.equal(escapeHtml(null), '');
});

test('renderMarkdown never injects raw HTML from model text', () => {
  const html = renderMarkdown('Hello <img src=x onerror=alert(1)> **bold** <script>x</script>');
  assert.ok(!html.includes('<img'), html);
  assert.ok(!html.includes('<script'), html);
  assert.ok(html.includes('&lt;img src=x onerror=alert(1)&gt;'), html);
  assert.ok(html.includes('<strong>bold</strong>'), html);
});

test('renderMarkdown renders headings, lists, italics and inline code', () => {
  const html = renderMarkdown('# Title\n## Sub\n- one\n- *two*\n1. first\n2. `x<y`');
  assert.match(html, /<h2 class="md-h md-h1">Title<\/h2>/);
  assert.match(html, /<h3 class="md-h md-h2">Sub<\/h3>/);
  assert.match(html, /<ul class="md-list"><li>one<\/li><li><em>two<\/em><\/li><\/ul>/);
  assert.match(html, /<ol class="md-list"><li>first<\/li><li><code>x&lt;y<\/code><\/li><\/ol>/);
});

test('renderMarkdown renders simple pipe tables with escaped cells', () => {
  const html = renderMarkdown('| A | B |\n| --- | :-: |\n| 1 | <i>2</i> |');
  assert.match(html, /<table class="chat-table"><thead><tr><th>A<\/th><th>B<\/th><\/tr><\/thead>/);
  assert.match(html, /<td>1<\/td><td>&lt;i&gt;2&lt;\/i&gt;<\/td>/);
});

test('renderMarkdown only links http(s) URLs and keeps quotes out of href', () => {
  const ok = renderMarkdown('[site](https://example.com/a?b=1&c=2)');
  assert.match(ok, /<a href="https:\/\/example.com\/a\?b=1&amp;c=2" target="_blank" rel="noopener noreferrer">site<\/a>/);
  const bad = renderMarkdown('[x](javascript:alert(1))');
  assert.ok(!bad.includes('<a '), bad);
  const quote = renderMarkdown('[x](https://e.com/"onmouseover="alert(1))');
  assert.ok(!/href="[^"]*"onmouseover/.test(quote), quote);
});

test('renderMarkdown escapes code blocks, including an unterminated one while streaming', () => {
  assert.match(renderMarkdown('```\n<b>hi</b>\n```'), /<pre><code>&lt;b&gt;hi&lt;\/b&gt;<\/code><\/pre>/);
  assert.match(renderMarkdown('text\n```js\nlet a = "<";'), /<pre><code>let a = &quot;&lt;&quot;;<\/code><\/pre>/);
});

test('relativeTime gives friendly labels', () => {
  const now = Date.parse('2026-10-29T12:00:00Z');
  assert.equal(relativeTime('2026-10-29T11:59:30Z', now), 'just now');
  assert.equal(relativeTime('2026-10-29 11:50:00', now), '10 min ago');
  assert.equal(relativeTime('2026-10-29T09:00:00Z', now), '3 h ago');
  assert.equal(relativeTime('2026-10-28T09:00:00Z', now), 'yesterday');
  assert.equal(relativeTime('2026-10-26T09:00:00Z', now), '3 days ago');
  assert.equal(relativeTime(Date.parse('2026-10-29T11:00:00Z') / 1000, now), '1 h ago');
  assert.equal(relativeTime('', now), '');
});

test('humanToolLabel formats tools for user display', () => {
  assert.equal(
    humanToolLabel('word_insert_ooxml', { location: 'after:p1' }),
    'Writing content into Word in one native operation (after:p1)...'
  );
  assert.equal(
    humanToolLabel('word_insert_ooxml', {}),
    'Writing content into Word in one native operation (cursor)...'
  );
  assert.equal(
    humanToolLabel('__capture_document', {}),
    'Verifying the document (re-reading and checking quality)...'
  );
  assert.equal(
    humanToolLabel('word_format_text', { paragraph_index: 3 }),
    'Applying text formatting to paragraph #3...'
  );
  assert.equal(
    humanToolLabel('unknown_custom_tool', {}),
    'Executing unknown_custom_tool...'
  );
  assert.equal(
    humanToolLabel('word_delete_column', { table_index: 0, col_index: 2 }),
    'Deleting column #2 from table #0...'
  );
  assert.equal(
    humanToolLabel('word_delete_row', { table_index: 1, row_index: 3 }),
    'Deleting row #3 from table #1...'
  );
});

test('formatFamilyLabel formats model families with display names', () => {
  assert.equal(formatFamilyLabel('gemini-3.8-flash'), 'Gemini 3.8 Flash');
  assert.equal(formatFamilyLabel('claude-sonnet-5-5'), 'Claude Sonnet 5.5');
  assert.equal(formatFamilyLabel('gpt-oss-120b'), 'GPT-OSS 120B');
  assert.equal(formatFamilyLabel(''), 'Other');
  assert.equal(formatFamilyLabel(null), 'Other');

  // Custom labels dict override
  assert.equal(
    formatFamilyLabel('custom-fam', { 'custom-fam': 'Custom Family 2.0' }),
    'Custom Family 2.0'
  );

  // Model-level family_label
  const models = [
    { id: 'custom-m1', family: 'custom-fam-2', family_label: 'Custom Display Name' }
  ];
  assert.equal(formatFamilyLabel('custom-fam-2', null, models), 'Custom Display Name');

  // Fallback formatting
  assert.equal(formatFamilyLabel('mistral-large-2'), 'Mistral Large 2');
});
