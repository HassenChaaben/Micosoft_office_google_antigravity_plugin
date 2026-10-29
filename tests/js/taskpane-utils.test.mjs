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

const { escapeHtml, renderMarkdown, relativeTime } = loadUtils();

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
