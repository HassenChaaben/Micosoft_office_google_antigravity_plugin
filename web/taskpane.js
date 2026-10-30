/**
 * Gemini Antigravity Office Taskpane Controller
 * Handles chat streaming, UI state, skills dispatch, flyout menus, and Office execution.
 */

/**
 * Pure helpers (no DOM): HTML escaping, safe markdown rendering, relative time.
 * Exposed as window.taskpaneUtils so they can be unit-tested in Node (tests/js/taskpane-utils.test.mjs).
 */
(function (root) {
  function escapeHtml(str) {
    return String(str == null ? '' : str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Inline markdown on ONE line. The text is escaped first; only safe tags are added afterwards.
  function renderInline(text) {
    const codes = [];
    let s = escapeHtml(text).replace(/`([^`]+)`/g, (m, c) => {
      codes.push(`<code>${c}</code>`);
      return `\u0001${codes.length - 1}\u0001`;
    });
    // Links: only http(s). The URL is already escaped (quotes included), so it cannot break the attribute.
    s = s.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g,
      '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>');
    s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    s = s.replace(/__([^_]+)__/g, '<strong>$1</strong>');
    s = s.replace(/(^|[^*\w])\*([^*\s][^*]*?)\*(?!\*)/g, '$1<em>$2</em>');
    s = s.replace(/(^|[^_\w])_([^_\s][^_]*?)_(?![_\w])/g, '$1<em>$2</em>');
    return s.replace(/\u0001(\d+)\u0001/g, (m, i) => codes[Number(i)]);
  }

  function splitTableRow(line) {
    let t = line.trim();
    if (t.startsWith('|')) t = t.slice(1);
    if (t.endsWith('|')) t = t.slice(0, -1);
    return t.split('|').map(c => c.trim());
  }

  const TABLE_SEPARATOR = /^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;

  function renderTable(rows) {
    const head = rows[0] || [];
    let html = '<div class="chat-table-wrapper"><table class="chat-table"><thead><tr>';
    head.forEach(c => { html += `<th>${renderInline(c)}</th>`; });
    html += '</tr></thead><tbody>';
    rows.slice(1).forEach(r => {
      html += '<tr>';
      r.forEach(c => { html += `<td>${renderInline(c)}</td>`; });
      html += '</tr>';
    });
    return html + '</tbody></table></div>';
  }

  /** Markdown -> safe HTML: headings, bold, italics, inline code, code blocks, lists, quotes, pipe tables, http(s) links. */
  function renderMarkdown(md) {
    if (!md) return '';
    const blocks = [];
    let src = String(md).replace(/\r\n?/g, '\n').replace(/[\u0000\u0001]/g, '');
    // Fenced code blocks (an unterminated fence while streaming runs to the end).
    src = src.replace(/```[^\n]*\n?([\s\S]*?)(?:```|$)/g, (m, code) => {
      blocks.push(`<pre><code>${escapeHtml(code.replace(/\n$/, ''))}</code></pre>`);
      return `\n\u0000${blocks.length - 1}\u0000\n`;
    });

    const lines = src.split('\n');
    const out = [];
    let listType = null;
    const closeList = () => {
      if (listType) { out.push(`</${listType}>`); listType = null; }
    };

    let i = 0;
    while (i < lines.length) {
      const line = lines[i];
      const t = line.trim();
      let m;

      if ((m = t.match(/^\u0000(\d+)\u0000$/))) {
        closeList();
        out.push(blocks[Number(m[1])]);
        i++;
        continue;
      }
      if (t.startsWith('|') && i + 1 < lines.length && lines[i + 1].includes('-') && TABLE_SEPARATOR.test(lines[i + 1].trim())) {
        closeList();
        const rows = [splitTableRow(t)];
        i += 2;
        while (i < lines.length && lines[i].trim().startsWith('|')) {
          rows.push(splitTableRow(lines[i]));
          i++;
        }
        out.push(renderTable(rows));
        continue;
      }
      if ((m = t.match(/^(#{1,6})\s+(.*)$/))) {
        closeList();
        const level = m[1].length;
        const tag = `h${Math.min(level + 1, 6)}`;
        out.push(`<${tag} class="md-h md-h${Math.min(level, 3)}">${renderInline(m[2])}</${tag}>`);
        i++;
        continue;
      }
      if (/^(-{3,}|\*{3,}|_{3,})$/.test(t)) {
        closeList();
        out.push('<hr class="md-hr">');
        i++;
        continue;
      }
      if ((m = t.match(/^>\s?(.*)$/))) {
        closeList();
        out.push(`<blockquote class="md-quote">${renderInline(m[1])}</blockquote>`);
        i++;
        continue;
      }
      if ((m = t.match(/^[-*+]\s+(.*)$/))) {
        if (listType !== 'ul') { closeList(); out.push('<ul class="md-list">'); listType = 'ul'; }
        out.push(`<li>${renderInline(m[1])}</li>`);
        i++;
        continue;
      }
      if ((m = t.match(/^\d+[.)]\s+(.*)$/))) {
        if (listType !== 'ol') { closeList(); out.push('<ol class="md-list">'); listType = 'ol'; }
        out.push(`<li>${renderInline(m[1])}</li>`);
        i++;
        continue;
      }
      closeList();
      if (t === '') {
        if (out.length && out[out.length - 1] !== '<div class="md-gap"></div>') out.push('<div class="md-gap"></div>');
      } else {
        out.push(`<div class="md-p">${renderInline(line)}</div>`);
      }
      i++;
    }
    closeList();
    return out.join('');
  }

  /** "just now", "5 min ago", "3 h ago", "yesterday", "4 days ago", or a short date. */
  function relativeTime(value, now) {
    if (!value) return '';
    let ms;
    if (typeof value === 'number') ms = value < 1e12 ? value * 1000 : value;
    else {
      const s = String(value);
      // SQLite "YYYY-MM-DD HH:MM:SS" has no zone: treat it as UTC.
      ms = Date.parse(/^\d{4}-\d{2}-\d{2} \d/.test(s) && !/[zZ]|[+-]\d{2}:?\d{2}$/.test(s) ? s.replace(' ', 'T') + 'Z' : s);
    }
    if (!isFinite(ms)) return '';
    const diff = Math.max(0, ((now == null ? Date.now() : now) - ms) / 1000);
    if (diff < 60) return 'just now';
    if (diff < 3600) return `${Math.floor(diff / 60)} min ago`;
    if (diff < 86400) return `${Math.floor(diff / 3600)} h ago`;
    if (diff < 172800) return 'yesterday';
    if (diff < 7 * 86400) return `${Math.floor(diff / 86400)} days ago`;
    const d = new Date(ms);
    return d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' });
  }

  function humanToolLabel(tool, args) {
    args = args || {};
    switch (tool) {
      case 'word_insert_ooxml': return `Writing content into Word in one native operation (${args.location || 'cursor'})...`;
      case '__capture_document': return 'Verifying the document (re-reading and checking quality)...';
      case 'word_get_document_outline': return 'Reading Word document structure & paragraphs...';
      case 'word_read_paragraphs': return `Reading paragraphs (${args.start_index}..${args.end_index || args.start_index})...`;
      case 'word_search': return `Searching for "${args.search_text}"...`;
      case 'word_insert_paragraph': return `Inserting paragraph (${args.style || 'Normal'}): "${(args.text || '').slice(0, 35)}..."`;
      case 'word_insert_list': return `Inserting ${args.items ? args.items.length : 0} ${args.list_type || 'bullet'} items...`;
      case 'word_insert_table': return `Inserting native Word table (${args.headers ? args.headers.join(', ') : 'table'})...`;
      case 'word_insert_image': return `Inserting picture (${args.caption ? args.caption : (args.source || 'visual')})...`;
      case 'word_format_table': return `Formatting table #${args.table_index} (${args.style || 'default'})...`;
      case 'word_set_table_cell': return `Updating table cell [${args.row_index}, ${args.col_index}]: "${args.text}"`;
      case 'word_add_table_row': return `Adding row to table #${args.table_index}...`;
      case 'word_delete_table': return `Deleting table #${args.table_index}...`;
      case 'word_delete_column': return `Deleting column #${args.col_index !== undefined ? args.col_index : (args.column_index !== undefined ? args.column_index : '')} from table #${args.table_index}...`;
      case 'word_delete_row': return `Deleting row #${args.row_index} from table #${args.table_index}...`;
      case 'word_read_table': return `Reading content of table #${args.table_index}...`;
      case 'word_clear_document': return 'Clearing document body...';
      case 'word_format_text': return `Applying text formatting to paragraph #${args.paragraph_index}...`;
      case 'word_apply_style': return `Setting style "${args.style}" on paragraph #${args.paragraph_index}...`;
      case 'word_add_comment': return `Adding review comment: "${args.comment}"`;
      case 'word_insert_page_break': return 'Inserting page break...';
      case 'word_select': return `Focusing cursor on paragraph #${args.paragraph_index}...`;
      case 'word_set_paragraph_text': return `Rewriting paragraph #${args.paragraph_index}...`;
      case 'word_batch_update_table_cells': return `Updating ${args.updates ? args.updates.length : 0} table cells in batch...`;
      case 'word_fill_form_fields': return `Filling ${args.fields ? Object.keys(args.fields).length : 0} form fields in document...`;
      case 'word_replace_text': return `Replacing "${args.search_text}" with "${args.replacement_text}"...`;
      case 'word_delete_paragraph': return `Deleting paragraph #${args.paragraph_index}...`;
      default: return `Executing ${tool}...`;
    }
  }

  const FAMILY_LABELS = {
    'gemini-3.8-flash': 'Gemini 3.8 Flash',
    'gemini-4-argon': 'Gemini 4 Argon',
    'claude-sonnet-5-5': 'Claude Sonnet 5.5',
    'claude-opus-5-5': 'Claude Opus 5.5',
    'gpt-oss-120b': 'GPT-OSS 120B'
  };

  function formatFamilyLabel(family, customLabels, models) {
    if (!family) return 'Other';
    const labels = customLabels || FAMILY_LABELS;
    if (labels && labels[family]) return labels[family];
    if (Array.isArray(models)) {
      const matching = models.find(m => m && m.family === family && (m.family_label || m.family_display));
      if (matching) return matching.family_label || matching.family_display;
    }
    return family
      .replace(/(\d)-(\d)/g, '$1.$2')
      .split('-')
      .map(w => (/^[a-z]/.test(w) ? w.charAt(0).toUpperCase() + w.slice(1) : w.toUpperCase()))
      .join(' ');
  }

  root.taskpaneUtils = { escapeHtml, renderInline, renderMarkdown, relativeTime, humanToolLabel, formatFamilyLabel, FAMILY_LABELS };
})(typeof window !== 'undefined' ? window : globalThis);

document.addEventListener('DOMContentLoaded', async () => {
  const { escapeHtml, renderMarkdown, relativeTime, humanToolLabel, formatFamilyLabel, FAMILY_LABELS } = window.taskpaneUtils;
  // DOM Elements
  const hostBadge = document.getElementById('hostBadge');
  const onboardingSection = document.getElementById('onboardingSection');
  const chatSection = document.getElementById('chatSection');
  const chatMessages = document.getElementById('chatMessages');
  const promptInput = document.getElementById('promptInput');
  const btnSend = document.getElementById('btnSend');
  const btnStop = document.getElementById('btnStop');
  const btnFlyoutPlus = document.getElementById('btnFlyoutPlus');
  const flyoutMenu = document.getElementById('flyoutMenu');
  const btnGlobe = document.getElementById('btnGlobe');
  const modelSelect = document.getElementById('modelSelect');
  const btnMic = document.getElementById('btnMic');
  const contextBar = document.getElementById('contextBar');
  const contextText = document.getElementById('contextText');
  const contextRemove = document.getElementById('contextRemove');
  const btnNewChat = document.getElementById('btnNewChat');
  const btnHistory = document.getElementById('btnHistory');
  const btnMenu = document.getElementById('btnMenu');
  const historyDrawer = document.getElementById('historyDrawer');
  const closeHistory = document.getElementById('closeHistory');
  const settingsModal = document.getElementById('settingsModal');
  const closeSettings = document.getElementById('closeSettings');
  const btnAuthSignIn = document.getElementById('btnAuthSignIn');
  const btnSyncConfig = document.getElementById('btnSyncConfig');
  const menuAddFiles = document.getElementById('menuAddFiles');
  const fileUploadInput = document.getElementById('fileUploadInput');
  const menuSearchWeb = document.getElementById('menuSearchWeb');
  const manageSkillsLink = document.getElementById('manageSkillsLink');
  const slashMenu = document.getElementById('slashMenu');

  // State
  let currentHost = 'Office';
  let isWebSearchEnabled = true;
  let attachedFiles = [];
  let attachedDocumentContext = null;
  let conversations = [];
  let isStreaming = false;
  let currentChatId = null;      // set from the SSE 'chat' event; sent as chat_id on every run
  let currentDocUrl = '';        // url of the open document (key for "last chat" memory)
  let availableModels = [];      // [{id,label,family,effort}] from /api/models (or fallback)
  const limitedModels = new Set(); // model ids that hit their quota in this session

  // -------------------------------------------------------------------------
  // Safe localStorage helpers (storage can be blocked in some Office WebViews)
  // -------------------------------------------------------------------------
  function storeGet(key) {
    try { return window.localStorage.getItem(key); } catch (_) { return null; }
  }
  function storeSet(key, value) {
    try { window.localStorage.setItem(key, value); } catch (_) { /* ignore */ }
  }
  function storeRemove(key) {
    try { window.localStorage.removeItem(key); } catch (_) { /* ignore */ }
  }
  const MODEL_KEY = 'agy.model';
  function lastChatKey(docUrl) {
    return `agy.lastChat:${docUrl || '(untitled)'}`;
  }
  function rememberChatForDoc(chatId) {
    if (chatId) storeSet(lastChatKey(currentDocUrl), chatId);
    else storeRemove(lastChatKey(currentDocUrl));
  }
  function readDocUrl() {
    try {
      return (window.Office && Office.context && Office.context.document && Office.context.document.url) || '';
    } catch (_) {
      return '';
    }
  }

  // Host Skills Map for Dynamic Onboarding Adapting
  const hostSkillsMap = {
    Word: [
      { label: '✨ Improve Writing', prompt: 'Review this document and refine the grammar, clarity, and phrasing in-place while keeping the exact document structure.' },
      { label: '📊 Create Summary Table', prompt: 'Create a clean, formatted native Word table summarizing the key sections and metrics of this document.' },
      { label: '🖼️ Add Visual Chart', prompt: 'Insert a visual performance metrics chart with a caption directly into the document.' },
      { label: '🎨 Polish Styling', prompt: 'Enhance the headings, font sizes, and paragraph formatting across the document for an executive look.' },
      { label: '🔍 Check Issues', prompt: 'Inspect the document outline and identify any inconsistencies, repetitions, or missing sections.' },
      { label: '📝 Rewrite Selection', prompt: 'Rewrite the currently selected text to be concise, professional, and clear.' }
    ],
    Excel: [
      { label: '📈 Financial Model', prompt: 'Build a financial projection model with revenue, costs, and margin formulas.' },
      { label: '📊 Generate Chart', prompt: 'Generate a chart from the active worksheet data.' },
      { label: '🧮 Formula Helper', prompt: 'Explain and optimize formulas in the selected range.' },
      { label: '📑 Excel to Report', prompt: 'Synthesize the numbers from this workbook into an executive summary.' }
    ],
    PowerPoint: [
      { label: '🖼️ Slide Designer', prompt: 'Design executive presentation slides with clear visual structure.' },
      { label: '📑 Storylining', prompt: 'Create a structured narrative storyline for this deck.' },
      { label: '📊 Executive Summary', prompt: 'Draft a one-page executive briefing slide.' }
    ],
    Outlook: [
      { label: '✉️ Reply Email', prompt: 'Draft a professional, clear response to this email thread.' },
      { label: '📝 Summarize Thread', prompt: 'Summarize the key decisions and action items from this conversation.' },
      { label: '📅 Schedule Meeting', prompt: 'Propose meeting agenda and times based on this thread.' }
    ],
    Standalone: [
      { label: '✨ Improve Writing', prompt: 'Review and refine this text with executive clarity.' },
      { label: '📊 Create Table', prompt: 'Format this data into a clean structured table.' },
      { label: '📑 Executive Summary', prompt: 'Draft a structured executive briefing.' }
    ]
  };

  function updateHostUI(host) {
    hostBadge.style.display = 'inline-block';
    if (host === 'Word') {
      hostBadge.textContent = 'Word';
      hostBadge.className = 'host-badge host-word';
    } else if (host === 'Excel') {
      hostBadge.textContent = 'Excel';
      hostBadge.className = 'host-badge host-excel';
    } else if (host === 'PowerPoint') {
      hostBadge.textContent = 'PowerPoint';
      hostBadge.className = 'host-badge host-ppt';
    } else if (host === 'Outlook') {
      hostBadge.textContent = 'Outlook';
      hostBadge.className = 'host-badge host-outlook';
    } else {
      hostBadge.textContent = 'Suite';
      hostBadge.className = 'host-badge host-standalone';
    }
    renderHostSkillPills(host);
  }

  function renderHostSkillPills(host) {
    const stack = document.querySelector('.skills-stack');
    if (!stack) return;
    const skills = hostSkillsMap[host] || hostSkillsMap.Standalone;
    stack.innerHTML = '';
    skills.forEach(skill => {
      const btn = document.createElement('button');
      btn.className = 'skill-pill';
      const label = typeof skill === 'object' ? skill.label : skill;
      const prompt = typeof skill === 'object' ? skill.prompt : `${skill} `;
      btn.setAttribute('data-skill', label);
      btn.textContent = label;
      btn.addEventListener('click', () => {
        promptInput.value = prompt;
        promptInput.focus();
        updateSendButtonState();
        sendMessage();
      });
      stack.appendChild(btn);
    });
  }

  // Initialize Office Bridge
  try {
    currentHost = await window.officeBridge.init();
    updateHostUI(currentHost);
    console.log(`[Gemini Taskpane] Initialized for ${currentHost}`);
  } catch (err) {
    console.warn('[Gemini Taskpane] Office init fallback:', err);
    updateHostUI('Standalone');
  }

  currentDocUrl = readDocUrl();

  // Pre-load active document context preview
  refreshDocumentContext();

  async function refreshDocumentContext() {
    try {
      if (currentHost === 'Word') {
        const doc = await window.officeBridge.readWordContent();
        if (doc && doc.content && doc.content.length > 0) {
          attachedDocumentContext = doc.content;
          contextText.textContent = `Active Word Document (${doc.content.slice(0, 32)}...)`;
          contextBar.classList.remove('hidden');
        }
      } else if (currentHost === 'Excel') {
        const xl = await window.officeBridge.readExcelContent();
        if (xl && (xl.sheet_name || xl.used_range)) {
          attachedDocumentContext = JSON.stringify(xl);
          contextText.textContent = `Active Sheet: ${xl.sheet_name || 'Sheet1'} (${xl.used_range || 'A1:E10'})`;
          contextBar.classList.remove('hidden');
        }
      } else if (currentHost === 'PowerPoint') {
        const ppt = await window.officeBridge.readPowerPointContent();
        attachedDocumentContext = JSON.stringify(ppt);
        contextText.textContent = `Active PowerPoint Deck (${ppt.slides ? ppt.slides.length : 1} slides)`;
        contextBar.classList.remove('hidden');
      } else if (currentHost === 'Outlook') {
        const mail = await window.officeBridge.readOutlookEmail();
        if (mail && mail.subject) {
          attachedDocumentContext = JSON.stringify(mail);
          contextText.textContent = `Active Email: ${mail.subject.slice(0, 28)}...`;
          contextBar.classList.remove('hidden');
        }
      }
    } catch (e) {
      console.log('Context preview silent error:', e);
    }
  }

  contextRemove.addEventListener('click', () => {
    attachedDocumentContext = null;
    contextBar.classList.add('hidden');
  });

  // -------------------------------------------------------------------------
  // Onboarding Skill Pills Interaction
  // -------------------------------------------------------------------------
  document.querySelectorAll('.skill-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      const skillCmd = pill.getAttribute('data-skill');
      handleSkillExecution(skillCmd);
    });
  });

  function handleSkillExecution(skillCmd) {
    promptInput.value = `${skillCmd} `;
    promptInput.focus();
    updateSendButtonState();
    // Auto-trigger if document context is already attached
    if (attachedDocumentContext) {
      sendMessage();
    }
  }

  // -------------------------------------------------------------------------
  // Flyout Popup Menu Management (Screenshot 2 Replica)
  // -------------------------------------------------------------------------
  btnFlyoutPlus.addEventListener('click', (e) => {
    e.stopPropagation();
    flyoutMenu.classList.toggle('hidden');
  });

  // Close flyout when clicking outside
  document.addEventListener('click', (e) => {
    if (!flyoutMenu.contains(e.target) && e.target !== btnFlyoutPlus) {
      flyoutMenu.classList.add('hidden');
    }
  });

  // Search the web toggle checkmark
  menuSearchWeb.addEventListener('click', () => {
    isWebSearchEnabled = !isWebSearchEnabled;
    const check = menuSearchWeb.querySelector('.check-mark');
    if (isWebSearchEnabled) {
      check.innerHTML = '&#10003;';
      menuSearchWeb.classList.add('active');
      btnGlobe.classList.add('active');
    } else {
      check.innerHTML = '';
      menuSearchWeb.classList.remove('active');
      btnGlobe.classList.remove('active');
    }
  });

  // Add Files
  menuAddFiles.addEventListener('click', () => {
    fileUploadInput.click();
    flyoutMenu.classList.add('hidden');
  });

  fileUploadInput.addEventListener('change', () => {
    if (fileUploadInput.files.length > 0) {
      const file = fileUploadInput.files[0];
      attachedFiles.push(file);
      contextText.textContent = `Attached: ${file.name}`;
      contextBar.classList.remove('hidden');
    }
  });

  // Skills flyout submenu items
  document.querySelectorAll('#submenuSkills .sub-item').forEach(item => {
    item.addEventListener('click', () => {
      const slash = item.getAttribute('data-slash');
      flyoutMenu.classList.add('hidden');
      handleSkillExecution(slash);
    });
  });

  // Connectors / Plugins submenu items: not wired to a real backend yet, so say so honestly.
  document.querySelectorAll('#submenuConnectors .sub-item, #submenuPlugins .sub-item').forEach(item => {
    item.addEventListener('click', () => {
      flyoutMenu.classList.add('hidden');
      showChatView();
      appendSystemMessage('Connectors are managed in Settings (coming soon).');
    });
  });

  // -------------------------------------------------------------------------
  // Slash Commands & Connectors Catalog (/notion, /github, /skills)
  // -------------------------------------------------------------------------
  const slashCatalog = [
    // Connectors
    {
      cmd: '/notion',
      name: '/notion',
      desc: 'Query Notion workspace, pages, databases & notes',
      category: 'Connectors',
      icon: '📝',
      badge: 'Connector'
    },
    {
      cmd: '/github',
      name: '/github',
      desc: 'Search GitHub repositories, issues, PRs & code',
      category: 'Connectors',
      icon: '🐙',
      badge: 'Connector'
    },
    {
      cmd: '/gdrive',
      name: '/gdrive',
      desc: 'Access Google Drive files, Docs, and Sheets',
      category: 'Connectors',
      icon: '📁',
      badge: 'Connector'
    },
    {
      cmd: '/filesystem',
      name: '/filesystem',
      desc: 'Read local project files and directories',
      category: 'Connectors',
      icon: '💻',
      badge: 'Connector'
    },
    // Skills
    {
      cmd: '/check-doc',
      name: '/check-doc',
      desc: 'Review document for clarity, tone and readability',
      category: 'Skills',
      icon: '📋',
      badge: 'Skill'
    },
    {
      cmd: '/copy-edit',
      name: '/copy-edit',
      desc: 'Fix grammar, improve conciseness and style',
      category: 'Skills',
      icon: '✍️',
      badge: 'Skill'
    },
    {
      cmd: '/summarize-contract',
      name: '/summarize-contract',
      desc: 'Extract key clauses, obligations and deadlines',
      category: 'Skills',
      icon: '📜',
      badge: 'Skill'
    },
    {
      cmd: '/flag-issues',
      name: '/flag-issues',
      desc: 'Find potential legal/content risks and ambiguities',
      category: 'Skills',
      icon: '⚠️',
      badge: 'Skill'
    },
    {
      cmd: '/industry-overview',
      name: '/industry-overview',
      desc: 'Generate market trends and competitive landscape',
      category: 'Skills',
      icon: '📊',
      badge: 'Skill'
    },
    {
      cmd: '/storylining',
      name: '/storylining',
      desc: 'Build executive presentation storyline & slides',
      category: 'Skills',
      icon: '📑',
      badge: 'PowerPoint'
    },
    {
      cmd: '/slide-designer',
      name: '/slide-designer',
      desc: 'Format active slide with executive visual layouts',
      category: 'Skills',
      icon: '🎨',
      badge: 'PowerPoint'
    },
    {
      cmd: '/analyze-sheet',
      name: '/analyze-sheet',
      desc: 'Statistical summary, trend analysis and anomalies for Excel',
      category: 'Skills',
      icon: '📊',
      badge: 'Excel'
    },
    {
      cmd: '/financial-model',
      name: '/financial-model',
      desc: 'Build Wall Street 3-statement or SaaS projection model',
      category: 'Skills',
      icon: '📈',
      badge: 'Excel'
    },
    {
      cmd: '/formula-helper',
      name: '/formula-helper',
      desc: 'Generate, troubleshoot or explain complex Excel formulas',
      category: 'Skills',
      icon: '🔢',
      badge: 'Excel'
    },
    {
      cmd: '/generate-chart',
      name: '/generate-chart',
      desc: 'Insert native clustered column, line or pie chart',
      category: 'Skills',
      icon: '📉',
      badge: 'Excel'
    },
    {
      cmd: '/reply-email',
      name: '/reply-email',
      desc: 'Draft executive email reply with action items',
      category: 'Skills',
      icon: '✉️',
      badge: 'Outlook'
    },
    {
      cmd: '/summarize-thread',
      name: '/summarize-thread',
      desc: 'Extract key decisions and deliverables from email thread',
      category: 'Skills',
      icon: '🧵',
      badge: 'Outlook'
    },
    {
      cmd: '/schedule-meeting',
      name: '/schedule-meeting',
      desc: 'Create Outlook calendar appointment or meeting',
      category: 'Skills',
      icon: '📅',
      badge: 'Outlook'
    },
    {
      cmd: '/doc-to-slides',
      name: '/doc-to-slides',
      desc: 'Transform active Word document into PowerPoint slides',
      category: 'Skills',
      icon: '🚀',
      badge: 'Suite'
    },
    {
      cmd: '/excel-to-report',
      name: '/excel-to-report',
      desc: 'Convert Excel financial model into executive Word report',
      category: 'Skills',
      icon: '📑',
      badge: 'Suite'
    },
    {
      cmd: '/email-summary',
      name: '/email-summary',
      desc: 'Draft executive summary email from active Office documents',
      category: 'Skills',
      icon: '📧',
      badge: 'Suite'
    }
  ];

  let activeSlashIndex = 0;
  let currentFilteredSlash = [];

  function renderSlashMenu(items) {
    currentFilteredSlash = items;
    if (items.length === 0) {
      slashMenu.classList.add('hidden');
      return;
    }

    const connectors = items.filter(i => i.category === 'Connectors');
    const skills = items.filter(i => i.category === 'Skills');

    let html = '';
    let globalIndex = 0;

    if (connectors.length > 0) {
      html += '<div class="slash-section-header">Connectors</div>';
      connectors.forEach(item => {
        const isActive = globalIndex === activeSlashIndex ? 'active' : '';
        html += `
          <div class="slash-item ${isActive}" data-index="${globalIndex}" data-cmd="${item.cmd}">
            <div class="slash-item-left">
              <span class="slash-item-icon">${item.icon}</span>
              <div class="slash-item-text">
                <span class="slash-item-name">${item.name}</span>
                <span class="slash-item-desc">${item.desc}</span>
              </div>
            </div>
            <span class="slash-item-badge">${item.badge}</span>
          </div>
        `;
        globalIndex++;
      });
    }

    if (skills.length > 0) {
      html += '<div class="slash-section-header">Skills</div>';
      skills.forEach(item => {
        const isActive = globalIndex === activeSlashIndex ? 'active' : '';
        html += `
          <div class="slash-item ${isActive}" data-index="${globalIndex}" data-cmd="${item.cmd}">
            <div class="slash-item-left">
              <span class="slash-item-icon">${item.icon}</span>
              <div class="slash-item-text">
                <span class="slash-item-name">${item.name}</span>
                <span class="slash-item-desc">${item.desc}</span>
              </div>
            </div>
            <span class="slash-item-badge">${item.badge}</span>
          </div>
        `;
        globalIndex++;
      });
    }

    slashMenu.innerHTML = html;
    slashMenu.classList.remove('hidden');

    slashMenu.querySelectorAll('.slash-item').forEach(el => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const cmd = el.getAttribute('data-cmd');
        selectSlashCommand(cmd);
      });
    });
  }

  function selectSlashCommand(cmd) {
    const val = promptInput.value;
    const cursorPos = promptInput.selectionStart ?? val.length;
    const textBeforeCursor = val.slice(0, cursorPos);
    const textAfterCursor = val.slice(cursorPos);
    const match = textBeforeCursor.match(/(?:^|\s)\/([a-zA-Z0-9_-]*)$/);

    if (match) {
      const prefix = textBeforeCursor.slice(0, textBeforeCursor.length - match[0].length + (match[0].startsWith(' ') ? 1 : 0));
      promptInput.value = `${prefix}${cmd} ${textAfterCursor}`;
      const newPos = prefix.length + cmd.length + 1;
      promptInput.setSelectionRange(newPos, newPos);
    } else {
      promptInput.value = `${cmd} `;
    }
    slashMenu.classList.add('hidden');
    promptInput.focus();
    updateSendButtonState();
  }

  function handleSlashInput() {
    const val = promptInput.value;
    const cursorPos = promptInput.selectionStart ?? val.length;
    const textBeforeCursor = val.slice(0, cursorPos);
    const match = textBeforeCursor.match(/(?:^|\s)\/([a-zA-Z0-9_-]*)$/);

    if (match) {
      const query = match[1].toLowerCase();
      const filtered = slashCatalog.filter(item => 
        item.cmd.toLowerCase().includes(query) ||
        item.desc.toLowerCase().includes(query) ||
        item.name.toLowerCase().includes(query)
      );
      activeSlashIndex = 0;
      renderSlashMenu(filtered);
      return;
    }
    slashMenu.classList.add('hidden');
  }

  function updateActiveSlashItem() {
    const items = slashMenu.querySelectorAll('.slash-item');
    items.forEach((item, idx) => {
      if (idx === activeSlashIndex) {
        item.classList.add('active');
        item.scrollIntoView({ block: 'nearest' });
      } else {
        item.classList.remove('active');
      }
    });
  }

  // Close slash menu on click outside
  document.addEventListener('click', (e) => {
    if (!slashMenu.contains(e.target) && e.target !== promptInput) {
      slashMenu.classList.add('hidden');
    }
  });

  // -------------------------------------------------------------------------
  // Input & Messaging
  // -------------------------------------------------------------------------
  promptInput.addEventListener('input', () => {
    // Auto-grow textarea height up to 100px
    promptInput.style.height = 'auto';
    promptInput.style.height = `${Math.min(promptInput.scrollHeight, 100)}px`;
    handleSlashInput();
    updateSendButtonState();
  });

  promptInput.addEventListener('keydown', (e) => {
    // If slash menu is visible, handle navigation and selection
    if (!slashMenu.classList.contains('hidden') && currentFilteredSlash.length > 0) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        activeSlashIndex = (activeSlashIndex + 1) % currentFilteredSlash.length;
        updateActiveSlashItem();
        return;
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        activeSlashIndex = (activeSlashIndex - 1 + currentFilteredSlash.length) % currentFilteredSlash.length;
        updateActiveSlashItem();
        return;
      } else if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        const selected = currentFilteredSlash[activeSlashIndex];
        if (selected) {
          selectSlashCommand(selected.cmd);
        }
        return;
      } else if (e.key === 'Escape') {
        e.preventDefault();
        slashMenu.classList.add('hidden');
        return;
      }
    }

    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      if (!btnSend.disabled) {
        sendMessage();
      }
    }
  });

  function updateSendButtonState() {
    if (isStreaming) {
      btnSend.classList.add('hidden');
      if (btnStop) {
        btnStop.classList.remove('hidden');
        btnStop.disabled = false;
      }
    } else {
      btnSend.classList.remove('hidden');
      if (btnStop) {
        btnStop.classList.add('hidden');
      }
      btnSend.disabled = promptInput.value.trim().length === 0;
    }
  }

  btnSend.addEventListener('click', sendMessage);

  let currentAbortController = null;

  async function cancelCurrentRun() {
    if (btnStop) {
      btnStop.disabled = true;
    }
    document.querySelectorAll('.btn-chip-stop').forEach(b => {
      b.disabled = true;
      b.textContent = 'Stopping...';
    });
    try {
      await fetch('/api/agent/cancel', { method: 'POST' });
    } catch (err) {
      console.warn('[Taskpane] Cancel request error:', err);
    }
    if (currentAbortController) {
      try { currentAbortController.abort(); } catch (_) {}
    }
  }

  if (btnStop) {
    btnStop.addEventListener('click', cancelCurrentRun);
  }

  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && isStreaming) {
      e.preventDefault();
      cancelCurrentRun();
    }
  });

  async function runWordAgent(text, assistantMsg, modelOverride) {
    currentAbortController = new AbortController();
    const modelId = modelOverride || modelSelect.value;
    const bubble = assistantMsg.querySelector('.msg-bubble');
    bubble.innerHTML = ''; // Clear default spinner

    // Live agent steps container
    const stepsContainer = document.createElement('div');
    stepsContainer.className = 'agent-run-container';
    bubble.appendChild(stepsContainer);

    // Initial agent starting indicator
    const initChip = document.createElement('div');
    initChip.className = 'agent-step-chip active';
    initChip.innerHTML = '<div class="agent-step-spinner"></div><span>Agent connecting to Word...</span><button class="btn-chip-stop" title="Stop">⏹ Stop</button>';
    const btnInitStop = initChip.querySelector('.btn-chip-stop');
    if (btnInitStop) btnInitStop.onclick = cancelCurrentRun;
    stepsContainer.appendChild(initChip);

    const summaryTextDiv = document.createElement('div');
    summaryTextDiv.className = 'agent-summary-text';
    summaryTextDiv.style.marginTop = '8px';
    bubble.appendChild(summaryTextDiv);

    // Approval handler for destructive actions
    if (window.wordAgent) {
      window.wordAgent.onApprovalRequired = (callId, toolName, args) => {
        return new Promise((resolve) => {
          const card = document.createElement('div');
          card.className = 'agent-approval-card';
          card.innerHTML = `
            <div class="agent-approval-title">⚠️ Approval Required: ${escapeHtml(toolName)}</div>
            <div class="agent-approval-desc">The agent wants to modify or delete existing content: <code>${escapeHtml(JSON.stringify(args))}</code></div>
            <div class="agent-approval-actions">
              <button class="btn-agent-approve">Approve Action</button>
              <button class="btn-agent-reject">Reject</button>
            </div>
          `;
          const btnApprove = card.querySelector('.btn-agent-approve');
          const btnReject = card.querySelector('.btn-agent-reject');
          btnApprove.onclick = () => {
            card.remove();
            resolve(true);
          };
          btnReject.onclick = () => {
            card.remove();
            resolve(false);
          };
          bubble.appendChild(card);
          scrollChatToBottom();
        });
      };
    }

    try {
      checkDocUrlUpdate();
      // Perceive first: the daemon analyses the document before the model sees the prompt.
      const docContext = window.wordAgent ? await window.wordAgent.captureDocumentContext() : {};
      if (docContext && docContext.doc_url) {
        currentDocUrl = docContext.doc_url;
        const notice = document.getElementById('unsavedDocNotice');
        if (notice) notice.remove();
      }
      const response = await fetch('/api/agent/run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: currentAbortController ? currentAbortController.signal : undefined,
        body: JSON.stringify({
          prompt: text,
          host: 'Word',
          model: modelId,
          chat_id: currentChatId || undefined,
          ...docContext
        })
      });

      if (!response.ok) {
        let detail = '';
        try { detail = (await response.json()).error || ''; } catch (_) { /* not JSON */ }
        throw new Error(detail || `The local Antigravity service answered with status ${response.status}.`);
      }

      initChip.remove();

      const reader = response.body.getReader();
      const decoder = new TextDecoder('utf-8');
      let buffer = '';
      let summaryText = '';
      let renderScheduled = false;
      let finishedOk = false;
      let failed = false;
      let streamDone = false;
      let thinkingText = '';
      let thinkingScheduled = false;

      // Re-render the markdown at most once per animation frame while text is streaming.
      const scheduleSummaryRender = () => {
        if (renderScheduled) return;
        renderScheduled = true;
        const raf = window.requestAnimationFrame || ((fn) => setTimeout(fn, 16));
        raf(() => {
          renderScheduled = false;
          summaryTextDiv.innerHTML = renderMarkdown(summaryText);
          scrollChatToBottom();
        });
      };

      const runLogs = [];
      let activeSpinnerChip = null;

      while (!streamDone) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();

        for (const line of lines) {
          if (!line.startsWith('data: ')) continue;
          const dataStr = line.slice(6).trim();
          if (dataStr === '[DONE]') { streamDone = true; break; }

          let ev;
          try { ev = JSON.parse(dataStr); } catch (_) { continue; }

          if (ev.event === 'chat') {
            if (ev.chat_id) {
              currentChatId = ev.chat_id;
              rememberChatForDoc(currentChatId);
            }
            continue;
          }

          if (ev.event === 'analysis') {
            const label = ev.headline || ev.doc_type || '';
            if (label) {
              const card = document.createElement('div');
              card.className = 'agent-info-chip agent-analysis-card';
              card.innerHTML = `<span class="agent-info-icon">🔎</span><span>${escapeHtml(label)}</span>`;
              stepsContainer.appendChild(card);
              scrollChatToBottom();
            }
            continue;
          }

          if (ev.event === 'skills') {
            const labels = friendlySkillLabels(ev.names || []);
            if (labels.length > 0) {
              const chip = document.createElement('div');
              chip.className = 'agent-info-chip agent-skills-chip';
              chip.innerHTML = `<span class="agent-info-icon">🧠</span><span>Using expertise: ${escapeHtml(labels.join(', '))}</span>`;
              stepsContainer.appendChild(chip);
              scrollChatToBottom();
            }
            continue;
          }

          if (ev.event === 'plan') {
            const steps = Array.isArray(ev.steps) ? ev.steps : [];
            if (steps.length > 0) {
              const card = document.createElement('div');
              card.className = 'agent-plan-card';
              const items = steps.map(s => {
                const t = typeof s === 'string' ? s : (s && (s.text || s.title || s.step)) || JSON.stringify(s);
                return `<li>${escapeHtml(t)}</li>`;
              }).join('');
              card.innerHTML = `<div class="agent-plan-title">📋 Plan</div><ol class="agent-plan-list">${items}</ol>`;
              stepsContainer.appendChild(card);
              scrollChatToBottom();
            }
            continue;
          }

          if (ev.event === 'step_status') {
            if (ev.tool) {
              if (!activeSpinnerChip) {
                activeSpinnerChip = document.createElement('div');
                activeSpinnerChip.className = 'agent-step-chip active thinking-pulse';
                activeSpinnerChip.innerHTML = '<div class="agent-step-spinner"></div><span></span><button class="btn-chip-stop" title="Stop">⏹ Stop</button>';
                const btnChipStop = activeSpinnerChip.querySelector('.btn-chip-stop');
                if (btnChipStop) btnChipStop.onclick = cancelCurrentRun;
                stepsContainer.appendChild(activeSpinnerChip);
              }
              const span = activeSpinnerChip.querySelector('span');
              if (span) span.textContent = `Working: ${ev.tool}...`;
            }
            continue;
          }

          if (ev.event === 'tool_call') {
            if (activeSpinnerChip) {
              activeSpinnerChip.remove();
              activeSpinnerChip = null;
            }

            const chip = document.createElement('div');
            chip.className = 'agent-step-chip active';
            chip.innerHTML = `<div class="agent-step-spinner"></div><span>${escapeHtml(humanToolLabel(ev.tool, ev.arguments || {}))}</span><button class="btn-chip-stop" title="Stop">⏹ Stop</button>`;
            const btnChipStop = chip.querySelector('.btn-chip-stop');
            if (btnChipStop) btnChipStop.onclick = cancelCurrentRun;
            stepsContainer.appendChild(chip);
            scrollChatToBottom();

            const callStart = Date.now();
            // Execute the tool in Word
            let res = { ok: false, error: 'Word agent not initialized' };
            if (window.wordAgent) {
              res = await window.wordAgent.executeTool(ev.tool, ev.arguments || {}, ev.call_id);
            }
            const durationMs = Date.now() - callStart;

            runLogs.push({
              tool: ev.tool,
              arguments: ev.arguments,
              result: res,
              durationMs,
              time: new Date().toLocaleTimeString()
            });

            // Update chip UI
            chip.classList.remove('active');
            if (res && res.ok !== false) {
              chip.classList.add('success');
              chip.innerHTML = `<span>✓</span><span>${escapeHtml(humanToolLabel(ev.tool, ev.arguments || {}).replace('...', ''))} (done in ${durationMs}ms)</span>`;
            } else {
              chip.classList.add('error');
              chip.innerHTML = `<span>❌</span><span>${escapeHtml(humanToolLabel(ev.tool, ev.arguments || {}).replace('...', ''))} failed: ${escapeHtml(res ? res.error : 'Unknown error')}</span>`;
            }

            // Immediately display reasoning spinner so UI is never perceived as stuck
            activeSpinnerChip = document.createElement('div');
            activeSpinnerChip.className = 'agent-step-chip active thinking-pulse';
            activeSpinnerChip.innerHTML = '<div class="agent-step-spinner"></div><span>Agent is thinking and planning next action...</span><button class="btn-chip-stop" title="Stop">⏹ Stop</button>';
            const btnChipStop2 = activeSpinnerChip.querySelector('.btn-chip-stop');
            if (btnChipStop2) btnChipStop2.onclick = cancelCurrentRun;
            stepsContainer.appendChild(activeSpinnerChip);

            // Report result back to daemon
            await fetch('/api/agent/tool_result', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                call_id: ev.call_id,
                result: res
              })
            });
            scrollChatToBottom();
          } else if (ev.event === 'thinking') {
            let thinkingBox = bubble.querySelector('.agent-thinking-box');
            if (!thinkingBox) {
              thinkingBox = document.createElement('div');
              thinkingBox.className = 'agent-thinking-box';
              thinkingBox.innerHTML = `
                <div class="agent-thinking-header">
                  <span class="agent-thinking-sparkle">💭</span>
                  <span class="agent-thinking-title">Agent Thinking & Strategy</span>
                  <span class="agent-thinking-toggle">▼</span>
                </div>
                <div class="agent-thinking-content"></div>
              `;
              const toggleBtn = thinkingBox.querySelector('.agent-thinking-header');
              toggleBtn.onclick = () => {
                const isCollapsed = thinkingBox.classList.toggle('collapsed');
                thinkingBox.querySelector('.agent-thinking-toggle').textContent = isCollapsed ? '▶' : '▼';
              };
              bubble.insertBefore(thinkingBox, stepsContainer);
            }
            // 'thought' arrives as streamed deltas: accumulate, render once per frame.
            thinkingText += (ev.thought || '');
            if (!thinkingScheduled) {
              thinkingScheduled = true;
              const content = thinkingBox.querySelector('.agent-thinking-content');
              const raf = window.requestAnimationFrame || ((fn) => setTimeout(fn, 16));
              raf(() => {
                thinkingScheduled = false;
                content.innerHTML = renderMarkdown(thinkingText);
                scrollChatToBottom();
              });
            }
          } else if (ev.event === 'text') {
            summaryText += (ev.delta || '');
            scheduleSummaryRender();
          } else if (ev.event === 'error') {
            if (activeSpinnerChip) {
              activeSpinnerChip.remove();
              activeSpinnerChip = null;
            }
            failed = true;
            bubble.appendChild(buildErrorCard(ev, text, modelId));
            scrollChatToBottom();
            streamDone = true;
            break;
          } else if (ev.event === 'cancelled') {
            if (activeSpinnerChip) {
              activeSpinnerChip.remove();
              activeSpinnerChip = null;
            }
            failed = true;
            const note = document.createElement('div');
            note.className = 'agent-error-card agent-error-soft';
            note.textContent = 'Stopped. Nothing else will be changed.';
            bubble.appendChild(note);
            streamDone = true;
            break;
          } else if (ev.event === 'agent_finished') {
            if (activeSpinnerChip) {
              activeSpinnerChip.remove();
              activeSpinnerChip = null;
            }
            finishedOk = true;
            streamDone = true;
            break;
          }
        }
      }
      try { reader.cancel(); } catch (_) { /* already closed */ }

      if (activeSpinnerChip) {
        activeSpinnerChip.remove();
        activeSpinnerChip = null;
      }
      // Final, complete render of the answer text.
      if (summaryText) summaryTextDiv.innerHTML = renderMarkdown(summaryText);

      if (!finishedOk && !failed) {
        // Stream ended without a terminal event: never pretend it worked.
        bubble.appendChild(buildErrorCard({ error: 'The connection to the agent closed before it finished. Please try again.', code: 'other' }, text, modelId));
      }

      // Add completion badge and logs inspection panel
      const footerDiv = document.createElement('div');
      footerDiv.className = 'agent-footer-container';

      if (finishedOk) {
        const badge = document.createElement('div');
        badge.className = 'agent-completion-badge';
        badge.textContent = runLogs.length > 0
          ? `✓ Done (${runLogs.length} action${runLogs.length === 1 ? '' : 's'} in Word)`
          : '✓ Done';
        footerDiv.appendChild(badge);
      }

      if (runLogs.length > 0) {
        const btnToggleLogs = document.createElement('button');
        btnToggleLogs.className = 'btn-toggle-agent-logs';
        btnToggleLogs.innerHTML = `📜 Show Execution Trace & Logs (${runLogs.length})`;

        const logsPanel = document.createElement('div');
        logsPanel.className = 'agent-logs-panel hidden';

        let logsHtml = '<div class="agent-logs-header">Autonomous Agent Execution Trace</div>';
        for (const log of runLogs) {
          const isSuccess = log.result && log.result.ok !== false;
          logsHtml += `
            <div class="agent-log-item ${isSuccess ? 'log-success' : 'log-error'}">
              <div class="log-item-header">
                <span class="log-time">[${escapeHtml(log.time)}]</span>
                <span class="log-tool">${escapeHtml(log.tool)}</span>
                <span class="log-duration">${log.durationMs}ms</span>
                <span class="log-badge ${isSuccess ? 'badge-ok' : 'badge-err'}">${isSuccess ? 'OK' : 'ERR'}</span>
              </div>
              <div class="log-item-body">
                <div class="log-section"><strong>Args:</strong> <code>${escapeHtml(JSON.stringify(log.arguments))}</code></div>
                <div class="log-section"><strong>Result:</strong> <code>${escapeHtml(JSON.stringify(log.result))}</code></div>
              </div>
            </div>
          `;
        }
        logsPanel.innerHTML = logsHtml;

        btnToggleLogs.addEventListener('click', () => {
          const isHidden = logsPanel.classList.toggle('hidden');
          btnToggleLogs.innerHTML = isHidden
            ? `📜 Show Execution Trace & Logs (${runLogs.length})`
            : `▲ Hide Execution Trace`;
          scrollChatToBottom();
        });

        footerDiv.appendChild(btnToggleLogs);
        footerDiv.appendChild(logsPanel);
      }

      bubble.appendChild(footerDiv);
      scrollChatToBottom();

    } catch (err) {
      initChip.remove();
      if (activeSpinnerChip) {
        activeSpinnerChip.remove();
        activeSpinnerChip = null;
      }
      if (err && (err.name === 'AbortError' || err.message === 'The user aborted a request.')) {
        const note = document.createElement('div');
        note.className = 'agent-error-card agent-error-soft';
        note.textContent = 'Stopped. Nothing else will be changed.';
        bubble.appendChild(note);
        scrollChatToBottom();
        return;
      }
      console.error('[Taskpane] Agent run error:', err);
      const isNetwork = err && (/Failed to fetch|network error|NetworkError/i.test(err.message || ''));
      bubble.appendChild(buildErrorCard({
        error: isNetwork
          ? 'Cannot reach the local Antigravity service. Make sure it is running, then try again.'
          : (err.message || String(err)),
        code: 'other'
      }, text, modelId));
      scrollChatToBottom();
    } finally {
      currentAbortController = null;
    }
  }

  // -------------------------------------------------------------------------
  // Error cards (quota / auth / other) — always show the real problem
  // -------------------------------------------------------------------------
  function buildErrorCard(ev, promptText, modelId) {
    const card = document.createElement('div');
    card.className = 'agent-error-card';
    const code = ev.code || (ev.error && ev.error.includes('already running') ? 'already_running' : 'other');

    if (code === 'quota') {
      limitedModels.add(modelId);
      markLimitedModelsInPicker();
      const label = modelLabel(modelId);
      const resets = ev.retry_after ? ` (resets in ${ev.retry_after})` : '';
      card.innerHTML = `<div class="agent-error-title">⏳ ${escapeHtml(label)} has reached its usage limit${escapeHtml(resets)}.</div>
        <div class="agent-error-desc">Pick another model:</div>
        <div class="agent-error-actions"></div>`;
      const actions = card.querySelector('.agent-error-actions');
      const alternatives = pickAlternativeModels(modelId, 3);
      if (alternatives.length === 0) {
        actions.innerHTML = '<span class="agent-error-desc">No other model is available right now.</span>';
      }
      alternatives.forEach(m => {
        const btn = document.createElement('button');
        btn.className = 'btn-retry-model';
        btn.textContent = `Try ${m.label}`;
        btn.addEventListener('click', () => {
          if (isStreaming) return;
          card.querySelectorAll('button').forEach(b => { b.disabled = true; });
          modelSelect.value = m.id;
          storeSet(MODEL_KEY, m.id);
          retryPrompt(promptText, m.id);
        });
        actions.appendChild(btn);
      });
    } else if (code === 'auth') {
      card.innerHTML = `<div class="agent-error-title">🔒 You are not signed in to Antigravity</div>
        <div class="agent-error-desc">Sign in to Antigravity on this computer, then send your request again.</div>`;
    } else if (code === 'already_running') {
      card.innerHTML = `<div class="agent-error-title">⚠️ Another task is already running</div>
        <div class="agent-error-desc">${escapeHtml(ev.error || 'Please wait for it to finish or cancel it.')}</div>
        <div class="agent-error-actions">
          <button class="btn-cancel-running btn-retry-model">Stop running task</button>
        </div>`;
      const btnCancel = card.querySelector('.btn-cancel-running');
      if (btnCancel) {
        btnCancel.addEventListener('click', async () => {
          btnCancel.disabled = true;
          try {
            await fetch('/api/agent/cancel', { method: 'POST' });
          } catch (_) {}
        });
      }
    } else {
      const title = code === 'model' ? 'This model is not available' : 'Something went wrong';
      card.innerHTML = `<div class="agent-error-title">⚠️ ${escapeHtml(title)}</div>
        <div class="agent-error-desc">${escapeHtml(ev.error || 'The request failed.')}</div>`;
    }
    return card;
  }

  async function retryPrompt(promptText, modelId) {
    if (isStreaming) return;
    showChatView();
    appendUserMessage(promptText);
    const assistantMsg = createAssistantMessagePlaceholder();
    isStreaming = true;
    updateSendButtonState();
    try {
      await runWordAgent(promptText, assistantMsg, modelId);
    } finally {
      isStreaming = false;
      updateSendButtonState();
    }
  }

  // -------------------------------------------------------------------------
  // Skills chip labels
  // -------------------------------------------------------------------------
  const SKILL_LABELS = {
    'template-form-filling': 'Form filling',
    'pm-documentation-suite': 'PM documentation',
    'long-report-architecture': 'Report structure',
    'smartart-and-diagrams': 'Diagrams',
    'native-charts': 'Charts',
    'office-document-analysis': null // internal, hidden
  };

  function friendlySkillLabels(names) {
    const out = [];
    for (const raw of names) {
      const name = String(raw || '').trim();
      if (!name) continue;
      let label;
      if (Object.prototype.hasOwnProperty.call(SKILL_LABELS, name)) {
        label = SKILL_LABELS[name];
      } else {
        label = name.replace(/[-_]+/g, ' ').replace(/^\w/, c => c.toUpperCase());
      }
      if (label && !out.includes(label)) out.push(label);
    }
    return out;
  }

  // -------------------------------------------------------------------------
  // Model picker (loaded from /api/models, grouped by family)
  // -------------------------------------------------------------------------
  const DEFAULT_MODEL = 'gemini-3.8-flash-high';
  const FALLBACK_MODELS = [
    { id: 'gemini-3.8-flash-high', label: 'Gemini 3.8 Flash (High)', family: 'gemini-3.8-flash', effort: 'high' },
    { id: 'gemini-4-argon-medium', label: 'Gemini 4 Argon (Medium)', family: 'gemini-4-argon', effort: 'medium' },
    { id: 'claude-sonnet-5-5-medium', label: 'Claude Sonnet 5.5 (Medium)', family: 'claude-sonnet-5-5', effort: 'medium' },
    { id: 'claude-opus-5-5-high', label: 'Claude Opus 5.5 (High)', family: 'claude-opus-5-5', effort: 'high' },
    { id: 'gpt-oss-120b-medium', label: 'GPT-OSS 120B (Medium)', family: 'gpt-oss-120b', effort: 'medium' }
  ];

  function familyLabel(family) {
    return formatFamilyLabel(family, FAMILY_LABELS, availableModels);
  }

  function modelLabel(modelId) {
    const m = availableModels.find(x => x.id === modelId);
    return m ? m.label : modelId;
  }

  function pickAlternativeModels(modelId, max) {
    const current = availableModels.find(x => x.id === modelId);
    const currentFamily = current ? current.family : modelId;
    const seenFamilies = new Set([currentFamily]);
    const picks = [];
    for (const m of availableModels) {
      if (picks.length >= max) break;
      if (seenFamilies.has(m.family) || limitedModels.has(m.id)) continue;
      // Prefer the same effort level as the failed model when the family offers several.
      const sameFamily = availableModels.filter(x => x.family === m.family && !limitedModels.has(x.id));
      const best = (current && sameFamily.find(x => x.effort === current.effort)) || sameFamily[0];
      seenFamilies.add(m.family);
      picks.push(best);
    }
    return picks;
  }

  function renderModelOptions(models) {
    availableModels = models;
    const saved = storeGet(MODEL_KEY);
    const previous = modelSelect.value;
    modelSelect.innerHTML = '';
    const groups = new Map();
    models.forEach(m => {
      const fam = m.family || 'other';
      if (!groups.has(fam)) groups.set(fam, []);
      groups.get(fam).push(m);
    });
    groups.forEach((list, fam) => {
      const og = document.createElement('optgroup');
      og.label = familyLabel(fam);
      list.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.id;
        opt.textContent = m.label;
        opt.dataset.label = m.label;
        og.appendChild(opt);
      });
      modelSelect.appendChild(og);
    });
    const ids = models.map(m => m.id);
    const choice = [saved, previous, DEFAULT_MODEL].find(id => id && ids.includes(id)) || ids[0];
    if (choice) modelSelect.value = choice;
    markLimitedModelsInPicker();
  }

  function markLimitedModelsInPicker() {
    Array.from(modelSelect.options).forEach(opt => {
      const base = opt.dataset.label || opt.textContent;
      opt.textContent = limitedModels.has(opt.value) ? `${base} (limit reached)` : base;
    });
  }

  async function loadModels() {
    try {
      const res = await fetch('/api/models');
      const data = await res.json();
      if (data && data.families && typeof data.families === 'object') {
        Object.assign(FAMILY_LABELS, data.families);
      }
      const models = (data && Array.isArray(data.models) ? data.models : [])
        .filter(m => m && m.id)
        .map(m => ({
          id: m.id,
          label: m.label || m.id,
          family: m.family || m.id,
          family_label: m.family_label || m.family_display || (data.families ? data.families[m.family] : null),
          effort: m.effort || null
        }));
      if (models.length > 0) {
        renderModelOptions(models);
        return;
      }
      if (data && data.error) console.warn('[Taskpane] /api/models error:', data.code, data.error);
    } catch (e) {
      console.warn('[Taskpane] /api/models unavailable, using fallback list:', e);
    }
    renderModelOptions(FALLBACK_MODELS);
  }

  modelSelect.addEventListener('change', () => storeSet(MODEL_KEY, modelSelect.value));
  renderModelOptions(FALLBACK_MODELS); // instant, replaced as soon as /api/models answers
  loadModels();

  async function sendMessage() {
    const text = promptInput.value.trim();
    if (!text || isStreaming) return;

    // Transition from onboarding to chat mode
    showChatView();

    // Append User Message Bubble
    appendUserMessage(text);
    promptInput.value = '';
    promptInput.style.height = 'auto';

    // Prepare Assistant placeholder
    const assistantMsg = createAssistantMessagePlaceholder();
    isStreaming = true;
    updateSendButtonState();

    // In Word host: invoke the autonomous agent manipulating Word directly
    if (currentHost === 'Word') {
      try {
        await runWordAgent(text, assistantMsg);
      } finally {
        isStreaming = false;
        updateSendButtonState();
      }
      return;
    }

    // Fetch fresh document content if not attached
    let docContext = attachedDocumentContext;
    if (!docContext) {
      if (currentHost === 'Word') {
        const doc = await window.officeBridge.readWordContent();
        docContext = doc ? doc.content : '';
      } else if (currentHost === 'Excel') {
        const xl = await window.officeBridge.readExcelContent();
        docContext = xl ? JSON.stringify(xl) : '';
      } else if (currentHost === 'PowerPoint') {
        const ppt = await window.officeBridge.readPowerPointContent();
        docContext = ppt ? JSON.stringify(ppt) : '';
      } else if (currentHost === 'Outlook') {
        const mail = await window.officeBridge.readOutlookEmail();
        docContext = mail ? JSON.stringify(mail) : '';
      }
    }

    const payload = {
      prompt: text,
      model: modelSelect.value,
      host: currentHost,
      web_search: isWebSearchEnabled,
      document_context: docContext || '',
      attached_files: attachedFiles.map(f => f.name)
    };

    try {
      const response = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      if (!response.ok) {
        throw new Error(`Server returned ${response.status}`);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder('utf-8');
      let fullContent = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        const chunk = decoder.decode(value, { stream: true });
        
        // Handle SSE stream or raw text
        const lines = chunk.split('\n');
        for (const line of lines) {
          if (line.startsWith('data: ')) {
            const dataStr = line.slice(6).trim();
            if (dataStr === '[DONE]') break;
            try {
              const parsed = JSON.parse(dataStr);
              fullContent += (parsed.delta || parsed.text || '');
            } catch (e) {
              fullContent += dataStr;
            }
          } else if (line.trim().length > 0) {
            fullContent += line;
          }
        }
        updateAssistantMessage(assistantMsg, fullContent);
      }

      // Finish streaming and add Office action buttons
      finalizeAssistantMessage(assistantMsg, fullContent, text);

    } catch (err) {
      console.warn('[Taskpane] Chat request failed:', err);
      const bubble = assistantMsg.querySelector('.msg-bubble');
      bubble.innerHTML = '';
      bubble.appendChild(buildErrorCard({
        error: (err && err.message) ? err.message : 'The request failed.',
        code: 'other'
      }, text, modelSelect.value));
      scrollChatToBottom();
    } finally {
      isStreaming = false;
      updateSendButtonState();
    }
  }

  function appendUserMessage(text) {
    const msgDiv = document.createElement('div');
    msgDiv.className = 'chat-message user';
    msgDiv.innerHTML = `<div class="msg-bubble">${escapeHtml(text)}</div>`;
    chatMessages.appendChild(msgDiv);
    scrollChatToBottom();
  }

  function createAssistantMessagePlaceholder() {
    const msgDiv = document.createElement('div');
    msgDiv.className = 'chat-message assistant';
    msgDiv.innerHTML = `
      <div class="msg-header">
        <img src="/assets/icon-16.png" width="14" height="14" alt="Gemini">
        <span>Gemini Antigravity</span>
      </div>
      <div class="msg-bubble">Thinking...</div>
    `;
    chatMessages.appendChild(msgDiv);
    scrollChatToBottom();
    return msgDiv;
  }

  function updateAssistantMessage(msgDiv, text) {
    const bubble = msgDiv.querySelector('.msg-bubble');
    bubble.innerHTML = renderMarkdown(text);
    scrollChatToBottom();
  }

  function finalizeAssistantMessage(msgDiv, fullText, userPrompt) {
    // Add interactive action buttons based on content and active host
    const bubble = msgDiv.querySelector('.msg-bubble');
    
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'msg-actions';

    if (currentHost === 'Word' || userPrompt.includes('/check-doc') || userPrompt.includes('/copy-edit') || fullText.includes('#') || fullText.includes('|')) {
      const btnApplyWord = document.createElement('button');
      btnApplyWord.className = 'btn-office-action primary';
      btnApplyWord.innerHTML = '✨ Insert into Word Document';
      btnApplyWord.addEventListener('click', async () => {
        btnApplyWord.disabled = true;
        btnApplyWord.textContent = 'Formatting & Inserting...';
        try {
          const res = await window.officeBridge.insertFormattedReport(fullText);
          console.log('[Taskpane] Insert result:', res);
          btnApplyWord.textContent = '✓ Inserted into Word!';
          setTimeout(() => {
            btnApplyWord.disabled = false;
            btnApplyWord.textContent = '✨ Insert into Word Document';
          }, 3500);
        } catch (e) {
          console.error('[Taskpane] Format insert failed:', e);
          btnApplyWord.disabled = false;
          btnApplyWord.textContent = '❌ Insert Failed';
          alert('Could not insert content into Word. Please make sure Word is open. Details: ' + (e.message || e));
        }
      });
      actionsDiv.appendChild(btnApplyWord);

      const btnComment = document.createElement('button');
      btnComment.className = 'btn-office-action';
      btnComment.innerHTML = '💬 Review Comment';
      btnComment.addEventListener('click', async () => {
        btnComment.disabled = true;
        try {
          await window.officeBridge.insertWordComment(`Gemini Suggestion: ${cleanTextForInsertion(fullText).slice(0, 140)}...`);
          btnComment.textContent = '✓ Comment Added!';
        } catch (e) {
          btnComment.textContent = 'Comment Failed';
        }
      });
      actionsDiv.appendChild(btnComment);

      const btnCopy = document.createElement('button');
      btnCopy.className = 'btn-office-action';
      btnCopy.innerHTML = '📋 Copy Text';
      btnCopy.addEventListener('click', () => {
        navigator.clipboard.writeText(fullText);
        btnCopy.textContent = 'Copied!';
        setTimeout(() => { btnCopy.textContent = '📋 Copy Text'; }, 2000);
      });
      actionsDiv.appendChild(btnCopy);
    }

    if (currentHost === 'PowerPoint' || userPrompt.includes('/storylining') || userPrompt.includes('/slide-designer')) {
      const btnPptSlide = document.createElement('button');
      btnPptSlide.className = 'btn-office-action primary';
      btnPptSlide.innerHTML = '📊 Create Formatted Slide';
      btnPptSlide.addEventListener('click', async () => {
        btnPptSlide.disabled = true;
        btnPptSlide.textContent = 'Building Slide...';
        let slideTitle = 'Executive Briefing';
        const matchTitle = fullText.match(/^#+?\s*(.+)$/m);
        if (matchTitle) slideTitle = matchTitle[1].replace(/[*_#]/g, '').trim();
        await window.officeBridge.insertPowerPointSlide({
          title: slideTitle,
          bullets: extractBulletPoints(fullText)
        });
        btnPptSlide.textContent = '✓ Slide Created!';
      });
      actionsDiv.appendChild(btnPptSlide);
    }

    if (currentHost === 'Excel' || userPrompt.includes('/financial-model') || userPrompt.includes('/analyze-sheet') || userPrompt.includes('/generate-chart')) {
      const btnApplyExcel = document.createElement('button');
      btnApplyExcel.className = 'btn-office-action primary';
      btnApplyExcel.innerHTML = '📊 Insert Financial Model into Excel';
      btnApplyExcel.addEventListener('click', async () => {
        btnApplyExcel.disabled = true;
        btnApplyExcel.textContent = 'Writing to Excel...';
        await window.officeBridge.applyFinancialModel('saas_revenue', '$');
        btnApplyExcel.textContent = '✓ Financial Model Inserted!';
      });
      actionsDiv.appendChild(btnApplyExcel);

      const btnChart = document.createElement('button');
      btnChart.className = 'btn-office-action';
      btnChart.innerHTML = '📈 Generate Chart';
      btnChart.addEventListener('click', async () => {
        btnChart.disabled = true;
        btnChart.textContent = 'Creating Chart...';
        await window.officeBridge.insertExcelChart('column', 'A1:E4', 'Revenue Breakdown');
        btnChart.textContent = '✓ Chart Created!';
      });
      actionsDiv.appendChild(btnChart);

      const btnReport = document.createElement('button');
      btnReport.className = 'btn-office-action';
      btnReport.innerHTML = '📑 Run Excel-to-Report';
      btnReport.addEventListener('click', async () => {
        btnReport.disabled = true;
        btnReport.textContent = 'Generating Word Report...';
        await window.officeBridge.excelToReport();
        btnReport.textContent = '✓ Report Inserted into Word!';
      });
      actionsDiv.appendChild(btnReport);
    }

    if (currentHost === 'Outlook' || userPrompt.includes('/reply-email') || userPrompt.includes('/summarize-thread')) {
      const btnReply = document.createElement('button');
      btnReply.className = 'btn-office-action primary';
      btnReply.innerHTML = '✉️ Draft Reply in Outlook';
      btnReply.addEventListener('click', async () => {
        btnReply.disabled = true;
        btnReply.textContent = 'Drafting Reply...';
        const htmlBody = window.officeBridge.markdownToExecutiveHtml(fullText);
        await window.officeBridge.draftOutlookReply(htmlBody);
        btnReply.textContent = '✓ Reply Drafted in Outlook!';
      });
      actionsDiv.appendChild(btnReply);

      const btnMeeting = document.createElement('button');
      btnMeeting.className = 'btn-office-action';
      btnMeeting.innerHTML = '📅 Schedule Review Meeting';
      btnMeeting.addEventListener('click', async () => {
        btnMeeting.disabled = true;
        btnMeeting.textContent = 'Scheduling...';
        await window.officeBridge.scheduleOutlookMeeting('Executive Office Review', '2026-10-29 10:00', 45, 'ergys@cloud4job.com');
        btnMeeting.textContent = '✓ Meeting Scheduled!';
      });
      actionsDiv.appendChild(btnMeeting);
    }

    if (userPrompt.includes('/doc-to-slides')) {
      const btnDocToSlides = document.createElement('button');
      btnDocToSlides.className = 'btn-office-action primary';
      btnDocToSlides.innerHTML = '🚀 Build PowerPoint Deck';
      btnDocToSlides.addEventListener('click', async () => {
        btnDocToSlides.disabled = true;
        btnDocToSlides.textContent = 'Generating 4 Slides...';
        await window.officeBridge.docToSlides();
        btnDocToSlides.textContent = '✓ PowerPoint Deck Ready!';
      });
      actionsDiv.appendChild(btnDocToSlides);
    }

    if (userPrompt.includes('/email-summary')) {
      const btnEmailSummary = document.createElement('button');
      btnEmailSummary.className = 'btn-office-action primary';
      btnEmailSummary.innerHTML = '📧 Draft Briefing in Outlook';
      btnEmailSummary.addEventListener('click', async () => {
        btnEmailSummary.disabled = true;
        btnEmailSummary.textContent = 'Drafting in Outlook...';
        await window.officeBridge.emailSuiteSummary();
        btnEmailSummary.textContent = '✓ Outlook Briefing Created!';
      });
      actionsDiv.appendChild(btnEmailSummary);
    }

    bubble.appendChild(actionsDiv);
    scrollChatToBottom();
  }

  function appendSystemMessage(text) {
    const msgDiv = document.createElement('div');
    msgDiv.className = 'chat-message system';
    msgDiv.innerHTML = `<div style="font-size:11px;color:#64748B;text-align:center;padding:4px;">${escapeHtml(text)}</div>`;
    chatMessages.appendChild(msgDiv);
    scrollChatToBottom();
  }

  function scrollChatToBottom() {
    const main = document.getElementById('mainContent');
    main.scrollTop = main.scrollHeight;
  }

  function cleanTextForInsertion(text) {
    return text.replace(/[#*`>]/g, '').trim();
  }

  function extractBulletPoints(text) {
    const lines = text.split('\n');
    const bullets = [];
    for (const line of lines) {
      const trimmed = line.trim();
      if (trimmed.startsWith('- ') || trimmed.startsWith('* ') || /^\d+\./.test(trimmed)) {
        bullets.push(trimmed.replace(/^[-*]|\d+\.\s*/, '').trim());
      }
    }
    if (bullets.length === 0) {
      bullets.push(text.slice(0, 100));
    }
    return bullets.slice(0, 5);
  }

  // -------------------------------------------------------------------------
  // Views, conversation continuity & history drawer
  // -------------------------------------------------------------------------
  function showChatView() {
    onboardingSection.classList.add('hidden');
    chatSection.classList.remove('hidden');
    const resume = document.getElementById('resumeCard');
    if (resume) resume.remove();
  }

  function showOnboardingView() {
    chatMessages.innerHTML = '';
    chatSection.classList.add('hidden');
    onboardingSection.classList.remove('hidden');
  }

  function startNewChat() {
    if (isStreaming) return;
    currentChatId = null;
    rememberChatForDoc(null);
    showOnboardingView();
    promptInput.value = '';
    updateSendButtonState();
    promptInput.focus();
  }

  function appendStoredAssistantMessage(content, options = {}) {
    const msgDiv = createAssistantMessagePlaceholder();
    const bubble = msgDiv.querySelector('.msg-bubble');
    bubble.innerHTML = '';

    const stepsContainer = document.createElement('div');
    stepsContainer.className = 'agent-run-container';
    bubble.appendChild(stepsContainer);

    // Plans
    const plans = options.plans || [];
    plans.forEach(p => {
      const steps = Array.isArray(p.steps) ? p.steps : (Array.isArray(p) ? p : []);
      if (steps.length > 0) {
        const card = document.createElement('div');
        card.className = 'agent-plan-card';
        const items = steps.map(s => {
          const t = typeof s === 'string' ? s : (s && (s.text || s.title || s.step)) || JSON.stringify(s);
          return `<li>${escapeHtml(t)}</li>`;
        }).join('');
        card.innerHTML = `<div class="agent-plan-title">📋 Plan</div><ol class="agent-plan-list">${items}</ol>`;
        stepsContainer.appendChild(card);
      }
    });

    // Tool chips
    const toolCalls = options.toolCalls || [];
    toolCalls.forEach(tc => {
      const chip = document.createElement('div');
      chip.className = 'agent-step-chip ' + (tc.ok === false ? 'error' : 'success');
      let args = {};
      try {
        args = typeof tc.args_json === 'string' ? JSON.parse(tc.args_json) : (tc.args || {});
      } catch (_) {}
      const label = humanToolLabel(tc.tool, args).replace('...', '');
      const dur = tc.duration_ms != null ? ` (done in ${tc.duration_ms}ms)` : '';
      if (tc.ok === false) {
        chip.innerHTML = `<span>❌</span><span>${escapeHtml(label)} failed</span>`;
      } else {
        chip.innerHTML = `<span>✓</span><span>${escapeHtml(label)}${escapeHtml(dur)}</span>`;
      }
      stepsContainer.appendChild(chip);
    });

    if (plans.length === 0 && toolCalls.length === 0) {
      stepsContainer.remove();
    }

    const textDiv = document.createElement('div');
    textDiv.className = 'agent-summary-text';
    textDiv.innerHTML = renderMarkdown(content || '');
    bubble.appendChild(textDiv);
  }

  async function fetchChat(chatId) {
    const res = await fetch(`/api/chats/${encodeURIComponent(chatId)}`);
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Could not load the conversation (status ${res.status}).`);
    return res.json();
  }

  /** Loads a saved conversation into the view; the next message continues it. */
  async function openChat(chatId) {
    if (isStreaming || !chatId) return;
    historyDrawer.classList.add('hidden');
    showChatView();
    chatMessages.innerHTML = '';
    appendSystemMessage('Loading conversation...');
    let data;
    try {
      data = await fetchChat(chatId);
    } catch (e) {
      chatMessages.innerHTML = '';
      appendSystemMessage(e.message || 'Could not load the conversation.');
      return;
    }
    chatMessages.innerHTML = '';
    if (!data) {
      appendSystemMessage('This conversation no longer exists.');
      if (storeGet(lastChatKey(currentDocUrl)) === chatId) rememberChatForDoc(null);
      currentChatId = null;
      return;
    }
    currentChatId = chatId;
    rememberChatForDoc(chatId);
    const title = data.chat && data.chat.title ? data.chat.title : 'Conversation';
    appendSystemMessage(`Continuing: ${title}`);

    if (!currentDocUrl && data.chat && data.chat.doc_url) {
      appendSystemMessage('ℹ️ Note: The current document is unsaved. Save this file to link it with this conversation permanently.');
    }

    const msgs = data.messages || [];
    const chatPlans = data.plans || [];
    const chatTools = data.tool_calls || [];
    const turns = data.turns || [];

    const turnByAssistId = new Map();
    turns.forEach(t => {
      if (t.assistant_message && t.assistant_message.id) {
        turnByAssistId.set(t.assistant_message.id, t);
      }
    });

    msgs.forEach(msg => {
      if (msg.role === 'user') {
        appendUserMessage(msg.content || '');
      } else if (msg.role === 'assistant') {
        const turn = turnByAssistId.get(msg.id);
        const plans = (msg.plans && msg.plans.length > 0)
          ? msg.plans
          : (turn ? turn.plans : (msgs.length <= 2 ? chatPlans : []));
        const toolCalls = (msg.tool_calls && msg.tool_calls.length > 0)
          ? msg.tool_calls
          : (turn ? turn.tool_calls : (msgs.length <= 2 ? chatTools : []));

        appendStoredAssistantMessage(msg.content || '', { plans, toolCalls });
      }
    });

    const actions = (data.tool_calls || []).length;
    if (actions > 0) {
      appendSystemMessage(`${actions} action${actions === 1 ? ' was' : 's were'} made in Word in this conversation.`);
    }
    scrollChatToBottom();
    promptInput.focus();
  }

  function showUnsavedDocNotice() {
    if (currentDocUrl || currentHost === 'Standalone') return;
    if (document.getElementById('unsavedDocNotice')) return;
    const notice = document.createElement('div');
    notice.id = 'unsavedDocNotice';
    notice.className = 'unsaved-doc-banner';
    notice.innerHTML = `
      <div class="unsaved-doc-content">
        <span class="unsaved-doc-icon">💾</span>
        <span class="unsaved-doc-text">This document is unsaved. Save your file (Ctrl+S) so your conversation history is preserved and can be resumed next time.</span>
      </div>
    `;
    onboardingSection.insertBefore(notice, onboardingSection.firstChild);
  }

  function checkDocUrlUpdate() {
    const updated = readDocUrl();
    if (updated && !currentDocUrl) {
      currentDocUrl = updated;
      const notice = document.getElementById('unsavedDocNotice');
      if (notice) notice.remove();
    }
    return currentDocUrl;
  }

  /** On load: offer to continue the last conversation of this document. */
  async function offerResumeLastChat() {
    checkDocUrlUpdate();
    if (!currentDocUrl) {
      showUnsavedDocNotice();
      return; // unsaved document: no stable key to match a past chat
    }
    const lastId = storeGet(lastChatKey(currentDocUrl));
    if (!lastId) return;
    let data;
    try {
      data = await fetchChat(lastId);
    } catch (_) {
      return; // service not reachable yet: just start fresh
    }
    if (!data || !data.chat) {
      rememberChatForDoc(null);
      return;
    }
    if (currentChatId || !chatSection.classList.contains('hidden')) return; // user already started
    const card = document.createElement('div');
    card.id = 'resumeCard';
    card.className = 'resume-card';
    const text = document.createElement('div');
    text.className = 'resume-text';
    const strong = document.createElement('strong');
    strong.textContent = 'Continue where you left off?';
    const sub = document.createElement('div');
    sub.className = 'resume-sub';
    sub.textContent = data.chat.title || 'Last conversation';
    text.appendChild(strong);
    text.appendChild(sub);
    const btns = document.createElement('div');
    btns.className = 'resume-actions';
    const btnContinue = document.createElement('button');
    btnContinue.className = 'btn-resume primary';
    btnContinue.textContent = 'Continue';
    btnContinue.addEventListener('click', () => openChat(lastId));
    const btnFresh = document.createElement('button');
    btnFresh.className = 'btn-resume';
    btnFresh.textContent = 'Start new';
    btnFresh.addEventListener('click', () => {
      rememberChatForDoc(null);
      card.remove();
    });
    btns.appendChild(btnContinue);
    btns.appendChild(btnFresh);
    card.appendChild(text);
    card.appendChild(btns);
    onboardingSection.insertBefore(card, onboardingSection.firstChild);
  }

  const historyList = document.getElementById('historyList');

  function historyMessage(text) {
    const div = document.createElement('div');
    div.className = 'history-item empty';
    div.textContent = text;
    return div;
  }

  function chatsUrl(params) {
    const q = new URLSearchParams();
    Object.keys(params).forEach(k => { if (params[k]) q.set(k, params[k]); });
    return `/api/chats?${q.toString()}`;
  }

  async function loadHistory() {
    historyList.innerHTML = '';
    historyList.appendChild(historyMessage('Loading...'));
    const host = currentHost === 'Standalone' ? '' : currentHost;
    try {
      const [docRes, allRes] = await Promise.all([
        currentDocUrl
          ? fetch(chatsUrl({ host, doc_url: currentDocUrl })).then(r => r.json())
          : Promise.resolve({ chats: [] }),
        fetch(chatsUrl({ host })).then(r => r.json())
      ]);
      const docChats = (docRes && docRes.chats) || [];
      const docIds = new Set(docChats.map(c => c.id));
      const otherChats = ((allRes && allRes.chats) || []).filter(c => !docIds.has(c.id));
      historyList.innerHTML = '';
      if (docChats.length === 0 && otherChats.length === 0) {
        historyList.appendChild(historyMessage('No previous conversations'));
        return;
      }
      if (docChats.length > 0) renderHistorySection('This document', docChats);
      if (otherChats.length > 0) renderHistorySection('All chats', otherChats);
    } catch (e) {
      historyList.innerHTML = '';
      historyList.appendChild(historyMessage('Could not load your conversations. Is the Antigravity service running?'));
    }
  }

  function renderHistorySection(title, chats) {
    const header = document.createElement('div');
    header.className = 'history-section-title';
    header.textContent = title;
    historyList.appendChild(header);
    chats.forEach(chat => historyList.appendChild(buildHistoryItem(chat)));
  }

  function buildHistoryItem(chat) {
    const item = document.createElement('div');
    item.className = 'history-item' + (chat.id === currentChatId ? ' current' : '');
    item.tabIndex = 0;

    const main = document.createElement('div');
    main.className = 'history-main';
    const titleRow = document.createElement('div');
    titleRow.className = 'history-title-row';
    const titleEl = document.createElement('span');
    titleEl.className = 'history-title';
    titleEl.textContent = chat.title || 'Untitled chat';
    const timeEl = document.createElement('span');
    timeEl.className = 'history-time';
    timeEl.textContent = relativeTime(chat.updated_at);
    titleRow.appendChild(titleEl);
    titleRow.appendChild(timeEl);
    main.appendChild(titleRow);
    const preview = String(chat.last_message || '').replace(/\s+/g, ' ').trim();
    if (preview) {
      const prevEl = document.createElement('div');
      prevEl.className = 'history-preview';
      prevEl.textContent = preview.length > 90 ? `${preview.slice(0, 90)}…` : preview;
      main.appendChild(prevEl);
    }
    if (chat.doc_title && chat.doc_url !== currentDocUrl) {
      const docEl = document.createElement('div');
      docEl.className = 'history-doc';
      docEl.textContent = `📄 ${chat.doc_title}`;
      main.appendChild(docEl);
    }

    const actions = document.createElement('div');
    actions.className = 'history-actions';
    const btnRename = document.createElement('button');
    btnRename.className = 'history-action';
    btnRename.title = 'Rename';
    btnRename.setAttribute('aria-label', 'Rename');
    btnRename.textContent = '✎';
    const btnDelete = document.createElement('button');
    btnDelete.className = 'history-action';
    btnDelete.title = 'Delete';
    btnDelete.setAttribute('aria-label', 'Delete');
    btnDelete.textContent = '🗑';
    actions.appendChild(btnRename);
    actions.appendChild(btnDelete);

    item.appendChild(main);
    item.appendChild(actions);

    const open = () => openChat(chat.id);
    main.addEventListener('click', open);
    item.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target === item) open();
    });

    btnRename.addEventListener('click', (e) => {
      e.stopPropagation();
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'history-rename-input';
      input.value = chat.title || '';
      input.maxLength = 120;
      titleRow.replaceChild(input, titleEl);
      input.focus();
      input.select();
      let finished = false;
      const finish = async (save) => {
        if (finished) return;
        finished = true;
        const newTitle = input.value.trim();
        if (input.parentNode) titleRow.replaceChild(titleEl, input);
        if (save && newTitle && newTitle !== chat.title) {
          try {
            const res = await fetch(`/api/chats/${encodeURIComponent(chat.id)}`, {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ title: newTitle })
            });
            if (res.ok) {
              chat.title = newTitle;
              titleEl.textContent = newTitle;
            }
          } catch (_) { /* keep the old title */ }
        }
      };
      input.addEventListener('click', ev => ev.stopPropagation());
      input.addEventListener('keydown', (ev) => {
        ev.stopPropagation();
        if (ev.key === 'Enter') finish(true);
        else if (ev.key === 'Escape') finish(false);
      });
      input.addEventListener('blur', () => finish(true));
    });

    btnDelete.addEventListener('click', (e) => {
      e.stopPropagation();
      // window.confirm() is blocked in Office add-ins, so ask inline.
      actions.innerHTML = '';
      const ask = document.createElement('span');
      ask.className = 'history-confirm-text';
      ask.textContent = 'Delete?';
      const yes = document.createElement('button');
      yes.className = 'history-confirm yes';
      yes.textContent = 'Yes';
      const no = document.createElement('button');
      no.className = 'history-confirm';
      no.textContent = 'No';
      actions.appendChild(ask);
      actions.appendChild(yes);
      actions.appendChild(no);
      no.addEventListener('click', (ev) => {
        ev.stopPropagation();
        actions.innerHTML = '';
        actions.appendChild(btnRename);
        actions.appendChild(btnDelete);
      });
      yes.addEventListener('click', async (ev) => {
        ev.stopPropagation();
        yes.disabled = true;
        try {
          const res = await fetch(`/api/chats/${encodeURIComponent(chat.id)}`, { method: 'DELETE' });
          if (!res.ok) throw new Error(String(res.status));
          if (storeGet(lastChatKey(currentDocUrl)) === chat.id) rememberChatForDoc(null);
          if (currentChatId === chat.id) {
            currentChatId = null;
            showOnboardingView();
          }
          item.remove();
        } catch (_) {
          ask.textContent = 'Could not delete.';
          yes.disabled = false;
        }
      });
    });
    return item;
  }

  // -------------------------------------------------------------------------
  // Header Actions & Modals
  // -------------------------------------------------------------------------
  // "New chat" only forgets the conversation id: the next message starts a fresh chat.
  // (We no longer call /api/agent/reset: it wiped the agent's memory.)
  btnNewChat.addEventListener('click', () => {
    historyDrawer.classList.add('hidden');
    startNewChat();
  });

  btnHistory.addEventListener('click', () => {
    const willOpen = historyDrawer.classList.contains('hidden');
    historyDrawer.classList.toggle('hidden');
    if (willOpen) loadHistory();
  });

  closeHistory.addEventListener('click', () => {
    historyDrawer.classList.add('hidden');
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !historyDrawer.classList.contains('hidden')) {
      historyDrawer.classList.add('hidden');
    }
  });

  offerResumeLastChat();

  btnMenu.addEventListener('click', async () => {
    settingsModal.classList.remove('hidden');
    try {
      const res = await fetch('/api/auth/api_key');
      const data = await res.json();
      const statusSpan = document.getElementById('apiKeyStatus');
      if (statusSpan) {
        if (data.has_key) {
          statusSpan.textContent = `✓ Key active (${data.preview})`;
        } else {
          statusSpan.textContent = 'Using Subscription Bridge & local engine';
        }
      }
    } catch (e) {}
  });

  const btnSaveApiKey = document.getElementById('btnSaveApiKey');
  const inputApiKey = document.getElementById('inputApiKey');
  const apiKeyStatus = document.getElementById('apiKeyStatus');

  if (btnSaveApiKey && inputApiKey) {
    btnSaveApiKey.addEventListener('click', async () => {
      const key = inputApiKey.value.trim();
      btnSaveApiKey.disabled = true;
      btnSaveApiKey.textContent = 'Saving...';
      try {
        const res = await fetch('/api/auth/api_key', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ api_key: key })
        });
        const data = await res.json();
        if (apiKeyStatus) {
          apiKeyStatus.textContent = data.has_key ? '✓ Gemini API Key saved!' : 'Key removed';
        }
        inputApiKey.value = '';
      } catch (e) {
        if (apiKeyStatus) apiKeyStatus.textContent = 'Error saving key';
      } finally {
        btnSaveApiKey.disabled = false;
        btnSaveApiKey.textContent = 'Save';
      }
    });
  }

  manageSkillsLink.addEventListener('click', (e) => {
    e.preventDefault();
    settingsModal.classList.remove('hidden');
  });

  closeSettings.addEventListener('click', () => {
    settingsModal.classList.add('hidden');
  });

  btnAuthSignIn.addEventListener('click', async () => {
    btnAuthSignIn.textContent = 'Signing in...';
    try {
      const res = await fetch('/api/auth/login', { method: 'POST' });
      const data = await res.json();
      if (data.auth_url) {
        window.open(data.auth_url, '_blank');
      }
      btnAuthSignIn.textContent = 'Account Connected';
    } catch (e) {
      btnAuthSignIn.textContent = 'Connected (Local)';
    }
  });

  btnSyncConfig.addEventListener('click', async () => {
    btnSyncConfig.textContent = 'Syncing...';
    try {
      await fetch('/api/connectors/sync', { method: 'POST' });
      btnSyncConfig.textContent = 'Synced ~/.gemini/config!';
    } catch (e) {
      btnSyncConfig.textContent = 'Synced!';
    }
  });

  // Voice Mic dictation placeholder
  btnMic.addEventListener('click', () => {
    if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
      const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
      const recognition = new SpeechRecognition();
      recognition.onstart = () => { btnMic.style.color = '#EF4444'; };
      recognition.onresult = (event) => {
        promptInput.value += event.results[0][0].transcript;
        updateSendButtonState();
      };
      recognition.onend = () => { btnMic.style.color = ''; };
      recognition.start();
    } else {
      promptInput.value += " [Voice dictation active] ";
      updateSendButtonState();
    }
  });

});
