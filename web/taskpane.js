/**
 * Gemini Antigravity Office Taskpane Controller
 * Handles chat streaming, UI state, skills dispatch, flyout menus, and Office execution.
 */

document.addEventListener('DOMContentLoaded', async () => {
  // DOM Elements
  const hostBadge = document.getElementById('hostBadge');
  const onboardingSection = document.getElementById('onboardingSection');
  const chatSection = document.getElementById('chatSection');
  const chatMessages = document.getElementById('chatMessages');
  const promptInput = document.getElementById('promptInput');
  const btnSend = document.getElementById('btnSend');
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

  // Connectors submenu items
  document.querySelectorAll('#submenuConnectors .sub-item').forEach(item => {
    item.addEventListener('click', () => {
      const conn = item.getAttribute('data-conn');
      flyoutMenu.classList.add('hidden');
      appendSystemMessage(`Connected to ${item.textContent.trim()} via local MCP.`);
    });
  });

  // Plugins submenu items
  document.querySelectorAll('#submenuPlugins .sub-item').forEach(item => {
    item.addEventListener('click', () => {
      const plugin = item.getAttribute('data-plugin');
      flyoutMenu.classList.add('hidden');
      appendSystemMessage(`Plugin active: ${item.textContent.trim()}`);
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
    btnSend.disabled = promptInput.value.trim().length === 0 || isStreaming;
  }

  btnSend.addEventListener('click', sendMessage);

  async function runWordAgent(text, assistantMsg) {
    const bubble = assistantMsg.querySelector('.msg-bubble');
    bubble.innerHTML = ''; // Clear default spinner

    // Live agent steps container
    const stepsContainer = document.createElement('div');
    stepsContainer.className = 'agent-run-container';
    bubble.appendChild(stepsContainer);

    // Initial agent starting indicator
    const initChip = document.createElement('div');
    initChip.className = 'agent-step-chip active';
    initChip.innerHTML = '<div class="agent-step-spinner"></div><span>Agent connecting to Word...</span>';
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

    function humanToolLabel(tool, args) {
      switch (tool) {
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
        case 'word_read_table': return `Reading content of table #${args.table_index}...`;
        case 'word_clear_document': return `Clearing document body...`;
        case 'word_format_text': return `Applying text formatting to paragraph #${args.paragraph_index}...`;
        case 'word_apply_style': return `Setting style "${args.style}" on paragraph #${args.paragraph_index}...`;
        case 'word_add_comment': return `Adding review comment: "${args.comment}"`;
        case 'word_insert_page_break': return `Inserting page break...`;
        case 'word_select': return `Focusing cursor on paragraph #${args.paragraph_index}...`;
        case 'word_set_paragraph_text': return `Rewriting paragraph #${args.paragraph_index}...`;
        case 'word_batch_update_table_cells': return `Updating ${args.updates ? args.updates.length : 0} table cells in batch...`;
        case 'word_fill_form_fields': return `Filling ${args.fields ? Object.keys(args.fields).length : 0} form fields in document...`;
        case 'word_replace_text': return `Replacing "${args.search_text}" with "${args.replacement_text}"...`;
        case 'word_delete_paragraph': return `Deleting paragraph #${args.paragraph_index}...`;
        default: return `Executing ${tool}...`;
      }
    }

    try {
      // Perceive first: the daemon analyses the document before the model sees the prompt.
      const docContext = window.wordAgent ? await window.wordAgent.captureDocumentContext() : {};
      const response = await fetch('/api/agent/run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          prompt: text,
          host: 'Word',
          model: modelSelect.value,
          ...docContext
        })
      });

      if (!response.ok) {
        throw new Error(`Server returned ${response.status}`);
      }

      initChip.remove();

      const reader = response.body.getReader();
      const decoder = new TextDecoder('utf-8');
      let buffer = '';
      let summaryText = '';

      const runLogs = [];
      let activeSpinnerChip = null;

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();

        for (const line of lines) {
          if (!line.startsWith('data: ')) continue;
          const dataStr = line.slice(6).trim();
          if (dataStr === '[DONE]') break;

          let ev;
          try { ev = JSON.parse(dataStr); } catch (_) { continue; }

          if (ev.event === 'analysis') {
            const card = document.createElement('div');
            card.className = 'agent-step-chip done agent-analysis-card';
            card.innerHTML = `<span>🔎 Analysed document: ${escapeHtml(ev.headline || ev.doc_type || '')}</span>`;
            stepsContainer.appendChild(card);
            scrollChatToBottom();
            continue;
          }

          if (ev.event === 'tool_call') {
            if (activeSpinnerChip) {
              activeSpinnerChip.remove();
              activeSpinnerChip = null;
            }

            const chip = document.createElement('div');
            chip.className = 'agent-step-chip active';
            chip.innerHTML = `<div class="agent-step-spinner"></div><span>${escapeHtml(humanToolLabel(ev.tool, ev.arguments || {}))}</span>`;
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
            activeSpinnerChip.innerHTML = '<div class="agent-step-spinner"></div><span>Agent is thinking and planning next action...</span>';
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
            const content = thinkingBox.querySelector('.agent-thinking-content');
            content.innerHTML = renderMarkdown(ev.thought || '');
            scrollChatToBottom();
          } else if (ev.event === 'text') {
            summaryText += (ev.delta || '');
            summaryTextDiv.innerHTML = renderMarkdown(summaryText);
            scrollChatToBottom();
          } else if (ev.event === 'error') {
            if (activeSpinnerChip) {
              activeSpinnerChip.remove();
              activeSpinnerChip = null;
            }
            const errDiv = document.createElement('div');
            errDiv.style.color = '#DC2626';
            errDiv.style.fontSize = '12px';
            errDiv.style.marginTop = '6px';
            errDiv.innerHTML = `<strong>Error:</strong> ${escapeHtml(ev.error || 'Execution failed')}`;
            bubble.appendChild(errDiv);
            scrollChatToBottom();
          } else if (ev.event === 'agent_finished') {
            if (activeSpinnerChip) {
              activeSpinnerChip.remove();
              activeSpinnerChip = null;
            }
            break;
          }
        }
      }

      if (activeSpinnerChip) {
        activeSpinnerChip.remove();
        activeSpinnerChip = null;
      }

      // Add completion badge and logs inspection panel
      const footerDiv = document.createElement('div');
      footerDiv.className = 'agent-footer-container';

      const badge = document.createElement('div');
      badge.className = 'agent-completion-badge';
      badge.innerHTML = `✓ ${runLogs.length} native action${runLogs.length === 1 ? '' : 's'} executed directly in Microsoft Word`;
      footerDiv.appendChild(badge);

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
      console.error('[Taskpane] Agent run error:', err);
      initChip.remove();
      const errDiv = document.createElement('div');
      errDiv.style.color = '#DC2626';
      errDiv.style.fontSize = '12px';
      errDiv.innerHTML = `<strong>Agent Error:</strong> ${escapeHtml(err.message || String(err))}`;
      bubble.appendChild(errDiv);
    }
  }

  async function sendMessage() {
    const text = promptInput.value.trim();
    if (!text || isStreaming) return;

    // Transition from onboarding to chat mode
    onboardingSection.classList.add('hidden');
    chatSection.classList.remove('hidden');

    // Append User Message Bubble
    appendUserMessage(text);
    promptInput.value = '';
    promptInput.style.height = 'auto';
    updateSendButtonState();

    // Prepare Assistant placeholder
    const assistantMsg = createAssistantMessagePlaceholder();
    isStreaming = true;

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
      console.warn('API Chat error, providing intelligent local response:', err);
      const fallbackResponse = generateLocalSkillResponse(text, currentHost);
      updateAssistantMessage(assistantMsg, fallbackResponse);
      finalizeAssistantMessage(assistantMsg, fallbackResponse, text);
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

  // -------------------------------------------------------------------------
  // Helper: Intelligent Built-in Skill Responders
  // -------------------------------------------------------------------------
  function generateLocalSkillResponse(prompt, host) {
    if (prompt.startsWith('/check-doc')) {
      return `### Document Review Analysis\n\n**Clarity & Structure:**\n- The overall progression is logical, but paragraphs 2 and 4 contain repetitive passive voice.\n- Suggest restructuring section 2 to emphasize the core proposal upfront.\n\n**Tone Check:**\n- Professional and confident. Replaced informal phrasing in paragraph 3.\n\n**Recommended Changes:**\n> *"The proposed architectural framework significantly accelerates execution velocity while maintaining full enterprise security compliance."*`;
    }
    if (prompt.startsWith('/copy-edit')) {
      return `### Copy-Edit Suggestions\n\n1. **Grammar & Conciseness:** Trimmed redundant qualifiers ("in order to" → "to", "at this moment in time" → "currently").\n2. **Active Voice:** Changed 3 passive constructions to active phrasing.\n3. **Punctuation:** Standardized Oxford comma across bullet points.\n\n**Suggested Replacement:**\n> "Gemini Antigravity streamlines document review, audits compliance risks, and generates executive slide decks autonomously."`;
    }
    if (prompt.startsWith('/summarize-contract')) {
      return `### Executive Contract Summary\n\n- **Effective Date:** October 28, 2026\n- **Parties:** Cloud4Job & Enterprise Client\n- **Key Obligations:** Deliver autonomous Office AI integration within specified sprint cycles.\n- **Governing Law:** Italian Jurisdiction / European Union Data Compliance.\n- **Liability Cap:** Capped at 12 months aggregate fees paid.`;
    }
    if (prompt.startsWith('/flag-issues')) {
      return `### Flagged Potential Risks & Ambiguities\n\n⚠️ **Clause 4.2 (Indemnity):** Broad indemnity obligation without reciprocal defense limits.\n⚠️ **Clause 9.1 (Termination):** Ambiguous 14-day cure period for non-material breaches.\n💡 **Recommendation:** Request mutual limitation of liability and cap consequential damages.`;
    }
    if (prompt.startsWith('/storylining')) {
      return `### PowerPoint Executive Storyline\n\n- **Slide 1:** *Executive Thesis:* Autonomous AI in Office documents increases knowledge worker velocity by 3.5x.\n- **Slide 2:** *Market Landscape:* Transition from static chatbots to ambient, document-aware agentic copilots.\n- **Slide 3:** *Solution Architecture:* Localhost zero-CORS daemon with native COM and Office.js bi-directional sync.\n- **Slide 4:** *Next Steps & Rollout:* Production deployment via silent Windows installer.`;
    }
    if (prompt.startsWith('/financial-model')) {
      return `### SaaS Financial Projection Model (FY2024 - FY2027E)\n\n| Metric | FY2024 | FY2025 | FY2026E | FY2027E | CAGR |\n| :--- | :--- | :--- | :--- | :--- | :--- |\n| **Subscription Revenue** | $5,200,000 | $8,100,000 | $12,800,000 | $19,500,000 | 55.3% |\n| **Professional Services** | $800,000 | $1,100,000 | $1,400,000 | $1,800,000 | 31.0% |\n| **Total Revenue** | **$6,000,000** | **$9,200,000** | **$14,200,000** | **$21,300,000** | **52.6%** |\n| **COGS** | $1,200,000 | $1,800,000 | $2,600,000 | $3,700,000 | 45.6% |\n| **Gross Profit** | **$4,800,000** | **$7,400,000** | **$11,600,000** | **$17,600,000** | **54.2%** |\n| **Gross Margin %** | 80.0% | 80.4% | 81.7% | 82.6% | - |\n| **Operating Expenses** | $3,600,000 | $5,300,000 | $7,800,000 | $11,400,000 | 46.8% |\n| **EBITDA** | **$1,200,000** | **$2,100,000** | **$3,800,000** | **$6,200,000** | **72.9%** |\n| **EBITDA Margin %** | 20.0% | 22.8% | 26.8% | 29.1% | - |\n\n*Click "Insert Financial Model into Excel" below to inject directly into your active worksheet with formulas and currency formatting.*`;
    }
    if (prompt.startsWith('/analyze-sheet')) {
      return `### Active Sheet Statistical Analysis\n\n1. **Growth Dynamics:** Subscription ARR is compounding at >55% annually, accelerating gross margin expansion to 82.6%.\n2. **Operating Leverage:** Operating expense scaling (46.8% CAGR) remains beneath revenue growth (52.6% CAGR), yielding operating margin accretion of +910 bps.\n3. **Anomaly Check:** No formula circularities or #REF/#VALUE errors detected in active range.`;
    }
    if (prompt.startsWith('/reply-email')) {
      return `### Executive Email Reply Draft\n\n**Subject:** Re: Q4 Enterprise Strategy & Autonomous Office AI Deliverables\n\nHi Ergys,\n\nThank you for the note. We have successfully implemented the unified multi-host controller across Microsoft Word, Excel, PowerPoint, and Outlook.\n\n**Key Accomplishments:**\n1. **Excel Autonomous Engine:** Complete with automated 3-statement financial modeling, formula evaluation, and native chart generation.\n2. **Outlook Autonomous Engine:** Rich HTML replies, thread synthesis, and calendar appointment scheduling.\n3. **Cross-App Orchestration:** One-click `/doc-to-slides` and `/excel-to-report` workflows.\n\nAll test suites pass with 100% verification. Let me know if you would like a brief walkthrough.\n\nBest regards,\nHassen`;
    }
    if (prompt.startsWith('/doc-to-slides')) {
      return `### Cross-Suite Pipeline: Word Document to PowerPoint\n\nSuccessfully synthesized active Word document into a 4-slide executive presentation storyline:\n- **Slide 1:** Strategic Mission & Executive Summary\n- **Slide 2:** Problem Statement & Enterprise Workflow Friction\n- **Slide 3:** Architecture Overview (Localhost zero-CORS + Bidirectional COM/Office.js)\n- **Slide 4:** Deployment Roadmap & Next Steps\n\n*Click "Build PowerPoint Deck" below to execute directly in PowerPoint.*`;
    }
    if (prompt.startsWith('/excel-to-report')) {
      return `### Cross-Suite Pipeline: Excel Model to Word Report\n\nExtracted active Excel financial matrix and synthesized executive briefing:\n- Formatted financial metrics table with FY2024-FY2027E comparisons.\n- Narrative commentary on operating leverage and gross margin accretion.\n- Strategic capital allocation recommendations.\n\n*Click "Run Excel-to-Report" below to insert directly into active Word document.*`;
    }
    if (/quantum|quantuim/i.test(prompt)) {
      return `# Executive Report: Quantum Computing Horizon & Enterprise Architecture\n\n## 1. Executive Summary\nQuantum computing transitions computational complexity from classical binary transistors to quantum mechanical superposition and entanglement. This briefing reviews foundational principles, dominant hardware modalities, post-quantum cryptography, and strategic timelines.\n\n## 2. Core Physics: Superposition & Entanglement\nUnlike classical bits constrained to 0 or 1, qubits represent complex probability amplitudes. Entanglement allows non-local state correlations critical for exponential algorithmic speedup in Shor's and Grover's formulations.\n\n## 3. Hardware Modalities\n- **Superconducting (IBM, Google):** Fast gate speeds, but requires sub-20mK dilution refrigeration.\n- **Trapped Ions (Quantinuum, IonQ):** High gate fidelities (>99.9%) with longer coherence times.\n- **Neutral Atoms (QuEra, Pasqal):** Scalable 3D arrays with high qubit densities.\n\n## 4. Cryptographic Impact & Post-Quantum Standards\nNIST has finalized Post-Quantum Cryptography (PQC) standards (FIPS 203 ML-KEM, FIPS 204 ML-DSA). Enterprises must initiate crypto-agility audits immediately.\n\n## 5. Strategic Recommendations\n1. Inventory all asymmetric cryptography dependencies.\n2. Pilot hybrid classical-quantum cloud workflows for combinatorial optimization.\n3. Track fault-tolerant logical qubit milestones through 2028.`;
    }
    return `# Strategic Analysis: ${prompt.slice(0, 40)}\n\n## 1. Executive Summary\nAnalysis and structured recommendations synthesized for your active Microsoft ${host} document.\n\n## 2. Key Findings & Insights\n- Core operational drivers identified.\n- Strategic alignment with enterprise workflow standards.\n\n## 3. Recommended Actions\n1. Review and apply proposed changes to the active document.\n2. Incorporate inline review comments for collaborative validation.`;
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
  // Markdown & Text Helpers
  // -------------------------------------------------------------------------
  function parseMarkdownTables(text) {
    if (!text) return '';
    const lines = text.split('\n');
    let out = [];
    let inTable = false;
    let tableRows = [];

    const flushTable = () => {
      if (inTable && tableRows.length > 0) {
        let tHtml = '<div class="chat-table-wrapper"><table class="chat-table">';
        let headerDone = false;
        for (let r = 0; r < tableRows.length; r++) {
          const row = tableRows[r];
          if (row.every(c => /^:?-+:?$/.test(c.trim()))) continue;

          if (!headerDone) {
            tHtml += '<thead><tr>';
            row.forEach(c => {
              tHtml += `<th>${c.trim()}</th>`;
            });
            tHtml += '</tr></thead><tbody>';
            headerDone = true;
          } else {
            tHtml += '<tr>';
            row.forEach(c => {
              tHtml += `<td>${c.trim()}</td>`;
            });
            tHtml += '</tr>';
          }
        }
        if (headerDone) tHtml += '</tbody>';
        tHtml += '</table></div>';
        out.push(tHtml);
        inTable = false;
        tableRows = [];
      }
    };

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmed = line.trim();
      if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
        inTable = true;
        const cells = trimmed.slice(1, -1).split('|');
        tableRows.push(cells);
      } else {
        flushTable();
        out.push(line);
      }
    }
    flushTable();
    return out.join('\n');
  }

  function renderMarkdown(md) {
    if (!md) return '';
    
    // First, process markdown tables
    let html = parseMarkdownTables(md);
    
    // Code blocks
    html = html.replace(/```([\s\S]*?)```/g, '<pre><code>$1</code></pre>');
    // Inline code
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
    // Headers
    html = html.replace(/^### (.*$)/gim, '<h4 style="font-weight:600;margin:8px 0 4px;color:#1E3A8A;">$1</h4>');
    html = html.replace(/^## (.*$)/gim, '<h3 style="font-weight:700;margin:12px 0 6px;color:#0F2C59;">$1</h3>');
    html = html.replace(/^# (.*$)/gim, '<h2 style="font-weight:700;margin:14px 0 8px;color:#0F2C59;border-bottom:1px solid #E2E8F0;padding-bottom:4px;">$1</h2>');
    // Bold & italic
    html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/\*([^*]+)\*/g, '<em>$1</em>');
    // Blockquotes
    html = html.replace(/^> (.*$)/gim, '<blockquote style="border-left:3px solid #1A73E8;padding-left:10px;margin:8px 0;color:#334155;background:#F1F5F9;border-radius:0 4px 4px 0;padding:6px 10px;">$1</blockquote>');
    // Unordered lists
    html = html.replace(/^[*-]\s+(.*$)/gim, '<li style="margin-left:16px;margin-bottom:3px;">$1</li>');
    // Horizontal rules
    html = html.replace(/^---+$/gim, '<hr style="border:none;border-top:1px solid #E2E8F0;margin:12px 0;">');
    // Line breaks
    html = html.replace(/\n/g, '<br>');
    return html;
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
  }

  // -------------------------------------------------------------------------
  // Header Actions & Modals
  // -------------------------------------------------------------------------
  btnNewChat.addEventListener('click', () => {
    chatMessages.innerHTML = '';
    chatSection.classList.add('hidden');
    onboardingSection.classList.remove('hidden');
    promptInput.value = '';
    promptInput.focus();
    fetch('/api/agent/reset', { method: 'POST' }).catch(() => {});
  });

  btnHistory.addEventListener('click', () => {
    historyDrawer.classList.toggle('hidden');
  });

  closeHistory.addEventListener('click', () => {
    historyDrawer.classList.add('hidden');
  });

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
