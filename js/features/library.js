/*
 * Document library: every document autosaves to IndexedDB in this browser.
 * FILE → New offers templates, FILE → Open lists recent documents, and
 * FILE → Options holds the user name and AutoCorrect settings.
 */
(function () {
  'use strict';
  const { $, docEl } = Ink;
  const CURRENT_KEY = 'inkwell:current';
  const OPTIONS_KEY = 'inkwell:options';
  const LEGACY_KEY = 'inkwell:autosave';
  let db = null;

  // ------------------------------------------------------------ options

  const DEFAULT_OPTIONS = { author: 'Author', autocorrect: { quotes: true, symbols: true, lists: true, links: true } };
  function loadOptions() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(OPTIONS_KEY) || '{}') || {}; } catch (e) { /* defaults */ }
    return {
      author: typeof saved.author === 'string' && saved.author.trim() ? saved.author.trim().slice(0, 60) : DEFAULT_OPTIONS.author,
      autocorrect: Object.assign({}, DEFAULT_OPTIONS.autocorrect, saved.autocorrect || {}),
    };
  }
  Ink.options = loadOptions();

  function saveOptions() {
    try { localStorage.setItem(OPTIONS_KEY, JSON.stringify(Ink.options)); } catch (e) { /* not persisted */ }
  }

  const form = $('#optionsForm');
  function fillOptions() {
    form.author.value = Ink.options.author;
    for (const k of Object.keys(DEFAULT_OPTIONS.autocorrect)) form[k].checked = !!Ink.options.autocorrect[k];
  }
  form.addEventListener('input', () => {
    Ink.options.author = form.author.value.trim().slice(0, 60) || 'Author';
    for (const k of Object.keys(DEFAULT_OPTIONS.autocorrect)) Ink.options.autocorrect[k] = form[k].checked;
    saveOptions();
  });
  form.addEventListener('submit', (e) => e.preventDefault());

  // ------------------------------------------------------------ storage

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) { reject(new Error('IndexedDB unavailable')); return; }
      const req = indexedDB.open('inkwell', 1);
      req.onupgradeneeded = () => req.result.createObjectStore('docs', { keyPath: 'id' });
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function store(mode, fn) {
    return new Promise((resolve, reject) => {
      const tx = db.transaction('docs', mode);
      const req = fn(tx.objectStore('docs'));
      tx.oncomplete = () => resolve(req && req.result);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }
  const put = (doc) => store('readwrite', (s) => s.put(doc));
  const get = (id) => store('readonly', (s) => s.get(id));
  const all = () => store('readonly', (s) => s.getAll());
  const remove = (id) => store('readwrite', (s) => s.delete(id));

  const newId = () => 'd' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  function setCurrent(id) {
    Ink.state.docId = id;
    try { localStorage.setItem(CURRENT_KEY, id); } catch (e) { /* not persisted */ }
  }

  async function persist() {
    const data = Ink.docData();
    data.id = data.id || newId();
    data.updated = Date.now();
    data.words = ((docEl.innerText || '').match(/\S+/g) || []).length;
    setCurrent(data.id);
    try {
      if (db) await put(data);
      else localStorage.setItem(LEGACY_KEY, JSON.stringify(data));
      Ink.setSaveState('Saved in this browser');
    } catch (e) {
      Ink.setSaveState('Autosave unavailable');
    }
  }

  // ------------------------------------------------------------ templates

  const p = (t) => '<p>' + t + '</p>';
  const TEMPLATES = {
    blank: { name: 'Blank document', desc: 'Start from an empty page', make: () => ({ title: 'Document1', html: '<p><br></p>' }) },
    welcome: {
      name: 'Welcome to Inkwell', desc: 'A quick tour of what you can do',
      make: () => ({
        title: 'Welcome to Inkwell',
        footer: { text: 'Page {PAGE} of {PAGES}', align: 'center' },
        html: '<p class="title">Welcome to Inkwell</p>' +
          p('Inkwell is a word processor modelled on Microsoft Word 2013. It reads and writes real <b>.docx</b> files, so anything you write here opens in Word, LibreOffice or Google Docs.') +
          '<h1>Write and format</h1>' +
          '<ul><li>Format text: <b>bold</b>, <i>italic</i>, <u>underline</u>, <s>strikethrough</s>, <span style="color: #c00000;">colour</span>, ' +
          '<span style="background-color: #ffff00;">highlight</span>, H<sub>2</sub>O and x<sup>2</sup></li>' +
          '<li>Apply and modify styles, or create your own from the Styles menu</li>' +
          '<li>Use the Paragraph dialog or drag the ruler markers to set indents and spacing</li>' +
          '<li>Copy formatting with the Format Painter, and change case with Aa</li></ul>' +
          '<h1>Lay out pages</h1>' +
          '<ol><li>Headers, footers and page numbers (Insert tab)</li><li>Margins, orientation, paper size and columns (Page Layout tab)</li>' +
          '<li>A table of contents that keeps its page numbers current (References tab)</li></ol>' +
          '<h1>Review with others</h1>' +
          p('Select text and choose <b>Review → New Comment</b>, or turn on <b>Track Changes</b> to record every insertion and deletion. Both are saved into the .docx file.') +
          '<h2>Handy shortcuts</h2>' +
          '<table><tbody><tr><td><b>Shortcut</b></td><td><b>Action</b></td></tr>' +
          '<tr><td>Ctrl+S</td><td>Save as .docx</td></tr><tr><td>Ctrl+F / Ctrl+H</td><td>Find / Replace</td></tr>' +
          '<tr><td>Ctrl+Alt+1, 2, 3</td><td>Heading 1, 2, 3</td></tr><tr><td>Ctrl+Enter</td><td>Page break</td></tr></tbody></table>' +
          '<blockquote>Everything you type is saved in this browser automatically.</blockquote>',
      }),
    },
    resume: {
      name: 'Resume', desc: 'One page, clean and simple',
      make: () => ({
        title: 'Resume',
        html: '<p class="title">Your Name</p>' + p('City, Country · you@example.com · (555) 123-4567') +
          '<h1>Experience</h1>' + p('<b>Job Title</b>, Company — <i>2022 to present</i>') +
          '<ul><li>Describe an achievement with a number in it</li><li>Another responsibility or result</li></ul>' +
          p('<b>Previous Title</b>, Company — <i>2019 to 2022</i>') + '<ul><li>What you built, led or improved</li></ul>' +
          '<h1>Education</h1>' + p('<b>Degree</b>, University — <i>2019</i>') +
          '<h1>Skills</h1>' + p('Skill one · Skill two · Skill three · Skill four'),
      }),
    },
    letter: {
      name: 'Business letter', desc: 'Block format with date and signature',
      make: () => ({
        title: 'Letter',
        html: p('Your Name<br>Street Address<br>City, Postcode') +
          p(new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })) +
          p('Recipient Name<br>Company<br>Street Address<br>City, Postcode') + p('Dear Recipient,') +
          p('Start with the reason you are writing. Keep the first paragraph short and specific.') +
          p('Use the next paragraph for detail: what you need, by when, and why it matters.') +
          p('Thank you for your time. I look forward to hearing from you.') + p('Sincerely,') + p('<br>') + p('Your Name'),
      }),
    },
    report: {
      name: 'Report', desc: 'Title, contents, headings and page numbers',
      make: () => ({
        title: 'Report',
        header: { text: 'Report Title', align: 'right' },
        footer: { text: 'Page {PAGE} of {PAGES}', align: 'center' },
        differentFirst: true,
        html: '<p class="title">Report Title</p>' + p('Prepared by Your Name · ' + new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long' })) +
          '<div class="toc" contenteditable="false"></div>' + '<div class="page-break" contenteditable="false"></div>' +
          '<h1>Summary</h1>' + p('State the purpose of the report and its main finding in two or three sentences.') +
          '<h1>Background</h1>' + p('Explain the context a reader needs.') +
          '<h2>Scope</h2>' + p('What this report covers, and what it does not.') +
          '<h1>Findings</h1>' + p('Present the evidence. Use tables and figures where they help.') +
          '<table><tbody><tr><td><b>Measure</b></td><td><b>Result</b></td></tr><tr><td>First</td><td>—</td></tr><tr><td>Second</td><td>—</td></tr></tbody></table>' +
          '<h1>Recommendations</h1><ol><li>First recommendation</li><li>Second recommendation</li></ol>',
      }),
    },
    notes: {
      name: 'Meeting notes', desc: 'Agenda, decisions and action items',
      make: () => ({
        title: 'Meeting Notes',
        html: '<p class="title">Meeting Notes</p>' + p('<b>Date:</b> ' + new Date().toLocaleDateString() + '   <b>Attendees:</b> ') +
          '<h1>Agenda</h1><ol><li>Topic one</li><li>Topic two</li></ol>' +
          '<h1>Decisions</h1><ul><li>Decision and who made it</li></ul>' +
          '<h1>Action items</h1><ul data-bullet="check"><li>Task — owner — due date</li></ul>',
      }),
    },
    essay: {
      name: 'Essay', desc: 'Double-spaced with a running header',
      make: () => ({
        title: 'Essay',
        header: { text: 'Surname {PAGE}', align: 'right' },
        html: '<p style="line-height: 2;">Your Name<br>Instructor<br>Course<br>' + new Date().toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' }) + '</p>' +
          '<p style="text-align: center; line-height: 2;">Title of the Essay</p>' +
          '<p style="text-indent: 48px; line-height: 2;">Begin with a sentence that draws the reader in, then narrow toward your thesis.</p>' +
          '<p style="text-indent: 48px; line-height: 2;">Each body paragraph makes one point and supports it with evidence.</p>',
      }),
    },
  };

  async function newFromTemplate(key) {
    if (Ink.state.docId) await persist();
    const t = (TEMPLATES[key] || TEMPLATES.blank).make();
    const id = newId();
    setCurrent(id);
    Ink.loadDocument(Object.assign({ id }, t));
    await persist();
  }

  function renderTemplates() {
    $('#templateList').innerHTML = Object.entries(TEMPLATES).map(([key, t]) =>
      '<button class="template" data-cmd="newFrom" data-arg="' + key + '"><span class="thumb thumb-' + key + '"></span>' +
      '<strong>' + Ink.escHtml(t.name) + '</strong><span>' + Ink.escHtml(t.desc) + '</span></button>').join('');
  }

  async function renderRecent() {
    const list = $('#recentList');
    if (!db) { list.innerHTML = '<p class="muted">Recent documents need browser storage, which isn’t available here.</p>'; return; }
    const docs = (await all()).sort((a, b) => (b.updated || 0) - (a.updated || 0));
    list.innerHTML = docs.length ? docs.map((d) =>
      '<div class="recent-item' + (d.id === Ink.state.docId ? ' current' : '') + '">' +
      '<button class="recent-open" data-cmd="openDoc" data-arg="' + Ink.escHtml(d.id) + '"><strong>' + Ink.escHtml(d.title || 'Untitled') + '</strong>' +
      '<span>' + (d.updated ? new Date(d.updated).toLocaleString() : '') + ' · ' + (d.words || 0).toLocaleString() + ' words' +
      (d.id === Ink.state.docId ? ' · open now' : '') + '</span></button>' +
      '<button class="recent-del" data-cmd="deleteDoc" data-arg="' + Ink.escHtml(d.id) + '" title="Remove from this browser" aria-label="Remove">×</button></div>').join('')
      : '<p class="muted">No saved documents yet.</p>';
  }

  async function openDoc(id) {
    if (id === Ink.state.docId) { Ink.closeBackstage(); return; }
    await persist();
    const doc = await get(id);
    if (!doc) { Ink.toast('That document is no longer available'); renderRecent(); return; }
    Ink.closeBackstage();
    setCurrent(id);
    Ink.loadDocument(doc);
    Ink.toast('Opened ' + (doc.title || 'document'));
  }

  async function deleteDoc(id) {
    const doc = await get(id);
    const ok = await Ink.dialog('Remove Document', '<p>Remove “' + Ink.escHtml((doc && doc.title) || 'Untitled') +
      '” from this browser? This can’t be undone. Download it first if you want to keep a copy.</p>', 'Remove');
    if (!ok) return;
    await remove(id);
    if (id === Ink.state.docId) {
      Ink.state.docId = null;
      await newFromTemplate('blank');
      Ink.openBackstage('open');
    }
    renderRecent();
  }

  Ink.command('newFrom', (key) => { Ink.closeBackstage(); newFromTemplate(key); });
  Ink.command('newDoc', () => newFromTemplate('blank'));
  Ink.command('openDoc', openDoc);
  Ink.command('deleteDoc', deleteDoc);
  Ink.shortcut((e) => e.altKey && e.key.toLowerCase() === 'n', () => newFromTemplate('blank'));

  Ink.on('backstage', (panel) => {
    if (panel === 'open') renderRecent();
    if (panel === 'new') renderTemplates();
    if (panel === 'options') fillOptions();
  });
  // A file opened from disk becomes a new document in the library.
  Ink.on('beforeOpen', (data) => {
    if (Ink.state.docId) persist();
    data.id = newId();
    setCurrent(data.id);
  });
  window.addEventListener('pagehide', () => { persist(); });

  Ink.persist = persist;
  Ink.boot = async function () {
    try { db = await openDb(); } catch (e) { db = null; }
    // Documents autosaved by Inkwell 1 move into the library.
    let legacy = null;
    try { legacy = JSON.parse(localStorage.getItem(LEGACY_KEY) || 'null'); } catch (e) { /* ignore */ }
    if (db && legacy && typeof legacy.html === 'string') {
      legacy.id = legacy.id || newId();
      legacy.html = Ink.sanitize(legacy.html, true);
      await put(legacy);
      try { localStorage.removeItem(LEGACY_KEY); localStorage.setItem(CURRENT_KEY, legacy.id); } catch (e) { /* ignore */ }
    }
    let currentId = null;
    try { currentId = localStorage.getItem(CURRENT_KEY); } catch (e) { /* none */ }
    let doc = null;
    if (db && currentId) { try { doc = await get(currentId); } catch (e) { doc = null; } }
    if (!db && legacy) doc = legacy;
    if (doc && typeof doc.html === 'string') {
      setCurrent(doc.id || newId());
      Ink.loadDocument(Object.assign({}, doc, { html: Ink.sanitize(doc.html, true) }));
      Ink.toast('Restored your last session');
    } else {
      await newFromTemplate('welcome');
    }
  };
})();
