/*
 * Navigation pane headings outline, automatic table of contents and
 * captions (References tab).
 */
(function () {
  'use strict';
  const { $, $$, docEl } = Ink;
  const pane = $('#findPane');

  const headings = () => $$('h1, h2, h3', docEl).filter((h) => !h.closest('.toc') && h.textContent.trim());

  function goTo(el) {
    el.scrollIntoView({ block: 'start', behavior: 'smooth' });
    docEl.focus({ preventScroll: true });
    Ink.placeCaret(el);
  }

  // ------------------------------------------------------------ navigation pane

  let headingsTimer = null;
  function renderHeadings() {
    if (pane.hidden) return;
    const list = headings();
    $('#navHeadings').innerHTML = list.length
      ? list.map((h, i) => '<button class="nav-heading lvl-' + h.tagName[1] + '" data-heading="' + i + '">' + Ink.escHtml(h.textContent.trim()) + '</button>').join('')
      : '<p class="muted nav-empty">Apply the Heading 1–3 styles to see an outline of your document here.</p>';
  }
  const renderSoon = () => { clearTimeout(headingsTimer); headingsTimer = setTimeout(renderHeadings, 250); };

  $('#navHeadings').addEventListener('click', (e) => {
    const b = e.target.closest('[data-heading]');
    const h = b && headings()[+b.dataset.heading];
    if (h) goTo(h);
  });

  function setNavPane(on) {
    if (on) {
      pane.hidden = false;
      $('#replaceArea').hidden = true;
      Ink.showNavTab('headings');
      renderHeadings();
    } else {
      Ink.closeFind();
    }
    $('#navBtn').classList.toggle('on', on);
  }

  Ink.command('navPane', () => setNavPane(pane.hidden));
  Ink.on('navClosed', () => $('#navBtn').classList.remove('on'));
  Ink.on('navTab', (name) => {
    $('#navBtn').classList.toggle('on', !pane.hidden);
    if (name === 'headings') renderHeadings();
  });
  Ink.on('change', renderSoon);

  // ------------------------------------------------------------ table of contents

  function tocHtml() {
    const list = headings();
    return '<p class="toc-title">Contents</p>' + (list.length
      ? list.map((h, i) => '<p data-level="' + h.tagName[1] + '" data-index="' + i + '"><span class="toc-text">' + Ink.escHtml(h.textContent.trim()) +
        '</span><span class="toc-dots"></span><span class="toc-page">' + Ink.pageOf(h) + '</span></p>').join('')
      : '<p class="toc-empty">No headings yet. Apply Heading 1–3 styles and the table fills in.</p>');
  }

  function refreshTocs() {
    const tocs = $$('.toc', docEl);
    if (!tocs.length) return false;
    const html = tocHtml();
    let changed = false;
    for (const t of tocs) {
      if (t.innerHTML !== html) { t.innerHTML = html; changed = true; }
    }
    return changed;
  }

  // Page numbers depend on layout, so the table follows every pagination.
  let refreshing = false;
  Ink.on('paginated', () => {
    if (refreshing) return;
    refreshing = true;
    try { if (refreshTocs()) Ink.paginate(); } finally { refreshing = false; }
  });

  docEl.addEventListener('click', (e) => {
    const entry = e.target.closest('.toc p[data-index]');
    const h = entry && headings()[+entry.dataset.index];
    if (h) goTo(h);
  });

  Ink.command('insertToc', () => {
    if ($('.toc', docEl)) { Ink.run('updateToc'); return; }
    Ink.edit(() => Ink.exec('insertHTML', '<div class="toc" contenteditable="false">' + tocHtml() + '</div><p><br></p>'));
  });
  Ink.command('updateToc', () => {
    if (!$('.toc', docEl)) { Ink.toast('Insert a table of contents first (References → Table of Contents)'); return; }
    Ink.edit(refreshTocs);
    Ink.toast('Table of contents updated');
  });
  Ink.command('removeToc', () => Ink.edit(() => { for (const t of $$('.toc', docEl)) t.remove(); }));

  // ------------------------------------------------------------ captions

  async function insertCaption() {
    const img = Ink.getActiveImage();
    const s = getSelection();
    const anchor = img || (s.rangeCount && Ink.inDoc(s.anchorNode) ? s.anchorNode : null);
    const res = await Ink.dialog('Insert Caption',
      '<label>Label<select name="label"><option>Figure</option><option>Table</option><option>Equation</option></select></label>' +
      '<label>Caption<input type="text" name="text" placeholder="Describe it"></label>', 'Insert');
    if (!res) return;
    const n = $$('.caption', docEl).filter((c) => c.textContent.startsWith(res.label + ' ')).length + 1;
    Ink.edit(() => {
      const p = document.createElement('p');
      p.className = 'caption';
      p.textContent = res.label + ' ' + n + (res.text.trim() ? ': ' + res.text.trim() : '');
      let block = anchor ? (Ink.closestIn(anchor, 'table') || Ink.closestBlock(anchor)) : null;
      if (block && block.tagName === 'LI') block = block.closest('ul,ol');
      while (block && block.parentElement && block.parentElement !== docEl && !block.parentElement.classList.contains('section')) block = block.parentElement;
      if (block) block.after(p); else docEl.appendChild(p);
      Ink.placeCaret(p, true);
    }, { focus: false });
  }
  Ink.command('insertCaption', insertCaption);
})();
