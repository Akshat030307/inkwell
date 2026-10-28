/*
 * Print Layout: the document is one editable column laid over a stack of
 * page sheets. Blocks that would cross a page's bottom margin are pushed to
 * the next page by growing their top margin; sheets carry the header and
 * footer, with {PAGE} and {PAGES} filled in.
 */
(function () {
  'use strict';
  const { $, $$, docEl, state } = Ink;
  const GAP = 24;   // px of desk between sheets
  const sheets = $('#sheets');
  const hf = { header: { text: '', align: 'center' }, footer: { text: '', align: 'center' }, differentFirst: false };
  let pages = 1;
  let enabled = true;

  function geometry() {
    const L = state.layout;
    const pageH = L.h * 96;
    return { L, pageH, P: pageH + GAP, mt: L.t * 96, mb: L.b * 96 };
  }

  function strip(root) {
    for (const el of root.querySelectorAll('[data-pg]')) {
      el.style.marginTop = el.dataset.pgOrig || '';
      delete el.dataset.pg;
      delete el.dataset.pgOrig;
      if (!el.getAttribute('style')) el.removeAttribute('style');
    }
  }

  function unpaginate() {
    strip(docEl);
    docEl.style.minHeight = '';
  }

  function paginate() {
    unpaginate();
    if (!enabled) { renderSheets(); return; }
    const { L, pageH, P, mt, mb } = geometry();
    const units = [];
    for (const el of docEl.children) {
      if (el.tagName === 'UL' || el.tagName === 'OL') units.push(...Array.from(el.children).filter((c) => c.tagName === 'LI'));
      else units.push(el);
    }
    const rect = docEl.getBoundingClientRect();
    const scale = rect.width / (L.w * 96) || 1;
    const measured = units.map((el) => {
      const r = el.getBoundingClientRect();
      return { el, top: (r.top - rect.top) / scale, bottom: (r.bottom - rect.top) / scale };
    });
    let shift = 0, prevBottom = mt, forced = null, lastBottom = mt;
    for (const u of measured) {
      const top = u.top + shift, h = u.bottom - u.top;
      const k = Math.floor(top / P);
      const contentBottom = k * P + pageH - mb;
      let target = null;
      if (forced !== null) target = forced;
      else if (top >= contentBottom) target = (k + 1) * P + mt;
      else if (k > 0 && top < k * P + mt - 0.5) target = k * P + mt;
      else if (top + h > contentBottom + 0.5 && h <= pageH - mt - mb) target = (k + 1) * P + mt;
      forced = null;
      if (target !== null && target > top + 0.5) {
        const push = target - top;
        // The measured gap already includes collapsed margins; adding the
        // push to it moves this block down by exactly `push`.
        u.el.dataset.pgOrig = u.el.style.marginTop;
        u.el.dataset.pg = '';
        u.el.style.marginTop = (Math.round((u.top - prevBottom + push) * 100) / 100) + 'px';
        shift += push;
      }
      if (u.el.classList && u.el.classList.contains('page-break')) {
        forced = (Math.floor((u.top + shift) / P) + 1) * P + mt;
      }
      prevBottom = u.bottom;
      lastBottom = Math.max(lastBottom, u.bottom + shift);
    }
    pages = Math.max(1, Math.floor((lastBottom - 1) / P) + 1);
    docEl.style.minHeight = (pages * P - GAP) + 'px';
    renderSheets();
    Ink.emit('paginated');
  }

  const fill = (text, k) => Ink.escHtml(text).replace(/\{PAGE\}/g, String(k + 1)).replace(/\{PAGES\}/g, String(pages));

  function renderSheets() {
    const { pageH, P } = geometry();
    const count = enabled ? pages : 0;
    let html = '';
    for (let k = 0; k < count; k++) {
      const show = !(hf.differentFirst && k === 0);
      html += '<div class="sheet" style="top:' + k * P + 'px;height:' + pageH + 'px">';
      for (const kind of ['header', 'footer']) {
        const h = hf[kind];
        html += '<div class="hf hf-' + kind + (show && h.text ? '' : ' empty') + '" style="text-align:' + h.align + '">' +
          (show && h.text ? fill(h.text, k) : '') + '</div>';
      }
      html += '</div>';
    }
    if (sheets.innerHTML !== html) sheets.innerHTML = html;
  }

  function pageInfo() {
    const { L, P } = geometry();
    let current = 1;
    const s = getSelection();
    if (s.rangeCount && Ink.inDoc(s.anchorNode)) {
      const rect = docEl.getBoundingClientRect();
      const scale = rect.width / (L.w * 96) || 1;
      let caret = s.getRangeAt(0).getBoundingClientRect();
      if (!caret.height) caret = Ink.elementOf(s.anchorNode).getBoundingClientRect();
      current = Ink.clamp(Math.floor((caret.top - rect.top) / scale / P) + 1, 1, pages);
    }
    return { total: pages, current };
  }

  function pageOf(el) {
    const { L, P } = geometry();
    const rect = docEl.getBoundingClientRect();
    const scale = rect.width / (L.w * 96) || 1;
    return Ink.clamp(Math.floor((el.getBoundingClientRect().top - rect.top) / scale / P) + 1, 1, pages);
  }

  // Print uses the browser's own pagination; page margin boxes carry the
  // header and footer, with counter(page) for page numbers.
  function printCss() {
    const L = state.layout;
    const content = (text) => text.split(/(\{PAGE\}|\{PAGES\})/).filter(Boolean).map((p) =>
      p === '{PAGE}' ? 'counter(page)' : p === '{PAGES}' ? 'counter(pages)' : '"' + p.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\n\r]/g, ' ') + '"').join(' ');
    const box = (kind) => {
      const h = hf[kind];
      if (!h.text) return '';
      const side = h.align === 'left' ? 'left' : h.align === 'right' ? 'right' : 'center';
      return '@' + (kind === 'header' ? 'top' : 'bottom') + '-' + side + ' { content: ' + content(h.text) + '; font: 10pt Calibri, Carlito, sans-serif; color: #444; }';
    };
    const none = ['top', 'bottom'].map((v) => ['left', 'center', 'right'].map((s) => '@' + v + '-' + s + ' { content: none; }').join(' ')).join(' ');
    $('#printPage').textContent = '@page { size: ' + L.w + 'in ' + L.h + 'in; margin: ' + L.t + 'in ' + L.r + 'in ' + L.b + 'in ' + L.l + 'in; ' +
      box('header') + ' ' + box('footer') + ' }' + (hf.differentFirst ? ' @page :first { ' + none + ' }' : '');
  }

  function changedHF() {
    renderSheets();
    printCss();
    Ink.changed();
  }

  async function editHF(kind) {
    const h = hf[kind];
    const label = kind === 'header' ? 'Header' : 'Footer';
    const res = await Ink.dialog('Edit ' + label,
      '<label>' + label + ' text<input type="text" name="text" value="' + Ink.escHtml(h.text) + '" placeholder="e.g. Quarterly Report"></label>' +
      '<p class="muted hint">Type {PAGE} for the page number and {PAGES} for the number of pages.</p>' +
      '<label>Alignment<select name="align">' + ['left', 'center', 'right'].map((a) =>
        '<option value="' + a + '"' + (h.align === a ? ' selected' : '') + '>' + a[0].toUpperCase() + a.slice(1) + '</option>').join('') + '</select></label>' +
      '<label class="check"><input type="checkbox" name="differentFirst"' + (hf.differentFirst ? ' checked' : '') + '> Different first page</label>', 'OK');
    if (!res) return;
    hf[kind] = { text: res.text.slice(0, 300), align: res.align };
    hf.differentFirst = !!res.differentFirst;
    changedHF();
  }

  function pageNumbers(arg) {
    const strip = (t) => t.replace(/Page \{PAGE\} of \{PAGES\}|\{PAGE\}|\{PAGES\}/g, '').replace(/\s{2,}/g, ' ').trim();
    hf.header.text = strip(hf.header.text);
    hf.footer.text = strip(hf.footer.text);
    if (arg !== 'remove') {
      const kind = arg.startsWith('top') ? 'header' : 'footer';
      const token = arg === 'bottom-xofy' ? 'Page {PAGE} of {PAGES}' : '{PAGE}';
      const h = hf[kind];
      h.text = h.text ? h.text + '    ' + token : token;
      h.align = arg.endsWith('right') ? 'right' : 'center';
    }
    changedHF();
  }

  Ink.command('editHF', editHF);
  Ink.command('removeHF', (kind) => { hf[kind] = { text: '', align: 'center' }; changedHF(); });
  Ink.command('pageNumbers', pageNumbers);

  // Double-click in a page's top or bottom margin edits its header/footer.
  docEl.addEventListener('dblclick', (e) => {
    if (!enabled || state.readOnly) return;
    const { L, pageH, P, mt, mb } = geometry();
    const rect = docEl.getBoundingClientRect();
    const scale = rect.width / (L.w * 96) || 1;
    const y = (e.clientY - rect.top) / scale;
    const local = y - Math.floor(y / P) * P;
    if (local < mt * 0.8) { e.preventDefault(); editHF('header'); }
    else if (local > pageH - mb * 0.8 && local <= pageH) { e.preventDefault(); editHF('footer'); }
  });

  window.addEventListener('beforeprint', () => { enabled = false; paginate(); });
  window.addEventListener('afterprint', () => { enabled = true; paginate(); });
  // Late-loading fonts and pictures change heights.
  docEl.addEventListener('load', () => paginate(), true);
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => paginate());

  Ink.paginate = paginate;
  Ink.unpaginate = unpaginate;
  Ink.pageInfo = pageInfo;
  Ink.pageOf = pageOf;
  Ink.pageGeometry = geometry;
  Ink.setPagination = (on) => { enabled = on; paginate(); };
  Ink.on('clean', strip);
  Ink.on('layoutChanged', printCss);
  Ink.on('collect', (data) => {
    data.header = Object.assign({}, hf.header);
    data.footer = Object.assign({}, hf.footer);
    data.differentFirst = hf.differentFirst;
  });
  Ink.on('load', (data) => {
    const pick = (h) => ({ text: String((h && h.text) || '').slice(0, 300), align: ['left', 'center', 'right'].includes(h && h.align) ? h.align : 'center' });
    hf.header = pick(data.header);
    hf.footer = pick(data.footer);
    hf.differentFirst = !!data.differentFirst;
    printCss();
  });
})();
