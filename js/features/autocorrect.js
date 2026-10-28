/*
 * AutoCorrect as you type: smart quotes, dashes and symbols, automatic lists
 * ("* " or "1. " at the start of a paragraph) and automatic hyperlinks.
 * Each correction is its own undo step, so Ctrl+Z restores what was typed.
 */
(function () {
  'use strict';
  const SYMBOLS = [['(c)', '©'], ['(r)', '®'], ['(tm)', '™'], ['...', '…'],
    ['-->', '→'], ['<--', '←'], ['==>', '⇒'], ['<=>', '⇔']];

  function caretTo(node, offset) {
    const r = document.createRange();
    r.setStart(node, offset);
    r.collapse(true);
    Ink.setRange(r);
  }

  // Replace `len` characters ending `end` characters before the caret.
  function replaceBefore(node, caret, len, text, end = 0) {
    Ink.record();
    const at = caret - end - len;
    node.data = node.data.slice(0, at) + text + node.data.slice(at + len);
    caretTo(node, caret - len + text.length);
  }

  function autoList(block, kind, format) {
    setTimeout(() => {
      Ink.edit(() => {
        block.innerHTML = '<br>';
        Ink.placeCaret(block);
        Ink.exec(kind === 'ul' ? 'insertUnorderedList' : 'insertOrderedList');
        const s = getSelection();
        const list = s.rangeCount && Ink.closestIn(s.anchorNode, kind);
        if (list && format) list.dataset.format = format;
      });
    }, 0);
  }

  function autoLink(node, caret, url) {
    Ink.record();
    const r = document.createRange();
    r.setStart(node, caret - 1 - url.length);
    r.setEnd(node, caret - 1);
    Ink.setRange(r);
    Ink.exec('createLink', /^www\./i.test(url) ? 'https://' + url : url);
    const a = Ink.closestIn(getSelection().anchorNode, 'a') || Ink.closestIn(getSelection().focusNode, 'a');
    const after = a && a.nextSibling;
    if (after && after.nodeType === 3) caretTo(after, Math.min(1, after.length));
    else getSelection().collapseToEnd();
  }

  Ink.on('input', (e) => {
    const o = Ink.options.autocorrect || {};
    if (e.inputType !== 'insertText' || !e.data || Ink.state.readOnly) return;
    const s = getSelection();
    if (!s.rangeCount || !s.isCollapsed) return;
    const node = s.anchorNode, caret = s.anchorOffset;
    if (!node || node.nodeType !== 3 || Ink.closestIn(node, 'del.tc')) return;
    const before = node.data.slice(0, caret);
    const ch = e.data.slice(-1);

    if (o.quotes && (ch === '"' || ch === "'")) {
      const prev = before.slice(-2, -1);
      const opening = !prev || /[\s(\[{—–-]/.test(prev);
      const smart = ch === '"' ? (opening ? '“' : '”') : (opening ? '‘' : '’');
      replaceBefore(node, caret, 1, smart);
      return;
    }

    if (o.symbols) {
      const lower = before.toLowerCase();
      for (const [k, v] of SYMBOLS) {
        if (lower.endsWith(k)) { replaceBefore(node, caret, k.length, v); return; }
      }
      // word--word becomes an em dash; " -- " or " - " between words an en dash.
      if (/[^\s-]--[^\s-]$/.test(before)) { replaceBefore(node, caret, 2, '—', 1); return; }
      if (/\S\s--\s\S$/.test(before)) { replaceBefore(node, caret, 2, '–', 2); return; }
      if (/\S\s-\s\S$/.test(before)) { replaceBefore(node, caret, 1, '–', 2); return; }
    }

    if (o.lists && ch === ' ') {
      const block = Ink.closestBlock(node);
      if (block && block.tagName === 'P' && !block.closest('li,td,th,.toc') && block.textContent.replace(/ /g, ' ') === before.replace(/ /g, ' ')) {
        const t = before.replace(/ /g, ' ');
        if (/^[*\-•] $/.test(t)) { autoList(block, 'ul'); return; }
        if (/^1[.)] $/.test(t)) { autoList(block, 'ol'); return; }
        if (/^a[.)] $/.test(t)) { autoList(block, 'ol', 'lower-alpha'); return; }
        if (/^A[.)] $/.test(t)) { autoList(block, 'ol', 'upper-alpha'); return; }
      }
    }

    if (o.links && (ch === ' ' || ch === ' ') && !Ink.closestIn(node, 'a')) {
      const m = before.match(/(?:^|[\s(])((?:https?:\/\/|www\.)[^\s<>"]*[^\s<>".,;:!?)'])[\s ]$/i);
      // Deferred: createLink can't run inside another editing command.
      if (m) setTimeout(() => { if (node.isConnected && node.data.slice(0, caret) === before) autoLink(node, caret, m[1]); }, 0);
    }
  });
})();
