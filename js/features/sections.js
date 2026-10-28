/*
 * Columns. A multi-column region is a <div class="section" data-cols="n">
 * holding top-level blocks; docx.js writes it as a continuous section.
 */
(function () {
  'use strict';
  const { docEl } = Ink;

  // The top-level child of the document that contains `node`.
  function topOf(node) {
    let n = Ink.elementOf(node);
    while (n && n.parentElement !== docEl) n = n.parentElement;
    return n;
  }

  function unwrap(section) {
    while (section.firstChild) section.parentNode.insertBefore(section.firstChild, section);
    section.remove();
  }

  function columns(arg) {
    const n = Ink.clamp(parseInt(arg, 10) || 1, 1, 3);
    Ink.edit(() => Ink.keepingSelection(() => {
      const s = getSelection();
      if (!s.rangeCount) return false;
      const r = s.getRangeAt(0);
      const section = Ink.closestIn(r.startContainer, '.section');
      if (section) {
        if (n === 1) unwrap(section); else section.dataset.cols = String(n);
        return true;
      }
      if (n === 1) return false;
      let first, last;
      if (r.collapsed) {
        // No selection: the whole run of blocks around the cursor, as Word
        // applies columns to the current section.
        first = last = topOf(r.startContainer);
        if (!first) return false;
        while (first.previousElementSibling && !first.previousElementSibling.classList.contains('section')) first = first.previousElementSibling;
        while (last.nextElementSibling && !last.nextElementSibling.classList.contains('section')) last = last.nextElementSibling;
      } else {
        first = topOf(r.startContainer);
        last = topOf(r.endContainer);
      }
      if (!first || !last || first.classList.contains('section')) return false;
      const div = document.createElement('div');
      div.className = 'section';
      div.dataset.cols = String(n);
      first.before(div);
      let el = first;
      while (el) {
        const next = el === last ? null : el.nextElementSibling;
        div.appendChild(el);
        el = next;
      }
      if (!div.nextElementSibling) div.after(Object.assign(document.createElement('p'), { innerHTML: '<br>' }));
      return true;
    }));
  }

  Ink.command('columns', columns);
  Ink.command('columnBreak', () => {
    const s = getSelection();
    if (!(s.rangeCount && Ink.closestIn(s.anchorNode, '.section')) && !(Ink.currentRange() && Ink.closestIn(Ink.currentRange().startContainer, '.section'))) {
      Ink.toast('Column breaks work inside columns — set Columns to Two or Three first');
      return;
    }
    Ink.edit(() => Ink.exec('insertHTML', '<div class="column-break" contenteditable="false"></div><p><br></p>'));
  });
})();
