/* Bullet, numbering and multilevel list libraries. */
(function () {
  'use strict';

  function listAt(kind) {
    const s = getSelection();
    return s.rangeCount ? Ink.closestIn(s.anchorNode, kind) : null;
  }

  // The outermost list of the same kind, so nested levels share one format.
  function outermost(list) {
    let top = list;
    for (let p = list.parentElement; p && Ink.inDoc(p); p = p.parentElement) {
      if (p.tagName === list.tagName) top = p;
    }
    return top;
  }

  function listFormat(arg) {
    Ink.edit(() => {
      if (arg === 'none') {
        const list = listAt('ul,ol');
        if (list) Ink.exec(list.tagName === 'UL' ? 'insertUnorderedList' : 'insertOrderedList');
        return;
      }
      const [kind, fmt] = arg.split(':');
      let list = listAt(kind);
      if (!list) {
        Ink.exec(kind === 'ul' ? 'insertUnorderedList' : 'insertOrderedList');
        list = listAt(kind);
      }
      if (!list) return;
      const top = outermost(list);
      if (kind === 'ul') {
        if (fmt === 'disc') delete top.dataset.bullet; else top.dataset.bullet = fmt;
      } else {
        if (fmt === 'decimal') delete top.dataset.format; else top.dataset.format = fmt;
      }
    });
  }

  Ink.command('listFormat', listFormat);
})();
