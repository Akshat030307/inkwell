/*
 * REVIEW tab: comments and Track Changes.
 *
 * Comments anchor to <span class="comment" data-cid>; their text lives in
 * document data and becomes word/comments.xml. Tracked insertions and
 * deletions are <ins class="tc"> / <del class="tc"> with author and date,
 * written as <w:ins> / <w:del>.
 */
(function () {
  'use strict';
  const { $, $$, docEl } = Ink;
  const pane = $('#commentsPane');
  let comments = {};
  let track = false;
  let showComments = true;
  let activeCid = null;

  const me = () => Ink.options.author || 'Author';
  const now = () => new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const anchors = (cid) => $$('span.comment[data-cid="' + String(cid).replace(/[^\w-]/g, '') + '"]', docEl);

  // Text nodes inside a range, split so each lies wholly within it.
  function textPieces(range) {
    const root = range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer;
    const nodes = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = walker.nextNode())) if (range.intersectsNode(n) && !n.parentElement.closest('.toc')) nodes.push(n);
    const pieces = [];
    const { startContainer: sc, startOffset: so, endContainer: ec, endOffset: eo } = range;
    for (let node of nodes) {
      const start = node === sc ? so : 0;
      const end = node === ec ? eo : node.length;
      if (start >= end) continue;
      if (end < node.length) node.splitText(end);
      if (start > 0) node = node.splitText(start);
      pieces.push(node);
    }
    return pieces;
  }

  function wrap(node, tag, attrs) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'className') el.className = v; else el.dataset[k] = v;
    }
    node.parentNode.insertBefore(el, node);
    el.appendChild(node);
    return el;
  }

  function unwrap(el) {
    while (el.firstChild) el.parentNode.insertBefore(el.firstChild, el);
    el.remove();
  }

  // ------------------------------------------------------------ comments

  function liveIds() {
    const seen = [];
    for (const span of $$('span.comment[data-cid]', docEl)) {
      const cid = span.dataset.cid;
      if (comments[cid] && !seen.includes(cid)) seen.push(cid);
    }
    return seen;
  }

  function when(iso) {
    const d = new Date(iso);
    return isNaN(d) ? '' : d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) + ' ' + d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  }

  function renderComments() {
    const ids = liveIds();
    pane.hidden = !showComments || !ids.length;
    $('#showCommentsBtn').classList.toggle('on', showComments);
    const focused = document.activeElement && document.activeElement.closest('#commentList') ? document.activeElement.dataset.cid : null;
    if (focused) return;   // don't rebuild under someone typing a comment
    $('#commentList').innerHTML = ids.map((cid) => {
      const c = comments[cid];
      return '<div class="comment-card' + (cid === activeCid ? ' active' : '') + (c.done ? ' done' : '') + '" data-card="' + cid + '">' +
        '<div class="cc-head"><strong>' + Ink.escHtml(c.author) + '</strong><span>' + when(c.date) + '</span></div>' +
        '<textarea data-cid="' + cid + '" rows="2" placeholder="Type your comment">' + Ink.escHtml(c.text) + '</textarea>' +
        '<div class="cc-actions"><button data-cc="done" data-cid="' + cid + '">' + (c.done ? 'Reopen' : 'Resolve') + '</button>' +
        '<button data-cc="delete" data-cid="' + cid + '">Delete</button></div></div>';
    }).join('');
  }

  function setActive(cid) {
    activeCid = cid;
    for (const s of $$('span.comment', docEl)) s.classList.toggle('active', !!cid && s.dataset.cid === cid);
    for (const c of $$('.comment-card', pane)) c.classList.toggle('active', c.dataset.card === cid);
  }

  function newComment() {
    const r = Ink.currentRange();
    if (!r || r.collapsed) { Ink.toast('Select the text you want to comment on'); return; }
    const range = r.cloneRange();
    const cid = 'c' + Date.now().toString(36);
    comments[cid] = { author: me(), date: now(), text: '', done: false };
    Ink.edit(() => {
      for (const p of textPieces(range)) wrap(p, 'span', { className: 'comment', cid });
    });
    if (!anchors(cid).length) {
      delete comments[cid];
      Ink.toast('Comments can’t be attached to that part of the document');
      return;
    }
    showComments = true;
    renderComments();
    setActive(cid);
    const ta = $('textarea[data-cid="' + cid + '"]', pane);
    if (ta) ta.focus();
  }

  function deleteComment(cid) {
    cid = cid || activeCid;
    if (!cid) { Ink.toast('Click inside a comment first'); return; }
    Ink.edit(() => { for (const s of anchors(cid)) unwrap(s); }, { focus: false });
    delete comments[cid];
    activeCid = null;
    renderComments();
  }

  function stepComment(dir) {
    const ids = liveIds();
    if (!ids.length) { Ink.toast('There are no comments'); return; }
    const i = ids.indexOf(activeCid);
    const cid = ids[(i + dir + ids.length) % ids.length] || ids[0];
    const span = anchors(cid)[0];
    if (span) {
      span.scrollIntoView({ block: 'center', behavior: 'smooth' });
      docEl.focus({ preventScroll: true });
      Ink.placeCaret(span);
    }
    setActive(cid);
  }

  pane.addEventListener('input', (e) => {
    const cid = e.target.dataset.cid;
    if (cid && comments[cid]) { comments[cid].text = e.target.value.slice(0, 2000); Ink.changed(); }
  });
  pane.addEventListener('focusout', () => setTimeout(renderComments, 0));
  pane.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-cc]');
    if (btn) {
      const cid = btn.dataset.cid;
      if (btn.dataset.cc === 'delete') deleteComment(cid);
      else { comments[cid].done = !comments[cid].done; Ink.changed(); renderComments(); }
      return;
    }
    const card = e.target.closest('[data-card]');
    if (card) {
      setActive(card.dataset.card);
      const span = anchors(card.dataset.card)[0];
      if (span) span.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  });

  // ------------------------------------------------------------ track changes

  function caretAt(node, offset) {
    const r = document.createRange();
    r.setStart(node, offset);
    r.collapse(true);
    Ink.setRange(r);
  }

  // Mark a range deleted. Text the current author inserted is really removed.
  function markDeleted(range) {
    const pieces = textPieces(range);
    let first = null, last = null;
    for (const p of pieces) {
      const ins = p.parentElement.closest('ins.tc');
      if (ins && ins.dataset.author === me()) {
        const anchor = document.createTextNode('');
        p.parentNode.insertBefore(anchor, p);
        p.remove();
        if (!ins.textContent) { ins.parentNode.insertBefore(anchor, ins); ins.remove(); }
        first = first || anchor;
        last = anchor;
        continue;
      }
      if (p.parentElement.closest('del.tc')) { first = first || p; last = p; continue; }
      // Grow an adjacent deletion by the same author instead of starting a
      // new one, so repeated Backspace makes a single change.
      const mine = (n) => n && n.nodeType === 1 && n.matches('del.tc') && n.dataset.author === me();
      let del;
      if (mine(p.previousSibling)) { del = p.previousSibling; del.appendChild(p); }
      else if (mine(p.nextSibling)) { del = p.nextSibling; del.insertBefore(p, del.firstChild); }
      else del = wrap(p, 'del', { className: 'tc', author: me(), date: now() });
      del.normalize();
      first = first || del;
      last = del;
    }
    return { first, last };
  }

  function insertTracked(text) {
    const s = getSelection();
    let r = s.getRangeAt(0);
    const del = Ink.closestIn(r.startContainer, 'del.tc');
    if (del) { r = document.createRange(); r.setStartAfter(del); r.collapse(true); }
    const ins = Ink.closestIn(r.startContainer, 'ins.tc');
    if (ins && ins.dataset.author === me() && r.startContainer.nodeType === 3) {
      r.startContainer.insertData(r.startOffset, text);
      caretAt(r.startContainer, r.startOffset + text.length);
      return;
    }
    const el = document.createElement('ins');
    el.className = 'tc';
    el.dataset.author = me();
    el.dataset.date = now();
    el.textContent = text;
    r.insertNode(el);
    caretAt(el.firstChild, text.length);
  }

  let paginateTimer = null;
  function afterTrackedEdit() {
    Ink.recordSoon();
    Ink.changed();
    clearTimeout(paginateTimer);
    paginateTimer = setTimeout(() => Ink.paginate(), 120);
  }

  docEl.addEventListener('beforeinput', (e) => {
    if (!track || Ink.state.readOnly) return;
    const t = e.inputType;
    if (t === 'insertText' || t === 'insertReplacementText') {
      const text = e.data != null ? e.data : (e.dataTransfer && e.dataTransfer.getData('text/plain')) || '';
      if (!text) return;
      e.preventDefault();
      const s = getSelection();
      if (!s.rangeCount) return;
      const r = s.getRangeAt(0);
      if (!r.collapsed) {
        const { last } = markDeleted(r);
        if (last) { const c = document.createRange(); c.setStartAfter(last); c.collapse(true); Ink.setRange(c); }
      }
      insertTracked(text);
      afterTrackedEdit();
    } else if (t.startsWith('delete')) {
      const targets = e.getTargetRanges ? e.getTargetRanges() : [];
      if (!targets.length) return;
      const sr = targets[0];
      const r = document.createRange();
      r.setStart(sr.startContainer, sr.startOffset);
      r.setEnd(sr.endContainer, sr.endOffset);
      if (r.collapsed || !r.toString()) return;   // paragraph joins stay untracked
      e.preventDefault();
      const { first, last } = markDeleted(r);
      const c = document.createRange();
      if (/Backward/.test(t) && first) c.setStartBefore(first);
      else if (last) c.setStartAfter(last);
      c.collapse(true);
      Ink.setRange(c);
      afterTrackedEdit();
    }
  });

  docEl.addEventListener('paste', (e) => {
    if (!track || !e.clipboardData) return;
    const text = e.clipboardData.getData('text/plain');
    e.preventDefault();
    e.stopImmediatePropagation();
    if (!text) return;
    const s = getSelection();
    if (s.rangeCount && !s.getRangeAt(0).collapsed) markDeleted(s.getRangeAt(0));
    insertTracked(text.replace(/\r\n?/g, '\n'));
    afterTrackedEdit();
  }, true);

  function setTrack(on) {
    track = on;
    $('#trackBtn').classList.toggle('on', on);
    $('#stTrack').hidden = !on;
  }

  const changes = () => $$('ins.tc, del.tc', docEl);
  const accept = (el) => (el.tagName === 'INS' ? unwrap(el) : el.remove());
  const reject = (el) => (el.tagName === 'INS' ? el.remove() : unwrap(el));

  function current() {
    const r = Ink.currentRange();
    return r ? changes().filter((el) => r.intersectsNode(el)) : [];
  }

  function resolve(fn, all) {
    const list = all ? changes() : current();
    if (!list.length) { Ink.toast(all ? 'There are no tracked changes' : 'Put the cursor in a tracked change'); return; }
    Ink.edit(() => { for (const el of list) fn(el); });
    Ink.toast((all ? list.length + ' change' + (list.length === 1 ? '' : 's') : 'Change') + (fn === accept ? ' accepted' : ' rejected'));
  }

  function stepChange(dir) {
    const list = changes();
    if (!list.length) { Ink.toast('There are no tracked changes'); return; }
    const r = Ink.currentRange();
    const caret = r && (dir > 0 ? r.endContainer : r.startContainer);
    let target = null;
    if (caret) {
      // Next: the first change after the caret. Previous: the last change
      // before it that doesn't contain it.
      const after = (el) => el.compareDocumentPosition(caret) & Node.DOCUMENT_POSITION_PRECEDING;
      const before = (el) => (el.compareDocumentPosition(caret) & Node.DOCUMENT_POSITION_FOLLOWING) && !el.contains(caret);
      target = dir > 0 ? list.find(after) : list.slice().reverse().find(before);
    }
    target = target || (dir > 0 ? list[0] : list[list.length - 1]);
    const sel = document.createRange();
    sel.selectNodeContents(target);
    docEl.focus({ preventScroll: true });
    Ink.setRange(sel);
    target.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }

  Ink.command('newComment', newComment);
  Ink.command('deleteComment', () => deleteComment());
  Ink.command('prevComment', () => stepComment(-1));
  Ink.command('nextComment', () => stepComment(1));
  Ink.command('showComments', () => { showComments = !showComments; renderComments(); });
  Ink.command('trackChanges', () => { setTrack(!track); Ink.changed(); Ink.toast('Track Changes ' + (track ? 'on' : 'off')); });
  Ink.command('showMarkup', () => {
    const final = document.body.classList.toggle('tc-final');
    $('#markupBtn').classList.toggle('on', !final);
  });
  Ink.command('acceptChange', () => resolve(accept, false));
  Ink.command('acceptAll', () => resolve(accept, true));
  Ink.command('rejectChange', () => resolve(reject, false));
  Ink.command('rejectAll', () => resolve(reject, true));
  Ink.command('prevChange', () => stepChange(-1));
  Ink.command('nextChange', () => stepChange(1));

  Ink.on('ui', (el) => {
    const span = Ink.closestIn(el, 'span.comment');
    if ((span ? span.dataset.cid : null) !== activeCid && (span || activeCid)) setActive(span ? span.dataset.cid : null);
  });
  Ink.on('change', () => { clearTimeout(renderComments.t); renderComments.t = setTimeout(renderComments, 200); });
  Ink.on('collect', (data) => {
    const live = {};
    for (const cid of liveIds()) live[cid] = Object.assign({}, comments[cid]);
    // Keep comments whose anchors were removed, so undo can bring them back.
    data.comments = Object.assign({}, comments, live);
    data.track = track;
  });
  Ink.on('load', (data) => {
    comments = {};
    for (const [cid, c] of Object.entries(data.comments || {})) {
      if (/^[\w-]{1,40}$/.test(cid) && c) comments[cid] = { author: String(c.author || 'Author').slice(0, 60), date: String(c.date || ''), text: String(c.text || '').slice(0, 2000), done: !!c.done };
    }
    activeCid = null;
    setTrack(!!data.track);
    document.body.classList.remove('tc-final');
    $('#markupBtn').classList.add('on');
    renderComments();
  });
  Ink.on('start', () => $('#markupBtn').classList.add('on'));
})();
