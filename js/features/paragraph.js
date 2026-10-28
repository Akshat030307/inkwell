/*
 * Paragraph formatting: the Paragraph dialog, ruler indent markers, Page
 * Layout indent/spacing fields, borders and shading, plus the Clipboard
 * group (Format Painter, cut/copy/paste) and Change Case.
 */
(function () {
  'use strict';
  const { $, $$, docEl, state } = Ink;
  const PX_PER_IN = 96, PX_PER_PT = 96 / 72;
  const toPx = (v) => {
    const n = parseFloat(v);
    if (!isFinite(n)) return 0;
    return /pt$/.test(v) ? n * PX_PER_PT : /in$/.test(v) ? n * PX_PER_IN : n;
  };
  const round = (v, d = 100) => Math.round(v * d) / d;

  // Measured paragraph formatting, with page layout spacing set aside.
  function measure(block) {
    Ink.unpaginate();
    try {
      const cs = getComputedStyle(block);
      return {
        align: { center: 'center', right: 'right', justify: 'justify' }[cs.textAlign] || 'left',
        left: parseFloat(cs.marginLeft) || 0, right: parseFloat(cs.marginRight) || 0, indent: parseFloat(cs.textIndent) || 0,
        before: parseFloat(cs.marginTop) || 0, after: parseFloat(cs.marginBottom) || 0,
        line: round(parseFloat(cs.lineHeight) / parseFloat(cs.fontSize)) || 1.08,
      };
    } finally {
      Ink.paginate();
    }
  }

  // ------------------------------------------------------------ Paragraph dialog

  async function paragraphDialog() {
    const blocks = Ink.selectedBlocks();
    if (!blocks.length) { Ink.toast('Put the cursor in a paragraph first'); return; }
    const m = measure(blocks[0]);
    const hanging = m.indent < 0 ? -m.indent : 0;
    const special = m.indent > 0 ? 'first' : m.indent < 0 ? 'hanging' : 'none';
    const opt = (v, cur, label) => '<option value="' + v + '"' + (v === cur ? ' selected' : '') + '>' + label + '</option>';
    const res = await Ink.dialog('Paragraph',
      '<div class="form-grid">' +
      '<label>Alignment<select name="align">' + opt('left', m.align, 'Left') + opt('center', m.align, 'Centered') +
        opt('right', m.align, 'Right') + opt('justify', m.align, 'Justified') + '</select></label><span></span>' +
      '<label>Indent left (in)<input type="number" name="left" step="0.1" min="0" max="6" value="' + round((m.left - hanging) / PX_PER_IN) + '"></label>' +
      '<label>Indent right (in)<input type="number" name="right" step="0.1" min="0" max="6" value="' + round(m.right / PX_PER_IN) + '"></label>' +
      '<label>Special<select name="special">' + opt('none', special, '(none)') + opt('first', special, 'First line') + opt('hanging', special, 'Hanging') + '</select></label>' +
      '<label>By (in)<input type="number" name="by" step="0.1" min="0" max="4" value="' + round(Math.abs(m.indent) / PX_PER_IN) + '"></label>' +
      '<label>Spacing before (pt)<input type="number" name="before" min="0" max="200" value="' + round(m.before / PX_PER_PT, 1) + '"></label>' +
      '<label>Spacing after (pt)<input type="number" name="after" min="0" max="200" value="' + round(m.after / PX_PER_PT, 1) + '"></label>' +
      '<label>Line spacing (multiple)<input type="number" name="line" min="0.5" max="5" step="0.01" value="' + m.line + '"></label>' +
      '</div>', 'OK');
    if (!res) return;
    const by = Math.max(0, parseFloat(res.by) || 0) * PX_PER_IN;
    const left = Math.max(0, parseFloat(res.left) || 0) * PX_PER_IN;
    const isHanging = res.special === 'hanging';
    const values = {
      textAlign: res.align === 'left' ? '' : res.align,
      // Word's hanging indent keeps the first line at "Left" and wraps the
      // rest further in, so the CSS left margin includes the hanging amount.
      marginLeft: left + (isHanging ? by : 0),
      marginRight: Math.max(0, parseFloat(res.right) || 0) * PX_PER_IN,
      textIndent: res.special === 'first' ? by : isHanging ? -by : 0,
      marginTop: Math.max(0, parseFloat(res.before) || 0) * PX_PER_PT,
      marginBottom: Math.max(0, parseFloat(res.after) || 0) * PX_PER_PT,
      lineHeight: Ink.clamp(parseFloat(res.line) || 1.08, 0.5, 5),
    };
    const before = { marginLeft: m.left, marginRight: m.right, textIndent: m.indent, marginTop: m.before, marginBottom: m.after, lineHeight: m.line };
    Ink.edit(() => {
      for (const b of blocks) {
        b.style.textAlign = values.textAlign;
        // Only fields the user changed become direct formatting.
        for (const k of ['marginLeft', 'marginRight', 'textIndent', 'marginTop', 'marginBottom']) {
          if (Math.abs(values[k] - before[k]) > 0.5) b.style[k] = values[k] ? round(values[k]) + 'px' : (k === 'marginTop' || k === 'marginBottom' ? '0px' : '');
        }
        if (Math.abs(values.lineHeight - before.lineHeight) > 0.004) b.style.lineHeight = String(values.lineHeight);
      }
    });
  }

  // ------------------------------------------------------------ Page Layout fields

  const fields = { layIndL: 'marginLeft', layIndR: 'marginRight', laySpB: 'marginTop', laySpA: 'marginBottom' };
  for (const [id, prop] of Object.entries(fields)) {
    $('#' + id).addEventListener('change', (e) => {
      const v = Math.max(0, parseFloat(e.target.value) || 0);
      const px = id.startsWith('layInd') ? v * PX_PER_IN : v * PX_PER_PT;
      // Zero spacing is still direct formatting (it overrides the style); zero indent is not.
      const isSpacing = prop === 'marginTop' || prop === 'marginBottom';
      Ink.edit(() => { for (const b of Ink.selectedBlocks()) b.style[prop] = px ? round(px) + 'px' : isSpacing ? '0px' : ''; });
    });
  }

  function syncFields(block) {
    if (!block || document.activeElement && document.activeElement.closest && document.activeElement.closest('.layout-fields')) return;
    const st = block.style;
    const top = block.dataset.pg !== undefined ? block.dataset.pgOrig || '' : st.marginTop;
    const cs = getComputedStyle(block);
    $('#layIndL').value = round(toPx(st.marginLeft) / PX_PER_IN);
    $('#layIndR').value = round(toPx(st.marginRight) / PX_PER_IN);
    $('#laySpB').value = round((top ? toPx(top) : block.dataset.pg !== undefined ? 0 : parseFloat(cs.marginTop)) / PX_PER_PT, 1);
    $('#laySpA').value = round(parseFloat(cs.marginBottom) / PX_PER_PT, 1);
  }

  // ------------------------------------------------------------ ruler

  const ruler = $('#ruler');
  ruler.innerHTML = '<div class="ruler-band"><div class="ruler-margin" data-side="l"></div><div class="ruler-margin" data-side="r"></div>' +
    '<div class="ruler-nums"></div><div class="rm rm-first" data-m="first" title="First Line Indent"></div>' +
    '<div class="rm rm-left" data-m="left" title="Left Indent"></div><div class="rm rm-right" data-m="right" title="Right Indent"></div></div>';
  const markers = { first: $('.rm-first', ruler), left: $('.rm-left', ruler), right: $('.rm-right', ruler) };

  function drawRuler() {
    const L = state.layout;
    $('.ruler-margin[data-side="l"]', ruler).style.width = L.l + 'in';
    $('.ruler-margin[data-side="r"]', ruler).style.width = L.r + 'in';
    let nums = '';
    for (let i = -Math.floor(L.l); i <= Math.floor(L.w - L.l); i++) {
      if (i === 0) continue;
      nums += '<span style="left:' + (L.l + i) + 'in">' + Math.abs(i) + '</span>';
    }
    $('.ruler-nums', ruler).innerHTML = nums;
  }

  let rulerBlock = null;
  function placeMarkers(block) {
    rulerBlock = block;
    const L = state.layout;
    for (const m of Object.values(markers)) m.hidden = !block;
    if (!block) return;
    const cs = getComputedStyle(block);
    const left = parseFloat(cs.marginLeft) || 0, indent = parseFloat(cs.textIndent) || 0, right = parseFloat(cs.marginRight) || 0;
    markers.left.style.left = (L.l * PX_PER_IN + left) + 'px';
    markers.first.style.left = (L.l * PX_PER_IN + left + indent) + 'px';
    markers.right.style.left = ((L.w - L.r) * PX_PER_IN - right) + 'px';
  }

  ruler.addEventListener('pointerdown', (e) => {
    const marker = e.target.closest('.rm');
    if (!marker || !rulerBlock) return;
    e.preventDefault();
    marker.setPointerCapture(e.pointerId);
    Ink.restoreRange();
    const blocks = Ink.selectedBlocks();
    const L = state.layout;
    const band = $('.ruler-band', ruler).getBoundingClientRect();
    const scale = band.width / (L.w * PX_PER_IN) || 1;
    const cs = getComputedStyle(rulerBlock);
    const start = { left: parseFloat(cs.marginLeft) || 0, indent: parseFloat(cs.textIndent) || 0 };
    const firstAbs = start.left + start.indent;
    const kind = marker.dataset.m;
    const move = (ev) => {
      // Snap to sixteenths of an inch, measured from the left margin.
      const x = Math.round(((ev.clientX - band.left) / scale - L.l * PX_PER_IN) / 6) * 6;
      for (const b of blocks) {
        if (kind === 'left') {
          const left = Math.max(0, x);
          b.style.marginLeft = left ? left + 'px' : '';
          b.style.textIndent = firstAbs - left ? (firstAbs - left) + 'px' : '';
        } else if (kind === 'first') {
          const indent = x - start.left;
          b.style.textIndent = indent ? indent + 'px' : '';
        } else {
          const right = Math.max(0, (L.w - L.l - L.r) * PX_PER_IN - x);
          b.style.marginRight = right ? right + 'px' : '';
        }
      }
      placeMarkers(rulerBlock);
    };
    const up = () => {
      marker.removeEventListener('pointermove', move);
      marker.removeEventListener('pointerup', up);
      Ink.edit(() => {});
    };
    marker.addEventListener('pointermove', move);
    marker.addEventListener('pointerup', up);
  });

  function setRuler(on) {
    ruler.hidden = !on;
    $('#rulerBtn').classList.toggle('on', on);
    try { localStorage.setItem('inkwell:ruler', on ? '1' : '0'); } catch (e) { /* private mode */ }
  }
  Ink.command('toggleRuler', () => setRuler(ruler.hidden));

  // ------------------------------------------------------------ borders & shading

  function border(side) {
    Ink.edit(() => {
      const blocks = Ink.selectedBlocks();
      const sides = side === 'box' ? ['Top', 'Right', 'Bottom', 'Left'] : side === 'none' ? [] : [side[0].toUpperCase() + side.slice(1)];
      const all = sides.length && blocks.every((b) => sides.every((s) => b.style['border' + s + 'Style'] === 'solid'));
      for (const b of blocks) {
        if (side === 'none') { for (const s of ['Top', 'Right', 'Bottom', 'Left']) b.style['border' + s] = ''; continue; }
        for (const s of sides) b.style['border' + s] = all ? '' : '1px solid #000';
      }
    });
  }

  let shade = '#D9E2F3';
  function shading(color) {
    if (color) shade = color === 'none' ? '' : color;
    Ink.edit(() => { for (const b of Ink.selectedBlocks()) b.style.backgroundColor = shade; });
  }

  // ------------------------------------------------------------ Format Painter

  let painter = null;
  const painterBtn = $('#painterBtn');

  function capture() {
    const s = getSelection();
    const el = s.rangeCount && Ink.inDoc(s.anchorNode) ? Ink.elementOf(s.anchorNode) : null;
    if (!el) return null;
    const cs = getComputedStyle(el);
    const block = Ink.closestBlock(el) || docEl;
    const p = {
      bold: parseInt(cs.fontWeight, 10) >= 600, italic: cs.fontStyle === 'italic', underline: false, strike: false, bg: '',
      color: cs.color, font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(),
      size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2,
    };
    for (let e = el; e && e !== block; e = e.parentElement) {
      const deco = e.style.textDecorationLine || e.style.textDecoration || '';
      if (e.tagName === 'U' || /underline/.test(deco)) p.underline = true;
      if (e.tagName === 'S' || e.tagName === 'STRIKE' || /line-through/.test(deco)) p.strike = true;
      if (!p.bg && e.style.backgroundColor) p.bg = e.style.backgroundColor;
    }
    return p;
  }

  function paint(p) {
    Ink.edit(() => {
      Ink.exec('removeFormat');
      const s = getSelection();
      const block = s.rangeCount ? Ink.closestBlock(s.anchorNode) || docEl : docEl;
      const base = getComputedStyle(block);
      if (p.bold !== parseInt(base.fontWeight, 10) >= 600) Ink.exec('bold');
      if (p.italic !== (base.fontStyle === 'italic')) Ink.exec('italic');
      if (p.underline) Ink.exec('underline');
      if (p.strike) Ink.exec('strikeThrough');
      if (p.color !== base.color) Ink.exec('foreColor', p.color);
      if (p.bg) Ink.exec('hiliteColor', p.bg);
      if (p.font !== base.fontFamily.split(',')[0].replace(/["']/g, '').trim()) Ink.exec('fontName', p.font);
      if (p.size !== Math.round(parseFloat(base.fontSize) * 0.75 * 2) / 2) Ink.setFontSize(p.size);
    });
  }

  function stopPainter() {
    painter = null;
    painterBtn.classList.remove('on');
    document.body.classList.remove('painting');
  }

  function startPainter(sticky) {
    const p = capture();
    if (!p) { Ink.toast('Click in some formatted text first'); return; }
    painter = { props: p, sticky };
    painterBtn.classList.add('on');
    document.body.classList.add('painting');
  }

  painterBtn.addEventListener('dblclick', () => startPainter(true));
  docEl.addEventListener('mouseup', () => {
    if (!painter) return;
    setTimeout(() => {
      const s = getSelection();
      if (!s.rangeCount || s.isCollapsed) return;
      paint(painter.props);
      if (!painter.sticky) stopPainter();
    }, 0);
  });
  Ink.on('escape', () => { if (painter) stopPainter(); });

  // ------------------------------------------------------------ Change Case

  function changeCase(mode) {
    const r = Ink.currentRange();
    if (!r || r.collapsed) { Ink.toast('Select the text to change'); return; }
    const range = r.cloneRange();
    Ink.edit(() => {
      const root = range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer;
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const parts = [];
      let n;
      while ((n = walker.nextNode())) {
        if (!range.intersectsNode(n)) continue;
        const a = n === range.startContainer ? range.startOffset : 0;
        const b = n === range.endContainer ? range.endOffset : n.length;
        if (b > a) parts.push({ n, a, b });
      }
      const text = parts.map((p) => p.n.data.slice(p.a, p.b)).join('');
      // Transform character by character so every text node keeps its length.
      const out = [];
      let sentenceStart = true, wordStart = true;
      for (const ch of text) {
        let c = ch;
        if (mode === 'upper') c = ch.toUpperCase();
        else if (mode === 'lower') c = ch.toLowerCase();
        else if (mode === 'toggle') c = ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase();
        else if (mode === 'title') c = wordStart ? ch.toUpperCase() : ch.toLowerCase();
        else if (mode === 'sentence') c = sentenceStart && /\p{L}/u.test(ch) ? ch.toUpperCase() : ch.toLowerCase();
        if (c.length !== ch.length) c = ch;
        if (/\p{L}/u.test(ch)) sentenceStart = false;
        if (/[.!?]/.test(ch)) sentenceStart = true;
        wordStart = /\s|[-(\["“]/.test(ch);
        out.push(c);
      }
      let i = 0;
      const joined = out.join('');
      for (const p of parts) {
        const len = p.b - p.a;
        p.n.data = p.n.data.slice(0, p.a) + joined.slice(i, i + len) + p.n.data.slice(p.b);
        i += len;
      }
    });
  }

  // ------------------------------------------------------------ clipboard

  async function paste() {
    Ink.restoreRange();
    try {
      if (navigator.clipboard && navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          if (item.types.includes('text/html')) {
            const html = await (await item.getType('text/html')).text();
            Ink.edit(() => Ink.exec('insertHTML', Ink.sanitize(html)));
            return;
          }
        }
      }
      const text = await navigator.clipboard.readText();
      if (text) Ink.edit(() => (/\n/.test(text) ? Ink.exec('insertHTML', Ink.textToHtml(text)) : Ink.exec('insertText', text)));
    } catch (e) {
      Ink.toast('Your browser blocked clipboard access — use Ctrl+V to paste');
    }
  }

  Ink.command('paragraphDialog', paragraphDialog);
  Ink.command('border', border);
  Ink.command('shading', (arg) => shading(arg));
  Ink.command('formatPainter', () => (painter ? stopPainter() : startPainter(false)));
  Ink.command('changeCase', changeCase);
  Ink.command('cut', () => { Ink.restoreRange(); Ink.flushHistory(); if (!Ink.exec('cut')) Ink.toast('Use Ctrl+X to cut'); else { Ink.record(); Ink.changed(); } });
  Ink.command('copy', () => { Ink.restoreRange(); if (!Ink.exec('copy')) Ink.toast('Use Ctrl+C to copy'); });
  Ink.command('paste', paste);

  Ink.on('start', () => {
    $('#shadingMenu').innerHTML = '<div class="menu-sub">Theme Colors</div>' + Ink.swatches(Ink.THEME_COLORS.concat(['#D9E2F3', '#FBE4D5', '#E2EFD9', '#FFF2CC', '#DEEAF6', '#EDEDED', '#F2F2F2', '#D0CECE', '#B4C6E7', '#C5E0B3']), 'shading') +
      '<div class="menu-sub">Standard Colors</div>' + Ink.swatches(Ink.STANDARD_COLORS, 'shading') +
      '<button class="wide-item" data-cmd="shading" data-arg="none"><i></i>No Color</button>';
    let on = true;
    try { on = localStorage.getItem('inkwell:ruler') !== '0'; } catch (e) { /* default on */ }
    setRuler(on);
    drawRuler();
  });
  Ink.on('layoutChanged', drawRuler);
  Ink.on('ui', (el) => {
    const block = Ink.closestBlock(el);
    placeMarkers(block);
    syncFields(block);
  });
  Ink.on('paginated', () => { if (rulerBlock && docEl.contains(rulerBlock)) placeMarkers(rulerBlock); });
})();
