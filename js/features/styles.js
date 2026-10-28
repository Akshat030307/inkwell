/*
 * Paragraph styles: Word 2013's built-in set plus custom styles. Definitions
 * drive the document CSS, the Styles gallery and word/styles.xml.
 */
(function () {
  'use strict';
  const { $, docEl } = Ink;
  const BUILTIN = ['normal', 'title', 'h1', 'h2', 'h3', 'quote'];
  const FONTS = ['Calibri', 'Calibri Light', 'Cambria', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Trebuchet MS', 'Courier New', 'Comic Sans MS'];
  const styleEl = document.head.appendChild(document.createElement('style'));
  let defs = defaults();

  function defaults() {
    const d = JSON.parse(JSON.stringify(Docx.DEFAULT_STYLES));
    d.custom = [];
    return d;
  }

  const family = (f) => {
    const name = String(f || 'Calibri').replace(/["\\;{}<>]/g, '');
    return '"' + name + '", ' + (/^Calibri/.test(name) ? 'Carlito, ' : '') + '"Segoe UI", Arial, sans-serif';
  };

  function props(d) {
    return 'font-family: ' + family(d.font) + '; font-size: ' + (+d.size || 11) + 'pt; color: ' + (d.color || '#000') + '; ' +
      'font-weight: ' + (d.bold ? 'bold' : 'normal') + '; font-style: ' + (d.italic ? 'italic' : 'normal') + '; ' +
      'line-height: ' + (+d.line || 1.08) + '; text-align: ' + (d.align || 'left') + '; ' +
      (d.letter ? 'letter-spacing: ' + d.letter + 'pt; ' : '') +
      'margin: ' + (+d.before || 0) + 'pt ' + (d.indent ? d.indent + 'in' : '0') + ' ' + (+d.after || 0) + 'pt ' + (d.indent ? d.indent + 'in' : '0') + ';';
  }

  function css() {
    const n = defs.normal;
    let s = '.doc { font-family: ' + family(n.font) + '; font-size: ' + n.size + 'pt; color: ' + n.color + '; line-height: ' + n.line + '; }\n';
    s += '.doc p, .doc div:not(.page-break):not(.column-break):not(.section):not(.toc) { ' + props(n) + ' }\n';
    s += '.doc ul, .doc ol { margin-bottom: ' + n.after + 'pt; }\n';
    s += '.doc .title { ' + props(defs.title) + ' }\n';
    s += '.doc h1 { ' + props(defs.h1) + ' }\n.doc h2 { ' + props(defs.h2) + ' }\n.doc h3 { ' + props(defs.h3) + ' }\n';
    s += '.doc blockquote { ' + props(defs.quote) + ' }\n';
    for (const c of defs.custom) s += '.doc [data-style="' + String(c.id).replace(/[^\w-]/g, '') + '"] { ' + props(c) + ' }\n';
    return s;
  }

  function apply() {
    styleEl.textContent = css();
    renderGallery();
    Ink.paginate();
  }

  const all = () => BUILTIN.map((k) => [k, defs[k]]).concat(defs.custom.map((c) => [c.id, c]));
  const find = (key) => (BUILTIN.includes(key) ? defs[key] : defs.custom.find((c) => c.id === key));

  function renderGallery() {
    $('#styleGallery').innerHTML = all().map(([key, d]) =>
      '<button data-cmd="blockStyle" data-arg="' + Ink.escHtml(key) + '" class="style-card" title="' + Ink.escHtml(d.name) + '">' +
      '<span class="sample" style="font-family: ' + Ink.escHtml(family(d.font)) + '; font-size: ' + Math.min(+d.size || 11, 20) + 'pt; color: ' +
      Ink.escHtml(d.color || '#000') + '; font-weight: ' + (d.bold ? 'bold' : 'normal') + '; font-style: ' + (d.italic ? 'italic' : 'normal') + '">' +
      (key === 'title' ? 'AaB' : key === 'normal' ? 'AaBbCcDc' : 'AaBbCc') + '</span><span>' + Ink.escHtml(d.name) + '</span></button>').join('');
    Ink.updateUi();
  }

  function currentBlock() {
    const s = getSelection();
    return s.rangeCount ? Ink.closestBlock(s.anchorNode) : null;
  }

  function styleForm(d, withName) {
    const pick = (v, list) => list.map((o) => '<option' + (String(o) === String(v) ? ' selected' : '') + '>' + o + '</option>').join('');
    return (withName ? '<label>Name<input type="text" name="name" required value="' + Ink.escHtml(d.name || '') + '"></label>' : '') +
      '<div class="form-grid">' +
      '<label>Font<select name="font">' + pick(d.font, FONTS.includes(d.font) ? FONTS : [d.font].concat(FONTS)) + '</select></label>' +
      '<label>Size (pt)<input type="number" name="size" min="4" max="96" step="0.5" value="' + d.size + '"></label>' +
      '<label>Color<input type="color" name="color" value="' + (d.color || '#000000') + '"></label>' +
      '<label>Alignment<select name="align">' + [['', 'Left'], ['center', 'Center'], ['right', 'Right'], ['justify', 'Justify']]
        .map(([v, t]) => '<option value="' + v + '"' + ((d.align || '') === v ? ' selected' : '') + '>' + t + '</option>').join('') + '</select></label>' +
      '<label>Space before (pt)<input type="number" name="before" min="0" max="200" value="' + (d.before || 0) + '"></label>' +
      '<label>Space after (pt)<input type="number" name="after" min="0" max="200" value="' + (d.after || 0) + '"></label>' +
      '<label>Line spacing<input type="number" name="line" min="0.5" max="5" step="0.01" value="' + (d.line || 1) + '"></label>' +
      '<div class="checks"><label><input type="checkbox" name="bold"' + (d.bold ? ' checked' : '') + '> Bold</label>' +
      '<label><input type="checkbox" name="italic"' + (d.italic ? ' checked' : '') + '> Italic</label></div></div>';
  }

  function fromForm(res, base) {
    return Object.assign({}, base, {
      font: res.font, size: Ink.clamp(parseFloat(res.size) || 11, 4, 96), color: res.color.toUpperCase(),
      align: res.align, before: Math.max(0, parseFloat(res.before) || 0), after: Math.max(0, parseFloat(res.after) || 0),
      line: Ink.clamp(parseFloat(res.line) || 1, 0.5, 5), bold: !!res.bold, italic: !!res.italic,
    }, res.name ? { name: res.name.trim().slice(0, 40) } : {});
  }

  async function modifyStyle() {
    const key = Ink.styleKey(currentBlock());
    const d = find(key) || defs.normal;
    const res = await Ink.dialog('Modify Style: ' + d.name, styleForm(d, !BUILTIN.includes(key)), 'OK');
    if (!res) return;
    const next = fromForm(res, d);
    if (BUILTIN.includes(key)) defs[key] = next;
    else defs.custom = defs.custom.map((c) => (c.id === key ? next : c));
    apply();
    Ink.changed();
  }

  // A new style takes the current paragraph's formatting, as Word does.
  async function createStyle() {
    const block = currentBlock();
    if (!block) { Ink.toast('Put the cursor in a paragraph first'); return; }
    let d;
    Ink.unpaginate();
    try {
      const cs = getComputedStyle(block);
      const rgb = cs.color.match(/\d+/g) || [0, 0, 0];
      d = {
        name: 'Style ' + (defs.custom.length + 1),
        font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(),
        size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2,
        color: '#' + rgb.slice(0, 3).map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase(),
        bold: parseInt(cs.fontWeight, 10) >= 600, italic: cs.fontStyle === 'italic',
        align: { center: 'center', right: 'right', justify: 'justify' }[cs.textAlign] || '',
        before: Math.round(parseFloat(cs.marginTop) * 0.75), after: Math.round(parseFloat(cs.marginBottom) * 0.75),
        line: Math.round(parseFloat(cs.lineHeight) / parseFloat(cs.fontSize) * 100) / 100 || 1.08,
      };
    } finally {
      Ink.paginate();
    }
    const res = await Ink.dialog('Create New Style from Formatting', styleForm(d, true), 'Create');
    if (!res || !res.name.trim()) return;
    const def = fromForm(res, { id: 's' + Date.now().toString(36) });
    defs.custom.push(def);
    apply();
    Ink.run('blockStyle', def.id);
    Ink.toast('Created style "' + def.name + '"');
  }

  function deleteStyle() {
    const key = Ink.styleKey(currentBlock());
    if (BUILTIN.includes(key)) { Ink.toast('Built-in styles can be modified but not deleted'); return; }
    const d = find(key);
    if (!d) return;
    defs.custom = defs.custom.filter((c) => c.id !== key);
    Ink.edit(() => {
      for (const el of Ink.$$('[data-style="' + key.replace(/[^\w-]/g, '') + '"]', docEl)) delete el.dataset.style;
    });
    apply();
    Ink.toast('Deleted style "' + d.name + '"');
  }

  Ink.command('modifyStyle', modifyStyle);
  Ink.command('createStyle', createStyle);
  Ink.command('deleteStyle', deleteStyle);
  Ink.command('resetStyles', () => {
    const custom = defs.custom;
    defs = defaults();
    defs.custom = custom;
    apply();
    Ink.changed();
    Ink.toast('Built-in styles restored');
  });

  Ink.styleCss = css;
  Ink.styleDefs = () => defs;
  Ink.on('collect', (data) => { data.styles = JSON.parse(JSON.stringify(defs)); });
  Ink.on('load', (data) => {
    const base = defaults();
    const s = data.styles || {};
    for (const k of BUILTIN) if (s[k]) base[k] = Object.assign({}, base[k], s[k], { name: base[k].name });
    base.custom = Array.isArray(s.custom) ? s.custom.filter((c) => c && c.id && c.name) : [];
    defs = base;
    apply();
  });
  Ink.on('start', apply);
})();
