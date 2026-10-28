/*
 * Inkwell DOCX — a minimal Office Open XML (ECMA-376) reader and writer.
 *
 * Mirrors the structure Microsoft Word itself writes:
 *   [Content_Types].xml, _rels/.rels, docProps/{core,app}.xml,
 *   word/document.xml, word/styles.xml, word/numbering.xml, word/settings.xml,
 *   word/header1.xml, word/footer1.xml, word/comments.xml,
 *   word/_rels/document.xml.rels, word/media/*
 *
 * Word's document model: a body is a flat list of paragraphs (<w:p>) and
 * tables (<w:tbl>). A paragraph holds runs (<w:r>); every run is a stretch of
 * text sharing one set of character properties (<w:rPr>). Lists are ordinary
 * paragraphs that point at numbering.xml via <w:numPr>. Sections end at a
 * paragraph carrying <w:sectPr>; the body's own sectPr closes the last one.
 * The editor's DOM is converted to and from that model here.
 */
(function (global) {
  'use strict';

  const NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  };
  const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';

  // Unit conversions. CSS px are 1/96 inch; Word uses twips (1/1440 inch),
  // half-points for font sizes and EMUs (1/914400 inch) for drawings.
  const TWIPS_PER_INCH = 1440;
  const PX_TO_TWIPS = 15;
  const PX_TO_EMU = 9525;

  // Word 2013's built-in paragraph styles. Sizes and spacing in points,
  // indents in inches, line spacing as a multiple.
  const DEFAULT_STYLES = {
    normal: { name: 'Normal', font: 'Calibri', size: 11, color: '#000000', bold: false, italic: false, before: 0, after: 8, line: 1.08, align: '' },
    title: { name: 'Title', font: 'Calibri Light', size: 28, color: '#000000', bold: false, italic: false, before: 0, after: 0, line: 1, align: '', letter: -0.5 },
    h1: { name: 'Heading 1', font: 'Calibri Light', size: 16, color: '#2E74B5', bold: false, italic: false, before: 12, after: 0, line: 1.08, align: '' },
    h2: { name: 'Heading 2', font: 'Calibri Light', size: 13, color: '#2E74B5', bold: false, italic: false, before: 2, after: 0, line: 1.08, align: '' },
    h3: { name: 'Heading 3', font: 'Calibri Light', size: 12, color: '#1F4D78', bold: false, italic: false, before: 2, after: 0, line: 1.08, align: '' },
    quote: { name: 'Quote', font: 'Calibri', size: 11, color: '#404040', bold: false, italic: true, before: 10, after: 8, line: 1.08, align: 'center', indent: 0.6 },
  };
  const BUILTIN_BY_NAME = { normal: 'normal', title: 'title', 'heading 1': 'h1', 'heading 2': 'h2', 'heading 3': 'h3', quote: 'quote' };
  const FIXED_STYLE_NAMES = /^(list paragraph|caption|toc heading|toc \d|header|footer|heading [4-9]|intense quote|no spacing|default paragraph font|normal table|table grid|hyperlink|comment .*|annotation .*|balloon text)$/;

  const BLOCK_STYLE = { H1: 'Heading1', H2: 'Heading2', H3: 'Heading3', BLOCKQUOTE: 'Quote' };
  const STYLE_BLOCK = {
    'title': { tag: 'p', cls: 'title' },
    'heading 1': { tag: 'h1' }, 'heading 2': { tag: 'h2' }, 'heading 3': { tag: 'h3' },
    'heading 4': { tag: 'h3' }, 'heading 5': { tag: 'h3' }, 'heading 6': { tag: 'h3' },
    'quote': { tag: 'blockquote' }, 'intense quote': { tag: 'blockquote' },
    'caption': { tag: 'p', cls: 'caption' },
  };
  // <w:highlight> only accepts a fixed palette.
  const HIGHLIGHT = {
    yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff',
    red: '#ff0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
    darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
    lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
  };
  const HIGHLIGHT_NAME = Object.fromEntries(Object.entries(HIGHLIGHT).map(([k, v]) => [v, k]));
  const SAFE_HREF = /^(https?:|mailto:|#)/i;
  const BLOCK_TAGS = /^(P|H[1-6]|DIV|BLOCKQUOTE|LI|UL|OL|TABLE|PRE|HR|SECTION|ARTICLE|HEADER|FOOTER|FIGURE)$/;

  // List formats: bullet glyphs per level, and numFmt cycles for ordered lists.
  const BULLETS = {
    disc: ['•', 'o', '▪'], circle: ['o', '▪', '•'], square: ['▪', '•', 'o'],
    check: ['✓'], arrow: ['➢'], dash: ['–'],
  };
  const ORDERED = {
    decimal: ['decimal', 'lowerLetter', 'lowerRoman'], 'lower-alpha': ['lowerLetter', 'lowerRoman', 'decimal'],
    'upper-alpha': ['upperLetter', 'lowerLetter', 'lowerRoman'], 'lower-roman': ['lowerRoman', 'lowerLetter', 'decimal'],
    'upper-roman': ['upperRoman', 'upperLetter', 'decimal'], legal: ['decimal'],
  };

  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const xmlDecl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  function toHex(color) {
    if (!color) return null;
    const m = color.match(/rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/);
    if (m) {
      if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
      return [m[1], m[2], m[3]].map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase();
    }
    const h = color.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (!h) return null;
    return (h[1].length === 3 ? h[1].replace(/./g, '$&$&') : h[1]).toUpperCase();
  }

  // Inline CSS lengths in the editor are px, pt or in.
  function toPx(v) {
    if (!v) return 0;
    const n = parseFloat(v);
    if (!isFinite(n)) return 0;
    if (/pt$/.test(v)) return n * 96 / 72;
    if (/in$/.test(v)) return n * 96;
    return n;
  }

  // ---------------------------------------------------------------- styles.xml

  function rPrOf(d) {
    let x = '';
    if (d.font) x += '<w:rFonts w:ascii="' + esc(d.font) + '" w:hAnsi="' + esc(d.font) + '" w:cs="' + esc(d.font) + '"/>';
    if (d.bold) x += '<w:b/><w:bCs/>';
    if (d.italic) x += '<w:i/><w:iCs/>';
    if (d.color && toHex(d.color)) x += '<w:color w:val="' + toHex(d.color) + '"/>';
    if (d.letter) x += '<w:spacing w:val="' + Math.round(d.letter * 20) + '"/><w:kern w:val="28"/>';
    if (d.size) x += '<w:sz w:val="' + Math.round(d.size * 2) + '"/><w:szCs w:val="' + Math.round(d.size * 2) + '"/>';
    return x ? '<w:rPr>' + x + '</w:rPr>' : '';
  }

  // pPr children in schema order: pre (keepNext...), spacing, ind, mid
  // (contextualSpacing), jc, post (outlineLvl).
  function pPrOf(d, o = {}) {
    let x = o.pre || '';
    x += '<w:spacing w:before="' + Math.round((d.before || 0) * 20) + '" w:after="' + Math.round((d.after || 0) * 20) +
      '" w:line="' + Math.round((d.line || 1) * 240) + '" w:lineRule="auto"/>';
    if (d.indent) x += '<w:ind w:left="' + Math.round(d.indent * TWIPS_PER_INCH) + '" w:right="' + Math.round(d.indent * TWIPS_PER_INCH) + '"/>';
    x += o.mid || '';
    const jc = { center: 'center', right: 'right', justify: 'both' }[d.align];
    if (jc) x += '<w:jc w:val="' + jc + '"/>';
    x += o.post || '';
    return '<w:pPr>' + x + '</w:pPr>';
  }

  function paraStyle(id, name, d, o = {}) {
    return '<w:style w:type="paragraph"' + (id === 'Normal' ? ' w:default="1"' : '') + ' w:styleId="' + id + '">' +
      '<w:name w:val="' + esc(name) + '"/>' + (id === 'Normal' ? '' : '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>') +
      '<w:qFormat/>' + pPrOf(d, o) + rPrOf(d) + '</w:style>';
  }

  function customStyleId(def) {
    return 'Custom' + String(def.id).replace(/[^A-Za-z0-9]/g, '');
  }

  function stylesXml(styles) {
    const s = Object.assign({}, DEFAULT_STYLES, styles || {});
    const n = s.normal;
    let x = xmlDecl + '<w:styles xmlns:w="' + NS.w + '">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="' + esc(n.font) + '" w:eastAsia="' + esc(n.font) + '" w:hAnsi="' + esc(n.font) + '" w:cs="' + esc(n.font) + '"/>' +
      '<w:sz w:val="' + Math.round(n.size * 2) + '"/><w:szCs w:val="' + Math.round(n.size * 2) + '"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="' + Math.round(n.after * 20) + '" w:line="' + Math.round(n.line * 240) + '" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';
    x += paraStyle('Normal', 'Normal', n);
    x += paraStyle('Title', 'Title', s.title, { mid: '<w:contextualSpacing/>' });
    [1, 2, 3].forEach((i) => {
      x += paraStyle('Heading' + i, 'heading ' + i, s['h' + i], { pre: '<w:keepNext/><w:keepLines/>', post: '<w:outlineLvl w:val="' + (i - 1) + '"/>' });
    });
    x += paraStyle('Quote', 'Quote', s.quote);
    for (const c of (styles && styles.custom) || []) x += paraStyle(customStyleId(c), c.name, c);
    x += '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:qFormat/>' +
      '<w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>';
    [1, 2, 3].forEach((i) => {
      x += '<w:style w:type="paragraph" w:styleId="TOC' + i + '"><w:name w:val="toc ' + i + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
        '<w:pPr><w:spacing w:after="100"/><w:ind w:left="' + (i - 1) * 220 + '"/></w:pPr></w:style>';
    });
    x += '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>' +
      '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/>' +
      '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/>' +
      '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/>' +
      '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
      '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/>' +
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>' +
      ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((b) => '<w:' + b + ' w:val="single" w:sz="4" w:space="0" w:color="auto"/>').join('') +
      '</w:tblBorders></w:tblPr></w:style></w:styles>';
    return x;
  }

  // ---------------------------------------------------------------- writer

  class Writer {
    constructor(root, opts) {
      this.root = root;
      this.opts = opts;
      this.layout = opts.layout;
      this.rels = [];          // document.xml.rels entries
      this.media = [];         // { path, base64 }
      this.abstracts = [];     // list format keys, index = abstractNumId
      this.nums = [];          // numbering instances: abstract index
      this.sharedNums = {};    // bullet format -> numId
      this.drawingId = 0;
      this.revId = 1;
      this.sections = 0;
      this.inDel = false;
      this.customIds = {};
      for (const c of (opts.styles && opts.styles.custom) || []) this.customIds[c.id] = customStyleId(c);
      this.prepareComments();
      const hasText = (hf) => hf && hf.text && hf.text.trim();
      if (hasText(opts.header)) this.headerRid = this.addRel('header', 'header1.xml');
      if (hasText(opts.footer)) this.footerRid = this.addRel('footer', 'footer1.xml');
    }

    addRel(type, target, external) {
      const id = 'rId' + (this.rels.length + 10);
      this.rels.push({ id, type, target, external });
      return id;
    }

    // Comment anchors may be split into several spans; the range starts at
    // the first and ends after the last one, in document order.
    prepareComments() {
      this.comments = [];
      this.commentSpans = new Map();
      const known = this.opts.comments || {};
      const seen = {};
      for (const span of this.root.querySelectorAll('span.comment[data-cid]')) {
        const cid = span.dataset.cid;
        if (!known[cid]) continue;
        if (!seen[cid]) {
          seen[cid] = { id: this.comments.length, first: span, last: span, data: known[cid] };
          this.comments.push(seen[cid]);
        }
        seen[cid].last = span;
        this.commentSpans.set(span, seen[cid]);
      }
      if (this.comments.length) this.commentsRid = this.addRel('comments', 'comments.xml');
    }

    abstractFor(kind, fmt) {
      const key = kind + ':' + fmt;
      let i = this.abstracts.indexOf(key);
      if (i < 0) { this.abstracts.push(key); i = this.abstracts.length - 1; }
      return i;
    }

    newNum(abstract) {
      this.nums.push(abstract);
      return this.nums.length;
    }

    contentWidthTwips() {
      const L = this.layout;
      return Math.round((L.w - L.l - L.r) * TWIPS_PER_INCH);
    }

    // Body: top-level blocks in order; multi-column sections are closed by a
    // paragraph carrying their sectPr, as Word does.
    body() {
      let xml = '';
      let segment = [];
      let pending = false;
      const flush = () => {
        if (segment.length) { xml += this.nodes(segment, null, null); pending = true; }
        segment = [];
      };
      for (const child of this.root.childNodes) {
        if (child.nodeType === 1 && child.classList.contains('section')) {
          flush();
          if (pending) xml += this.sectionBreak(1);
          xml += this.nodes(Array.from(child.childNodes), null, null);
          xml += this.sectionBreak(Math.min(3, Math.max(1, parseInt(child.dataset.cols, 10) || 1)));
          pending = false;
        } else {
          segment.push(child);
        }
      }
      flush();
      if (!/<\/w:p>$|<w:p\/>$/.test(xml)) xml += '<w:p/>';
      return xml;
    }

    sectionBreak(cols) {
      const type = this.sections ? 'continuous' : null;
      const x = '<w:p><w:pPr>' + this.sectPr(cols, type, this.sections === 0) + '</w:pPr></w:p>';
      this.sections++;
      return x;
    }

    // Emit a sequence of sibling nodes. Runs of inline nodes become one
    // paragraph formatted by `blockEl`; block elements are emitted recursively.
    nodes(list, blockEl, listInfo) {
      let out = '';
      let inline = [];
      const flush = () => {
        if (inline.length) out += this.paragraph(blockEl, inline, listInfo);
        inline = [];
      };
      for (const child of list) {
        if (child.nodeType === 1 && BLOCK_TAGS.test(child.tagName)) {
          flush();
          out += this.block(child, listInfo);
        } else if (child.nodeType === 1 || (child.nodeType === 3 && child.data.length)) {
          if (child.nodeType === 3 && !inline.length && !blockEl && !/\S/.test(child.data)) continue;
          inline.push(child);
        }
      }
      flush();
      return out;
    }

    block(el, list) {
      const tag = el.tagName;
      const cls = el.classList;
      if (cls.contains('page-break')) return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      if (cls.contains('column-break')) return '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
      if (cls.contains('toc')) return this.toc(el);
      if (cls.contains('section')) return this.nodes(Array.from(el.childNodes), null, null);
      if (tag === 'UL' || tag === 'OL') return this.list(el, list);
      if (tag === 'TABLE') return this.table(el);
      if (tag === 'HR') return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>';
      const xml = this.nodes(Array.from(el.childNodes), el, list);
      return xml || this.paragraph(el, [], list);
    }

    list(el, parent) {
      const ordered = el.tagName === 'OL';
      const level = parent ? Math.min(parent.level + 1, 8) : 0;
      let numId;
      if (parent && parent.ordered === ordered) {
        numId = parent.numId;   // nested list of the same kind continues its parent's numbering
      } else if (ordered) {
        numId = this.newNum(this.abstractFor('ol', ORDERED[el.dataset.format] ? el.dataset.format : 'decimal'));
      } else {
        const fmt = BULLETS[el.dataset.bullet] ? el.dataset.bullet : 'disc';
        numId = this.sharedNums[fmt] || (this.sharedNums[fmt] = this.newNum(this.abstractFor('ul', fmt)));
      }
      const info = { numId, level, ordered };
      let out = '';
      for (const li of el.children) {
        if (li.tagName === 'LI') out += this.nodes(Array.from(li.childNodes), li, info) || this.paragraph(li, [], info);
        else out += this.block(li, info);
      }
      return out;
    }

    paragraph(blockEl, nodes, list) {
      const pPr = this.paragraphProps(blockEl, list);
      const base = blockEl || this.root;
      let runs = '';
      for (const n of nodes) runs += this.inline(n, base);
      return '<w:p>' + (pPr ? '<w:pPr>' + pPr + '</w:pPr>' : '') + runs + '</w:p>';
    }

    paragraphProps(el, list) {
      let x = '';
      let style = null;
      const isBlock = el && el !== this.root && !/^(TD|TH)$/.test(el.tagName);
      if (el && el.tagName) {
        style = BLOCK_STYLE[el.tagName] || null;
        if (el.classList.contains('title')) style = 'Title';
        if (el.classList.contains('caption')) style = 'Caption';
        if (el.dataset && el.dataset.style && this.customIds[el.dataset.style]) style = this.customIds[el.dataset.style];
      }
      if (list && (!style || style === 'Title')) style = 'ListParagraph';
      if (style) x += '<w:pStyle w:val="' + style + '"/>';
      if (list) x += '<w:numPr><w:ilvl w:val="' + list.level + '"/><w:numId w:val="' + list.numId + '"/></w:numPr>';
      if (!isBlock) return x;

      const st = el.style;
      const sides = ['Top', 'Left', 'Bottom', 'Right'].filter((s) => st['border' + s + 'Style'] && st['border' + s + 'Style'] !== 'none');
      if (sides.length) {
        x += '<w:pBdr>' + sides.map((s) => '<w:' + s.toLowerCase() + ' w:val="single" w:sz="4" w:space="1" w:color="auto"/>').join('') + '</w:pBdr>';
      }
      const shade = toHex(st.backgroundColor);
      if (shade) x += '<w:shd w:val="clear" w:color="auto" w:fill="' + shade + '"/>';

      let spacing = '';
      if (st.marginTop) spacing += ' w:before="' + Math.round(toPx(st.marginTop) * PX_TO_TWIPS) + '"';
      if (st.marginBottom) spacing += ' w:after="' + Math.round(toPx(st.marginBottom) * PX_TO_TWIPS) + '"';
      const lh = parseFloat(st.lineHeight);
      if (lh > 0 && !/px|pt|em|%/.test(st.lineHeight)) spacing += ' w:line="' + Math.round(lh * 240) + '" w:lineRule="auto"';
      if (spacing) x += '<w:spacing' + spacing + '/>';

      let ind = '';
      const left = toPx(st.marginLeft), right = toPx(st.marginRight), first = toPx(st.textIndent);
      if (!list && left > 0) ind += ' w:left="' + Math.round(left * PX_TO_TWIPS) + '"';
      if (right > 0) ind += ' w:right="' + Math.round(right * PX_TO_TWIPS) + '"';
      if (first > 0) ind += ' w:firstLine="' + Math.round(first * PX_TO_TWIPS) + '"';
      if (first < 0) ind += ' w:hanging="' + Math.round(-first * PX_TO_TWIPS) + '"';
      if (ind) x += '<w:ind' + ind + '/>';

      const align = st.textAlign || el.getAttribute('align') || '';
      const jc = { center: 'center', right: 'right', justify: 'both' }[align.toLowerCase()];
      if (jc) x += '<w:jc w:val="' + jc + '"/>';
      return x;
    }

    inline(node, base, link) {
      if (node.nodeType === 3) return this.textRun(node, base, link);
      if (node.nodeType !== 1) return '';
      const tag = node.tagName;
      const children = (lnk) => Array.from(node.childNodes).map((c) => this.inline(c, base, lnk)).join('');
      if (tag === 'BR') return '<w:r><w:br/></w:r>';
      if (tag === 'IMG') return this.image(node);
      if (tag === 'SCRIPT' || tag === 'STYLE') return '';

      if ((tag === 'INS' || tag === 'DEL') && node.classList.contains('tc')) {
        const name = tag === 'INS' ? 'w:ins' : 'w:del';
        const was = this.inDel;
        this.inDel = tag === 'DEL';
        // Revisions may only hold runs, so links inside them are written as text.
        const inner = children(link || node);
        this.inDel = was;
        return '<' + name + ' w:id="' + (this.revId++) + '" w:author="' + esc(node.dataset.author || 'Author') + '"' +
          (node.dataset.date ? ' w:date="' + esc(node.dataset.date) + '"' : '') + '>' + inner + '</' + name + '>';
      }

      if (tag === 'SPAN' && this.commentSpans.has(node)) {
        const c = this.commentSpans.get(node);
        let x = c.first === node ? '<w:commentRangeStart w:id="' + c.id + '"/>' : '';
        x += children(link);
        if (c.last === node) {
          x += '<w:commentRangeEnd w:id="' + c.id + '"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>' +
            '<w:commentReference w:id="' + c.id + '"/></w:r>';
        }
        return x;
      }

      const href = tag === 'A' && node.getAttribute('href');
      if (href && SAFE_HREF.test(href) && !link) {
        const inner = children(node);
        // In-document links target a bookmark; others go through a relationship.
        const target = href.startsWith('#')
          ? 'w:anchor="' + esc(href.slice(1)) + '"'
          : 'r:id="' + this.addRel('hyperlink', href, true) + '"';
        return '<w:hyperlink ' + target + ' w:history="1">' + inner + '</w:hyperlink>';
      }
      return children(link);
    }

    textRun(node, base, link) {
      const text = node.data.replace(/ /g, ' ');
      if (!text) return '';
      const rPr = this.runProps(node.parentElement, base, link);
      const t = this.inDel ? 'w:delText' : 'w:t';
      let content = '';
      text.split(/(\t|\n)/).forEach((part) => {
        if (part === '\t') content += '<w:tab/>';
        else if (part === '\n') content += '<w:br/>';
        else if (part) content += '<' + t + ' xml:space="preserve">' + esc(part) + '</' + t + '>';
      });
      return '<w:r>' + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') + content + '</w:r>';
    }

    // Character properties: emit only what differs from the paragraph's own
    // style, the way Word stores direct formatting on top of a style.
    runProps(el, base, link) {
      const isLink = !!link && link.tagName === 'A';
      const cs = getComputedStyle(el);
      const bs = getComputedStyle(isLink ? link : base);
      let x = isLink ? '<w:rStyle w:val="Hyperlink"/>' : '';

      const family = (s) => s.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (family(cs) !== family(bs)) {
        const f = esc(family(cs));
        x += '<w:rFonts w:ascii="' + f + '" w:hAnsi="' + f + '" w:cs="' + f + '"/>';
      }
      if ((parseInt(cs.fontWeight, 10) >= 600) !== (parseInt(bs.fontWeight, 10) >= 600)) {
        x += parseInt(cs.fontWeight, 10) >= 600 ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/>';
      }
      if ((cs.fontStyle === 'italic') !== (bs.fontStyle === 'italic')) {
        x += cs.fontStyle === 'italic' ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/>';
      }

      // Decorations, highlight and vertical alignment are not inherited in
      // computed style, so walk the element chain up to the paragraph.
      // Tracked-change markup and comment anchors are not formatting.
      let underline = false, strike = false, vert = null, fill = null;
      for (let e = el; e && e !== base && e !== link; e = e.parentElement) {
        const t = e.tagName;
        if ((t === 'INS' || t === 'DEL') && e.classList.contains('tc')) continue;
        const deco = (e.style.textDecorationLine || e.style.textDecoration || '');
        if (t === 'U' || t === 'INS' || /underline/.test(deco)) underline = true;
        if (t === 'S' || t === 'STRIKE' || t === 'DEL' || /line-through/.test(deco)) strike = true;
        if (!vert && (t === 'SUP' || e.style.verticalAlign === 'super')) vert = 'superscript';
        if (!vert && (t === 'SUB' || e.style.verticalAlign === 'sub')) vert = 'subscript';
        if (!fill && e.style.backgroundColor) fill = toHex(e.style.backgroundColor);
      }
      // Children of <w:rPr> must follow the schema's sequence or Word
      // reports the file as corrupt.
      if (strike) x += '<w:strike/>';
      const color = toHex(cs.color);
      if (color && color !== toHex(bs.color)) x += '<w:color w:val="' + color + '"/>';
      const size = Math.round(parseFloat(cs.fontSize) * 0.75 * 2);
      const baseSize = Math.round(parseFloat(bs.fontSize) * 0.75 * 2);
      // sub/superscript shrink the glyphs themselves; Word does that via vertAlign.
      if (size !== baseSize && !vert) x += '<w:sz w:val="' + size + '"/><w:szCs w:val="' + size + '"/>';
      const named = fill && HIGHLIGHT_NAME['#' + fill.toLowerCase()];
      if (named) x += '<w:highlight w:val="' + named + '"/>';
      if (underline && !isLink) x += '<w:u w:val="single"/>';
      if (fill && !named) x += '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>';
      if (vert) x += '<w:vertAlign w:val="' + vert + '"/>';
      return x;
    }

    image(img) {
      const m = (img.getAttribute('src') || '').match(/^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i);
      if (!m) return '';
      const ext = m[1].toLowerCase().replace('jpg', 'jpeg');
      const n = this.media.length + 1;
      const path = 'media/image' + n + '.' + ext;
      this.media.push({ path: 'word/' + path, base64: m[2] });
      const rid = this.addRel('image', path, false);
      const w = img.width || img.naturalWidth || 200;
      const h = img.height || img.naturalHeight || 150;
      const cx = Math.round(w * PX_TO_EMU), cy = Math.round(h * PX_TO_EMU);
      const id = ++this.drawingId;
      const docPr = '<wp:docPr id="' + id + '" name="Picture ' + id + '"' + (img.alt ? ' descr="' + esc(img.alt) + '"' : '') + '/>';
      const graphic = '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="' + id + '" name="image' + n + '.' + ext + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
        '<pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>';
      const wrap = img.dataset.wrap;
      if (wrap === 'left' || wrap === 'right' || wrap === 'block') {
        // Floating picture anchored to its paragraph, as Word's Wrap Text does.
        const align = wrap === 'block' ? 'center' : wrap;
        return '<w:r><w:drawing><wp:anchor distT="0" distB="91440" distL="114300" distR="114300" simplePos="0" ' +
          'relativeHeight="' + (251659264 + id) + '" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
          '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>' + align + '</wp:align></wp:positionH>' +
          '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
          '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
          (wrap === 'block' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapSquare wrapText="bothSides"/>') +
          docPr + '<wp:cNvGraphicFramePr/>' + graphic + '</wp:anchor></w:drawing></w:r>';
      }
      return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' + docPr + graphic + '</wp:inline></w:drawing></w:r>';
    }

    table(table) {
      const rows = Array.from(table.rows);
      const cols = Math.max(1, ...rows.map((r) => Array.from(r.cells).reduce((s, c) => s + (c.colSpan || 1), 0)));
      const colW = Math.floor(this.contentWidthTwips() / cols);
      let x = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>' +
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>';
      for (let i = 0; i < cols; i++) x += '<w:gridCol w:w="' + colW + '"/>';
      x += '</w:tblGrid>';
      for (const row of rows) {
        x += '<w:tr>';
        for (const cell of row.cells) {
          const span = cell.colSpan || 1;
          x += '<w:tc><w:tcPr><w:tcW w:w="' + colW * span + '" w:type="dxa"/>' +
            (span > 1 ? '<w:gridSpan w:val="' + span + '"/>' : '') + '</w:tcPr>';
          let inner = this.nodes(Array.from(cell.childNodes), cell, null);
          // Word requires every cell to end with a paragraph.
          if (!/<\/w:p>$|<w:p\/>$/.test(inner)) inner += '<w:p/>';
          x += inner + '</w:tc>';
        }
        x += '</w:tr>';
      }
      return x + '</w:tbl>';
    }

    // A table of contents is written as TOC1-3 paragraphs with a
    // right-aligned dot-leader tab before each page number, inside the
    // content control Word uses to mark a table of contents.
    toc(el) {
      const run = (t) => '<w:r><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
      let x = '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>';
      const title = el.querySelector('.toc-title');
      if (title) x += '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>' + run(title.textContent) + '</w:p>';
      for (const p of el.querySelectorAll('[data-level]')) {
        const text = (p.querySelector('.toc-text') || p).textContent;
        const page = (p.querySelector('.toc-page') || {}).textContent || '';
        x += '<w:p><w:pPr><w:pStyle w:val="TOC' + p.dataset.level + '"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="' +
          this.contentWidthTwips() + '"/></w:tabs></w:pPr>' + run(text) + '<w:r><w:tab/></w:r>' + run(page) + '</w:p>';
      }
      if (!el.querySelector('[data-level]') && !title) x += '<w:p/>';
      return x + '</w:sdtContent></w:sdt>';
    }

    sectPr(cols, type, first) {
      const L = this.layout;
      const t = (v) => Math.round(v * TWIPS_PER_INCH);
      let x = '<w:sectPr>';
      if (this.headerRid) x += '<w:headerReference w:type="default" r:id="' + this.headerRid + '"/>';
      if (this.footerRid) x += '<w:footerReference w:type="default" r:id="' + this.footerRid + '"/>';
      if (type) x += '<w:type w:val="' + type + '"/>';
      x += '<w:pgSz w:w="' + t(L.w) + '" w:h="' + t(L.h) + '"' + (L.w > L.h ? ' w:orient="landscape"' : '') + '/>' +
        '<w:pgMar w:top="' + t(L.t) + '" w:right="' + t(L.r) + '" w:bottom="' + t(L.b) + '" w:left="' + t(L.l) +
        '" w:header="720" w:footer="720" w:gutter="0"/>' +
        '<w:cols' + (cols > 1 ? ' w:num="' + cols + '"' : '') + ' w:space="720"/>';
      if (first && this.opts.differentFirst) x += '<w:titlePg/>';
      return x + '<w:docGrid w:linePitch="360"/></w:sectPr>';
    }

    // Header/footer text may hold {PAGE} and {PAGES}, written as fields.
    headerFooterXml(kind, hf) {
      const tag = kind === 'header' ? 'w:hdr' : 'w:ftr';
      const jc = { left: 'left', center: 'center', right: 'right' }[hf.align] || 'left';
      const runs = hf.text.split(/(\{PAGE\}|\{PAGES\})/).map((part) => {
        if (part === '{PAGE}') return '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
        if (part === '{PAGES}') return '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
        return part ? '<w:r><w:t xml:space="preserve">' + esc(part) + '</w:t></w:r>' : '';
      }).join('');
      return xmlDecl + '<' + tag + ' xmlns:w="' + NS.w + '" xmlns:r="' + NS.r + '"><w:p><w:pPr><w:jc w:val="' + jc + '"/></w:pPr>' +
        runs + '</w:p></' + tag + '>';
    }

    commentsXml() {
      let x = xmlDecl + '<w:comments xmlns:w="' + NS.w + '">';
      for (const c of this.comments) {
        const author = c.data.author || 'Author';
        const initials = author.split(/\s+/).map((s) => s[0] || '').join('').slice(0, 3).toUpperCase();
        x += '<w:comment w:id="' + c.id + '" w:author="' + esc(author) + '" w:date="' + esc(c.data.date || new Date().toISOString()) +
          '" w:initials="' + esc(initials) + '">' +
          String(c.data.text || '').split('\n').map((line) => '<w:p><w:r><w:t xml:space="preserve">' + esc(line) + '</w:t></w:r></w:p>').join('') +
          '</w:comment>';
      }
      return x + '</w:comments>';
    }

    numberingXml() {
      let abs = '';
      this.abstracts.forEach((key, a) => {
        const [kind, fmt] = key.split(':');
        abs += '<w:abstractNum w:abstractNumId="' + a + '"><w:multiLevelType w:val="' + (fmt === 'legal' ? 'multilevel' : 'hybridMultilevel') + '"/>';
        for (let l = 0; l < 9; l++) {
          let numFmt, text, font = '';
          if (kind === 'ul') {
            const glyphs = BULLETS[fmt];
            numFmt = 'bullet';
            text = glyphs[l % glyphs.length];
            if (text === 'o') font = 'Courier New';
          } else if (fmt === 'legal') {
            numFmt = 'decimal';
            text = Array.from({ length: l + 1 }, (_, i) => '%' + (i + 1)).join('.') + '.';
          } else {
            const cycle = ORDERED[fmt];
            numFmt = cycle[l % cycle.length];
            text = '%' + (l + 1) + '.';
          }
          abs += '<w:lvl w:ilvl="' + l + '"><w:start w:val="1"/><w:numFmt w:val="' + numFmt + '"/>' +
            (fmt === 'legal' ? '<w:isLgl/>' : '') + '<w:lvlText w:val="' + text + '"/><w:lvlJc w:val="left"/>' +
            '<w:pPr><w:ind w:left="' + 720 * (l + 1) + '" w:hanging="' + (fmt === 'legal' ? 360 + 140 * l : 360) + '"/></w:pPr>' +
            (font ? '<w:rPr><w:rFonts w:ascii="' + font + '" w:hAnsi="' + font + '" w:cs="' + font + '"/></w:rPr>' : '') +
            '</w:lvl>';
        }
        abs += '</w:abstractNum>';
      });
      let nums = '';
      this.nums.forEach((a, i) => {
        const ordered = this.abstracts[a].startsWith('ol:');
        nums += '<w:num w:numId="' + (i + 1) + '"><w:abstractNumId w:val="' + a + '"/>' +
          (ordered ? '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' : '') + '</w:num>';
      });
      return xmlDecl + '<w:numbering xmlns:w="' + NS.w + '">' + abs + nums + '</w:numbering>';
    }
  }

  async function write(root, opts) {
    const w = new Writer(root, opts);
    const body = w.body();
    const zip = new JSZip();
    // Word never writes directory entries into the package; neither do we.
    const put = (path, data, o) => zip.file(path, data, Object.assign({ createFolders: false }, o));
    const exts = new Set(w.media.map((m) => m.path.split('.').pop()));
    const mime = { png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };

    put('[Content_Types].xml', xmlDecl +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      Array.from(exts).map((e) => '<Default Extension="' + e + '" ContentType="' + mime[e] + '"/>').join('') +
      '<Override PartName="/word/document.xml" ContentType="' + CT + 'document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="' + CT + 'styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="' + CT + 'numbering+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="' + CT + 'settings+xml"/>' +
      (w.headerRid ? '<Override PartName="/word/header1.xml" ContentType="' + CT + 'header+xml"/>' : '') +
      (w.footerRid ? '<Override PartName="/word/footer1.xml" ContentType="' + CT + 'footer+xml"/>' : '') +
      (w.commentsRid ? '<Override PartName="/word/comments.xml" ContentType="' + CT + 'comments+xml"/>' : '') +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '</Types>');

    put('_rels/.rels', xmlDecl + '<Relationships xmlns="' + NS.rel + '">' +
      '<Relationship Id="rId1" Type="' + REL_TYPE + 'officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="' + REL_TYPE + 'extended-properties" Target="docProps/app.xml"/>' +
      '</Relationships>');

    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    put('docProps/core.xml', xmlDecl +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(opts.title || '') + '</dc:title><dc:creator>' + esc(opts.author || 'Inkwell') + '</dc:creator>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified></cp:coreProperties>');
    put('docProps/app.xml', xmlDecl +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
      '<Application>Inkwell</Application><AppVersion>2.0000</AppVersion></Properties>');

    put('word/document.xml', xmlDecl +
      '<w:document xmlns:w="' + NS.w + '" xmlns:r="' + NS.r + '" xmlns:wp="' + NS.wp + '" xmlns:a="' + NS.a + '" xmlns:pic="' + NS.pic + '">' +
      '<w:body>' + body + w.sectPr(1, w.sections ? 'continuous' : null, w.sections === 0) + '</w:body></w:document>');
    put('word/styles.xml', stylesXml(opts.styles));
    put('word/numbering.xml', w.numberingXml());
    put('word/settings.xml', xmlDecl + '<w:settings xmlns:w="' + NS.w + '">' +
      (opts.track ? '<w:trackRevisions/>' : '') +
      '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
      '</w:settings>');
    if (w.headerRid) put('word/header1.xml', w.headerFooterXml('header', opts.header));
    if (w.footerRid) put('word/footer1.xml', w.headerFooterXml('footer', opts.footer));
    if (w.commentsRid) put('word/comments.xml', w.commentsXml());

    put('word/_rels/document.xml.rels', xmlDecl + '<Relationships xmlns="' + NS.rel + '">' +
      '<Relationship Id="rId1" Type="' + REL_TYPE + 'styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="' + REL_TYPE + 'numbering" Target="numbering.xml"/>' +
      '<Relationship Id="rId3" Type="' + REL_TYPE + 'settings" Target="settings.xml"/>' +
      w.rels.map((r) => '<Relationship Id="' + r.id + '" Type="' + REL_TYPE + r.type + '" Target="' + esc(r.target) + '"' +
        (r.external ? ' TargetMode="External"' : '') + '/>').join('') +
      '</Relationships>');

    w.media.forEach((m) => put(m.path, m.base64, { base64: true }));

    return zip.generateAsync({ type: 'blob', mimeType: CT + 'document', compression: 'DEFLATE' });
  }

  // ---------------------------------------------------------------- reader

  const kids = (el, name) => Array.from(el ? el.children : []).filter((c) => !name || c.localName === name);
  const kid = (el, name) => kids(el, name)[0] || null;
  const wattr = (el, name) => (el ? el.getAttributeNS(NS.w, name) || el.getAttribute('w:' + name) : null);
  const onOff = (el) => !!el && !/^(0|false|off|none)$/i.test(wattr(el, 'val') || '');
  const THEME_FONTS = { minorHAnsi: 'Calibri', minorAscii: 'Calibri', majorHAnsi: 'Calibri Light', majorAscii: 'Calibri Light' };

  // Style properties from a <w:style> or docDefaults: pPr/rPr into our units.
  function styleProps(pPr, rPr) {
    const p = {};
    const fonts = kid(rPr, 'rFonts');
    const font = wattr(fonts, 'ascii') || wattr(fonts, 'hAnsi') || THEME_FONTS[wattr(fonts, 'asciiTheme')];
    if (font) p.font = font;
    const sz = wattr(kid(rPr, 'sz'), 'val');
    if (sz) p.size = parseInt(sz, 10) / 2;
    const color = wattr(kid(rPr, 'color'), 'val');
    if (color && color !== 'auto') p.color = '#' + color.toUpperCase();
    if (kid(rPr, 'b')) p.bold = onOff(kid(rPr, 'b'));
    if (kid(rPr, 'i')) p.italic = onOff(kid(rPr, 'i'));
    const sp = kid(pPr, 'spacing');
    if (wattr(sp, 'before')) p.before = parseInt(wattr(sp, 'before'), 10) / 20;
    if (wattr(sp, 'after')) p.after = parseInt(wattr(sp, 'after'), 10) / 20;
    const rule = wattr(sp, 'lineRule');
    if (wattr(sp, 'line') && (!rule || rule === 'auto')) p.line = Math.round(parseInt(wattr(sp, 'line'), 10) / 240 * 100) / 100;
    const jc = wattr(kid(pPr, 'jc'), 'val');
    if (jc) p.align = { center: 'center', right: 'right', end: 'right', both: 'justify' }[jc] || '';
    const ind = wattr(kid(pPr, 'ind'), 'left');
    if (ind) p.indent = parseInt(ind, 10) / TWIPS_PER_INCH;
    return p;
  }

  class Reader {
    constructor(zip) {
      this.zip = zip;
      this.openComments = [];
      this.custom = {};
    }

    async xml(path) {
      const f = this.zip.file(path);
      if (!f) return null;
      return new DOMParser().parseFromString(await f.async('string'), 'application/xml');
    }

    async readRels(partPath) {
      const doc = await this.xml(partPath.replace(/[^/]*$/, '') + '_rels/' + partPath.split('/').pop() + '.rels');
      const rels = {};
      if (doc) {
        for (const r of doc.getElementsByTagNameNS(NS.rel, 'Relationship')) {
          rels[r.getAttribute('Id')] = { target: r.getAttribute('Target'), external: r.getAttribute('TargetMode') === 'External', type: r.getAttribute('Type') };
        }
      }
      return rels;
    }

    async load() {
      const ct = await this.xml('_rels/.rels');
      let main = 'word/document.xml';
      if (ct) {
        for (const r of ct.getElementsByTagNameNS(NS.rel, 'Relationship')) {
          if (/\/officeDocument$/.test(r.getAttribute('Type'))) main = r.getAttribute('Target').replace(/^\//, '');
        }
      }
      this.base = main.replace(/[^/]*$/, '');
      this.doc = await this.xml(main);
      if (!this.doc) throw new Error('This file has no word/document.xml — is it a .docx?');
      this.rels = await this.readRels(main);
      await this.loadStyles();
      await this.loadNumbering();
      await this.loadComments();
      const core = await this.xml('docProps/core.xml');
      const titleEl = core && core.getElementsByTagNameNS('http://purl.org/dc/elements/1.1/', 'title')[0];
      this.title = titleEl ? titleEl.textContent.trim() : '';
    }

    async loadStyles() {
      this.styleNames = {};
      this.styleDefs = {};
      this.rawStyles = {};
      this.resolveStyle = () => Object.assign({}, DEFAULT_STYLES.normal);
      const styles = await this.xml(this.base + 'styles.xml');
      if (!styles) return;
      const dd = styles.getElementsByTagNameNS(NS.w, 'docDefaults')[0];
      const defaults = dd ? styleProps(dd.getElementsByTagNameNS(NS.w, 'pPr')[0], dd.getElementsByTagNameNS(NS.w, 'rPr')[0]) : {};
      const raw = this.rawStyles;
      for (const s of styles.getElementsByTagNameNS(NS.w, 'style')) {
        const id = wattr(s, 'styleId');
        const name = wattr(kid(s, 'name'), 'val') || id;
        this.styleNames[id] = name.toLowerCase();
        if (wattr(s, 'type') !== 'paragraph') continue;
        raw[id] = { name, basedOn: wattr(kid(s, 'basedOn'), 'val'), props: styleProps(kid(s, 'pPr'), kid(s, 'rPr')), isDefault: wattr(s, 'default') === '1' };
      }
      // Resolve basedOn chains down to docDefaults, as Word does.
      const resolve = (id, depth = 0) => {
        const r = raw[id];
        if (!r || depth > 10) return Object.assign({}, defaults);
        return Object.assign(resolve(r.basedOn, depth + 1), r.props);
      };
      this.resolveStyle = (id) => Object.assign({ before: 0, after: 0, line: 1, bold: false, italic: false, color: '#000000', align: '' }, resolve(id));
      for (const id of Object.keys(raw)) {
        const key = BUILTIN_BY_NAME[raw[id].name.toLowerCase()] || (raw[id].isDefault ? 'normal' : null);
        if (key && !this.styleDefs[key]) {
          this.styleDefs[key] = Object.assign({}, DEFAULT_STYLES[key], this.resolveStyle(id), { name: DEFAULT_STYLES[key].name });
        }
      }
    }

    // A paragraph style that isn't built in becomes an Inkwell custom style.
    customStyle(styleId) {
      if (this.custom[styleId]) return this.custom[styleId].id;
      const r = this.rawStyles[styleId];
      if (!r) return null;
      const def = Object.assign({}, DEFAULT_STYLES.normal, this.resolveStyle(styleId), { id: 'x' + styleId.replace(/[^A-Za-z0-9]/g, ''), name: r.name });
      delete def.indent;
      this.custom[styleId] = def;
      return def.id;
    }

    async loadNumbering() {
      this.numFmt = {};
      const numbering = await this.xml(this.base + 'numbering.xml');
      if (!numbering) return;
      const abstract = {};
      for (const a of numbering.getElementsByTagNameNS(NS.w, 'abstractNum')) {
        const lv = {};
        for (const l of kids(a, 'lvl')) lv[wattr(l, 'ilvl')] = { fmt: wattr(kid(l, 'numFmt'), 'val'), text: wattr(kid(l, 'lvlText'), 'val') || '' };
        abstract[wattr(a, 'abstractNumId')] = lv;
      }
      for (const n of numbering.getElementsByTagNameNS(NS.w, 'num')) {
        this.numFmt[wattr(n, 'numId')] = abstract[wattr(kid(n, 'abstractNumId'), 'val')] || {};
      }
    }

    async loadComments() {
      this.comments = {};
      const target = Object.values(this.rels).find((r) => /\/comments$/.test(r.type));
      const doc = target && await this.xml(this.base + target.target);
      if (!doc) return;
      for (const c of doc.getElementsByTagNameNS(NS.w, 'comment')) {
        const text = kids(c, 'p').map((p) => Array.from(p.getElementsByTagNameNS(NS.w, 't')).map((t) => t.textContent).join('')).join('\n');
        this.comments['c' + wattr(c, 'id')] = { author: wattr(c, 'author') || 'Author', date: wattr(c, 'date') || '', text, done: false };
      }
    }

    async headerFooter(sect, kind) {
      const refs = kids(sect, kind + 'Reference');
      const ref = refs.find((r) => wattr(r, 'type') === 'default') || refs[0];
      const rel = ref && this.rels[ref.getAttributeNS(NS.r, 'id')];
      const doc = rel && await this.xml(this.base + rel.target);
      if (!doc) return null;
      const state = { text: '' };
      const lines = [];
      let align = '';
      for (const p of Array.from(doc.getElementsByTagNameNS(NS.w, 'p'))) {
        state.text = '';
        this.hfRuns(p, state);
        if (state.text.trim()) {
          lines.push(state.text.trim());
          if (!align) align = { center: 'center', right: 'right', end: 'right' }[wattr(kid(kid(p, 'pPr'), 'jc'), 'val')] || 'left';
        }
      }
      return lines.length ? { text: lines.join('  '), align: align || 'left' } : null;
    }

    // Plain text of a header/footer paragraph with PAGE/NUMPAGES as tokens.
    hfRuns(parent, state) {
      for (const c of kids(parent)) {
        const n = c.localName;
        if (n === 'fldSimple') {
          const instr = (wattr(c, 'instr') || '').trim().split(/\s+/)[0];
          if (instr === 'PAGE') state.text += '{PAGE}';
          else if (instr === 'NUMPAGES') state.text += '{PAGES}';
          else this.hfRuns(c, state);
        } else if (n === 'r') {
          for (const x of kids(c)) {
            if (x.localName === 'fldChar') {
              const type = wattr(x, 'fldCharType');
              if (type === 'begin') state.fld = { instr: '', result: false };
              else if (type === 'separate' && state.fld) {
                state.fld.result = true;
                const instr = state.fld.instr.trim().split(/\s+/)[0];
                if (instr === 'PAGE') state.text += '{PAGE}';
                if (instr === 'NUMPAGES') state.text += '{PAGES}';
                state.fld.skip = instr === 'PAGE' || instr === 'NUMPAGES';
              } else if (type === 'end') state.fld = null;
            } else if (x.localName === 'instrText' && state.fld) state.fld.instr += x.textContent;
            else if (x.localName === 't' && !(state.fld && (!state.fld.result || state.fld.skip))) state.text += x.textContent;
            else if (x.localName === 'tab') state.text += '  ';
          }
        } else if (n === 'hyperlink' || n === 'ins' || n === 'smartTag' || n === 'sdtContent') {
          this.hfRuns(c, state);
        } else if (n === 'sdt') {
          this.hfRuns(kid(c, 'sdtContent'), state);
        }
      }
    }

    async media(rid) {
      const rel = this.rels[rid];
      if (!rel || rel.external) return null;
      const path = rel.target.startsWith('/') ? rel.target.slice(1) : this.base + rel.target;
      const f = this.zip.file(path.replace(/\/[^/]+\/\.\.\//g, '/'));
      if (!f) return null;
      const ext = path.split('.').pop().toLowerCase();
      const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' }[ext];
      if (!mime) return null;
      return 'data:' + mime + ';base64,' + (await f.async('base64'));
    }

    async toHtml() {
      await this.load();
      const body = this.doc.getElementsByTagNameNS(NS.w, 'body')[0];
      const out = document.createElement('div');
      this.sectionStart = 0;
      await this.blocks(body, out, true);
      const sect = kid(body, 'sectPr');
      this.endSection(out, sect);
      if (!out.firstChild) out.innerHTML = '<p><br></p>';
      return {
        html: out.innerHTML,
        title: this.title,
        layout: this.layout(sect),
        styles: Object.assign({}, this.styleDefs, { custom: Object.values(this.custom) }),
        header: sect ? await this.headerFooter(sect, 'header') : null,
        footer: sect ? await this.headerFooter(sect, 'footer') : null,
        differentFirst: !!kid(sect, 'titlePg'),
        comments: this.comments,
      };
    }

    layout(sect) {
      const sz = kid(sect, 'pgSz'), mar = kid(sect, 'pgMar');
      const inch = (el, a, d) => (el && wattr(el, a) ? parseInt(wattr(el, a), 10) / TWIPS_PER_INCH : d);
      return {
        w: inch(sz, 'w', 8.5), h: inch(sz, 'h', 11),
        t: inch(mar, 'top', 1), r: inch(mar, 'right', 1), b: inch(mar, 'bottom', 1), l: inch(mar, 'left', 1),
      };
    }

    // Close a section: blocks since the previous break get wrapped in a
    // multi-column container when the section has more than one column.
    endSection(out, sect) {
      const cols = parseInt(wattr(kid(sect, 'cols'), 'num') || '1', 10);
      const nodes = Array.from(out.childNodes).slice(this.sectionStart);
      if (cols > 1 && nodes.length) {
        const div = document.createElement('div');
        div.className = 'section';
        div.dataset.cols = String(Math.min(cols, 3));
        out.insertBefore(div, nodes[0]);
        nodes.forEach((n) => div.appendChild(n));
      }
      this.sectionStart = out.childNodes.length;
    }

    // Body-level elements. Consecutive list paragraphs are regrouped into
    // nested <ul>/<ol> trees using their ilvl.
    async blocks(parent, out, top) {
      let stack = [];
      for (const el of kids(parent)) {
        const name = el.localName;
        if (name === 'p') {
          const pPr = kid(el, 'pPr');
          const numPr = kid(pPr, 'numPr');
          const numId = numPr && wattr(kid(numPr, 'numId'), 'val');
          if (numId && numId !== '0') {
            const level = parseInt(wattr(kid(numPr, 'ilvl'), 'val') || '0', 10);
            const lv = (this.numFmt[numId] || {})[level] || { fmt: 'decimal', text: '' };
            const listTag = lv.fmt === 'bullet' ? 'ul' : 'ol';
            const li = document.createElement('li');
            await this.paragraphInto(el, li);
            while (stack.length > level + 1) stack.pop();
            while (stack.length < level + 1) {
              const list = document.createElement(listTag);
              this.listFormat(list, numId);
              if (stack.length) {
                const owner = stack[stack.length - 1].lastElementChild || stack[stack.length - 1].appendChild(document.createElement('li'));
                owner.appendChild(list);
              } else out.appendChild(list);
              stack.push(list);
            }
            if (stack[level].tagName.toLowerCase() !== listTag && level === 0) {
              const list = document.createElement(listTag);
              this.listFormat(list, numId);
              out.appendChild(list);
              stack = [list];
            }
            stack[level].appendChild(li);
          } else {
            stack = [];
            const sect = kid(pPr, 'sectPr');
            const hasContent = el.getElementsByTagNameNS(NS.w, 't').length || el.getElementsByTagNameNS(NS.w, 'drawing').length;
            if (!(sect && !hasContent)) await this.paragraph(el, out);
            if (sect && top) this.endSection(out, sect);
          }
        } else if (name === 'tbl') {
          stack = [];
          out.appendChild(await this.table(el));
        } else if (name === 'sdt') {
          stack = [];
          const gallery = el.getElementsByTagNameNS(NS.w, 'docPartGallery')[0];
          if (gallery && /table of contents/i.test(wattr(gallery, 'val') || '')) {
            // Inkwell rebuilds the contents from the document's headings.
            const toc = document.createElement('div');
            toc.className = 'toc';
            toc.contentEditable = 'false';
            out.appendChild(toc);
          } else {
            await this.blocks(kid(el, 'sdtContent'), out, top);
          }
        }
      }
    }

    listFormat(list, numId) {
      const lv = this.numFmt[numId] || {};
      const l0 = lv[0] || {}, l1 = lv[1] || {};
      if (list.tagName === 'UL') {
        const glyph = (l0.text || '').trim();
        const bullet = { '✓': 'check', '✔': 'check', '➢': 'arrow', '➤': 'arrow', '–': 'dash', '-': 'dash', '▪': 'square', '§': 'square', 'o': 'circle' }[glyph];
        if (bullet) list.dataset.bullet = bullet;
      } else if (/%1\.%2/.test(l1.text || '')) {
        list.dataset.format = 'legal';
      } else {
        const fmt = { lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman' }[l0.fmt];
        if (fmt) list.dataset.format = fmt;
      }
    }

    async paragraph(p, out) {
      const pPr = kid(p, 'pPr');
      const styleId = wattr(kid(pPr, 'pStyle'), 'val');
      const styleName = this.styleNames[styleId] || '';
      const map = STYLE_BLOCK[styleName] || { tag: 'p' };
      const el = document.createElement(map.tag);
      if (map.cls) el.className = map.cls;
      if (styleId && !STYLE_BLOCK[styleName] && !BUILTIN_BY_NAME[styleName] && !FIXED_STYLE_NAMES.test(styleName)) {
        const custom = this.customStyle(styleId);
        if (custom) el.dataset.style = custom;
      }
      const jc = wattr(kid(pPr, 'jc'), 'val');
      const align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[jc];
      if (align) el.style.textAlign = align;
      const spacing = kid(pPr, 'spacing');
      const line = wattr(spacing, 'line');
      const rule = wattr(spacing, 'lineRule');
      if (line && (!rule || rule === 'auto') && line !== '259') el.style.lineHeight = String(Math.round(parseInt(line, 10) / 240 * 100) / 100);
      const twipsPx = (v) => Math.round(parseInt(v, 10) / PX_TO_TWIPS) + 'px';
      if (wattr(spacing, 'before')) el.style.marginTop = twipsPx(wattr(spacing, 'before'));
      if (wattr(spacing, 'after') && wattr(spacing, 'after') !== '160') el.style.marginBottom = twipsPx(wattr(spacing, 'after'));
      const ind = kid(pPr, 'ind');
      const left = wattr(ind, 'left') || wattr(ind, 'start');
      if (left && parseInt(left, 10) > 0 && map.tag !== 'blockquote') el.style.marginLeft = twipsPx(left);
      const right = wattr(ind, 'right') || wattr(ind, 'end');
      if (right && parseInt(right, 10) > 0 && map.tag !== 'blockquote') el.style.marginRight = twipsPx(right);
      if (wattr(ind, 'firstLine')) el.style.textIndent = twipsPx(wattr(ind, 'firstLine'));
      if (wattr(ind, 'hanging')) el.style.textIndent = '-' + twipsPx(wattr(ind, 'hanging'));
      const bdr = kid(pPr, 'pBdr');
      for (const side of ['top', 'left', 'bottom', 'right']) {
        const b = kid(bdr, side);
        if (b && !/^(none|nil)$/.test(wattr(b, 'val') || '')) el.style['border' + side[0].toUpperCase() + side.slice(1)] = '1px solid #000';
      }
      const fill = wattr(kid(pPr, 'shd'), 'fill');
      if (fill && /^[0-9A-F]{6}$/i.test(fill)) el.style.backgroundColor = '#' + fill;

      const breaks = await this.paragraphInto(p, el);
      if (!el.firstChild) el.appendChild(document.createElement('br'));
      // A page break inside a paragraph splits it: content, then the break.
      if (breaks.only) {
        out.appendChild(pageBreak());
        return;
      }
      out.appendChild(el);
      for (let i = 0; i < breaks.count; i++) out.appendChild(pageBreak());
    }

    // Fill `el` with the paragraph's runs. Returns how many page breaks it held.
    async paragraphInto(p, el) {
      const state = { count: 0, text: false };
      await this.runs(p, el, state);
      return { count: state.count, only: state.count > 0 && !state.text };
    }

    async runs(parent, el, state) {
      for (const c of kids(parent)) {
        const n = c.localName;
        if (n === 'r') await this.run(c, el, state);
        else if (n === 'hyperlink') {
          const rel = this.rels[c.getAttributeNS(NS.r, 'id')];
          const anchor = wattr(c, 'anchor');
          const href = rel ? rel.target : anchor ? '#' + anchor : '';
          // Only web, mail and in-document links survive; others become plain text.
          const a = document.createElement(SAFE_HREF.test(href) ? 'a' : 'span');
          if (a.tagName === 'A') a.setAttribute('href', href);
          await this.runs(c, a, state);
          el.appendChild(a);
        } else if (n === 'ins' || n === 'del') {
          const mark = document.createElement(n);
          mark.className = 'tc';
          mark.dataset.author = wattr(c, 'author') || 'Author';
          if (wattr(c, 'date')) mark.dataset.date = wattr(c, 'date');
          await this.runs(c, mark, state);
          if (mark.firstChild) el.appendChild(mark);
        } else if (n === 'commentRangeStart') {
          if (this.comments['c' + wattr(c, 'id')]) this.openComments.push('c' + wattr(c, 'id'));
        } else if (n === 'commentRangeEnd') {
          this.openComments = this.openComments.filter((id) => id !== 'c' + wattr(c, 'id'));
        } else if (n === 'smartTag' || n === 'fldSimple' || n === 'customXml') {
          await this.runs(c, el, state);
        } else if (n === 'sdt') {
          await this.runs(kid(c, 'sdtContent'), el, state);
        }
      }
    }

    async run(r, el, state) {
      const rPr = kid(r, 'rPr');
      let target = el;
      const wrap = (tag) => { const t = document.createElement(tag); target.appendChild(t); target = t; return t; };
      if (this.openComments.length) {
        const span = wrap('span');
        span.className = 'comment';
        span.dataset.cid = this.openComments[this.openComments.length - 1];
      }
      if (rPr) {
        const style = [];
        const fonts = kid(rPr, 'rFonts');
        const font = wattr(fonts, 'ascii') || wattr(fonts, 'hAnsi');
        if (font) style.push('font-family: "' + font + '"');
        const sz = wattr(kid(rPr, 'sz'), 'val');
        if (sz) style.push('font-size: ' + parseInt(sz, 10) / 2 + 'pt');
        const color = wattr(kid(rPr, 'color'), 'val');
        if (color && color !== 'auto') style.push('color: #' + color);
        const hl = wattr(kid(rPr, 'highlight'), 'val');
        const shd = wattr(kid(rPr, 'shd'), 'fill');
        if (hl && HIGHLIGHT[hl]) style.push('background-color: ' + HIGHLIGHT[hl]);
        else if (shd && shd !== 'auto') style.push('background-color: #' + shd);
        if (style.length) wrap('span').setAttribute('style', style.join('; '));
        if (onOff(kid(rPr, 'b'))) wrap('b');
        if (onOff(kid(rPr, 'i'))) wrap('i');
        const u = kid(rPr, 'u');
        if (u && wattr(u, 'val') !== 'none') wrap('u');
        if (onOff(kid(rPr, 'strike')) || onOff(kid(rPr, 'dstrike'))) wrap('s');
        const va = wattr(kid(rPr, 'vertAlign'), 'val');
        if (va === 'superscript') wrap('sup');
        if (va === 'subscript') wrap('sub');
      }
      let text = '';
      const flush = () => { if (text) { target.appendChild(document.createTextNode(text)); state.text = true; text = ''; } };
      for (const c of kids(r)) {
        const n = c.localName;
        if (n === 't' || n === 'delText') text += c.textContent;
        else if (n === 'tab') text += '\t';
        else if (n === 'noBreakHyphen') text += '‑';
        else if (n === 'br' || n === 'cr') {
          if (wattr(c, 'type') === 'page') { state.count++; continue; }
          flush();
          target.appendChild(document.createElement('br'));
        } else if (n === 'drawing') {
          flush();
          const img = await this.drawing(c);
          if (img) { target.appendChild(img); state.text = true; }
        }
      }
      flush();
      // Drop formatting wrappers that ended up empty.
      while (target !== el && !target.firstChild) {
        const parent = target.parentNode;
        parent.removeChild(target);
        target = parent;
      }
    }

    async drawing(d) {
      const blip = d.getElementsByTagNameNS(NS.a, 'blip')[0];
      const src = blip && await this.media(blip.getAttributeNS(NS.r, 'embed'));
      if (!src) return null;
      const img = document.createElement('img');
      img.src = src;
      const extent = d.getElementsByTagNameNS(NS.wp, 'extent')[0];
      if (extent) {
        img.width = Math.round(parseInt(extent.getAttribute('cx'), 10) / PX_TO_EMU);
        img.height = Math.round(parseInt(extent.getAttribute('cy'), 10) / PX_TO_EMU);
      }
      const docPr = d.getElementsByTagNameNS(NS.wp, 'docPr')[0];
      img.alt = (docPr && docPr.getAttribute('descr')) || '';
      const anchor = d.getElementsByTagNameNS(NS.wp, 'anchor')[0];
      if (anchor) {
        const has = (n) => anchor.getElementsByTagNameNS(NS.wp, n).length > 0;
        const posH = anchor.getElementsByTagNameNS(NS.wp, 'positionH')[0];
        const align = posH && posH.getElementsByTagNameNS(NS.wp, 'align')[0];
        const side = align ? align.textContent.trim() : 'left';
        if (has('wrapTopAndBottom') || side === 'center') img.dataset.wrap = 'block';
        else if (has('wrapSquare') || has('wrapTight') || has('wrapThrough')) img.dataset.wrap = side === 'right' ? 'right' : 'left';
      }
      return img;
    }

    async table(tbl) {
      const table = document.createElement('table');
      const tbody = table.appendChild(document.createElement('tbody'));
      const open = {};   // grid column -> cell that a vertical merge continues
      for (const tr of kids(tbl, 'tr')) {
        const row = tbody.appendChild(document.createElement('tr'));
        let col = 0;
        for (const tc of kids(tr, 'tc')) {
          const tcPr = kid(tc, 'tcPr');
          const span = parseInt(wattr(kid(tcPr, 'gridSpan'), 'val') || '1', 10);
          const vMerge = kid(tcPr, 'vMerge');
          if (vMerge && wattr(vMerge, 'val') !== 'restart' && open[col]) {
            open[col].rowSpan += 1;
            col += span;
            continue;
          }
          const td = row.appendChild(document.createElement('td'));
          if (span > 1) td.colSpan = span;
          if (vMerge) open[col] = td; else delete open[col];
          await this.blocks(tc, td, false);
          // Single-paragraph cells hold their text directly, like the editor does.
          if (td.children.length === 1 && td.firstElementChild.tagName === 'P' && !td.firstElementChild.getAttribute('style')) {
            const p = td.firstElementChild;
            while (p.firstChild) td.insertBefore(p.firstChild, p);
            td.removeChild(p);
          }
          col += span;
        }
      }
      return table;
    }
  }

  function pageBreak() {
    const d = document.createElement('div');
    d.className = 'page-break';
    d.contentEditable = 'false';
    return d;
  }

  async function read(arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    return new Reader(zip).toHtml();
  }

  global.Docx = { read, write, DEFAULT_STYLES, BULLETS, ORDERED };
})(window);
