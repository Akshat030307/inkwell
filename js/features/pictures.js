/*
 * PICTURE contextual tab: text wrapping, alignment, exact size, rotate/flip,
 * crop and alt text. Rotation and cropping re-encode the picture on a canvas.
 */
(function () {
  'use strict';
  const { $ } = Ink;
  const PX_PER_IN = 96;
  const img = () => Ink.getActiveImage();

  function showSize(el) {
    if (!el) return;
    $('#imgW').value = Math.round(el.width / PX_PER_IN * 100) / 100;
    $('#imgH').value = Math.round(el.height / PX_PER_IN * 100) / 100;
  }

  Ink.on('image', (el) => {
    Ink.setContextTab('picture', !!el);
    showSize(el);
  });

  function setSize(w, h) {
    const el = img();
    if (!el) return;
    Ink.edit(() => {
      const max = Ink.contentWidthPx();
      if (w > max) { h = h * max / w; w = max; }
      el.width = Math.max(8, Math.round(w));
      el.height = Math.max(8, Math.round(h));
      Ink.selectImage(el);
    }, { focus: false });
    showSize(el);
    Ink.positionImgHandle();
  }

  // Width and height stay in proportion, as with Word's "Lock aspect ratio".
  $('#imgW').addEventListener('change', (e) => {
    const el = img();
    if (!el) return;
    const w = Math.max(0.1, parseFloat(e.target.value) || 0) * PX_PER_IN;
    setSize(w, w * el.height / el.width);
  });
  $('#imgH').addEventListener('change', (e) => {
    const el = img();
    if (!el) return;
    const h = Math.max(0.1, parseFloat(e.target.value) || 0) * PX_PER_IN;
    setSize(h * el.width / el.height, h);
  });

  async function source(el) {
    const src = new Image();
    src.src = el.src;
    await src.decode();
    return src;
  }

  // Redraw the picture through a canvas and swap in the result.
  async function transform(draw) {
    const el = img();
    if (!el) return;
    const src = await source(el);
    const canvas = document.createElement('canvas');
    const size = draw(canvas, canvas.getContext('2d'), src, el);
    const jpeg = /^data:image\/jpe?g/i.test(el.src);
    const url = canvas.toDataURL(jpeg ? 'image/jpeg' : 'image/png', 0.92);
    Ink.edit(() => {
      el.src = url;
      el.width = Math.round(size.w);
      el.height = Math.round(size.h);
      Ink.selectImage(el);
    }, { focus: false });
    showSize(el);
  }

  function rotate(arg) {
    transform((c, g, src, el) => {
      const w = src.naturalWidth, h = src.naturalHeight;
      if (arg === 'flip') {
        c.width = w; c.height = h;
        g.translate(w, 0);
        g.scale(-1, 1);
        g.drawImage(src, 0, 0);
        return { w: el.width, h: el.height };
      }
      c.width = h; c.height = w;
      g.translate(h / 2, w / 2);
      g.rotate((parseInt(arg, 10) || 90) * Math.PI / 180);
      g.drawImage(src, -w / 2, -h / 2);
      const max = Ink.contentWidthPx();
      const dw = el.height, dh = el.width;
      return dw > max ? { w: max, h: dh * max / dw } : { w: dw, h: dh };
    });
  }

  async function crop() {
    const el = img();
    if (!el) return;
    const res = await Ink.dialog('Crop Picture',
      '<p class="muted hint">Trim each edge by a percentage of the picture.</p><div class="form-grid">' +
      ['top', 'bottom', 'left', 'right'].map((s) => '<label>' + s[0].toUpperCase() + s.slice(1) + ' (%)<input type="number" name="' + s +
        '" min="0" max="45" step="1" value="0"></label>').join('') + '</div>', 'Crop');
    if (!res) return;
    const pct = (k) => Ink.clamp(parseFloat(res[k]) || 0, 0, 45) / 100;
    const t = pct('top'), b = pct('bottom'), l = pct('left'), r = pct('right');
    if (!(t || b || l || r)) return;
    transform((c, g, src, image) => {
      const w = src.naturalWidth, h = src.naturalHeight;
      c.width = Math.max(1, Math.round(w * (1 - l - r)));
      c.height = Math.max(1, Math.round(h * (1 - t - b)));
      g.drawImage(src, w * l, h * t, c.width, c.height, 0, 0, c.width, c.height);
      return { w: image.width * (1 - l - r), h: image.height * (1 - t - b) };
    });
  }

  async function altText() {
    const el = img();
    if (!el) return;
    const res = await Ink.dialog('Alt Text',
      '<label>Description<textarea name="alt" rows="3" placeholder="Describe the picture for people who can\'t see it">' + Ink.escHtml(el.alt || '') + '</textarea></label>', 'OK');
    if (res) Ink.edit(() => { el.alt = res.alt.trim().slice(0, 500); }, { focus: false });
  }

  function wrap(arg) {
    const el = img();
    if (!el) return;
    Ink.edit(() => {
      if (arg === 'inline') delete el.dataset.wrap; else el.dataset.wrap = arg;
      Ink.selectImage(el);
    }, { focus: false });
  }

  function align(arg) {
    const el = img();
    if (!el) return;
    if (el.dataset.wrap === 'left' || el.dataset.wrap === 'right') { wrap(arg === 'center' ? 'block' : arg); return; }
    Ink.edit(() => {
      const block = Ink.closestBlock(el);
      if (block) block.style.textAlign = arg === 'left' ? '' : arg;
      if (el.dataset.wrap === 'block') el.style.margin = arg === 'left' ? '6px auto 6px 0' : arg === 'right' ? '6px 0 6px auto' : '';
      Ink.selectImage(el);
    }, { focus: false });
  }

  async function resetSize() {
    const el = img();
    if (!el) return;
    const src = await source(el);
    setSize(src.naturalWidth, src.naturalHeight);
  }

  Ink.command('wrap', wrap);
  Ink.command('imgAlign', align);
  Ink.command('rotateImage', rotate);
  Ink.command('cropImage', crop);
  Ink.command('altText', altText);
  Ink.command('resetImage', resetSize);
})();
