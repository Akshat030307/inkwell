/*
 * One-click PDF export. The laid-out pages are rendered to a canvas with
 * html2canvas and placed one per PDF page with jsPDF; both libraries load
 * only when first needed. Text in the PDF is an image, not selectable —
 * Print → Save as PDF gives selectable text.
 */
(function () {
  'use strict';
  const { $ } = Ink;

  function loadScript(src) {
    if (document.querySelector('script[src="' + src + '"]')) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error('could not load ' + src));
      document.head.appendChild(s);
    });
  }

  async function exportPdf() {
    Ink.closeBackstage();
    Ink.toast('Creating PDF…');
    const L = Ink.state.layout;
    const { pageH, P } = Ink.pageGeometry();
    const pages = Ink.pageInfo().total;
    const zoom = Ink.pageWrap.style.zoom;
    try {
      await loadScript('vendor/html2canvas.min.js');
      await loadScript('vendor/jspdf.umd.min.js');
      Ink.selectImage(null);
      getSelection().removeAllRanges();
      document.body.classList.add('exporting');
      Ink.pageWrap.style.zoom = 1;
      Ink.paginate();
      // Keep the canvas under browsers' ~32k pixel height limit.
      const scale = Math.min(2, 30000 / (pages * P));
      const canvas = await html2canvas($('#pageStack'), { scale, backgroundColor: '#ffffff', logging: false });
      const { jsPDF } = window.jspdf;
      const orientation = L.w > L.h ? 'landscape' : 'portrait';
      const pdf = new jsPDF({ unit: 'in', format: [L.w, L.h], orientation });
      const pw = Math.round(L.w * 96 * scale), ph = Math.round(pageH * scale);
      for (let k = 0; k < pages; k++) {
        const c = document.createElement('canvas');
        c.width = pw;
        c.height = ph;
        const g = c.getContext('2d');
        g.fillStyle = '#fff';
        g.fillRect(0, 0, pw, ph);
        g.drawImage(canvas, 0, Math.round(k * P * scale), pw, ph, 0, 0, pw, ph);
        if (k) pdf.addPage([L.w, L.h], orientation);
        pdf.addImage(c.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, L.w, L.h);
      }
      pdf.save(Ink.fileName('.pdf'));
      Ink.toast('Saved ' + Ink.fileName('.pdf'));
    } catch (e) {
      console.error(e);
      Ink.toast('PDF export failed: ' + e.message);
    } finally {
      document.body.classList.remove('exporting');
      Ink.pageWrap.style.zoom = zoom;
      Ink.paginate();
    }
  }

  Ink.command('exportPdf', exportPdf);
})();
