/* VIEW tab: Read Mode, Print Layout and dark mode. */
(function () {
  'use strict';
  const { $, docEl } = Ink;

  function readMode(on) {
    document.body.classList.toggle('read-mode', on);
    Ink.state.readOnly = on;
    docEl.contentEditable = on ? 'false' : 'true';
    $('#readBar').hidden = !on;
    $('#readTitle').textContent = $('#docTitle').value;
    $('#printLayoutBtn').classList.toggle('on', !on);
    Ink.selectImage(null);
    Ink.setPagination(!on);
    if (on) { Ink.closeMenus(); getSelection().removeAllRanges(); }
    else docEl.focus({ preventScroll: true });
  }

  function setDark(on) {
    document.body.classList.toggle('dark', on);
    $('#darkBtn').classList.toggle('on', on);
    try { localStorage.setItem('inkwell:dark', on ? '1' : '0'); } catch (e) { /* private mode */ }
  }

  Ink.command('readMode', () => readMode(true));
  Ink.command('exitReadMode', () => readMode(false));
  Ink.command('printLayout', () => readMode(false));
  Ink.command('darkMode', () => setDark(!document.body.classList.contains('dark')));
  Ink.on('escape', () => { if (Ink.state.readOnly) readMode(false); });
  Ink.on('start', () => {
    let dark = false;
    try { dark = localStorage.getItem('inkwell:dark') === '1'; } catch (e) { /* default light */ }
    setDark(dark);
  });
})();
