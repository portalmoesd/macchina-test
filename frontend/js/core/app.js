/**
 * App shell — minimal. No sidebar, no auth. Initialises i18n and wires the
 * optional topbar language toggle for the page chrome.
 */
const App = {
  async init() {
    await I18n.init();
    this._wireLangToggle();
  },

  _wireLangToggle() {
    const wrap = document.getElementById('siteLangToggle');
    if (!wrap) return;
    const refresh = () => {
      const cur = I18n.getLocale();
      wrap.querySelectorAll('[data-lang]').forEach(btn => {
        const active = btn.dataset.lang === cur;
        btn.classList.toggle('is-active', active);
        btn.setAttribute('aria-pressed', active ? 'true' : 'false');
      });
    };
    wrap.querySelectorAll('[data-lang]').forEach(btn => {
      btn.addEventListener('click', async () => {
        await I18n.setLocale(btn.dataset.lang);
        refresh();
      });
    });
    refresh();
  },
};
