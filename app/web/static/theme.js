// Тема интерфейса: по умолчанию — как в системе, выбор пользователя сохраняется в браузере.
(function () {
  'use strict';
  var KEY = 'foxford-theme';
  var root = document.documentElement;

  function savedTheme() {
    try {
      var v = localStorage.getItem(KEY);
      return v === 'dark' || v === 'light' ? v : null;
    } catch (e) {
      return null;
    }
  }

  function systemTheme() {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  function currentTheme() { return savedTheme() || systemTheme(); }

  function updateButtons() {
    var dark = root.getAttribute('data-theme') === 'dark';
    document.querySelectorAll('[data-theme-toggle]').forEach(function (b) {
      b.textContent = dark ? 'Светлая' : 'Тёмная';
      b.setAttribute('aria-pressed', String(dark));
      b.setAttribute('aria-label', dark ? 'Включить светлую тему' : 'Включить тёмную тему');
      b.title = dark ? 'Включить светлую тему' : 'Включить тёмную тему';
    });
  }

  function apply() {
    var theme = currentTheme();
    root.setAttribute('data-theme', theme);
    root.style.colorScheme = theme;
    updateButtons();
  }

  function save(theme) {
    try { localStorage.setItem(KEY, theme); } catch (e) {}
    apply();
  }

  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-theme-toggle]');
    if (b) save(root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
  });

  var media = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;
  if (media) {
    var onSystemChange = function () { if (!savedTheme()) apply(); };
    if (media.addEventListener) media.addEventListener('change', onSystemChange);
    else if (media.addListener) media.addListener(onSystemChange);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', updateButtons);
  apply();
})();
