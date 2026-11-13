// Версия «в браузере» (browser/README.md): форма → обработка в этой вкладке → результат.
// Форма — копия сайта (app/web/static/app.js); вместо сервера и очереди — поток worker.js с Python внутри.
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var meta = null;          // шаблоны, шаги, тексты памятки — из meta.json (его делает browser/build.py)
  var state = { template: null, grade: null, file: null, choices: {} };
  var worker = null;        // поток с Python: один на одну презентацию
  var job = null;           // что сейчас обрабатывается: { filename, lesson }
  var workerFailed = false;
  var downloadUrl = null;

  function el(tag, attrs, text) {
    var e = document.createElement(tag);
    for (var k in attrs || {}) e.setAttribute(k, attrs[k]);
    if (text != null) e.textContent = text;
    return e;
  }

  function show(view) {
    ['form', 'progress', 'error', 'result'].forEach(function (v) { $('view-' + v).hidden = v !== view; });
    window.scrollTo(0, 0);
  }

  // ---------------------------------------------------------------- форма (как на сайте)
  function currentTemplate() {
    return meta.templates.filter(function (t) { return t.id === state.template; })[0];
  }

  // шаблоны — вкладками; шаблоны с "group" (индивидуальные) — выпадающим списком в той же полосе
  function renderTemplates() {
    var box = $('templates'), groups = {};
    meta.templates.forEach(function (t) {
      if (!t.group) {
        var b = el('button', { type: 'button', class: 'tab', 'data-id': t.id }, t.name);
        b.addEventListener('click', function () { selectTemplate(t.id); });
        box.appendChild(b);
        return;
      }
      var g = groups[t.group];
      if (!g) {
        g = groups[t.group] = el('div', { class: 'dd' });
        var btn = el('button', { type: 'button', class: 'tab dd__btn', 'aria-haspopup': 'true', 'aria-expanded': 'false',
                                 'data-group': t.group });
        btn.appendChild(el('span', { class: 'dd__label' }, t.group));
        btn.appendChild(el('span', { class: 'dd__arrow', 'aria-hidden': 'true' }, '▾'));
        btn.addEventListener('click', function (e) { e.stopPropagation(); toggleMenu(g); });
        g.appendChild(btn);
        g.appendChild(el('div', { class: 'dd__menu', role: 'menu', hidden: '' }));
        box.appendChild(g);
      }
      var item = el('button', { type: 'button', class: 'dd__item', role: 'menuitemradio', 'data-id': t.id }, t.name);
      item.addEventListener('click', function () { closeMenus(); selectTemplate(t.id); });
      g.querySelector('.dd__menu').appendChild(item);
    });
    document.addEventListener('click', closeMenus);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeMenus(); });
    box.hidden = meta.templates.length < 2;  // выбирать не из чего — полосу не показываем
    selectTemplate(meta.templates[0].id);
  }

  function toggleMenu(g) {
    var open = g.querySelector('.dd__menu').hidden;
    closeMenus();
    g.querySelector('.dd__menu').hidden = !open;
    g.querySelector('.dd__btn').setAttribute('aria-expanded', String(open));
  }

  function closeMenus() {
    document.querySelectorAll('#templates .dd').forEach(function (g) {
      g.querySelector('.dd__menu').hidden = true;
      g.querySelector('.dd__btn').setAttribute('aria-expanded', 'false');
    });
  }

  // пропорции слайдов шаблона (4 / 3 или 16 / 9) — для плиток выбора
  function setRatio(r) { document.body.style.setProperty('--ratio', r || '4 / 3'); }

  function selectTemplate(id) {
    state.template = id;
    var t = currentTemplate();
    setRatio(t.ratio);
    document.querySelectorAll('#templates .tab[data-id]').forEach(function (b) {
      b.setAttribute('aria-pressed', String(b.getAttribute('data-id') === id));
    });
    // выпадающий список: выбран шаблон из него — кнопка в рамке и с его названием
    document.querySelectorAll('#templates .dd').forEach(function (g) {
      var btn = g.querySelector('.dd__btn'), group = btn.getAttribute('data-group'), on = t.group === group;
      btn.setAttribute('aria-pressed', String(on));
      btn.querySelector('.dd__label').textContent = on ? group + ' · ' + t.name : group;
      g.querySelectorAll('.dd__item').forEach(function (i) {
        i.setAttribute('aria-checked', String(i.getAttribute('data-id') === id));
      });
    });
    // индивидуальным шаблонам класс, номер урока и предмет не нужны: теги переносятся со старого титула
    ['row-grade', 'row-lesson', 'row-subject'].forEach(function (r) { $(r).hidden = !t.lesson; });
    var box = $('grades'); box.innerHTML = '';
    t.grades.forEach(function (g) {
      var b = el('button', { type: 'button', 'data-grade': g, 'aria-label': g + ' класс' }, g);
      b.addEventListener('click', function () { selectGrade(g); });
      box.appendChild(b);
    });
    if (t.grades.length) selectGrade(t.grades.indexOf(state.grade) >= 0 ? state.grade : t.grades[0]);
    renderChoices();
  }

  // варианты слайдов шаблона (титул, прощание…): строка формы с миниатюрами на каждый.
  // Необязательный выбор (optional) начинается с плитки «Как в старой»: слайд останется из старой презентации;
  // выбор с "auto_label" — с плитки «По экзамену»: вариант подберётся сам (контакты ОГЭ / ЕГЭ)
  function renderChoices() {
    var t = currentTemplate(), rows = $('choice-rows');
    rows.innerHTML = ''; state.choices = {};
    var first = t.lesson ? 5 : 2;  // номера строк формы идут после «Файла» (и класса, урока, предмета)
    t.choices.forEach(function (c, ci) {
      var row = el('div', { class: 'row' });
      var lab = el('div', { class: 'row__label mono muted' }, ('0' + (first + ci)).slice(-2));
      lab.appendChild(el('b', null, c.label));
      var body = el('div'), grid = el('div', { class: 'titles', role: 'radiogroup', 'aria-label': c.label });
      var zero = c.optional ? 'Как в старой' : c.auto_label;
      if (zero) {
        var old = el('button', { type: 'button', class: 'title-opt title-opt--old', role: 'radio', 'data-n': 0,
                                 'aria-label': c.label + ': ' + zero.toLowerCase() });
        old.appendChild(el('i', { class: 'mono' }, zero));
        old.addEventListener('click', function () { selectChoice(grid, c.key, 0); });
        grid.appendChild(old);
      }
      c.slides.forEach(function (n, i) {
        var b = el('button', { type: 'button', class: 'title-opt', role: 'radio', 'data-n': n,
                               'aria-label': c.label + ' ' + (i + 1) });
        b.appendChild(el('img', { src: 'thumbs/' + t.id + '/' + n + '.jpg?v=' + t.thumbs_v, alt: '', loading: 'lazy' }));
        b.appendChild(el('span', { class: 'mono' }, ('0' + (i + 1)).slice(-2)));
        b.addEventListener('click', function () { selectChoice(grid, c.key, n); });
        grid.appendChild(b);
      });
      body.appendChild(grid);
      if (c.hint) body.appendChild(el('p', { class: 'hint mono muted' }, c.hint));
      row.appendChild(lab); row.appendChild(body); rows.appendChild(row);
      selectChoice(grid, c.key, zero ? 0 : c.default);
    });
  }

  function selectChoice(grid, key, n) {
    if (n) state.choices[key] = n; else delete state.choices[key];
    grid.querySelectorAll('.title-opt').forEach(function (b) {
      b.setAttribute('aria-checked', String(+b.getAttribute('data-n') === n));
    });
  }

  function selectGrade(g) {
    state.grade = g;
    document.querySelectorAll('#grades button').forEach(function (b) {
      b.setAttribute('aria-pressed', String(b.getAttribute('data-grade') === g));
    });
    $('lesson-prefix').textContent = g + '.';
  }

  function sizeMb(n) { return (n / 1024 / 1024).toFixed(1).replace('.', ',') + ' МБ'; }

  function setFile(f) {
    formError('');
    if (!f) return;
    if (!/\.pptx$/i.test(f.name)) { formError('Нужен файл .pptx. Старый формат .ppt сначала пересохраните в .pptx.'); return; }
    if (f.size > meta.max_upload_mb * 1024 * 1024) { formError('Файл больше ' + meta.max_upload_mb + ' МБ.'); return; }
    state.file = f;
    $('drop').classList.add('has-file');
    $('drop-main').textContent = f.name;
    $('drop-sub').textContent = sizeMb(f.size) + ' · нажмите, чтобы заменить';
    // «5.21 Религия….pptx», «Урок 9.23 Конституция….pptx» → класс и номер урока
    var m = /(?:^|\D)(\d{1,2})\.(\d{1,2})(?!\d)/.exec(f.name);
    if (m && currentTemplate().grades.indexOf(m[1]) >= 0) { selectGrade(m[1]); $('lesson').value = String(+m[2]); }
  }

  function dropHint() { return 'или нажмите, чтобы выбрать файл · до ' + meta.max_upload_mb + ' МБ'; }

  function formError(msg) { $('form-error').textContent = msg; $('form-error').hidden = !msg; }

  function bindForm() {
    var drop = $('drop'), input = $('file');
    input.addEventListener('change', function () { setFile(input.files[0]); });
    drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    ['dragenter', 'dragover'].forEach(function (t) {
      drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add('is-over'); });
    });
    ['dragleave', 'drop'].forEach(function (t) {
      drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove('is-over'); });
    });
    drop.addEventListener('drop', function (e) { setFile(e.dataTransfer.files[0]); });
    $('lesson').addEventListener('input', function () { this.value = this.value.replace(/\D/g, '').slice(0, 2); });
    $('form').addEventListener('submit', function (e) { e.preventDefault(); submit(); });
  }

  // ---------------------------------------------------------------- поток с Python
  // Запускается заранее, пока человек заполняет форму: Python и библиотеки (~15 МБ) успевают загрузиться
  function startWorker() {
    worker = new Worker('worker.js?v=' + meta.version, { type: 'module' });
    worker.onmessage = function (e) { onWorker(e.data); };
    worker.onerror = function (e) {
      console.error(e);
      workerFailed = true;  // поток не запустился (обрыв сети, старый браузер): при переносе попробуем ещё раз
      if (job) showError('Не удалось запустить конвертер в этом браузере. Обновите браузер (Chrome, Edge, Яндекс Браузер, ' +
                         'Safari, Firefox) или воспользуйтесь сайтом.', e.message || '');
    };
    worker.postMessage({ type: 'init', meta: { version: meta.version, packages: meta.packages, wheels: meta.wheels } });
  }

  function restartWorker() { if (worker) worker.terminate(); workerFailed = false; startWorker(); }

  function submit() {
    var n = $('lesson').value, t = currentTemplate(), lesson = '';
    if (!state.file) return formError('Выберите файл презентации.');
    if (t.lesson) {
      if (!n || +n < 1) return formError('Укажите номер урока.');
      lesson = state.grade + '.' + (+n);
    }
    formError('');
    job = { filename: state.file.name.normalize('NFC'), lesson: lesson };
    showProgress({ status: 'load' });
    if (workerFailed) restartWorker();
    worker.postMessage({ type: 'convert', file: state.file, template: state.template, lesson: lesson, choices: state.choices,
                         subject: t.lesson ? $('subject').value.trim() : '', base: t.bases[t.lesson ? state.grade : ''] });
  }

  function onWorker(m) {
    if (m.status === 'ready') return;
    if (!job) return;
    if (m.status === 'done') { showResult(m.result, m.file); restartWorker(); return; }
    if (m.status === 'error') {
      if (m.trace) console.error(m.trace);
      var memory = /memory|alloc/i.test(m.trace || '');
      showError(m.error || (memory
        ? 'Браузеру не хватило памяти на эту презентацию. Закройте лишние вкладки и попробуйте ещё раз или воспользуйтесь сайтом.'
        : 'Что-то пошло не так при обработке.'), m.trace);
      restartWorker();
      return;
    }
    showProgress(m);
  }

  // ---------------------------------------------------------------- ход работы
  function showProgress(m) {
    show('progress');
    $('p-file').textContent = (job.lesson ? 'Урок ' + job.lesson + ' · ' : '') + job.filename;
    var steps = [['load', 'Загрузка конвертера и шаблона']].concat(meta.steps), now = 0, pct = 0.02;
    if (m.status === 'base') {
      $('p-status').textContent = 'Загрузка шаблона · ' + Math.round(m.pct * 100) + ' %';
      pct = 0.05 + 0.2 * m.pct;
    } else if (m.status === 'step') {
      now = steps.map(function (s) { return s[0]; }).indexOf(m.step);
      $('p-status').textContent = now >= 0 ? steps[now][1] : 'Обработка';
      pct = 0.25 + 0.75 * Math.max(now - 1, 0) / meta.steps.length;
    } else {
      $('p-status').textContent = 'Загрузка конвертера';
    }
    $('p-sub').textContent = now === 0
      ? 'В первый раз это дольше: загружается около 65 МБ. Потом браузер берёт их из своей памяти.'
      : 'Обычно это занимает 1–3 минуты. Файл обрабатывается на вашем компьютере.';
    $('p-bar').style.width = Math.round(pct * 100) + '%';
    var ol = $('p-steps'); ol.innerHTML = '';
    steps.forEach(function (s, i) {
      var li = el('li', { class: i < now ? 'is-done' : i === now ? 'is-now' : '' });
      li.appendChild(el('span', { class: 'mono' }, ('0' + (i + 1)).slice(-2)));
      li.appendChild(el('span', null, s[1]));
      li.appendChild(el('span', { class: 'mark', 'aria-hidden': 'true' }));
      ol.appendChild(li);
    });
  }

  function showError(msg, trace) {
    show('error');
    $('e-file').textContent = job ? (job.lesson ? 'Урок ' + job.lesson + ' · ' : '') + job.filename : '';
    $('e-message').textContent = msg;
    var c = $('e-contact'); c.innerHTML = '';
    if (meta.contact) {
      c.appendChild(document.createTextNode('Если ошибка повторяется, сообщите автору: '));
      var isMail = /@/.test(meta.contact) && !/^@/.test(meta.contact);
      c.appendChild(isMail ? el('a', { href: 'mailto:' + meta.contact }, meta.contact) : el('b', null, meta.contact));
    }
    // журнала, как на сайте, здесь нет — подробности для автора можно скопировать прямо со страницы
    $('e-details').hidden = !trace; $('e-trace').textContent = trace || '';
    job = null;
  }

  // ---------------------------------------------------------------- результат
  function capital(t) { return t.charAt(0).toUpperCase() + t.slice(1); }

  function showResult(r, file) {
    show('result');
    $('r-meta').textContent = [r.template_name, r.subject, r.lesson && 'урок ' + r.lesson, r.slides.length + ' сл.',
      String(r.size_mb).replace('.', ',') + ' МБ'].filter(Boolean).join(' · ');
    $('r-file').textContent = job.filename;
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = URL.createObjectURL(new Blob([file],
      { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
    $('r-download').href = downloadUrl;
    $('r-download').download = job.filename.replace(/\.pptx$/i, '') + '_новый.pptx';

    // превью нет — все замечания списком: сначала общие, потом по слайдам
    var notes = r.checks.concat(r.stickers.map(function (s) { return { slide: s.slide, text: 'стикер: ' + s.text }; }));
    notes.sort(function (a, b) { return (a.slide || 0) - (b.slide || 0); });
    var box = $('r-checks'); box.innerHTML = '';
    if (!notes.length) box.appendChild(el('p', { class: 'empty' }, 'Конвертер ничего не отметил.'));
    else {
      var ul = el('ul');
      notes.forEach(function (it) {
        var li = el('li');
        if (it.slide) li.appendChild(el('span', { class: 'slide-no' }, 'Слайд ' + it.slide));
        li.appendChild(document.createTextNode(capital(it.text)));
        ul.appendChild(li);
      });
      box.appendChild(ul);
    }

    var words = Object.keys(r.missing); box = $('r-words');
    box.innerHTML = '';
    if (!words.length) box.appendChild(el('p', { class: 'empty' }, 'Все слова из старой презентации на месте.'));
    else {
      box.appendChild(el('p', { style: 'margin:0' }, 'Эти слова есть в старой презентации, но не нашлись в новой:'));
      var w = el('div', { class: 'words' });
      words.forEach(function (k) { w.appendChild(el('span', null, r.missing[k] > 1 ? k + ' ×' + r.missing[k] : k)); });
      box.appendChild(w);
      box.appendChild(el('p', { class: 'muted', style: 'margin:12px 0 0;font-size:13px' },
        'Если это кнопки старого шаблона («Вернуться к навигации» и т. п.) или слова, разбитые в исходнике на части, — всё в порядке.'));
    }

    var list = $('r-manual'); list.innerHTML = '';
    meta.manual_checks.forEach(function (t) {
      var li = el('li'), lab = el('label');
      lab.appendChild(el('input', { type: 'checkbox' })); lab.appendChild(el('span', null, t));
      li.appendChild(lab); list.appendChild(li);
    });
    job = null;
  }

  // ---------------------------------------------------------------- запуск
  function reset() {
    job = null; state.file = null;
    $('file').value = ''; $('lesson').value = ''; $('subject').value = '';
    $('drop').classList.remove('has-file');
    $('drop-main').textContent = 'Перетащите .pptx сюда';
    $('drop-sub').textContent = dropHint();
    formError(''); show('form');
  }

  // страницу нельзя закрывать, пока идёт обработка: всё происходит в этой вкладке
  window.addEventListener('beforeunload', function (e) { if (job) { e.preventDefault(); e.returnValue = ''; } });

  fetch('meta.json', { cache: 'no-cache' }).then(function (r) { return r.json(); }).then(function (d) {
    meta = d;
    $('drop-sub').textContent = dropHint();
    meta.subjects.forEach(function (s) { $('subjects').appendChild(el('option', { value: s })); });
    renderTemplates(); bindForm();
    document.querySelectorAll('[data-new]').forEach(function (b) { b.addEventListener('click', reset); });
    startWorker();
  }).catch(function () { document.body.textContent = 'Страница не загрузилась. Обновите её; если не помогло — напишите автору.'; });
})();
