// Конвертер презентаций: форма → загрузка → очередь → результат.
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var meta = null;          // шаблоны, шаги, тексты памятки — с сервера
  var state = { template: null, grade: null, file: null, choices: {} };
  var job = null;           // последнее состояние задачи
  var pollTimer = null;
  var slides = [];          // имена картинок слайдов результата
  var notesBySlide = {};    // номер слайда → замечания
  var viewerIndex = 0;

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

  function api(method, url, body) {
    return fetch(url, { method: method, body: body, credentials: 'same-origin' }).then(function (r) {
      if (r.status === 401) { location.href = '/login'; throw new Error('auth'); }
      return r.json().then(function (d) { if (!r.ok) throw new Error(d.error || 'Ошибка сервера.'); return d; });
    });
  }

  // ---------------------------------------------------------------- форма
  function currentTemplate() {
    return meta.templates.filter(function (t) { return t.id === state.template; })[0];
  }

  function renderTemplates() {
    var box = $('templates');
    meta.templates.forEach(function (t) {
      var b = el('button', { type: 'button', class: 'tab', 'data-id': t.id }, t.name);
      b.addEventListener('click', function () { selectTemplate(t.id); });
      box.appendChild(b);
    });
    box.hidden = meta.templates.length < 2;  // выбирать не из чего — полосу не показываем
    selectTemplate(meta.templates[0].id);
  }

  function selectTemplate(id) {
    state.template = id;
    document.querySelectorAll('#templates .tab').forEach(function (b) {
      b.setAttribute('aria-pressed', String(b.getAttribute('data-id') === id));
    });
    var grades = currentTemplate().grades;
    var box = $('grades'); box.innerHTML = '';
    grades.forEach(function (g) {
      var b = el('button', { type: 'button', 'data-grade': g, 'aria-label': g + ' класс' }, g);
      b.addEventListener('click', function () { selectGrade(g); });
      box.appendChild(b);
    });
    selectGrade(grades.indexOf(state.grade) >= 0 ? state.grade : grades[0]);
    renderChoices();
  }

  // варианты слайдов шаблона (титул, прощание…): строка формы с миниатюрами на каждый.
  // Необязательный выбор (optional) начинается с плитки «Как в старой»: слайд останется из старой презентации
  function renderChoices() {
    var t = currentTemplate(), rows = $('choice-rows');
    rows.innerHTML = ''; state.choices = {};
    t.choices.forEach(function (c, ci) {
      var row = el('div', { class: 'row' });
      var lab = el('div', { class: 'row__label mono muted' }, ('0' + (5 + ci)).slice(-2));
      lab.appendChild(el('b', null, c.label));
      var body = el('div'), grid = el('div', { class: 'titles', role: 'radiogroup', 'aria-label': c.label });
      if (c.optional) {
        var old = el('button', { type: 'button', class: 'title-opt title-opt--old', role: 'radio', 'data-n': 0,
                                 'aria-label': c.label + ': как в старой презентации' });
        old.appendChild(el('i', { class: 'mono' }, 'Как в старой'));
        old.addEventListener('click', function () { selectChoice(grid, c.key, 0); });
        grid.appendChild(old);
      }
      c.slides.forEach(function (n, i) {
        var b = el('button', { type: 'button', class: 'title-opt', role: 'radio', 'data-n': n,
                               'aria-label': c.label + ' ' + (i + 1) });
        b.appendChild(el('img', { src: '/api/templates/' + t.id + '/thumbs/' + n + '.jpg', alt: '' }));
        b.appendChild(el('span', { class: 'mono' }, ('0' + (i + 1)).slice(-2)));
        b.addEventListener('click', function () { selectChoice(grid, c.key, n); });
        grid.appendChild(b);
      });
      body.appendChild(grid);
      if (c.hint) body.appendChild(el('p', { class: 'hint mono muted' }, c.hint));
      row.appendChild(lab); row.appendChild(body); rows.appendChild(row);
      selectChoice(grid, c.key, c.optional ? 0 : c.default);
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

  function submit() {
    var n = $('lesson').value;
    if (!state.file) return formError('Выберите файл презентации.');
    if (!n || +n < 1) return formError('Укажите номер урока.');
    formError('');
    var lesson = state.grade + '.' + (+n);
    var fd = new FormData();
    fd.append('file', state.file); fd.append('template', state.template); fd.append('lesson', lesson);
    fd.append('choices', JSON.stringify(state.choices));
    fd.append('subject', $('subject').value.trim());

    showProgress({ filename: state.file.name, lesson: lesson, status: 'upload' });
    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/jobs');
    xhr.upload.onprogress = function (e) {
      if (e.lengthComputable) showProgress({ filename: state.file.name, lesson: lesson, status: 'upload', pct: e.loaded / e.total });
    };
    xhr.onload = function () {
      var d = {}; try { d = JSON.parse(xhr.responseText); } catch (_) {}
      if (xhr.status === 401) { location.href = '/login'; return; }
      if (xhr.status !== 200) { show('form'); formError(d.error || 'Не удалось загрузить файл.'); return; }
      location.hash = 'job=' + d.id;  // дальше — route() по событию hashchange
    };
    xhr.onerror = function () { show('form'); formError('Не удалось загрузить файл. Проверьте интернет и попробуйте ещё раз.'); };
    xhr.send(fd);
  }

  // ---------------------------------------------------------------- ход работы
  function showProgress(j) {
    show('progress');
    $('p-file').textContent = (j.lesson ? 'Урок ' + j.lesson + ' · ' : '') + j.filename;
    var steps = meta.steps, now = -1, pct = 0;
    if (j.status === 'upload') {
      $('p-status').textContent = 'Загрузка файла' + (j.pct != null ? ' · ' + Math.round(j.pct * 100) + ' %' : '');
      $('p-sub').textContent = 'Не закрывайте страницу, пока файл загружается.';
      pct = (j.pct || 0) * 0.15;
    } else if (j.status === 'queued') {
      $('p-status').textContent = 'В очереди';
      $('p-sub').textContent = j.ahead ? 'Перед вами презентаций: ' + j.ahead + '. Обработка одной занимает около минуты.' : 'Сейчас начнём.';
      pct = 0.15;
    } else {
      now = steps.map(function (s) { return s[0]; }).indexOf(j.step);
      $('p-status').textContent = now >= 0 ? steps[now][1] : 'Обработка';
      $('p-sub').textContent = 'Обычно это занимает около минуты.';
      pct = 0.15 + 0.85 * Math.max(now, 0) / steps.length;
    }
    $('p-bar').style.width = Math.round(pct * 100) + '%';
    var ol = $('p-steps'); ol.innerHTML = '';
    steps.forEach(function (s, i) {
      var li = el('li', { class: i < now ? 'is-done' : i === now ? 'is-now' : '' });
      li.appendChild(el('span', { class: 'mono' }, ('0' + (i + 1)).slice(-2)));
      li.appendChild(el('span', null, s[1]));
      li.appendChild(el('span', { class: 'mark', 'aria-hidden': 'true' }));
      ol.appendChild(li);
    });
    var waiting = j.status !== 'upload';
    $('p-note').textContent = waiting ? 'Можно закрыть страницу и вернуться по ссылке — результат хранится ' + hours() + '.' : '';
    $('copy-link').hidden = !waiting;
  }

  function hours() {
    var h = meta.ttl_hours;
    return h % 24 === 0 ? (h / 24 === 1 ? 'сутки' : h / 24 + ' сут.') : h + ' ч';
  }

  function poll(id) {
    clearTimeout(pollTimer);
    api('GET', '/api/jobs/' + id).then(function (j) {
      job = j;
      if (j.status === 'done') return showResult(j);
      if (j.status === 'error') return showError(j.error, j);
      showProgress(j);
      pollTimer = setTimeout(function () { poll(id); }, 1500);
    }).catch(function (e) {
      if (e.message !== 'auth') showError(e.message);
    });
  }

  function showError(msg, j) {
    show('error');
    $('e-file').textContent = j ? 'Урок ' + j.lesson + ' · ' + j.filename : '';
    $('e-message').textContent = msg;
    var c = $('e-contact'); c.innerHTML = '';
    if (meta.contact) {
      c.appendChild(document.createTextNode('Если ошибка повторяется, сообщите автору: '));
      var isMail = /@/.test(meta.contact) && !/^@/.test(meta.contact);
      c.appendChild(isMail ? el('a', { href: 'mailto:' + meta.contact }, meta.contact) : el('b', null, meta.contact));
    }
  }

  // ---------------------------------------------------------------- результат
  function slideButton(n) {
    var b = el('button', { type: 'button', class: 'slide-link' }, 'Слайд ' + n);
    b.addEventListener('click', function () { openViewer(n - 1); });
    return b;
  }

  function noteList(items, emptyText) {
    if (!items.length) return el('p', { class: 'empty' }, emptyText);
    var ul = el('ul');
    items.forEach(function (it) {
      var li = el('li');
      if (it.slide) li.appendChild(slideButton(it.slide));
      li.appendChild(document.createTextNode(it.slide ? it.text : it.text.charAt(0).toUpperCase() + it.text.slice(1)));
      ul.appendChild(li);
    });
    return ul;
  }

  function showResult(j) {
    var r = j.result;
    show('result');
    $('r-meta').textContent = r.template_name + ' · ' + (r.subject || '') + ' · урок ' + r.lesson + ' · ' +
      String(r.size_mb).replace('.', ',') + ' МБ';
    $('r-file').textContent = j.filename;
    $('r-download').href = '/api/jobs/' + j.id + '/download';

    notesBySlide = {};
    r.stickers.forEach(function (s) { (notesBySlide[s.slide] = notesBySlide[s.slide] || []).push('Стикер: ' + s.text); });
    r.checks.forEach(function (c) { (notesBySlide[c.slide] = notesBySlide[c.slide] || []).push(c.text); });

    $('r-stickers').replaceChildren(noteList(r.stickers, 'Стикеры добавлять не нужно.'));
    $('r-checks').replaceChildren(noteList(r.checks, 'Конвертер ничего не отметил.'));

    var words = Object.keys(r.missing), box = $('r-words');
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

    var ul = $('r-manual'); ul.innerHTML = '';
    meta.manual_checks.forEach(function (t) {
      var li = el('li'), lab = el('label');
      lab.appendChild(el('input', { type: 'checkbox' })); lab.appendChild(el('span', null, t));
      li.appendChild(lab); ul.appendChild(li);
    });

    slides = r.previews;
    $('r-count').textContent = slides.length + ' шт.';
    var grid = $('r-grid'); grid.innerHTML = '';
    slides.forEach(function (name, i) {
      var b = el('button', { type: 'button', class: 'thumb', 'aria-label': 'Слайд ' + (i + 1) });
      b.appendChild(el('img', { src: '/api/jobs/' + j.id + '/thumbs/' + name, alt: '', loading: 'lazy' }));
      var cap = el('span', { class: 'thumb__cap mono' });
      cap.appendChild(el('span', null, ('0' + (i + 1)).slice(-2)));
      var n = notesBySlide[i + 1];
      if (n) cap.appendChild(el('span', { class: 'flag', title: n.join('\n') }, n.some(function (t) { return /^Стикер/.test(t); }) ? 'стикер' : 'проверить'));
      b.appendChild(cap);
      b.addEventListener('click', function () { openViewer(i); });
      grid.appendChild(b);
    });
  }

  // ---------------------------------------------------------------- просмотр слайда
  function openViewer(i) {
    viewerIndex = Math.max(0, Math.min(slides.length - 1, i));
    $('v-img').src = '/api/jobs/' + job.id + '/slides/' + slides[viewerIndex];
    $('v-img').alt = 'Слайд ' + (viewerIndex + 1);
    $('v-title').textContent = 'Слайд ' + (viewerIndex + 1) + ' / ' + slides.length;
    $('v-notes').textContent = (notesBySlide[viewerIndex + 1] || []).join(' · ');
    $('v-prev').disabled = viewerIndex === 0;
    $('v-next').disabled = viewerIndex === slides.length - 1;
    if ($('viewer').hidden) { $('viewer').hidden = false; document.body.style.overflow = 'hidden'; $('v-close').focus(); }
  }

  function closeViewer() { $('viewer').hidden = true; document.body.style.overflow = ''; }

  function bindViewer() {
    $('v-close').addEventListener('click', closeViewer);
    $('v-prev').addEventListener('click', function () { openViewer(viewerIndex - 1); });
    $('v-next').addEventListener('click', function () { openViewer(viewerIndex + 1); });
    document.addEventListener('keydown', function (e) {
      if ($('viewer').hidden) return;
      if (e.key === 'Escape') closeViewer();
      if (e.key === 'ArrowLeft') openViewer(viewerIndex - 1);
      if (e.key === 'ArrowRight') openViewer(viewerIndex + 1);
    });
  }

  // ---------------------------------------------------------------- запуск
  function reset() {
    clearTimeout(pollTimer); job = null; state.file = null;
    history.replaceState(null, '', location.pathname);
    $('file').value = ''; $('lesson').value = ''; $('subject').value = '';
    $('drop').classList.remove('has-file');
    $('drop-main').textContent = 'Перетащите .pptx сюда';
    $('drop-sub').textContent = dropHint();
    formError(''); show('form');
  }

  function route() {
    var m = /job=([0-9a-f]{32})/.exec(location.hash);
    if (m) poll(m[1]); else show('form');
  }

  api('GET', '/api/meta').then(function (d) {
    meta = d;
    $('drop-sub').textContent = dropHint();
    meta.subjects.forEach(function (s) { $('subjects').appendChild(el('option', { value: s })); });
    renderTemplates(); bindForm(); bindViewer();
    document.querySelectorAll('[data-new]').forEach(function (b) { b.addEventListener('click', reset); });
    $('copy-link').addEventListener('click', function () {
      var b = this;
      navigator.clipboard.writeText(location.href).then(function () {
        b.textContent = 'Ссылка скопирована';
        setTimeout(function () { b.textContent = 'Скопировать ссылку на результат'; }, 2000);
      });
    });
    window.addEventListener('hashchange', route);
    route();
  }).catch(function (e) { if (e.message !== 'auth') document.body.textContent = 'Сайт временно недоступен.'; });
})();
