// Версия «в браузере»: конвертер (Python) работает в отдельном потоке, чтобы страница не замирала.
// Python здесь — Pyodide; сам перенос делает тот же код, что на сайте (browser/runner.py → app/pipeline.py).
// Один поток — одна презентация: после неё страница запускает новый, чтобы память браузера освободилась.
import { loadPyodide } from './pyodide/pyodide.mjs';

const here = (path) => new URL(path, import.meta.url).href;
const post = (m, transfer) => self.postMessage(m, transfer || []);
const FONTS = '/usr/share/fonts/truetype/open-sans/';  // путь, по которому converter/engine.py ищет шрифт
let py = null, v = '';

// файл по адресу → байты; onpct — доля загруженного (для больших файлов: база шаблона ~50 МБ)
async function download(url, onpct) {
  const r = await fetch(url);
  if (!r.ok) throw new Error('не загрузился файл ' + url + ' (' + r.status + ')');
  const total = +r.headers.get('Content-Length') || 0;
  if (!onpct || !total || !r.body) return new Uint8Array(await r.arrayBuffer());
  const chunks = [];
  let size = 0;
  for (const reader = r.body.getReader(); ;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value); size += value.length;
    onpct(Math.min(1, size / total));
  }
  const out = new Uint8Array(size);
  let at = 0;
  for (const c of chunks) { out.set(c, at); at += c.length; }
  return out;
}

async function init(meta) {
  v = '?v=' + meta.version;
  py = await loadPyodide({ indexURL: here('pyodide/'), packageBaseUrl: here('pyodide/') });
  await py.loadPackage(meta.packages.concat(meta.wheels.map(here)), { messageCallback() {} });
  py.FS.mkdirTree(FONTS);
  for (const f of ['OpenSans-Regular.ttf', 'OpenSans-Bold.ttf']) py.FS.writeFile(FONTS + f, await download(here('fonts/' + f)));
  py.FS.mkdirTree('/srv');
  py.unpackArchive((await download(here('code.zip') + v)).buffer, 'zip', { extractDir: '/srv' });
  py.runPython('import sys; sys.path.insert(0, "/srv/browser"); import runner');
}

async function convert(m) {
  post({ status: 'base', pct: 0 });
  const base = await download(here(m.base) + v, (pct) => post({ status: 'base', pct: pct }));
  py.FS.mkdirTree('/srv/' + m.base.replace(/\/[^/]*$/, ''));
  py.FS.writeFile('/srv/' + m.base, base);
  py.FS.mkdirTree('/work');
  py.FS.writeFile('/work/source.pptx', new Uint8Array(await m.file.arrayBuffer()));
  const run = py.runPython('runner.convert');
  const out = JSON.parse(run.callKwargs('/work/source.pptx', '/work', m.file.name, m.template, m.lesson, {
    choices: JSON.stringify(m.choices), subject: m.subject,
    progress: (step) => post({ status: 'step', step: step }),
  }));
  if (out.error) return post({ status: 'error', error: out.error, trace: out.trace || '' });
  const file = py.FS.readFile('/work/result.pptx');
  post({ status: 'done', result: out.result, file: file }, [file.buffer]);
}

let ready = null;
self.onmessage = (e) => {
  const m = e.data;
  if (m.type === 'init') {
    ready = init(m.meta);
    ready.then(() => post({ status: 'ready' }), () => {});  // ошибка загрузки покажется при переносе
    return;
  }
  ready.then(() => convert(m)).catch((err) => post({ status: 'error', error: '', trace: String(err && err.stack || err) }));
};
