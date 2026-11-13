// Проверка версии «в браузере» без браузера: тот же Pyodide запускается в Node и прогоняет тестовые
// презентации (browser/tests/run.sh). На каждую презентацию — свежий Python, как на странице.
import fs from 'node:fs';
import { loadPyodide } from '/srv/dist/browser/pyodide/pyodide.mjs';

const D = '/srv/dist/browser/';
const meta = JSON.parse(fs.readFileSync(D + 'meta.json', 'utf8'));
const args = process.argv.slice(2);
const only = args.includes('--only') ? args[args.indexOf('--only') + 1] : null;
const templates = args.filter((a, i) => !a.startsWith('--') && args[i - 1] !== '--only');

async function python() {
  const py = await loadPyodide({ indexURL: D + 'pyodide/', packageBaseUrl: D + 'pyodide/' });
  await py.loadPackage(meta.packages.concat(meta.wheels.map((w) => D + w)), { messageCallback() {} });
  py.FS.mkdirTree('/usr/share/fonts/truetype/open-sans');
  for (const f of ['OpenSans-Regular.ttf', 'OpenSans-Bold.ttf'])
    py.FS.writeFile('/usr/share/fonts/truetype/open-sans/' + f, fs.readFileSync(D + 'fonts/' + f));
  py.FS.mkdirTree('/srv');
  py.FS.mount(py.FS.filesystems.NODEFS, { root: '/srv' }, '/srv');  // проект — как в Docker
  py.runPython('import sys; sys.path.insert(0, "/srv/browser/tests"); import regress');
  return py;
}

let failed = false;
const summary = [];
for (const t of templates.length ? templates : meta.templates.map((x) => x.id)) {
  const cases = JSON.parse((await python()).runPython(`regress.cases(${JSON.stringify(t)})`));
  for (const c of cases) {
    if (only && only !== c.key) continue;
    const py = await python();
    let r;
    try {
      r = JSON.parse(py.globals.get('regress').run(t, c.src, c.key, c.lesson));
    } catch (e) {
      r = { status: 'СБОЙ: ' + String(e).split('\n').slice(-3).join(' ') };
    }
    const heap = Math.round(py._module.HEAPU8.length / 1e6);
    if (r.status !== 'совпадает') failed = true;
    console.log(`${t} · ${c.key} (${c.mb} МБ): ${r.status} | ${r.seconds ?? '—'} с | память Python ${heap} МБ | ` +
                `слайдов ${r.slides ?? '—'} | MISSING ${JSON.stringify(r.missing ?? {})}`);
    summary.push([t, c.key, r.status]);
  }
}
console.log('\nСводка:');
for (const [t, k, s] of summary) console.log(' ', t, '·', k, '—', s.slice(0, 200));
process.exit(failed ? 1 : 0);
