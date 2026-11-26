const BASE = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
let py;
let ready;
let chain = Promise.resolve();

async function ensurePy() {
  if (ready) return ready;
  ready = (async () => {
    postMessage({ type: 'status', text: 'Python 실행 환경 준비 중…' });
    importScripts(BASE + 'pyodide.js');
    py = await loadPyodide({ indexURL: BASE });
    await py.loadPackage(['numpy', 'micropip']);
    await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
    const response = await fetch('./model.py');
    if (!response.ok) throw Error('model.py를 불러올 수 없습니다.');
    await py.runPythonAsync(await response.text());
  })();
  return ready;
}

async function handle({ id, type, payload }) {
  try {
    let result;
    if (type === 'mapping') {
      await ensurePy();
      postMessage({ type: 'status', text: '분류 연결표 읽는 중…' });
      py.globals.set('raw_input', new Uint8Array(payload));
      await py.runPythonAsync('raw_bytes = raw_input.to_py()');
      result = await py.runPythonAsync('load_mapping(raw_bytes)');
      await py.runPythonAsync('del raw_input, raw_bytes');
    } else if (type === 'init') {
      await ensurePy();
      postMessage({ type: 'status', text: payload.levelKr + ' 자료 읽기·계수 검증 중…' });
      py.globals.set('raw_input', new Uint8Array(payload.bytes));
      py.globals.set('lvl_input', payload.level);
      await py.runPythonAsync('raw_bytes = raw_input.to_py()');
      result = await py.runPythonAsync('initialize(lvl_input, raw_bytes)');
      await py.runPythonAsync('del raw_input, raw_bytes');
    } else if (type === 'explore') {
      py.globals.set('lvl_input', payload.level);
      py.globals.set('sheet_input', payload.sheet);
      py.globals.set('row_input', payload.row || '');
      result = py.runPython('explore(lvl_input, sheet_input, row_input)');
    } else if (type === 'simulate') {
      py.globals.set('lvl_input', payload.level);
      py.globals.set('scenario_input', JSON.stringify(payload.shocks));
      result = py.runPython('simulate(lvl_input, scenario_input)');
    } else if (type === 'analysis') {
      py.globals.set('which_input', payload.which);
      py.globals.set('args_input', JSON.stringify(payload.args || {}));
      result = py.runPython('analysis(which_input, args_input)');
    } else {
      throw Error('알 수 없는 요청: ' + type);
    }
    postMessage({ id, ok: true, result: JSON.parse(result) });
  } catch (error) {
    postMessage({ id, ok: false, error: String(error) });
  }
}

onmessage = e => { chain = chain.then(() => handle(e.data)); };
