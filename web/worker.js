// 브라우저 안에서 Python(Pyodide)을 띄워 benefit_extractor.py 를 그대로 실행하는 작업자.
// 화면(index.html)과 메시지로 주고받습니다:
//   화면 → {type:"run", files:[{path, data}], code:{name, data}|null}
//   작업자 → {type:"status"|"log"|"done"|"error", ...}

const PYODIDE_VERSION = "0.26.4";
const CDN = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;
// 같은 곳에 pyodide/ 폴더를 함께 올리면(인터넷 CDN이 막힌 곳) 그 파일을 씁니다.
const LOCAL = new URL("pyodide/", self.location).href;
const PURE_WHEELS = [
  "openpyxl-3.1.5-py2.py3-none-any.whl",
  "et_xmlfile-2.0.0-py3-none-any.whl",
  "pdfminer.six-20231228-py3-none-any.whl",
  "pdfplumber-0.11.4-py3-none-any.whl",
];
const PANDAS_WHEEL = "pandas-2.2.0-cp312-cp312-pyodide_2024_0_wasm32.whl";
const PANDAS_PARTS = ["pandas.whl.part0", "pandas.whl.part1"];

const status = (text) => postMessage({ type: "status", text });
const log = (m) => postMessage({ type: "log", m: String(m) });

async function fetchOk(url, opts) {
  const r = await fetch(url, opts);
  if (!r.ok) throw new Error(`${url} (${r.status})`);
  return r;
}

async function hasLocalBundle() {
  try {
    return (await fetch(LOCAL + "pyodide-lock.json", { method: "HEAD" })).ok;
  } catch {
    return false;
  }
}

async function fetchText(...urls) {
  for (const u of urls) {
    try {
      return await (await fetchOk(new URL(u, self.location).href)).text();
    } catch {}
  }
  throw new Error(`${urls[0]} 을(를) 불러오지 못했습니다.`);
}

async function boot() {
  const local = await hasLocalBundle();
  const base = local ? LOCAL : CDN;
  status("Python 불러오는 중");
  importScripts(base + "pyodide.js");
  const py = await loadPyodide({ indexURL: base });

  status("pandas · openpyxl · pdfplumber 설치 중");
  const deps = ["numpy", "python-dateutil", "pytz", "six", "tzdata", "micropip", "cryptography", "charset-normalizer"];
  if (local) {
    await py.loadPackage(deps.concat(PURE_WHEELS.map((w) => LOCAL + w)));
    // pandas 휠은 올릴 수 있는 파일 크기 제한 때문에 둘로 나눠 둠 → 합쳐서 설치
    const parts = await Promise.all(PANDAS_PARTS.map(async (p) => new Uint8Array(await (await fetchOk(LOCAL + p)).arrayBuffer())));
    const whl = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
    parts.reduce((off, p) => (whl.set(p, off), off + p.length), 0);
    py.FS.writeFile("/tmp/" + PANDAS_WHEEL, whl);
    const micropip = py.pyimport("micropip");
    await micropip.install.callKwargs("emfs:/tmp/" + PANDAS_WHEEL, { deps: false });
  } else {
    await py.loadPackage(deps.concat(["pandas"]));
    const micropip = py.pyimport("micropip");
    await micropip.install(["openpyxl==3.1.5", "pdfminer.six==20231228"]);
    // pdfplumber 의 그림 기능(pypdfium2)은 쓰지 않으므로 의존성 없이 설치
    await micropip.install.callKwargs("pdfplumber==0.11.4", { deps: false });
  }

  status("추출 프로그램 불러오는 중");
  py.FS.mkdirTree("/app");
  py.FS.writeFile("/app/benefit_extractor.py", await fetchText("benefit_extractor.py", "../benefit_extractor.py"));
  py.FS.writeFile("/app/web_runner.py", await fetchText("web_runner.py"));
  py.runPython("import sys; sys.path.insert(0, '/app'); import web_runner");
  return py;
}

const ready = boot().then(
  (py) => {
    postMessage({ type: "ready" });
    return py;
  },
  (e) => {
    postMessage({ type: "fatal", message: String(e && e.message ? e.message : e) });
    throw e;
  }
);

let jobNo = 0;

self.onmessage = async (ev) => {
  const msg = ev.data;
  if (msg.type !== "run") return;
  let py;
  try {
    py = await ready;
  } catch {
    return;
  }
  const dir = `/work/job${++jobNo}`;
  try {
    const inDir = dir + "/약관";
    const outDir = dir + "/결과";
    py.FS.mkdirTree(inDir);
    py.FS.mkdirTree(outDir);
    for (const f of msg.files) {
      const p = inDir + "/" + f.path.replace(/^\/+/, "");
      py.FS.mkdirTree(p.slice(0, p.lastIndexOf("/")));
      py.FS.writeFile(p, new Uint8Array(f.data));
    }
    let codePath = null;
    if (msg.code) {
      codePath = dir + "/" + msg.code.name;
      py.FS.writeFile(codePath, new Uint8Array(msg.code.data));
    }
    const runJob = py.globals.get("web_runner").run_job;
    // 브라우저 안 임시 경로(/work/job1/...)는 화면에 보이지 않게 뗌
    const tidy = (s) => String(s).split(outDir + "/").join("").split(inDir + "/").join("");
    const res = runJob.callKwargs(inDir, outDir, codePath, { log: (m) => log(tidy(m)) });
    const [path, sheets] = res.toJs();
    res.destroy();
    runJob.destroy();
    const bytes = py.FS.readFile(path);
    const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    postMessage({ type: "done", name: path.split("/").pop(), xlsx: buf, sheets: tidy(sheets) }, [buf]);
  } catch (e) {
    let m = String(e && e.message ? e.message : e);
    // Python 오류는 마지막 줄(예: RuntimeError: Word/PDF 파일을 찾지 못했습니다.)만 보여줌
    const lines = m.trim().split("\n");
    if (lines.length > 1 && /Error|Exception/.test(lines[lines.length - 1])) m = lines[lines.length - 1].replace(/^\w+(Error|Exception):\s*/, "");
    postMessage({ type: "error", message: m });
  } finally {
    try {
      py.runPython(`import shutil; shutil.rmtree(${JSON.stringify(dir)}, ignore_errors=True)`);
    } catch {}
  }
};
