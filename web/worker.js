// 브라우저 안에서 Python(Pyodide)을 띄워 benefit_extractor2.py 를 그대로 실행하는 작업자.
// 화면(index.html)과 메시지로 주고받습니다:
//   화면 → {type:"run", files:[{path, data}], code:{name, data}|null}
//   작업자 → {type:"status"|"log"|"done"|"error", ...}

const PYODIDE_VERSION = "0.26.4";
const CDN = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/`;
// 같은 곳에 pyodide/ 묶음(make_bundle.py 로 만든 것)을 함께 올리면 인터넷 CDN 대신 그 파일을 씁니다.
const LOCAL = new URL("pyodide/", self.location).href;

// 묶음 모드에는 cryptography 가 없음: pdfminer 는 AES 암호화 PDF 에만 쓰므로 그때만 오류를 내는 대역을 둠
const CRYPTO_STUB = {
  "cryptography/__init__.py": "",
  "cryptography/hazmat/__init__.py": "",
  "cryptography/hazmat/backends/__init__.py": "def default_backend():\n    return None\n",
  "cryptography/hazmat/primitives/__init__.py": "",
  "cryptography/hazmat/primitives/ciphers/__init__.py":
    "def _no(*a, **k):\n    raise RuntimeError('AES로 암호화된 PDF는 이 화면에서 읽을 수 없습니다. 암호를 푼 PDF나 Word 파일로 올려 주세요.')\n" +
    "class _NS:\n    def __getattr__(self, name):\n        return _no\n" +
    "Cipher = _no\nalgorithms = _NS()\nmodes = _NS()\n",
};

const status = (text) => postMessage({ type: "status", text });
const log = (m) => postMessage({ type: "log", m: String(m) });

async function fetchOk(url, opts) {
  const r = await fetch(url, opts);
  if (!r.ok) throw new Error(`${url} (${r.status})`);
  return r;
}

async function localBundle() {
  try {
    const r = await fetch(LOCAL + "bundle.json");
    return r.ok ? await r.json() : null;
  } catch {
    return null;
  }
}

// base64 글자 조각들 → 원래 파일 바이트
async function partsToBytes(item) {
  const chunks = await Promise.all(item.parts.map(async (p) => {
    const bin = atob(await (await fetchOk(LOCAL + p)).text());
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }));
  const all = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  chunks.reduce((off, c) => (all.set(c, off), off + c.length), 0);
  return all;
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
  const bundle = await localBundle();
  status("1/3");
  let py;
  if (bundle) {
    importScripts(LOCAL + "pyodide.js");
    const stdlib = URL.createObjectURL(new Blob([await partsToBytes(bundle.stdlib)]));
    py = await loadPyodide({ indexURL: LOCAL, stdLibURL: stdlib });
    status("2/3");
    for (const w of bundle.wheels) py.unpackArchive(await partsToBytes(w), "wheel");
    for (const [path, src] of Object.entries(CRYPTO_STUB)) {
      py.FS.mkdirTree("/app/" + path.slice(0, path.lastIndexOf("/")));
      py.FS.writeFile("/app/" + path, src);
    }
  } else {
    importScripts(CDN + "pyodide.js");
    py = await loadPyodide({ indexURL: CDN });
    status("2/3");
    await py.loadPackage(["pandas", "micropip", "cryptography", "charset-normalizer"]);
    const micropip = py.pyimport("micropip");
    await micropip.install(["openpyxl==3.1.5", "pdfminer.six==20231228"]);
    // pdfplumber 의 그림 기능(pypdfium2)은 쓰지 않으므로 의존성 없이 설치
    await micropip.install.callKwargs("pdfplumber==0.11.4", { deps: false });
  }

  status("3/3");
  py.FS.mkdirTree("/app");
  py.FS.writeFile("/app/benefit_extractor2.py", await fetchText("benefit_extractor2.py", "../benefit_extractor2.py"));
  py.FS.writeFile("/app/web_runner.py", await fetchText("web_runner.py"));
  py.runPython("import sys, importlib; sys.path.insert(0, '/app'); importlib.invalidate_caches(); import web_runner");
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
