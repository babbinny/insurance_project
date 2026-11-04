# -*- coding: utf-8 -*-
"""
인터넷(jsDelivr CDN)을 쓸 수 없는 곳에 web/ 을 올릴 때 쓰는 Pyodide 묶음 만들기.

  python make_bundle.py <Pyodide 0.26.4 파일 폴더> <휠 폴더> web/pyodide

- <Pyodide 폴더>: pyodide-0.26.4.tar.bz2 를 푼 pyodide/ 폴더
    (https://github.com/pyodide/pyodide/releases/tag/0.26.4)
- <휠 폴더>: PyPI 에서 받은 openpyxl / et_xmlfile / pdfminer.six / pdfplumber 휠
    pip download --no-deps openpyxl==3.1.5 et_xmlfile==2.0.0 pdfminer.six==20231228 pdfplumber==0.11.4 -d <휠 폴더>
- .js / .wasm / .json 은 그대로 두고, .zip / .whl 은 웹서버가 막는 경우가 많아
  base64 글자 파일(.txt, 15MB 이하 조각)로 바꿉니다. worker.js 가 bundle.json 을 보고 다시 합칩니다.
"""
import base64
import glob
import json
import os
import shutil
import sys

PART_BYTES = 11 * 1024 * 1024          # base64 로 바꾸면 약 15MB
COPY = ["pyodide.js", "pyodide.asm.js", "pyodide.asm.wasm", "pyodide-lock.json"]
PYODIDE_WHEELS = ["numpy-", "pandas-", "python_dateutil-", "pytz-", "six-", "tzdata-", "charset_normalizer-"]
PYPI_WHEELS = ["openpyxl-", "et_xmlfile-", "pdfminer.six-", "pdfplumber-"]


def encode(src, out_dir):
    name = os.path.basename(src)
    with open(src, "rb") as f:
        data = f.read()
    parts = []
    for i in range(0, len(data), PART_BYTES):
        part = f"{name}.{len(parts)}.txt"
        with open(os.path.join(out_dir, part), "w") as f:
            f.write(base64.b64encode(data[i:i + PART_BYTES]).decode())
        parts.append(part)
    return {"name": name, "parts": parts}


def pick(folder, prefix):
    found = glob.glob(os.path.join(folder, prefix + "*.whl"))
    if len(found) != 1:
        raise SystemExit(f"{folder} 에서 {prefix}*.whl 을 하나만 찾아야 합니다: {found}")
    return found[0]


def main(pyodide_dir, wheel_dir, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    for f in COPY:
        shutil.copy(os.path.join(pyodide_dir, f), out_dir)
    manifest = {
        "stdlib": encode(os.path.join(pyodide_dir, "python_stdlib.zip"), out_dir),
        "wheels": [encode(pick(pyodide_dir, p), out_dir) for p in PYODIDE_WHEELS]
                  + [encode(pick(wheel_dir, p), out_dir) for p in PYPI_WHEELS],
    }
    with open(os.path.join(out_dir, "bundle.json"), "w") as f:
        json.dump(manifest, f, indent=1)
    print(f"{out_dir}: 파일 {len(os.listdir(out_dir))}개")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    main(*sys.argv[1:])
