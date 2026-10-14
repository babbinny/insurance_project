# -*- coding: utf-8 -*-
"""
웹페이지(web/index.html)에서 benefit_extractor.run() 을 부르는 연결 코드.
- 브라우저 안의 Python(Pyodide)에서 실행됩니다. 약관 파일은 서버로 올라가지 않습니다.
- 추출 규칙은 benefit_extractor.py 그대로 사용하고, 결과 엑셀을 화면 미리보기용 JSON 으로도 돌려줍니다.
"""
import json
import math

import pandas as pd

import benefit_extractor as be


def _cell(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if hasattr(v, "item"):          # numpy 숫자 → 파이썬 숫자
        v = v.item()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float, str, bool)):
        return v
    return str(v)


def sheets_json(xlsx_path):
    """결과 엑셀의 시트들 → [{name, columns, rows}] (엑셀에 적힌 값 그대로: '01' 은 '01')"""
    sheets = pd.read_excel(xlsx_path, sheet_name=None, dtype=object)
    out = []
    for name, df in sheets.items():
        out.append({"name": name, "columns": [str(c) for c in df.columns],
                    "rows": [[_cell(v) for v in row] for row in df.itertuples(index=False)]})
    return json.dumps(out, ensure_ascii=False)


def run_job(src, out_dir, code_xlsx=None, log=print):
    """Spyder 화면의 [실행] 과 같은 일: 추출 → 엑셀 저장. (엑셀 경로, 미리보기 JSON)"""
    path = be.run(src, out_dir, code_xlsx or None, log=log)
    return path, sheets_json(path)
