# insurance_project
## 웹페이지로 실행하기 (`web/`)

Spyder에서 F5로 띄우던 창과 같은 화면을 웹페이지로 만든 것입니다. `benefit_extractor2.py` 를 고치지 않고
브라우저 안의 Python(Pyodide)으로 그대로 실행하므로, 약관 파일은 서버로 올라가지 않습니다.

- `web/index.html` : 화면 (약관 폴더/파일 · 급부코드표 선택 → 실행/중단 → 로그 → 결과 시트 미리보기 → 엑셀 내려받기)
- `web/worker.js` : 브라우저 안에서 Python·pandas·openpyxl·pdfplumber 를 불러와 `run()` 을 실행
- `web/web_runner.py` : `run()` 결과 엑셀을 화면 미리보기용으로 바꾸는 연결 코드
- `web/sample/sample_policy.docx`, `web/sample/sample_revision.xlsx` : 시연용 가상 약관과 개정 양식 예제

페이지에는 탭이 두 개 있습니다.

- **약관에서 추출**: 상품 약관(PDF·Word) → 통계테이블 초안 (`benefit_extractor2.py`)
- **개정 후 자동화**: 기존 통계테이블 + 상품코드 변경 → 개정 후 통계테이블 (`stat_builder.py`).
- **약관 정보**: 상품 약관 → 급부별 지급사유 · 지급금액 · 면책기간 · 적립금 지급 여부 요약 (`benefit_extractor2.py` 의 추출 결과를 화면용으로 정리)
  예제 양식은 `web/sample/sample_revision.xlsx` 입니다. 관계 시트는 [주보종명 | 개별상품명 | SP_SPCD | SP_SPCD1] 순서이며,
  상품명은 SP_SPCD 바로 왼쪽 두 칼럼에서 순서대로 읽습니다 (칼럼명이 '주보종명', '개별상품명'이면 위치와 상관없이 그 칼럼).

로컬에서 보기 (저장소 폴더에서):

```
python -m http.server 8000
```

브라우저에서 `http://localhost:8000/web/` 을 엽니다. 처음 열 때 Python 실행 파일(약 40MB)을 인터넷(jsDelivr)에서 받습니다.
인터넷이 막힌 곳에서는 `web/make_bundle.py` 로 Pyodide 묶음을 `web/pyodide/` 에 만들어 두면 그 파일을 씁니다 (사용법은 파일 맨 위 설명).

회사 로고를 화면 왼쪽 위에 넣으려면 공식 로고 이미지를 `web/logo.png` 로 두면 됩니다 (없으면 글자 제목만 보입니다).

웹에서는 PC의 Word를 쓸 수 없어 문서보안(DRM) 파일과 `.doc` 는 읽지 못합니다. 이런 파일은 Spyder에서 실행하세요.
