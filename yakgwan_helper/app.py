# -*- coding: utf-8 -*-
"""
약관 맞춤 요약 웹사이트 (서버)
----------------------------------------------------------------
실행:  uvicorn app:app --reload        →  브라우저에서 http://127.0.0.1:8000
필요:  ANTHROPIC_API_KEY 환경변수 (.env 파일에 적어도 됩니다)
"""
import os
import time
import uuid
from pathlib import Path

import anthropic
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

BASE = Path(__file__).parent


def _load_env():
    """.env 파일(KEY=VALUE)을 읽어 환경변수로 넣습니다. 이미 있는 값은 그대로 둡니다."""
    env = BASE / ".env"
    if env.exists():
        for line in env.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_env()

import pdf_reader  # noqa: E402  (.env 를 먼저 읽은 뒤 불러옵니다)
import summarizer  # noqa: E402

MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "50"))
APP_PASSWORD = os.getenv("APP_PASSWORD", "")   # 인터넷에 공개할 때 남이 내 API 키를 못 쓰게 하는 비밀번호
DOC_TTL_SECONDS = 60 * 60 * 3        # 업로드한 약관은 3시간 뒤 메모리에서 지움
MAX_DOCS = 20                        # 서버 메모리에 동시에 보관할 약관 수

app = FastAPI(title="약관 맞춤 요약")
app.mount("/static", StaticFiles(directory=BASE / "static"), name="static")

@app.middleware("http")
async def check_password(request, call_next):
    if APP_PASSWORD and request.url.path in ("/api/upload", "/api/summarize"):
        if request.headers.get("x-app-password", "") != APP_PASSWORD:
            return JSONResponse({"detail": "비밀번호가 틀렸습니다."}, status_code=401)
    return await call_next(request)


# doc_id → {"filename", "pages", "chars", "tokens", "scanned", "pdf", "created"}
DOCS = {}


def _cleanup():
    now = time.time()
    for k in [k for k, d in DOCS.items() if now - d["created"] > DOC_TTL_SECONDS]:
        DOCS.pop(k, None)
    while len(DOCS) > MAX_DOCS:
        DOCS.pop(min(DOCS, key=lambda k: DOCS[k]["created"]))


@app.get("/")
def index():
    return FileResponse(BASE / "static" / "index.html")


@app.get("/api/health")
def health():
    return {"ok": True, "api_key": bool(os.getenv("ANTHROPIC_API_KEY")),
            "password": bool(APP_PASSWORD), "model": summarizer.MODEL}


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    data = await file.read()
    if not data[:5] == b"%PDF-":
        raise HTTPException(400, "PDF 파일만 올릴 수 있습니다.")
    if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(400, f"파일이 너무 큽니다 (최대 {MAX_UPLOAD_MB}MB).")

    try:
        pages = pdf_reader.extract_pages(data)
    except Exception as e:
        raise HTTPException(400, f"PDF를 읽지 못했습니다: {e}")

    scanned = pdf_reader.is_scanned(pages)
    if scanned and (len(pages) > summarizer.MAX_PDF_PAGES or len(data) > summarizer.MAX_PDF_BYTES):
        raise HTTPException(400, "글자를 읽을 수 없는 스캔 PDF인데 너무 큽니다 "
                                 f"(스캔본은 {summarizer.MAX_PDF_PAGES}쪽·30MB 이하만 가능). "
                                 "텍스트가 있는 PDF를 올려 주세요.")

    full = summarizer.pages_to_text(pages, len(pages), False)
    tokens = 0 if scanned else await summarizer.count_tokens(full)

    _cleanup()
    doc_id = uuid.uuid4().hex
    DOCS[doc_id] = {
        "filename": file.filename or "약관.pdf",
        "pages": pages,
        "chars": len(full),
        "tokens": tokens,
        "scanned": scanned,
        "pdf": data if scanned else None,
        "created": time.time(),
    }
    return {
        "doc_id": doc_id,
        "filename": file.filename,
        "page_count": len(pages),
        "title": pdf_reader.guess_title(pages),
        "tokens": tokens,
        "scanned": scanned,
        "will_trim": (not scanned) and tokens > summarizer.MAX_DOC_TOKENS,
    }


class SummarizeRequest(BaseModel):
    doc_id: str
    mode: str = "general"            # general | personal | question
    profile: dict | None = None
    question: str = ""


@app.post("/api/summarize")
async def summarize(req: SummarizeRequest):
    doc = DOCS.get(req.doc_id)
    if not doc:
        raise HTTPException(404, "약관을 찾을 수 없습니다. 다시 업로드해 주세요.")
    if req.mode not in ("general", "personal", "question"):
        raise HTTPException(400, "mode 값이 올바르지 않습니다.")
    if req.mode == "question" and not req.question.strip():
        raise HTTPException(400, "질문을 입력해 주세요.")
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise HTTPException(500, "서버에 ANTHROPIC_API_KEY가 설정되지 않았습니다. README를 참고하세요.")

    async def gen():
        try:
            async for chunk in summarizer.stream_summary(doc, req.mode, req.profile, req.question):
                yield chunk
        except anthropic.AuthenticationError:
            yield "\n\n> ❌ API 키가 올바르지 않습니다. .env 의 ANTHROPIC_API_KEY 를 확인하세요."
        except anthropic.RateLimitError:
            yield "\n\n> ❌ 요청이 너무 많습니다. 잠시 후 다시 시도하세요."
        except anthropic.BadRequestError as e:
            yield f"\n\n> ❌ 요청 오류: {e.message}"
        except anthropic.APIStatusError as e:
            yield f"\n\n> ❌ AI 서버 오류({e.status_code}). 잠시 후 다시 시도하세요."
        except anthropic.APIConnectionError:
            yield "\n\n> ❌ AI 서버에 연결하지 못했습니다. 인터넷 연결을 확인하세요."

    return StreamingResponse(gen(), media_type="text/plain; charset=utf-8",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})
