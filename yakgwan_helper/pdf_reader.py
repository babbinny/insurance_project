# -*- coding: utf-8 -*-
"""약관 PDF에서 페이지별 텍스트를 뽑습니다."""
import io
import re

import pdfplumber

# 페이지당 글자가 이보다 적으면 스캔(이미지) PDF로 봅니다
SCANNED_CHARS_PER_PAGE = 40


def _clean(text):
    text = text.replace("\x00", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_pages(pdf_bytes):
    """PDF 바이트 → [{"page": 1, "text": "..."}, ...]"""
    pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            try:
                text = page.extract_text() or ""
            except Exception:
                text = ""
            pages.append({"page": i, "text": _clean(text)})
    return pages


def is_scanned(pages):
    if not pages:
        return True
    total = sum(len(p["text"]) for p in pages)
    return total / len(pages) < SCANNED_CHARS_PER_PAGE


def guess_title(pages):
    """첫 몇 페이지에서 상품명처럼 보이는 줄을 찾습니다."""
    for p in pages[:5]:
        for line in p["text"].splitlines():
            line = line.strip()
            if 6 <= len(line) <= 60 and "보험" in line and "약관" not in line[:2]:
                return line
    return ""
