# -*- coding: utf-8 -*-
"""약관 텍스트 + 가입자 정보 → Claude 에게 요약을 요청합니다."""
import base64
import os
import re

import anthropic

# =====================================================================
# 설정 - .env 파일이나 환경변수로 바꿀 수 있습니다
# =====================================================================
MODEL = os.getenv("CLAUDE_MODEL", "claude-opus-5-5")
EFFORT = os.getenv("CLAUDE_EFFORT", "medium")            # low | medium | high | xhigh | max
MAX_OUTPUT_TOKENS = int(os.getenv("MAX_OUTPUT_TOKENS", "32000"))
# 한 번에 보낼 약관 분량(토큰). 넘으면 관련 있는 페이지만 골라서 보냅니다.
MAX_DOC_TOKENS = int(os.getenv("MAX_DOC_TOKENS", "500000"))
# 스캔 PDF를 그대로 보낼 때의 한도 (API 제한: 600쪽, 요청 32MB)
MAX_PDF_PAGES = 600
MAX_PDF_BYTES = 30 * 1024 * 1024

# 약관 앞부분(요약서·목차·가이드북)은 항상 포함
ALWAYS_FIRST_PAGES = 15

# 누구에게나 중요한 약관 키워드 (페이지 고를 때 점수)
CORE_KEYWORDS = {
    "보험금을 지급하지 않는": 6, "보험금을 지급하지 아니하는": 6, "면책": 5, "감액": 4,
    "계약 전 알릴 의무": 6, "계약전 알릴의무": 6, "알릴 의무": 4, "고지": 3,
    "보장개시일": 4, "지급사유": 4, "지급금액": 3, "보험금의 지급": 3,
    "해지환급금": 4, "무해지": 4, "저해지": 4, "납입면제": 4, "갱신": 3,
    "청약의 철회": 3, "청약철회": 3, "품질보증": 2, "보험기간": 2, "가입나이": 3,
    "소멸시효": 2, "부담보": 4, "할증": 3, "특약": 1, "보장내용": 3, "주요내용": 3,
}

# 가입 목적별로 더 볼 키워드
PURPOSE_KEYWORDS = {
    "건강보험": ["진단", "입원", "수술", "통원", "질병", "상해", "중증"],
    "암보험": ["암", "유사암", "소액암", "제자리암", "경계성종양", "갑상선암", "90일"],
    "실손보험": ["실손", "자기부담", "비급여", "급여", "통원", "입원의료비"],
    "사망보험": ["사망", "종신", "정기", "수익자", "재해사망", "자살"],
    "연금·저축": ["연금", "적립", "공시이율", "최저보증", "만기", "중도인출"],
    "투자(변액)": ["변액", "펀드", "특별계정", "원금", "최저보증", "투자", "수수료"],
    "어린이보험": ["어린이", "태아", "선천", "자녀"],
    "치아보험": ["치아", "보철", "임플란트", "크라운", "충전"],
    "운전자보험": ["운전", "교통사고", "벌금", "형사합의", "변호사"],
    "간병·치매": ["간병", "치매", "장기요양", "인지"],
}

SYSTEM_PROMPT = """너는 한국 보험 약관을 소비자 입장에서 읽어 주는 분석가다.

규칙:
- 반드시 제공된 약관 내용만 근거로 답한다. 약관에 없는 내용은 추측하지 말고 "약관에서 확인되지 않음"이라고 쓴다.
- 중요한 문장 끝에는 근거 페이지를 [p.12] 형식으로 붙인다. (문서의 "=== p.N ===" 표시 기준)
- 어려운 보험 용어는 괄호로 쉽게 풀어 쓴다. 예: 면책기간(가입 후 보장이 안 되는 기간)
- 금액·기간·비율(예: 50% 감액, 90일 면책)은 약관에 적힌 숫자를 그대로 쓴다.
- 가입자에게 불리할 수 있는 조건(면책, 감액, 갱신 시 보험료 인상, 해지환급금 손실, 고지의무 위반 시 계약 해지 등)은 빠뜨리지 말고 눈에 띄게 표시한다.
- 가입 심사 기준(인수 기준)은 보통 약관에 없다. 가입 가능 여부를 단정하지 말고, 약관상 근거와 보험사에 확인할 점을 나눠서 쓴다.
- 법률·의학적 판단을 확정하지 말고, 최종 확인은 보험사/설계사에게 하도록 안내한다.
- 형식: 한국어 마크다운. 제목(##), 짧은 목록, 필요할 때 표를 쓴다. 군더더기 없이 핵심만."""

GENERAL_TASK = """위 약관을 처음 보는 소비자가 꼭 알아야 할 내용만 요약해 줘. 아래 순서를 따르고, 약관에 없는 항목은 짧게 "확인되지 않음"으로 표시해.

## 1. 한눈에 보기
상품명, 상품 종류, 보험기간/납입기간, 갱신형·비갱신형, 해지환급금 유형(표준형/무해지/저해지), 가입나이 — 표로.

## 2. 무엇을 보장하나
주계약과 주요 특약의 급부를 표로 (급부명 | 언제 받나 | 얼마 받나 | 횟수). 너무 많으면 핵심 위주로 묶어서.

## 3. 보장이 안 되거나 줄어드는 경우 ⚠️
보험금을 지급하지 않는 사유, 면책기간, 감액기간, 보장 제외 대상(예: 유사암), 주의할 정의(용어 정의 중 소비자에게 불리할 수 있는 것).

## 4. 가입할 때 주의할 점
계약 전 알릴 의무(고지의무) 내용과 위반 시 결과.

## 5. 보험료·해지·환급
납입면제 조건, 갱신 시 보험료 변동, 중도 해지 시 손해, 만기환급 여부.

## 6. 내 권리
청약철회, 품질보증해지, 보험금 청구 방법, 청구 기한(소멸시효).

## 7. 가입 전 체크리스트
이 상품에서 특히 확인해야 할 것 5~8개."""

PERSONAL_TASK = """아래 가입 예정자 정보에 맞춰, 이 사람이 위 약관에서 꼭 알아야 할 내용만 골라 요약해 줘.
일반적인 설명보다 "이 사람에게 해당하는지"를 중심으로 쓰고, 해당 없는 내용은 과감히 뺀다.

### 가입 예정자 정보
{profile}

### 출력 순서
## 1. 핵심 요약
이 사람에게 이 상품이 어떤 의미인지 3~5줄. (목적에 맞는지, 가장 큰 주의점)

## 2. 내 건강 이력과 관련된 부분 ⚠️
입력한 지병·수술·치료 이력이 (a) 계약 전 알릴 의무(고지) 대상인지 — 약관의 질문 기간(예: 최근 3개월/1년/5년)과 비교해서, (b) 면책·감액·부담보·보장 제외로 이어질 수 있는지, (c) 숨기면 어떤 결과가 생기는지. 이력이 없으면 짧게.

## 3. 목적에 맞는 보장 / 부족한 보장
가입 목적 기준으로 이 상품이 잘 보장하는 것과 보장하지 않는 것을 표로.

## 4. 나이·성별·생활습관 관련
가입나이 범위, 갱신 시 보험료 인상 가능성, 성별·직업·운전·흡연 등과 관련된 약관 조항.

## 5. 돈 문제
보험료 납입 기간, 납입면제, 해지 시 손해, (투자·저축 목적이면) 원금 손실 가능성·최저보증·수수료.

## 6. 보험사·설계사에게 꼭 물어볼 질문
이 사람 상황에서 약관만으로는 알 수 없어 직접 확인해야 할 질문 5~8개."""

QUESTION_TASK = """위 약관을 근거로 아래 질문에 답해 줘. 짧고 명확하게, 근거 페이지를 붙여서.
{profile_part}
### 질문
{question}"""


# =====================================================================
# 가입자 정보 → 문장
# =====================================================================
def profile_to_text(p):
    if not p:
        return ""
    lines = []

    def add(label, value):
        if value not in (None, "", [], False):
            if isinstance(value, list):
                value = ", ".join(value)
            lines.append(f"- {label}: {value}")

    add("나이", f"만 {p.get('age')}세" if p.get("age") else "")
    add("성별", p.get("gender"))
    add("가입 목적", p.get("purposes"))
    add("지병", p.get("diseases"))
    add("지병(기타)", p.get("disease_other"))
    if p.get("surgery") == "있음":
        add("수술 이력", p.get("surgery_detail") or "있음(상세 미입력)")
    elif p.get("surgery") == "없음":
        add("수술 이력", "없음")
    for key, label in [
        ("q_3months", "최근 3개월 이내 진찰·검사로 질병 확정/의심, 치료, 입원, 수술, 투약"),
        ("q_1year", "최근 1년 이내 추가검사(재검사) 소견"),
        ("q_5years", "최근 5년 이내 입원·수술·7일 이상 치료·30일 이상 투약"),
        ("q_cancer", "최근 5년 이내 암·심장·뇌혈관 등 중대질병 진단/치료"),
        ("pregnant", "현재 임신 여부"),
    ]:
        add(label, p.get(key))
    add("위 질문 관련 상세", p.get("health_detail"))
    add("흡연", p.get("smoking"))
    add("음주", p.get("drinking"))
    add("직업", p.get("job"))
    add("위험한 직업·취미(오토바이, 등산, 잠수 등)", p.get("risky"))
    add("운전", p.get("driving"))
    add("월 보험료 예산", p.get("budget"))
    add("이미 가입한 보험", p.get("existing"))
    add("기타 궁금한 점", p.get("note"))
    return "\n".join(lines) if lines else "- (입력 없음)"


def profile_keywords(p):
    words = []
    if not p:
        return words
    for purpose in p.get("purposes") or []:
        words += PURPOSE_KEYWORDS.get(purpose, [purpose])
    for d in p.get("diseases") or []:
        words += [w for w in re.split(r"[\s/()·,]+", d) if len(w) >= 2]
    for key in ("disease_other", "surgery_detail", "health_detail", "job", "risky"):
        words += [w for w in re.split(r"[\s/()·,]+", p.get(key) or "") if len(w) >= 2]
    if p.get("surgery") == "있음":
        words.append("수술")
    if p.get("smoking") in ("흡연", "과거 흡연"):
        words.append("흡연")
    if p.get("pregnant") == "예":
        words += ["임신", "출산", "태아"]
    return list(dict.fromkeys(words))  # 중복 제거, 순서 유지


# =====================================================================
# 약관 페이지 고르기 (너무 길 때만)
# =====================================================================
def select_pages(doc, profile):
    """토큰 한도 안에 들어가도록 중요한 페이지를 고릅니다. → (pages, 생략 여부)"""
    pages = doc["pages"]
    tok_per_char = doc["tokens"] / max(1, doc["chars"])
    if doc["tokens"] <= MAX_DOC_TOKENS:
        return pages, False

    extra = profile_keywords(profile)

    def score(p):
        t = p["text"].replace(" ", "")
        s = sum(w * t.count(k.replace(" ", "")) for k, w in CORE_KEYWORDS.items())
        s += sum(3 * t.count(k) for k in extra)
        return s

    chosen, used = set(), 0
    for p in pages[:ALWAYS_FIRST_PAGES]:
        chosen.add(p["page"])
        used += len(p["text"]) * tok_per_char
    for p in sorted(pages[ALWAYS_FIRST_PAGES:], key=score, reverse=True):
        cost = len(p["text"]) * tok_per_char
        if used + cost > MAX_DOC_TOKENS:
            continue
        if score(p) == 0:
            break
        chosen.add(p["page"])
        used += cost
    return [p for p in pages if p["page"] in chosen], True


def pages_to_text(pages, total_pages, omitted):
    head = f"[보험 약관 원문 · 전체 {total_pages}쪽"
    if omitted:
        head += f" 중 관련성이 높은 {len(pages)}쪽만 발췌. 빠진 쪽의 내용은 '발췌본에서 확인되지 않음'으로 표시할 것"
    head += "]\n\n"
    return head + "\n\n".join(f"=== p.{p['page']} ===\n{p['text']}" for p in pages)


# =====================================================================
# Claude 호출
# =====================================================================
_client = None


def client():
    global _client
    if _client is None:
        _client = anthropic.AsyncAnthropic()
    return _client


async def count_tokens(text):
    """약관 텍스트의 토큰 수. 실패하면 글자 수로 어림잡습니다."""
    try:
        r = await client().messages.count_tokens(
            model=MODEL, messages=[{"role": "user", "content": text}])
        return r.input_tokens
    except Exception:
        return int(len(text) * 1.1)


def _doc_block(doc, profile):
    """약관 본문 블록. 같은 약관이면 매번 같은 내용이라 프롬프트 캐시가 적용됩니다."""
    if doc["scanned"]:
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf",
                       "data": base64.standard_b64encode(doc["pdf"]).decode()},
            "title": doc["filename"],
            "cache_control": {"type": "ephemeral"},
        }, False
    pages, omitted = select_pages(doc, profile)
    return {
        "type": "text",
        "text": pages_to_text(pages, len(doc["pages"]), omitted),
        "cache_control": {"type": "ephemeral"},
    }, omitted


def build_task(mode, profile, question=""):
    if mode == "personal":
        return PERSONAL_TASK.format(profile=profile_to_text(profile))
    if mode == "question":
        part = ""
        if profile:
            part = "\n### 질문자 정보 (참고)\n" + profile_to_text(profile) + "\n"
        return QUESTION_TASK.format(profile_part=part, question=question.strip())
    return GENERAL_TASK


async def stream_summary(doc, mode, profile=None, question=""):
    """요약 결과를 조각(text) 단위로 흘려보냅니다."""
    block, omitted = _doc_block(doc, profile)
    if omitted:
        yield ("> ℹ️ 약관이 매우 길어서 관련성이 높은 페이지만 골라 분석했습니다. "
               "빠진 부분이 궁금하면 아래 '추가 질문'으로 물어보세요.\n\n")

    async with client().beta.messages.stream(
        model=MODEL,
        max_tokens=MAX_OUTPUT_TOKENS,
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",                 # 안전 필터에 걸리면 서버가 다른 모델로 다시 시도
        thinking={"type": "adaptive"},
        output_config={"effort": EFFORT},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": [
            block,
            {"type": "text", "text": build_task(mode, profile, question)},
        ]}],
    ) as stream:
        async for text in stream.text_stream:
            yield text
        final = await stream.get_final_message()

    if final.stop_reason == "max_tokens":
        yield "\n\n> ⚠️ 답변이 너무 길어 중간에 끊겼습니다. 추가 질문으로 나머지를 물어보세요."
    elif final.stop_reason == "refusal":
        yield "\n\n> ⚠️ 이 요청은 처리할 수 없었습니다. 입력 내용을 바꿔서 다시 시도해 주세요."
