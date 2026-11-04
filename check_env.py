# -*- coding: utf-8 -*-
"""
내부망 실행 환경 점검 (파일이 왜 안 열리는지 확인용)
----------------------------------------------------------------
- 폴더를 고르면 그 안의 약관(.docx/.doc/.pdf)·엑셀(.xlsx/.xls)·프로그램(.py) 파일을 하나씩 열어 보고
  '정상 / 문서보안(DRM) 암호화 의심 / 구버전 형식' 등을 알려줍니다.
- Python·pandas·openpyxl·pdfplumber·win32com 설치 여부와, PC의 Word·Excel을 쓸 수 있는지도 확인합니다.
- 결과는 화면에 나오고, 같은 폴더에 '환경점검_날짜시간.txt' 로도 저장됩니다.
- Spyder에서 F5로 실행하면 폴더 선택 창이 뜹니다.
"""
import os
import sys
import zipfile
import datetime

CHECK_EXT = (".docx", ".docm", ".doc", ".xlsx", ".xlsm", ".xls", ".pdf", ".csv", ".py")


def file_kind(path):
    """파일 앞부분으로 실제 형식 판단 → (판정, 설명)"""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
        size = os.path.getsize(path)
    except Exception as e:
        return "열 수 없음", f"파일을 읽지 못함: {e}"
    ext = os.path.splitext(path)[1].lower()
    if size == 0:
        return "빈 파일", "크기가 0 — 다운로드나 복사가 덜 됐을 수 있음"
    if head.startswith(b"PK"):
        if ext in (".docx", ".docm", ".xlsx", ".xlsm"):
            try:
                with zipfile.ZipFile(path) as z:
                    z.namelist()
                return "정상", "일반 Office 형식 (파이썬이 바로 읽을 수 있음)"
            except zipfile.BadZipFile:
                return "손상", "zip 형식이 깨짐 — 파일을 다시 받거나 다시 저장해 보세요"
        return "정상", "zip 형식"
    if head.startswith(b"%PDF"):
        return "정상", "일반 PDF"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        if ext in (".doc", ".xls"):
            return "구버전 형식", "Office 97-2003 형식 — Word/Excel로 열어서 읽어야 함"
        return "암호화/구버전 의심", "Office 암호 걸린 파일이거나 문서보안(DRM) 또는 확장자만 바뀐 구버전 파일"
    if ext in (".py", ".csv"):
        try:
            with open(path, "rb") as f:
                data = f.read(4096)
            if b"\x00" in data:
                return "암호화 의심", "글자 파일인데 이상한 바이트가 있음 — 문서보안(DRM)이 걸렸을 수 있음"
            data.decode("utf-8")
            return "정상", "글자 파일"
        except UnicodeDecodeError:
            return "정상(인코딩 확인)", "UTF-8 이 아님 (cp949 등) — 대부분 문제없음"
    return "암호화 의심", f"알 수 없는 형식(앞 8바이트 {head.hex()}) — 문서보안(DRM) 암호화 가능성이 큼"


def check_packages():
    out = [f"Python {sys.version.split()[0]} ({sys.executable})"]
    for name in ("pandas", "openpyxl", "pdfplumber", "pdfminer", "win32com"):
        try:
            m = __import__(name)
            out.append(f"  {name}: 설치됨 {getattr(m, '__version__', '')}")
        except ImportError:
            hint = {"pdfplumber": " → PDF 약관을 읽으려면 필요", "win32com": " → 문서보안 파일을 Word/Excel로 읽으려면 필요"}
            out.append(f"  {name}: 없음{hint.get(name, '')}")
    return out


def check_office():
    out = []
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return ["  Word/Excel 확인 불가 (win32com 없음)"]
    pythoncom.CoInitialize()
    try:
        for prog, label in (("Word.Application", "Word"), ("Excel.Application", "Excel")):
            try:
                app = win32com.client.DispatchEx(prog)
                ver = app.Version
                app.Quit()
                out.append(f"  {label}: 사용 가능 (버전 {ver})")
            except Exception as e:
                out.append(f"  {label}: 사용 불가 ({e})")
    finally:
        pythoncom.CoUninitialize()
    return out


def run(folder, log=print):
    lines = [f"환경 점검 {datetime.datetime.now():%Y-%m-%d %H:%M}", f"폴더: {folder}", "", "[설치된 도구]"]
    lines += check_packages()
    lines += ["", "[Word/Excel]"] + check_office()
    lines += ["", "[파일]"]
    files = []
    for dp, _, fns in os.walk(folder):
        for fn in fns:
            if fn.lower().endswith(CHECK_EXT) and not fn.startswith("환경점검_"):
                files.append(os.path.join(dp, fn))
    if not files:
        lines.append("  점검할 파일이 없습니다.")
    counts = {}
    for p in sorted(files):
        if os.path.basename(p).startswith("~$"):
            kind, why = "임시파일", "Office가 열어 둔 동안 생기는 잠금 파일 — 고르지 마세요"
        else:
            kind, why = file_kind(p)
        counts[kind] = counts.get(kind, 0) + 1
        tag = "" if kind.startswith("정상") else "  ← 확인"
        name = os.path.relpath(p, folder)
        lines.append(f"  [{kind}] {name} — {why}{tag}")
    lines += ["", "[요약] " + ", ".join(f"{k} {v}개" for k, v in counts.items())]
    out = os.path.join(folder, f"환경점검_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt")
    try:
        with open(out, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        lines.append(f"결과 파일: {out}")
    except Exception as e:
        lines.append(f"(결과 파일 저장 실패: {e})")
    for line in lines:
        log(line)
    return out


if __name__ == "__main__":
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        d = filedialog.askdirectory(title="점검할 폴더 (약관·엑셀·py 파일이 있는 곳)")
        root.destroy()
    except Exception:
        d = input("점검할 폴더 경로: ").strip()
    if d:
        run(d)
