# -*- coding: utf-8 -*-
"""
유형별 통계테이블 일괄 반영 프로그램 (기준 유형 1개 → 나머지 유형)
----------------------------------------------------------------
간편심사유형처럼 약관은 같고 코드(주보종·특약)만 다른 유형이 여러 개 있을 때,
기준 유형 1개만 직접 검토·수정하면 나머지 유형의 값을 똑같이 맞춥니다.

칼럼 (A열부터 순서대로, 1행은 칼럼명 — 이름은 상관없고 순서만 맞으면 됨)
    0 상품명(간단)  1 주보종 코드  2 특약 코드  3 급부 코드  4 산출기준  5 지급율
    6 산출주기      7 적립금 지급여부  8 면책기간  9 기간1  10 기간2

처리
    - 0번(상품명) + 1·2번(대상 유형의 코드)이 같은 줄들을 한 블록으로 보고, 기준 유형의 같은 상품 줄과 비교
      (줄끼리는 위에서부터 순서대로 짝지음 — 같은 줄은 그대로, 다른 줄은 '값 변경'으로 봄)
        · 값 변경: 3~10번을 기준 유형 값으로 고침 (1·2번은 대상 유형 값 그대로) → 노란색
        · 기준에 없는 줄(예: 자동입력 3줄 → 검토 후 2줄): 대상에서도 뺌 → [변경내역]에 '삭제'
        · 기준에만 있는 줄(검토하며 새로 추가한 줄): 대상에 넣지 않음 → [확인필요]에 넣을 값과 함께 표시
    - 바뀔 것이 없는 줄도 모두 그대로 결과 시트에 들어감 (수정된 줄만 색 표시)
    - 기준에는 있는데 대상에 없는 상품 → 패스
    - 대상에는 있는데 기준에 없는 상품 → 그대로 두고 [확인필요]
    - 기준 유형 안에서 같은 상품명이 1·2번 조합 여러 개에 붙어 있고 내용이 서로 다르면 반영하지 않고 [확인필요]

입력
    - 엑셀 1개: 첫 번째 시트 = 기준 유형, 그 뒤 시트들 = 대상 유형
    - 또는 기준 엑셀 1개 + 대상 엑셀 여러 개 (각 파일의 첫 번째 시트)

결과 (원본은 바꾸지 않음): <기준파일명>_유형반영_날짜시간.xlsx 한 파일
    [기준 유형 시트] [대상 유형 시트들 — 전체 줄, 수정된 줄은 노란색(바뀐 칸은 진하게)]
    [요약] [확인필요] [변경내역]

- 문서보안(DRM)이 걸린 엑셀은 PC의 Excel로 열어 값만 읽습니다 (stat_builder.py 와 같은 방식).
- Spyder에서 F5로 실행하면 창이 뜹니다.
"""
import os
import queue
import zipfile
import datetime
import threading
import difflib
from collections import OrderedDict

import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.exceptions import InvalidFileException

# =====================================================================
# 설정
# =====================================================================

HEADER_ROWS = 1                 # 위에서 몇 행이 칼럼명인지 (그 아래부터 데이터)
FIRST_COL = 1                   # 0번(상품명) 칼럼 위치 (1 = A열)
COL_NAME, COL_MAIN, COL_SP = 0, 1, 2
COPY_COLS = list(range(3, 11))  # 기준 유형 값으로 맞출 칼럼 (3~10번)
N_COLS = 11                     # 0~10번
ROW_FILL = "FFF2CC"             # 수정된 줄 색
CELL_FILL = "FFD966"            # 그 줄에서 실제로 바뀐 칸 색

# =====================================================================
# 엑셀 읽기 도우미 (stat_builder.py 와 같음)
# =====================================================================


def norm(v):
    """비교용 값 정리: 앞뒤 공백 제거, 1.0 → 1"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _com_value(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def read_with_excel(path):
    """PC의 Excel로 열어 시트별 값만 읽어 openpyxl 통합문서로 만듦 (문서보안/구버전 파일용)"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError("일반 엑셀(.xlsx) 형식이 아니고(문서보안/구버전 .xls 의심), win32com 이 없어 Excel로도 읽지 못했습니다.")
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        xb = app.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=True)
        try:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            for xs in xb.Worksheets:
                ws = wb.create_sheet(title=str(xs.Name)[:31])
                used = xs.UsedRange
                vals = used.Value2
                if vals is None:
                    continue
                if not isinstance(vals, tuple):
                    vals = ((vals,),)
                r0, c0 = used.Row, used.Column
                for i, row in enumerate(vals):
                    for j, v in enumerate(row):
                        if v is not None and v != "":
                            ws.cell(r0 + i, c0 + j).value = _com_value(v)
        finally:
            xb.Close(False)
        return wb
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def load_input(path, log=print):
    try:
        return openpyxl.load_workbook(path)
    except (zipfile.BadZipFile, InvalidFileException):
        log(f"  (참고) {os.path.basename(path)}: 일반 엑셀 형식이 아니어서 PC의 Excel로 열어 값만 읽습니다.")
    try:
        return read_with_excel(path)
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"Excel로 열기 실패: {e}")


def read_rows(ws):
    """데이터 행 읽기 → [{"r": 행번호, "v": [그 행의 모든 칸 값]}] (0번 상품명이 빈 행은 '빈 행'으로 표시)"""
    width = max(ws.max_column, FIRST_COL + N_COLS - 1)
    rows = []
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, width + 1)]
        if all(norm(x) == "" for x in vals):
            continue
        rows.append({"r": r, "v": vals})
    return rows, width


def col(row_vals, i):
    """0~10번 칼럼 값"""
    return row_vals[FIRST_COL - 1 + i]


def sig(row_vals):
    """3~10번 값 묶음 (비교용)"""
    return tuple(norm(col(row_vals, i)) for i in COPY_COLS)


# =====================================================================
# 핵심 처리
# =====================================================================


def reference_sets(ref_rows, log):
    """기준 유형: 상품명 → 그 상품의 3~10번 줄 목록 (위에서부터 순서대로). 판단 못 하는 상품은 따로"""
    blocks = OrderedDict()                  # 상품명 → {(주보종, 특약): [행 값...]}
    for row in ref_rows:
        name = norm(col(row["v"], COL_NAME))
        if not name:
            continue
        pair = (norm(col(row["v"], COL_MAIN)), norm(col(row["v"], COL_SP)))
        blocks.setdefault(name, OrderedDict()).setdefault(pair, []).append(row["v"])
    ok, bad = OrderedDict(), OrderedDict()
    for name, pairs in blocks.items():
        lists = list(pairs.values())
        first = [sig(v) for v in lists[0]]
        if all([sig(v) for v in lst] == first for lst in lists[1:]):
            ok[name] = [[col(v, i) for i in COPY_COLS] for v in lists[0]]
        else:
            bad[name] = list(pairs.keys())
    log(f"기준 유형: 상품 {len(blocks)}개, 줄 {len(ref_rows)}개"
        + (f" (1·2번 조합끼리 내용이 달라 반영 보류: {len(bad)}개)" if bad else ""))
    return ok, bad


def sync_sheet(ws, label, ref_ok, ref_bad, heads, changes, checks, log):
    """대상 유형 시트 하나를 기준 유형에 맞춰 다시 씀 (전체 줄 유지, 수정된 줄만 색) → 요약 dict"""
    rows, width = read_rows(ws)
    groups = OrderedDict()                  # (상품, 주보종, 특약) → 원래 행들
    for row in rows:
        v = row["v"]
        groups.setdefault((norm(col(v, COL_NAME)), norm(col(v, COL_MAIN)), norm(col(v, COL_SP))), []).append(row)
    out = []                                # [(행 값, 바뀐 칼럼 번호 집합 또는 None)]
    done, flagged = set(), set()
    st = OrderedDict([("원래 줄", len(rows)), ("결과 줄", 0), ("그대로", 0), ("값 변경 줄", 0), ("삭제 줄", 0),
                      ("추가 안 한 줄(확인)", 0), ("기준에 없는 상품", 0), ("기준에만 있어 패스한 상품", 0)])

    def check(name, main, sp, why, vals=None):
        d = OrderedDict([("유형", label), ("상품명", name or "(비어 있음)"), ("주보종 코드", main), ("특약 코드", sp),
                         ("사유", why)])
        for i in range(N_COLS):
            d[heads[i]] = "" if vals is None else vals[i]
        checks.append(d)

    for row in rows:
        v = row["v"]
        name, main, sp = norm(col(v, COL_NAME)), norm(col(v, COL_MAIN)), norm(col(v, COL_SP))
        if name not in ref_ok:
            out.append((v, None))
            st["그대로"] += 1
            if name not in flagged:
                flagged.add(name)
                st["기준에 없는 상품"] += 1
                check(name, main, sp, "기준 유형 안에서 1·2번 조합별 내용이 서로 다름 → 그대로 둠" if name in ref_bad
                      else "기준 유형에 같은 상품명이 없음 → 그대로 둠")
            continue
        key = (name, main, sp)
        if key in done:
            continue
        done.add(key)
        old = [x["v"] for x in groups[key]]
        ref = ref_ok[name]
        sm = difflib.SequenceMatcher(None, [tuple(norm(x) for x in r) for r in ref], [sig(x) for x in old], autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                for k in range(j2 - j1):
                    out.append((old[j1 + k], None))
                    st["그대로"] += 1
                continue
            n = min(i2 - i1, j2 - j1)
            for k in range(n):                      # 값 변경: 3~10번만 기준 값으로
                before = old[j1 + k]
                new = list(before)
                diff = []
                for c, val in zip(COPY_COLS, ref[i1 + k]):
                    if norm(col(before, c)) != norm(val):
                        diff.append(f"{heads[c]}: {norm(col(before, c)) or '(빈칸)'} → {norm(val) or '(빈칸)'}")
                    new[FIRST_COL - 1 + c] = val
                out.append((new, {c for c, val in zip(COPY_COLS, ref[i1 + k]) if norm(col(before, c)) != norm(val)}))
                st["값 변경 줄"] += 1
                changes.append(OrderedDict([("유형", label), ("상품명", name), ("주보종 코드", main), ("특약 코드", sp),
                                            ("결과 행번호", HEADER_ROWS + len(out)), ("구분", "값 변경"),
                                            ("바뀐 내용", " / ".join(diff))]))
            for k in range(n, i2 - i1):             # 기준에만 있는 줄 → 넣지 않고 확인
                sugg = list(old[0])
                for c, val in zip(COPY_COLS, ref[i1 + k]):
                    sugg[FIRST_COL - 1 + c] = val
                st["추가 안 한 줄(확인)"] += 1
                check(name, main, sp, "기준 유형에만 있는 줄 → 추가하지 않음 (넣을 값은 오른쪽 칸)",
                      [col(sugg, i) for i in range(N_COLS)])
            for k in range(n, j2 - j1):             # 기준에 없는 줄 → 뺌
                gone = old[j1 + k]
                st["삭제 줄"] += 1
                changes.append(OrderedDict([("유형", label), ("상품명", name), ("주보종 코드", main), ("특약 코드", sp),
                                            ("결과 행번호", ""), ("구분", "삭제(기준에 없는 줄)"),
                                            ("바뀐 내용", " / ".join(f"{heads[c]}: {norm(col(gone, c))}" for c in COPY_COLS))]))
    names_here = {norm(col(r["v"], COL_NAME)) for r in rows}
    st["기준에만 있어 패스한 상품"] = sum(1 for n in ref_ok if n not in names_here)

    # ---- 다시 쓰기 (칼럼명 행은 그대로, 그 아래 전체 줄을 새로 채움)
    row_fill, cell_fill, blank = PatternFill("solid", fgColor=ROW_FILL), PatternFill("solid", fgColor=CELL_FILL), PatternFill()
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        for c in range(1, width + 1):
            ws.cell(r, c).value = None
            ws.cell(r, c).fill = blank
    for i, (vals, changed) in enumerate(out):
        r = HEADER_ROWS + 1 + i
        for c, val in enumerate(vals, start=1):
            cell = ws.cell(r, c)
            cell.value = val
            if changed is not None:
                cell.fill = cell_fill if (c - FIRST_COL) in changed else row_fill
    st["결과 줄"] = len(out)
    log(f"  [{label}] 줄 {st['원래 줄']} → {st['결과 줄']} · 그대로 {st['그대로']} · 값 변경 {st['값 변경 줄']} · "
        f"삭제 {st['삭제 줄']} · 추가 안 함(확인) {st['추가 안 한 줄(확인)']} · 기준에 없는 상품 {st['기준에 없는 상품']}")
    return st


def _unique(path):
    base, ext = os.path.splitext(path)
    n, out = 2, path
    while os.path.exists(out):
        out = f"{base}_{n}{ext}"
        n += 1
    return out


def _copy_values(src, dst):
    for row in src.iter_rows():
        for cell in row:
            if cell.value is not None:
                dst.cell(cell.row, cell.column).value = cell.value


def _report(wb, title, data, heads):
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    heads = list(data[0].keys()) if data else heads
    ws.append(heads)
    for d in data:
        ws.append([d.get(h, "") for h in heads])
    for c in range(1, len(heads) + 1):
        ws.cell(1, c).font = Font(bold=True)
        width = max(len(str(ws.cell(r, c).value or "")) for r in range(1, min(ws.max_row, 300) + 1))
        ws.column_dimensions[ws.cell(1, c).column_letter].width = min(max(10, width * 1.4), 70)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def run(ref_path, target_paths, out_dir, log=print):
    log(f"기준: {os.path.basename(ref_path)}")
    ref_wb = load_input(ref_path, log)
    if target_paths:                        # 파일 여러 개 → 한 통합문서에 시트로 모음
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ref_ws = wb.create_sheet(ref_wb.worksheets[0].title[:31])
        _copy_values(ref_wb.worksheets[0], ref_ws)
        targets = []
        for p in target_paths:
            src = load_input(p, log).worksheets[0]
            title = os.path.splitext(os.path.basename(p))[0][:31]
            while title in wb.sheetnames:
                title = (title[:28] + "_2")[:31]
            ws = wb.create_sheet(title)
            _copy_values(src, ws)
            targets.append(ws)
    else:                                   # 엑셀 1개 → 첫 시트 기준, 나머지 대상
        if len(ref_wb.worksheets) < 2:
            raise RuntimeError("기준 엑셀에 대상 유형 시트를 함께 넣거나, 대상 엑셀을 골라 주세요.")
        wb = ref_wb
        ref_ws, targets = wb.worksheets[0], wb.worksheets[1:]
    ref_rows, width = read_rows(ref_ws)
    if not ref_rows:
        raise RuntimeError("기준 유형 시트에 데이터가 없습니다.")
    heads = [norm(ref_ws.cell(HEADER_ROWS, FIRST_COL + i).value) or f"{i}번" for i in range(N_COLS)]
    heads = [h if heads.count(h) == 1 else f"{i}번 {h}" for i, h in enumerate(heads)]
    ref_ok, ref_bad = reference_sets(ref_rows, log)

    changes, checks, summary = [], [], []
    for ws in targets:
        st = sync_sheet(ws, ws.title, ref_ok, ref_bad, heads, changes, checks, log)
        summary.append(OrderedDict([("유형(시트)", ws.title), *st.items()]))
    for name, pairs in ref_bad.items():
        d = OrderedDict([("유형", ref_ws.title), ("상품명", name), ("주보종 코드", ", ".join(p[0] for p in pairs)),
                         ("특약 코드", ", ".join(p[1] for p in pairs)),
                         ("사유", "기준 유형 안에서 1·2번 조합별 3~10번 내용이 다름 → 반영 보류")])
        checks.append(d)

    _report(wb, "요약", summary, ["유형(시트)"])
    _report(wb, "확인필요", checks, ["유형", "상품명", "주보종 코드", "특약 코드", "사유"])
    _report(wb, "변경내역", changes, ["유형", "상품명", "주보종 코드", "특약 코드", "결과 행번호", "구분", "바뀐 내용"])
    stem = os.path.splitext(os.path.basename(ref_path))[0]
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = _unique(os.path.join(out_dir, f"{stem}_유형반영_{stamp}.xlsx"))
    wb.save(out)
    log(f"결과: {os.path.basename(out)} (값 변경 {sum(s['값 변경 줄'] for s in summary)}줄 · "
        f"삭제 {sum(s['삭제 줄'] for s in summary)}줄 · 확인필요 {len(checks)}건)")
    return out


# =====================================================================
# 화면
# =====================================================================

def open_file(p):
    try:
        os.startfile(p)
    except Exception:
        pass


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("유형별 통계테이블 일괄 반영 (기준 유형 → 나머지 유형)")
    root.geometry("800x460")
    root.lift()
    root.attributes("-topmost", True)
    root.after(800, lambda: root.attributes("-topmost", False))
    v_ref, v_out = tk.StringVar(), tk.StringVar()
    targets = []
    q = queue.Queue()

    def pick_ref():
        f = filedialog.askopenfilename(title="검토를 마친 기준 유형 엑셀", filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if f:
            v_ref.set(f)
            if not v_out.get():
                v_out.set(os.path.dirname(f))

    def pick_targets():
        fs = filedialog.askopenfilenames(title="반영할 대상 유형 엑셀 (여러 개 선택)", filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if fs:
            targets[:] = list(fs)
            v_tg.set(f"{len(targets)}개: " + ", ".join(os.path.basename(x) for x in targets))

    def pick_out():
        d = filedialog.askdirectory()
        if d:
            v_out.set(d)

    v_tg = tk.StringVar(value="(비워 두면 기준 엑셀의 두 번째 시트부터를 대상 유형으로 봄)")
    for r, (label, var, cmd) in enumerate([("기준 유형 엑셀", v_ref, pick_ref), ("대상 유형 엑셀", v_tg, pick_targets),
                                           ("결과 저장 위치", v_out, pick_out)]):
        tk.Label(root, text=label, width=14, anchor="w").grid(row=r, column=0, padx=8, pady=6, sticky="w")
        tk.Entry(root, textvariable=var, width=70).grid(row=r, column=1, padx=4, sticky="ew")
        tk.Button(root, text="찾아보기", command=cmd).grid(row=r, column=2, padx=8)

    logbox = scrolledtext.ScrolledText(root, height=14)
    logbox.grid(row=4, column=0, columnspan=3, padx=8, pady=8, sticky="nsew")
    root.grid_rowconfigure(4, weight=1)
    root.grid_columnconfigure(1, weight=1)

    def log(m):
        q.put(("log", m))

    def poll():
        try:
            while True:
                kind, m = q.get_nowait()
                if kind == "log":
                    logbox.insert("end", m + "\n")
                    logbox.see("end")
                elif kind == "done":
                    btn.config(state="normal")
                    if messagebox.askyesno("완료", "반영이 끝났습니다. 결과 엑셀을 열까요?"):
                        open_file(m)
                elif kind == "error":
                    btn.config(state="normal")
                    messagebox.showerror("오류", m)
        except queue.Empty:
            pass
        root.after(200, poll)

    def start():
        ref, out = v_ref.get().strip(), v_out.get().strip()
        if not os.path.isfile(ref):
            messagebox.showwarning("확인", "기준 유형 엑셀을 선택해 주세요.")
            return
        if not os.path.isdir(out):
            messagebox.showwarning("확인", "결과 저장 위치를 선택해 주세요.")
            return
        btn.config(state="disabled")
        logbox.delete("1.0", "end")
        log("반영 시작...")

        def work():
            try:
                q.put(("done", run(ref, list(targets), out, log)))
            except PermissionError:
                q.put(("error", "파일이 엑셀에서 열려 있으면 닫고 다시 실행해 주세요."))
            except Exception as e:
                q.put(("error", str(e)))

        threading.Thread(target=work, daemon=True).start()

    btn = tk.Button(root, text="실행", width=16, command=start)
    btn.grid(row=3, column=0, columnspan=3, pady=6)
    poll()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
