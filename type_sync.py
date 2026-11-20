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
    1) 0번(상품명)으로 기준 유형과 대상 유형을 짝지음 (띄어쓰기 차이는 무시)
    2) 같은 상품 안에서 특약(1·2번 조합)끼리 짝지음 — 코드는 유형마다 다르므로
       각 특약에 들어 있는 급부코드(3번) 구성이 가장 비슷한 것끼리 짝지음 (특약이 하나뿐이면 바로 짝)
    3) 짝지은 특약 안에서 줄을 위에서부터 비교 (같은 급부코드끼리 먼저 짝지음)
        · 값 변경: 3~10번을 기준 유형 값으로 고침 (1·2번은 대상 유형 값 그대로) → 노란색(바뀐 칸은 진하게)
        · 기준에만 있는 줄(검토하며 추가한 줄): 대상에도 추가 (1·2번은 그 특약 값) → 초록색
        · 기준에 없는 줄(예: 자동입력 3줄 → 검토 후 2줄): 대상에서도 뺌 → [변경내역]에 '삭제'
    - 바뀔 것이 없는 줄도 모두 그대로 결과 시트에 들어감 (수정·추가된 줄만 색 표시)
    - 대상 유형에 아예 없는 상품, 짝을 못 찾은 특약 → 추가하지 않고 [확인필요]
    - 대상 유형에만 있는 상품·특약 → 그대로 두고 [확인필요]

입력
    - 엑셀 1개: 첫 번째 시트 = 기준 유형, 그 뒤 시트들 = 대상 유형
    - 또는 기준 엑셀 1개 + 대상 엑셀 여러 개 (각 파일의 첫 번째 시트)

결과 (원본은 바꾸지 않음): <기준파일명>_유형반영_날짜시간.xlsx 한 파일
    [기준 유형 시트] [대상 유형 시트들 — 전체 줄, 수정된 줄은 노란색(바뀐 칸은 진하게), 추가된 줄은 초록색]
    [요약] [확인필요] [변경내역]

- 문서보안(DRM)이 걸린 엑셀은 PC의 Excel로 열어 값만 읽습니다 (stat_builder.py 와 같은 방식).
- Spyder에서 F5로 실행하면 창이 뜹니다.
"""
import os
import queue
import zipfile
import datetime
import threading
import re
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
ADD_FILL = "C6EFCE"             # 추가된 줄 색

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


def name_key(v):
    """상품명 비교용: 띄어쓰기·앞뒤 공백 무시"""
    return re.sub(r"\s+", "", norm(v))


def by_product(rows):
    """상품명 → {(주보종, 특약): [행...]} (시트 순서 그대로)"""
    out = OrderedDict()
    for row in rows:
        v = row["v"]
        pair = (norm(col(v, COL_MAIN)), norm(col(v, COL_SP)))
        out.setdefault(name_key(col(v, COL_NAME)), OrderedDict()).setdefault(pair, []).append(row)
    return out


def match_groups(ref_groups, tgt_groups):
    """같은 상품 안에서 기준 특약 ↔ 대상 특약 짝짓기 (급부코드 구성이 가장 비슷한 것끼리) → {대상 조합: 기준 조합}"""
    if len(ref_groups) == 1 and len(tgt_groups) == 1:
        return {next(iter(tgt_groups)): next(iter(ref_groups))}
    def codes(g):
        return [norm(col(x["v"], 3)) for x in g]
    scores = []
    for rp, rg in ref_groups.items():
        for tp, tg in tgt_groups.items():
            s1 = difflib.SequenceMatcher(None, codes(rg), codes(tg), autojunk=False).ratio()
            s2 = difflib.SequenceMatcher(None, [sig(x["v"]) for x in rg], [sig(x["v"]) for x in tg], autojunk=False).ratio()
            if s1 > 0:
                scores.append((s1 + s2 * 0.01, rp, tp))
    scores.sort(key=lambda t: -t[0])
    used_r, out = set(), {}
    for _, rp, tp in scores:
        if rp not in used_r and tp not in out:
            used_r.add(rp)
            out[tp] = rp
    return out


def align(ref_vals, tgt_vals):
    """특약 하나 안에서 줄 짝짓기 → [("same"|"change"|"add", 기준 줄 or None, 대상 줄 or None)] + 삭제된 대상 줄 목록"""
    sm = difflib.SequenceMatcher(None, [sig(x) for x in ref_vals], [sig(x) for x in tgt_vals], autojunk=False)
    res, gone = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            res += [("same", ref_vals[i1 + k], tgt_vals[j1 + k]) for k in range(i2 - i1)]
            continue
        rs, ts_ = list(range(i1, i2)), list(range(j1, j2))
        pair = {}
        for i in rs:                                # 같은 급부코드끼리 먼저
            for j in ts_:
                if j not in pair.values() and norm(col(ref_vals[i], 3)) == norm(col(tgt_vals[j], 3)):
                    pair[i] = j
                    break
        left = [j for j in ts_ if j not in pair.values()]
        for i in rs:                                # 남은 것은 순서대로
            if i not in pair and left:
                pair[i] = left.pop(0)
        for i in rs:
            res.append(("change", ref_vals[i], tgt_vals[pair[i]]) if i in pair else ("add", ref_vals[i], None))
        gone += [tgt_vals[j] for j in ts_ if j not in pair.values()]
    return res, gone


def sync_sheet(ws, label, ref_prod, heads, changes, checks, log):
    """대상 유형 시트 하나를 기준 유형에 맞춰 다시 씀 (전체 줄 유지, 수정·추가된 줄만 색) → 요약 dict"""
    rows, width = read_rows(ws)
    tgt_prod = by_product(rows)
    out = []                                # [(행 값, 표시)] 표시: None | "add" | 바뀐 칼럼 번호 집합
    st = OrderedDict([("원래 줄", len(rows)), ("결과 줄", 0), ("그대로", 0), ("값 변경 줄", 0), ("추가 줄", 0),
                      ("삭제 줄", 0), ("기준에 없는 상품·특약(확인)", 0), ("이 유형에 없는 상품·특약(확인)", 0)])

    def check(name, main, sp, why):
        checks.append(OrderedDict([("유형", label), ("상품명", name or "(비어 있음)"), ("주보종 코드", main),
                                   ("특약 코드", sp), ("사유", why)]))

    def log_change(kind, name, main, sp, row_no, text):
        changes.append(OrderedDict([("유형", label), ("상품명", name), ("주보종 코드", main), ("특약 코드", sp),
                                    ("결과 행번호", row_no), ("구분", kind), ("바뀐 내용", text)]))

    matches = {}                            # 상품 → {대상 조합: 기준 조합}
    for nk, tg in tgt_prod.items():
        if nk in ref_prod:
            matches[nk] = match_groups(ref_prod[nk], tg)
            for rp in ref_prod[nk]:
                if rp not in matches[nk].values():
                    st["이 유형에 없는 상품·특약(확인)"] += 1
                    nm = norm(col(ref_prod[nk][rp][0]["v"], COL_NAME))
                    check(nm, "", "", f"기준 유형의 특약(주보종 {rp[0]}, 특약 {rp[1]}, {len(ref_prod[nk][rp])}줄)과 "
                                      f"짝이 되는 특약이 이 유형에 없음 → 추가하지 않음")
    for nk, rg in ref_prod.items():
        if nk not in tgt_prod:
            st["이 유형에 없는 상품·특약(확인)"] += 1
            nm = norm(col(next(iter(rg.values()))[0]["v"], COL_NAME))
            check(nm, "", "", f"기준 유형에는 있는데 이 유형에 없는 상품 (기준 {sum(len(g) for g in rg.values())}줄) → 추가하지 않음")

    done = set()
    for row in rows:
        v = row["v"]
        nk, pair = name_key(col(v, COL_NAME)), (norm(col(v, COL_MAIN)), norm(col(v, COL_SP)))
        name = norm(col(v, COL_NAME))
        rp = matches.get(nk, {}).get(pair)
        if rp is None:                      # 기준에 없는 상품·특약 → 그대로
            out.append((v, None))
            st["그대로"] += 1
            if (nk, pair) not in done:
                done.add((nk, pair))
                st["기준에 없는 상품·특약(확인)"] += 1
                check(name, pair[0], pair[1], "기준 유형에 같은 상품명이 없음 → 그대로 둠" if nk not in ref_prod
                      else "기준 유형에서 짝이 되는 특약을 찾지 못함 → 그대로 둠")
            continue
        if (nk, pair) in done:
            continue
        done.add((nk, pair))
        tgt_vals = [x["v"] for x in tgt_prod[nk][pair]]
        res, gone = align([x["v"] for x in ref_prod[nk][rp]], tgt_vals)
        for kind, rv, tv in res:
            if kind == "same":
                out.append((tv, None))
                st["그대로"] += 1
                continue
            base = list(tv if tv is not None else tgt_vals[0])     # 1·2번과 11번 이후 칸은 대상 값
            diff, cols = [], set()
            for c in COPY_COLS:
                if tv is not None and norm(col(tv, c)) != norm(col(rv, c)):
                    diff.append(f"{heads[c]}: {norm(col(tv, c)) or '(빈칸)'} → {norm(col(rv, c)) or '(빈칸)'}")
                    cols.add(c)
                base[FIRST_COL - 1 + c] = col(rv, c)
            if kind == "change":
                out.append((base, cols))
                st["값 변경 줄"] += 1
                log_change("값 변경", name, pair[0], pair[1], HEADER_ROWS + len(out), " / ".join(diff))
            else:
                out.append((base, "add"))
                st["추가 줄"] += 1
                log_change("추가", name, pair[0], pair[1], HEADER_ROWS + len(out),
                           " / ".join(f"{heads[c]}: {norm(col(base, c))}" for c in COPY_COLS))
        for g in gone:
            st["삭제 줄"] += 1
            log_change("삭제(기준에 없는 줄)", name, pair[0], pair[1], "",
                       " / ".join(f"{heads[c]}: {norm(col(g, c))}" for c in COPY_COLS))

    # ---- 다시 쓰기 (칼럼명 행은 그대로, 그 아래 전체 줄을 새로 채움)
    row_fill, cell_fill, blank = PatternFill("solid", fgColor=ROW_FILL), PatternFill("solid", fgColor=CELL_FILL), PatternFill()
    add_fill = PatternFill("solid", fgColor=ADD_FILL)
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        for c in range(1, width + 1):
            ws.cell(r, c).value = None
            ws.cell(r, c).fill = blank
    for i, (vals, mark) in enumerate(out):
        r = HEADER_ROWS + 1 + i
        for c, val in enumerate(vals, start=1):
            cell = ws.cell(r, c)
            cell.value = val
            if mark == "add":
                cell.fill = add_fill
            elif mark is not None:
                cell.fill = cell_fill if (c - FIRST_COL) in mark else row_fill
    st["결과 줄"] = len(out)
    log(f"  [{label}] 줄 {st['원래 줄']} → {st['결과 줄']} · 그대로 {st['그대로']} · 값 변경 {st['값 변경 줄']} · "
        f"추가 {st['추가 줄']} · 삭제 {st['삭제 줄']} · 확인필요 "
        f"{st['기준에 없는 상품·특약(확인)'] + st['이 유형에 없는 상품·특약(확인)']}")
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
    ref_prod = by_product(ref_rows)
    log(f"기준 유형({ref_ws.title}): 상품 {len(ref_prod)}개 · 특약 {sum(len(g) for g in ref_prod.values())}개 · 줄 {len(ref_rows)}개")

    changes, checks, summary = [], [], []
    for ws in targets:
        st = sync_sheet(ws, ws.title, ref_prod, heads, changes, checks, log)
        summary.append(OrderedDict([("유형(시트)", ws.title), *st.items()]))

    _report(wb, "요약", summary, ["유형(시트)"])
    _report(wb, "확인필요", checks, ["유형", "상품명", "주보종 코드", "특약 코드", "사유"])
    _report(wb, "변경내역", changes, ["유형", "상품명", "주보종 코드", "특약 코드", "결과 행번호", "구분", "바뀐 내용"])
    stem = os.path.splitext(os.path.basename(ref_path))[0]
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = _unique(os.path.join(out_dir, f"{stem}_유형반영_{stamp}.xlsx"))
    wb.save(out)
    log(f"결과: {os.path.basename(out)} (값 변경 {sum(s['값 변경 줄'] for s in summary)}줄 · "
        f"추가 {sum(s['추가 줄'] for s in summary)}줄 · 삭제 {sum(s['삭제 줄'] for s in summary)}줄 · 확인필요 {len(checks)}건)")
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
