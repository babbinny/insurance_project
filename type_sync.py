# -*- coding: utf-8 -*-
"""
유형별 통계테이블 일괄 반영 프로그램 (기준 유형 1개 → 나머지 유형)
----------------------------------------------------------------
간편심사유형처럼 약관은 같고 코드(주보종·특약)만 다른 유형이 여러 개 있을 때,
기준 유형 1개만 직접 검토·수정하면 나머지 유형에 똑같이 반영합니다.

칼럼 (A열부터 순서대로, 1행은 칼럼명)
    0 상품명(간단)  1 주보종 코드  2 특약 코드  3 급부 코드  4 산출기준  5 지급율
    6 산출주기      7 적립금 지급여부  8 면책기간  9 기간1  10 기간2

처리
    - 0번(상품명)으로 기준 유형과 대상 유형을 짝지음
    - 대상 유형의 같은 상품 블록을 기준 유형의 그 상품 줄과 똑같이 다시 만듦
        · 3~10번: 기준 유형 값 그대로 (줄 추가 · 삭제 · 값 변경이 모두 반영됨)
        · 1·2번: 대상 유형의 그 블록 원래 값 그대로 (추가된 줄도 같은 1·2번을 물려받음)
        · 같은 상품이 대상 유형에서 1·2번 조합 여러 개에 붙어 있으면 조합마다 따로 만듦
    - 기준에는 있는데 대상에 없는 상품 → 패스 (새로 만들지 않음)
    - 대상에는 있는데 기준에 없는 상품 → 그대로 두고 '확인필요'
    - 기준 유형 안에서 같은 상품명이 1·2번 조합 여러 개에 붙어 있고 3~10번 내용이 서로 다르면
      어느 쪽이 정답인지 알 수 없으므로 반영하지 않고 '확인필요'

입력
    - 기준 엑셀 1개 + 대상 엑셀 여러 개 (각 파일의 첫 번째 시트)
    - 또는 기준 엑셀 1개만 고르면 그 파일의 '다른 시트들'을 대상 유형으로 봄

결과 (원본은 바꾸지 않음)
    - 대상 파일마다 '<파일명>_반영.xlsx' (기준 엑셀 1개만 고른 경우 '<파일명>_반영.xlsx' 하나)
        바뀌거나 추가된 줄은 노란색
    - 유형반영_요약_날짜시간.xlsx : [요약] [변경내역] [확인필요]

- 문서보안(DRM)이 걸린 엑셀은 PC의 Excel로 열어 값만 읽습니다 (stat_builder.py 와 같은 방식).
- Spyder에서 F5로 실행하면 창이 뜹니다.
"""
import os
import queue
import zipfile
import datetime
import threading
from collections import OrderedDict, Counter

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
CHANGED_FILL = "FFF2CC"         # 바뀌거나 추가된 줄 색

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
        first = Counter(sig(v) for v in lists[0])
        if all(Counter(sig(v) for v in lst) == first for lst in lists[1:]):
            ok[name] = [[col(v, i) for i in COPY_COLS] for v in lists[0]]
        else:
            bad[name] = list(pairs.keys())
    log(f"기준 유형: 상품 {len(blocks)}개, 줄 {len(ref_rows)}개"
        + (f" (1·2번 조합끼리 내용이 달라 반영 보류: {len(bad)}개)" if bad else ""))
    return ok, bad


def sync_sheet(ws, label, ref_ok, ref_bad, changes, checks, log):
    """대상 유형 시트 하나를 기준 유형에 맞춰 다시 씀 → 요약 dict"""
    rows, width = read_rows(ws)
    out = []                                # [(행 값, 바뀜 여부)]
    done = set()
    groups = OrderedDict()                  # (상품, 주보종, 특약) → 원래 행들
    for row in rows:
        v = row["v"]
        key = (norm(col(v, COL_NAME)), norm(col(v, COL_MAIN)), norm(col(v, COL_SP)))
        groups.setdefault(key, []).append(v)
    stat = Counter()
    flagged = set()
    for row in rows:
        v = row["v"]
        name = norm(col(v, COL_NAME))
        key = (name, norm(col(v, COL_MAIN)), norm(col(v, COL_SP)))
        if name not in ref_ok:
            out.append((v, False))
            if name and name not in flagged:
                flagged.add(name)
                why = ("기준 유형 안에서 1·2번 조합별 내용이 서로 다름 → 직접 확인" if name in ref_bad
                       else "기준 유형에 같은 상품명이 없음 → 그대로 둠")
                checks.append(OrderedDict([("유형", label), ("상품명", name or "(비어 있음)"), ("주보종 코드", key[1]),
                                           ("특약 코드", key[2]), ("행번호", row["r"]), ("사유", why)]))
                stat["확인필요"] += 1
            continue
        if key in done:
            continue
        done.add(key)
        old = groups[key]
        old_left = Counter(sig(x) for x in old)
        template = list(old[0])             # 1·2번과 11번 이후 칸은 이 블록의 첫 줄을 물려받음
        added = 0
        for ref_vals in ref_ok[name]:
            new = list(template)
            for i, val in zip(COPY_COLS, ref_vals):
                new[FIRST_COL - 1 + i] = val
            s = sig(new)
            changed = old_left[s] <= 0
            if changed:
                added += 1
                changes.append(OrderedDict([("유형", label), ("상품명", name), ("주보종 코드", key[1]), ("특약 코드", key[2]),
                                            ("구분", "추가/변경"), *[(f"{i}번", col(new, i)) for i in COPY_COLS]]))
            else:
                old_left[s] -= 1
            out.append((new, changed))
        removed = []                        # 기준에 없어서 빠지는 원래 줄
        for x in old:
            if old_left[sig(x)] > 0:
                old_left[sig(x)] -= 1
                removed.append(x)
        for x in removed:
            changes.append(OrderedDict([("유형", label), ("상품명", name), ("주보종 코드", key[1]), ("특약 코드", key[2]),
                                        ("구분", "삭제(기준에 없는 줄)"), *[(f"{i}번", col(x, i)) for i in COPY_COLS]]))
        stat["맞춘 블록"] += 1
        stat["추가/변경 줄"] += added
        stat["삭제 줄"] += len(removed)
        if added or removed:
            stat["바뀐 블록"] += 1
    names_here = {norm(col(r["v"], COL_NAME)) for r in rows}
    stat["기준에만 있어 패스한 상품"] = sum(1 for n in ref_ok if n not in names_here)

    # ---- 다시 쓰기 (칼럼명 행은 그대로, 그 아래만 새로 채움)
    yellow = PatternFill("solid", fgColor=CHANGED_FILL)
    for r in range(HEADER_ROWS + 1, ws.max_row + 1):
        for c in range(1, width + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill()
    for i, (vals, changed) in enumerate(out):
        r = HEADER_ROWS + 1 + i
        for c, val in enumerate(vals, start=1):
            cell = ws.cell(r, c)
            cell.value = val
            if changed:
                cell.fill = yellow
    stat["원래 줄"] = len(rows)
    stat["결과 줄"] = len(out)
    log(f"  [{label}] 줄 {len(rows)} → {len(out)} · 맞춘 블록 {stat['맞춘 블록']} (바뀐 블록 {stat['바뀐 블록']}) · "
        f"추가/변경 {stat['추가/변경 줄']}줄 · 삭제 {stat['삭제 줄']}줄 · 확인필요 {stat['확인필요']}건")
    return stat


def _unique(path):
    base, ext = os.path.splitext(path)
    n, out = 2, path
    while os.path.exists(out):
        out = f"{base}_{n}{ext}"
        n += 1
    return out


def run(ref_path, target_paths, out_dir, log=print):
    log(f"기준: {os.path.basename(ref_path)}")
    ref_wb = load_input(ref_path, log)
    ref_ws = ref_wb.worksheets[0]
    ref_rows, _ = read_rows(ref_ws)
    if not ref_rows:
        raise RuntimeError("기준 엑셀 첫 번째 시트에 데이터가 없습니다.")
    ref_ok, ref_bad = reference_sets(ref_rows, log)

    changes, checks, summary = [], [], []
    outputs = []
    if target_paths:
        jobs = []
        for p in target_paths:
            wb = load_input(p, log)
            stem = os.path.splitext(os.path.basename(p))[0]
            jobs.append((wb, [(wb.worksheets[0], stem)], _unique(os.path.join(out_dir, f"{stem}_반영.xlsx"))))
    else:
        if len(ref_wb.worksheets) < 2:
            raise RuntimeError("대상 엑셀을 고르거나, 기준 엑셀에 대상 유형 시트를 함께 넣어 주세요.")
        stem = os.path.splitext(os.path.basename(ref_path))[0]
        jobs = [(ref_wb, [(ws, ws.title) for ws in ref_wb.worksheets[1:]],
                 _unique(os.path.join(out_dir, f"{stem}_반영.xlsx")))]
    for wb, sheets, out in jobs:
        for ws, label in sheets:
            st = sync_sheet(ws, label, ref_ok, ref_bad, changes, checks, log)
            summary.append(OrderedDict([("유형", label), ("원래 줄", st["원래 줄"]), ("결과 줄", st["결과 줄"]),
                                        ("맞춘 상품 블록", st["맞춘 블록"]), ("바뀐 블록", st["바뀐 블록"]),
                                        ("추가/변경 줄", st["추가/변경 줄"]), ("삭제 줄", st["삭제 줄"]),
                                        ("기준에만 있어 패스한 상품", st["기준에만 있어 패스한 상품"]),
                                        ("확인필요", st["확인필요"])]))
        wb.save(out)
        outputs.append(out)
        log(f"  저장: {os.path.basename(out)}")

    for name, pairs in ref_bad.items():
        checks.append(OrderedDict([("유형", "기준"), ("상품명", name),
                                   ("주보종 코드", ", ".join(p[0] for p in pairs)), ("특약 코드", ", ".join(p[1] for p in pairs)),
                                   ("행번호", ""), ("사유", "기준 유형 안에서 1·2번 조합별 3~10번 내용이 다름 → 반영 보류")]))

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    rep = _unique(os.path.join(out_dir, f"유형반영_요약_{stamp}.xlsx"))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, data, empty in (("요약", summary, ["유형"]), ("변경내역", changes, ["유형", "상품명", "구분"]),
                               ("확인필요", checks, ["유형", "상품명", "사유"])):
        ws = wb.create_sheet(title)
        heads = list(data[0].keys()) if data else empty
        ws.append(heads)
        for d in data:
            ws.append([d.get(h, "") for h in heads])
        for c in range(1, len(heads) + 1):
            ws.cell(1, c).font = Font(bold=True)
            width = max(len(str(ws.cell(r, c).value or "")) for r in range(1, ws.max_row + 1))
            ws.column_dimensions[ws.cell(1, c).column_letter].width = min(max(10, width * 1.6), 60)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    wb.save(rep)
    log(f"요약: {os.path.basename(rep)} (변경 {len(changes)}건 · 확인필요 {len(checks)}건)")
    return rep, outputs


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

    v_tg = tk.StringVar(value="(고르지 않으면 기준 엑셀의 다른 시트들을 대상으로 봄)")
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
                    if messagebox.askyesno("완료", "반영이 끝났습니다. 요약 엑셀을 열까요?"):
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
                rep, _ = run(ref, list(targets), out, log)
                q.put(("done", rep))
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
