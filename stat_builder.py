# -*- coding: utf-8 -*-
"""
통계테이블 자동 작성 프로그램 (1단계: 상품코드 변경 이어받기)
----------------------------------------------------------------
입력: 아래 5개 시트가 들어 있는 엑셀 파일 (양식 파일 그대로)
    1번 시트  통계테이블 작성 대상  ← 여기에 결과가 채워짐
    2번 시트  통계테이블 기존      (개정 전 테이블)
    3번 시트  주보종코드-개별상품코드 관계 테이블 (주보종명, 개별상품명, SP_SPCD, SP_SPCD1)
              ※ 상품명 두 칼럼은 SP_SPCD 바로 왼쪽에 주보종명, 개별상품명 순서로 둡니다
    4번 시트  변경전,후 상품코드 (변경전, 변경후)
    5번 시트  급부코드 체계 (급부코드, 급부명)

처리:
    3번 시트의 (주보종코드, 개별상품코드) 마다
      → 4번 시트에서 변경전 개별상품코드를 찾고
      → 2번 시트에서 그 변경전 코드의 행들을 가져와
      → MAIN_PROD_CODE / SP_PROD_CODE 만 새 코드로 바꿔 1번 시트에 채움
    (급부코드·산출기준·지급율·지급주기·적립금·면책·경과년은 기존 값 그대로)

결과: 원본은 건드리지 않고 '통계테이블_자동작성_날짜시간.xlsx' 로 새로 저장
    [1번 시트]     채워진 통계테이블
    [검토_상품코드] 개별상품코드별 처리 상태 (기존 행 없음, 신규 등은 색 표시)
    [검토_요약]     상태별 건수

Spyder 에서 F5 로 실행하면 창이 뜹니다. (추가 설치 필요 없음)
"""
import os
import queue
import datetime
import threading
from collections import OrderedDict

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# =====================================================================
# 설정
# =====================================================================

# 시트 위치 (0 = 첫 번째 시트). 시트 이름이 바뀌어도 순서로 찾습니다.
SHEET_TARGET = 0
SHEET_OLD = 1
SHEET_REL = 2
SHEET_MAP = 3
SHEET_CODE = 4

# 통계테이블 칼럼 (영문 칼럼명이 적힌 행을 찾아서 그 아래부터 데이터로 봄)
TABLE_COLS = ["MAIN_PROD_CODE", "SP_PROD_CODE", "SCRP_RISK_PRM_CODE", "CALC_TYPE1", "PAY_RATE1",
              "PAY_TERM1", "END_RESRV", "MNGR_START_DAYS", "DUR_YR1", "DUR_YR2"]
REL_MAIN, REL_SP = "SP_SPCD", "SP_SPCD1"      # 관계 테이블 칼럼명
REL_MAIN_NAME, REL_SP_NAME = "주보종명", "개별상품명"   # 상품명 칼럼명 (검토 시트에도 이 이름으로 표시)
MAP_BEFORE, MAP_AFTER = "변경전", "변경후"      # 변경전/후 칼럼명
CODE_COL, CODE_NAME_COL = "급부코드", "급부명"

HEADER_SEARCH_ROWS = 10                       # 위에서 몇 행 안에서 칼럼명을 찾을지

# =====================================================================
# 엑셀 읽기 도우미
# =====================================================================


def norm(v):
    """코드값 정리: 앞뒤 공백 제거, 숫자로 저장된 코드도 문자로"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def find_header(ws, names, sheet_label):
    """칼럼명이 모두 들어 있는 행 찾기 → (행번호, {칼럼명: 열번호}, 그 행의 전체 칼럼)"""
    for r in range(1, min(ws.max_row, HEADER_SEARCH_ROWS) + 1):
        row = {norm(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if norm(ws.cell(r, c).value)}
        if all(n in row for n in names):
            return r, {n: row[n] for n in names}, row
    raise RuntimeError(f"[{sheet_label}] 시트에서 칼럼명 {names} 을(를) 찾지 못했습니다. "
                       f"(위에서 {HEADER_SEARCH_ROWS}행 안에 있어야 합니다)")


def read_rows(ws, hdr_row, cols):
    """칼럼명 행 아래 데이터 읽기. 코드 칼럼이 비어 있는 행은 건너뜀"""
    out = []
    first = list(cols.values())[0]
    for r in range(hdr_row + 1, ws.max_row + 1):
        if not norm(ws.cell(r, first).value):
            continue
        out.append({n: ws.cell(r, c).value for n, c in cols.items()})
    return out


def rel_name_cols(all_cols, code_cols):
    """관계 시트의 상품명 칼럼 → (주보종명 열, 개별상품명 열).
    1) 칼럼명이 '주보종명', '개별상품명' 이면 그 칼럼
    2) 아니면 SP_SPCD 바로 왼쪽 두 칼럼을 순서대로 (칼럼명이 둘 다 '상품명'이어도 위치로 구분)
    못 찾으면 None"""
    if REL_MAIN_NAME in all_cols and REL_SP_NAME in all_cols:
        return all_cols[REL_MAIN_NAME], all_cols[REL_SP_NAME]
    first = min(code_cols)
    a, b = first - 2, first - 1
    if a >= 1 and a not in code_cols and b not in code_cols:
        return a, b
    return None, None


def extra_cols(all_cols, used):
    """칼럼명 행에서 사용하지 않은 나머지 칼럼 (상품명/특약명 등) → 검토 시트에 함께 표시"""
    return {n: c for n, c in all_cols.items() if n not in used}


# =====================================================================
# 핵심 처리
# =====================================================================


def build(in_path, out_dir, log=print):
    wb = openpyxl.load_workbook(in_path)
    if len(wb.worksheets) < 4:
        raise RuntimeError("시트가 4개 이상 있어야 합니다 (작성대상, 기존, 관계, 변경전후).")
    ws_t, ws_o = wb.worksheets[SHEET_TARGET], wb.worksheets[SHEET_OLD]
    ws_r, ws_m = wb.worksheets[SHEET_REL], wb.worksheets[SHEET_MAP]
    ws_c = wb.worksheets[SHEET_CODE] if len(wb.worksheets) > SHEET_CODE else None

    t_hdr, t_cols, _ = find_header(ws_t, TABLE_COLS, ws_t.title)
    o_hdr, o_cols, _ = find_header(ws_o, TABLE_COLS, ws_o.title)
    r_hdr, r_cols, r_all = find_header(ws_r, [REL_MAIN, REL_SP], ws_r.title)
    m_hdr, m_cols, m_all = find_header(ws_m, [MAP_BEFORE, MAP_AFTER], ws_m.title)
    r_codes = [r_cols[REL_MAIN], r_cols[REL_SP]]
    r_name_main, r_name_sp = rel_name_cols(r_all, r_codes)
    r_used = set(r_codes) | {c for c in (r_name_main, r_name_sp) if c}
    r_extra = {n: c for n, c in r_all.items() if c not in r_used}     # 그 밖의 칼럼은 검토 시트 끝에 함께 표시
    if not r_name_main:
        log(f"  (참고) [{ws_r.title}] 시트에서 상품명 칼럼(SP_SPCD 왼쪽 두 칼럼: 주보종명, 개별상품명)을 "
            f"찾지 못해 상품명 없이 진행합니다.")

    old_rows = read_rows(ws_o, o_hdr, o_cols)
    rel = []
    for r in range(r_hdr + 1, ws_r.max_row + 1):
        main, sp = norm(ws_r.cell(r, r_cols[REL_MAIN]).value), norm(ws_r.cell(r, r_cols[REL_SP]).value)
        if main and sp:
            rel.append({"main": main, "sp": sp,
                        "main_name": norm(ws_r.cell(r, r_name_main).value) if r_name_main else "",
                        "sp_name": norm(ws_r.cell(r, r_name_sp).value) if r_name_sp else "",
                        "extra": {n: ws_r.cell(r, c).value for n, c in r_extra.items()}})
    maps = read_rows(ws_m, m_hdr, m_cols)
    log(f"기존 테이블 {len(old_rows)}행 / 관계 {len(rel)}건 / 변경전후 {len(maps)}건")

    code_names = {}
    if ws_c is not None:
        try:
            c_hdr, c_cols, _ = find_header(ws_c, [CODE_COL, CODE_NAME_COL], ws_c.title)
            for row in read_rows(ws_c, c_hdr, c_cols):
                code_names[norm(row[CODE_COL])] = norm(row[CODE_NAME_COL])
        except RuntimeError as e:
            log(f"  (참고) {e}")

    # 기존 테이블: 개별상품코드 → 주보종코드별 행 목록
    old_by_sp = OrderedDict()
    for row in old_rows:
        sp, main = norm(row["SP_PROD_CODE"]), norm(row["MAIN_PROD_CODE"])
        old_by_sp.setdefault(sp, OrderedDict()).setdefault(main, []).append(row)

    # 변경후 → 변경전 (여러 개일 수 있음)
    after_to_before = OrderedDict()
    before_seen = set()
    for m in maps:
        b, a = norm(m[MAP_BEFORE]), norm(m[MAP_AFTER])
        if a:
            after_to_before.setdefault(a, [])
            if b and b not in after_to_before[a]:
                after_to_before[a].append(b)
            before_seen.add(b)

    rel_sps = {x["sp"] for x in rel}
    out_rows, review = [], []
    for x in rel:
        main, sp = x["main"], x["sp"]
        befores = after_to_before.get(sp, [])
        note = []
        if befores:
            src_codes = befores
            how = "코드변경"
            if len(befores) > 1:
                note.append(f"변경전 코드가 {len(befores)}개 ({', '.join(befores)}) → 행을 합침")
        elif sp in old_by_sp:
            src_codes = [sp]
            how = "코드유지"            # 변경전후 표에 없지만 기존 테이블에 같은 코드가 있음
        else:
            src_codes = []
            how = "신규"

        rows, seen = [], set()
        for code in src_codes:
            by_main = old_by_sp.get(code)
            if not by_main:
                note.append(f"기존 테이블에 {code} 행 없음")
                continue
            if len(by_main) > 1:
                # 같은 개별상품코드가 여러 주보종에 있으면 첫 주보종 값 사용, 값이 다르면 표시
                vals = {tuple(tuple(norm(r[c]) for c in TABLE_COLS[2:]) for r in rs) for rs in by_main.values()}
                if len(vals) > 1:
                    note.append(f"{code}: 기존 주보종({', '.join(by_main)})마다 값이 달라 {list(by_main)[0]} 기준")
            for r in list(by_main.values())[0]:
                key = tuple(norm(r[c]) for c in TABLE_COLS[2:])
                if key in seen:
                    continue
                seen.add(key)
                new = dict(r)
                new["MAIN_PROD_CODE"], new["SP_PROD_CODE"] = main, sp
                rows.append(new)

        if not rows:
            status = "신규(기존 행 없음)" if how == "신규" else "기존 행 없음"
        elif note:
            status = "확인필요"
        else:
            status = "완료"
        out_rows += rows
        bens = OrderedDict()
        for r in rows:
            c = norm(r["SCRP_RISK_PRM_CODE"])
            bens[c] = code_names.get(c, "")
        review.append(OrderedDict([
            ("주보종코드", main), (REL_MAIN_NAME, x["main_name"]),
            ("개별상품코드(변경후)", sp), (REL_SP_NAME, x["sp_name"]),
            ("변경전코드", ", ".join(befores) if befores else ("(동일)" if how == "코드유지" else "")),
            ("처리", how), ("상태", status), ("작성행수", len(rows)),
            ("급부코드", ", ".join(bens)),
            ("급부명", ", ".join(v or "?" for v in bens.values())),
            ("비고", " / ".join(note)),
            *[(n, v) for n, v in x["extra"].items()],
        ]))

    # 변경전후 표에는 있는데 관계 테이블에 없는 변경후 코드
    for a, bs in after_to_before.items():
        if a not in rel_sps:
            review.append(OrderedDict([
                ("주보종코드", ""), (REL_MAIN_NAME, ""), ("개별상품코드(변경후)", a), (REL_SP_NAME, ""),
                ("변경전코드", ", ".join(bs)),
                ("처리", "-"), ("상태", "관계테이블에 없음"), ("작성행수", 0),
                ("급부코드", ""), ("급부명", ""),
                ("비고", "변경후 코드가 주보종-개별상품 관계 테이블에 없어 주보종코드를 알 수 없음"),
            ]))

    # ---- 1번 시트에 쓰기 (칼럼명 행 아래 기존 내용은 지우고 새로 채움)
    for r in range(t_hdr + 1, ws_t.max_row + 1):
        for c in t_cols.values():
            ws_t.cell(r, c).value = None
    for i, row in enumerate(out_rows):
        r = t_hdr + 1 + i
        for name, c in t_cols.items():
            v = row[name]
            cell = ws_t.cell(r, c)
            if name in ("MAIN_PROD_CODE", "SP_PROD_CODE", "SCRP_RISK_PRM_CODE", "CALC_TYPE1"):
                cell.value = norm(v)                  # 코드는 문자로 (01 의 0 유지)
                cell.number_format = "@"
            else:
                cell.value = v

    # ---- 검토 시트
    for name in ("검토_상품코드", "검토_요약"):
        if name in wb.sheetnames:
            del wb[name]
    ws_v = wb.create_sheet("검토_상품코드")
    heads = []
    for rv in review:
        for k in rv:
            if k not in heads:
                heads.append(k)
    ws_v.append(heads)
    colors = {"완료": None, "확인필요": "FFF2CC", "기존 행 없음": "F8CBAD",
              "신규(기존 행 없음)": "DDEBF7", "관계테이블에 없음": "F8CBAD"}
    for rv in review:
        ws_v.append([rv.get(h, "") for h in heads])
        color = colors.get(rv["상태"])
        if color:
            for c in range(1, len(heads) + 1):
                ws_v.cell(ws_v.max_row, c).fill = PatternFill("solid", fgColor=color)
    for c in range(1, len(heads) + 1):
        ws_v.cell(1, c).font = Font(bold=True)
        width = max(len(str(ws_v.cell(r, c).value or "")) for r in range(1, ws_v.max_row + 1))
        ws_v.column_dimensions[ws_v.cell(1, c).column_letter].width = min(max(10, width * 1.6), 60)
    ws_v.freeze_panes = "A2"
    ws_v.auto_filter.ref = ws_v.dimensions

    ws_s = wb.create_sheet("검토_요약")
    ws_s.append(["상태", "건수", "설명"])
    desc = {"완료": "기존 행을 그대로 이어받음",
            "확인필요": "이어받았지만 비고 확인 필요 (변경전 코드 여러 개 등)",
            "기존 행 없음": "변경전 코드는 있는데 기존 테이블에 행이 없음 → 약관 보고 직접 작성",
            "신규(기존 행 없음)": "변경전 코드도, 기존 행도 없음 → 신규 상품, 약관 보고 직접 작성",
            "관계테이블에 없음": "변경후 코드의 주보종코드를 알 수 없음 → 관계 테이블 확인"}
    counts = OrderedDict((k, 0) for k in desc)
    for rv in review:
        counts[rv["상태"]] = counts.get(rv["상태"], 0) + 1
    for k, n in counts.items():
        ws_s.append([k, n, desc.get(k, "")])
    ws_s.append([])
    ws_s.append(["작성된 행 수", len(out_rows), f"'{ws_t.title}' 시트"])
    for c, w in zip("ABC", (20, 10, 70)):
        ws_s.column_dimensions[c].width = w
    for c in range(1, 4):
        ws_s.cell(1, c).font = Font(bold=True)

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, f"통계테이블_자동작성_{stamp}.xlsx")
    n = 2
    while os.path.exists(out):
        out = os.path.join(out_dir, f"통계테이블_자동작성_{stamp}_{n}.xlsx")
        n += 1
    wb.save(out)

    for k, v in counts.items():
        if v:
            log(f"  {k}: {v}건")
    log(f"작성 행 수: {len(out_rows)}")
    log(f"결과 파일: {out}")
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
    root.title("통계테이블 자동 작성 (1단계: 상품코드 이어받기)")
    root.geometry("760x420")
    root.lift()
    root.attributes("-topmost", True)
    root.after(800, lambda: root.attributes("-topmost", False))
    v_in, v_out = tk.StringVar(), tk.StringVar()
    q = queue.Queue()

    def pick_in():
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if f:
            v_in.set(f)
            if not v_out.get():
                v_out.set(os.path.dirname(f))

    def pick_out():
        d = filedialog.askdirectory()
        if d:
            v_out.set(d)

    for r, (label, var, cmd) in enumerate([("입력 엑셀(양식)", v_in, pick_in), ("결과 저장 위치", v_out, pick_out)]):
        tk.Label(root, text=label, width=14, anchor="w").grid(row=r, column=0, padx=8, pady=6, sticky="w")
        tk.Entry(root, textvariable=var, width=66).grid(row=r, column=1, padx=4, sticky="ew")
        tk.Button(root, text="찾아보기", command=cmd).grid(row=r, column=2, padx=8)

    logbox = scrolledtext.ScrolledText(root, height=14)
    logbox.grid(row=3, column=0, columnspan=3, padx=8, pady=8, sticky="nsew")
    root.grid_rowconfigure(3, weight=1)
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
                    if messagebox.askyesno("완료", "작성이 끝났습니다. 결과 엑셀을 열까요?"):
                        open_file(m)
                elif kind == "error":
                    btn.config(state="normal")
                    messagebox.showerror("오류", m)
        except queue.Empty:
            pass
        root.after(200, poll)

    def start():
        src, out = v_in.get().strip(), v_out.get().strip()
        if not os.path.isfile(src):
            messagebox.showwarning("확인", "입력 엑셀 파일을 선택해 주세요.")
            return
        if not os.path.isdir(out):
            messagebox.showwarning("확인", "결과 저장 위치를 선택해 주세요.")
            return
        btn.config(state="disabled")
        logbox.delete("1.0", "end")
        log("작성 시작...")

        def work():
            try:
                q.put(("done", build(src, out, log)))
            except PermissionError:
                q.put(("error", "파일이 엑셀에서 열려 있으면 닫고 다시 실행해 주세요."))
            except Exception as e:
                q.put(("error", str(e)))

        threading.Thread(target=work, daemon=True).start()

    btn = tk.Button(root, text="실행", width=16, command=start)
    btn.grid(row=2, column=0, columnspan=3, pady=6)
    poll()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
