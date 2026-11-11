"""엑셀 기초 7강 — 시트 복사 · 보호 실측. 이 PC 엑셀(Microsoft 365)을 따로 띄워(DispatchEx) 시험 통합 문서를 %TEMP% 에 저장하고 잰다.
가상 '독서 모임 11월 회비' — '11월' 시트의 낸 금액 =IF(B3="O",설정!$B$1,0) 이 '설정' 시트의 월 회비(10,000)를 읽는다.
- 같은 파일 안 복사(Copy After): 새 시트 이름 · 식 · 합계
- 다른 파일로 복사(Copy, 새 통합 문서): 식이 원본 파일을 가리키는지(외부 연결) · LinkSources / 설정 시트와 같이 복사하면?
- 시트 복사 vs 범위 복사(붙여넣기): 열 너비
- 시트 보호: 기본 잠금 · 보호 뒤 입력 · 잠금 푼 칸만 입력 · 수식 숨기기 · 보호된 시트 삭제 · 이름 바꾸기 · 통합 문서 구조 보호 · 암호 · 복사본의 보호 상태
결과 measure.json (argv 'en' 이면 measure_en.json). 실행: python measure.py [en]  (pywin32 필요)"""
import datetime, json, os, sys, tempfile
import pythoncom
import win32com.client as win32

M = None  # 건너뛰는 앞 자리 선택 인자 — pythoncom.Missing·이름 붙인 인자(After=)는 새 통합 문서로 복사됐다(11/11 시험)

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}
S_MONTH, S_SET = ("Nov", "Settings") if EN else ("11월", "설정")
NAMES = (["Ava", "Ben", "Chloe", "Dan", "Emma", "Finn"] if EN else ["김하늘", "이도윤", "박서연", "최지우", "정민준", "한예린"])
PAID = ["O", "O", "X", "O", "X", "O"]
HEAD = ("Name", "Paid?", "Amount", "Note") if EN else ("이름", "납부 여부", "낸 금액", "비고")
WIDTHS = {"A": 12, "B": 10, "C": 12, "D": 22}


def err(f):
    try:
        v = f()
        return {"ok": True, "value": v}
    except Exception as e:  # noqa: BLE001
        msg = str(e)
        try:
            msg = e.excepinfo[2] or msg  # 엑셀이 띄우는 안내 문구
        except Exception:  # noqa: BLE001
            pass
        return {"ok": False, "error": (msg or "")[:200]}


def build(wb):
    ws_set = wb.Worksheets(1)
    ws_set.Name = S_SET
    ws_set.Range("A1").Value = "Monthly fee" if EN else "월 회비"
    ws_set.Range("B1").Value = 10000
    ws = wb.Worksheets.Add(ws_set)  # Before
    ws.Name = S_MONTH
    ws.Range("A1").Value = "Book club — November fees" if EN else "독서 모임 11월 회비"
    ws.Range("A2:D2").Value = (HEAD,)
    for i, (n, p) in enumerate(zip(NAMES, PAID), start=3):
        ws.Range(f"A{i}").Value = n
        ws.Range(f"B{i}").Value = p
        ws.Range(f"C{i}").Formula = f'=IF(B{i}="O",{S_SET}!$B$1,0)'
    ws.Range("A9").Value = "Total" if EN else "합계"
    ws.Range("C9").Formula = "=SUM(C3:C8)"
    ws.Range("C3:C9").NumberFormat = "#,##0"
    for c, w in WIDTHS.items():
        ws.Columns(c).ColumnWidth = w
    return ws


def new_book(before):
    """복사로 새로 생긴 통합 문서 — 복사 전 이름 목록에 없던 것"""
    for b in xl.Workbooks:
        if b.Name not in before:
            return b
    raise RuntimeError("new workbook not found")


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
tmp = tempfile.mkdtemp(prefix="xl7_")
try:
    xl.Visible = True  # 시트 두 개를 같이 고르기(Select)는 창이 보여야 된다
    xl.WindowState = -4143
    xl.Width, xl.Height = 760, 470
    xl.DisplayAlerts = False
    R["build"] = xl.Build
    wb = xl.Workbooks.Add()
    while wb.Worksheets.Count > 1:
        wb.Worksheets(wb.Worksheets.Count).Delete()
    ws = build(wb)
    src = os.path.join(tmp, "fees.xlsx" if EN else "회비장부.xlsx")
    wb.SaveAs(src, 51)
    R["source"] = {"file": os.path.basename(src), "c3": ws.Range("C3").Formula, "c9": ws.Range("C9").Text,
                   "sheets": [s.Name for s in wb.Worksheets]}

    # 1) 같은 파일 안 복사
    ws.Copy(M, ws)  # (Before 비움, After=ws)
    cp = wb.Worksheets(ws.Index + 1)
    R["copy_within"] = {"new_name": cp.Name, "c3": cp.Range("C3").Formula, "c9": cp.Range("C9").Text,
                        "widths": {c: cp.Columns(c).ColumnWidth for c in WIDTHS},
                        "sheets": [s.Name for s in wb.Worksheets]}
    cp.Name = "Dec" if EN else "12월"
    R["copy_within"]["renamed"] = cp.Name
    # 같은 이름으로 한 번 더 복사하면?
    ws.Copy(M, wb.Worksheets(wb.Worksheets.Count))
    R["copy_within"]["second_copy_name"] = wb.Worksheets(wb.Worksheets.Count).Name
    wb.Worksheets(wb.Worksheets.Count).Delete()

    # 2) 다른 파일(새 통합 문서)로 복사 — 시트 하나만
    before = [b.Name for b in xl.Workbooks]
    ws.Copy()
    nb = new_book(before)
    nws = nb.Worksheets(1)
    links = nb.LinkSources(1)
    R["copy_to_new_book"] = {"sheets": [s.Name for s in nb.Worksheets], "c3": nws.Range("C3").Formula,
                             "c9": nws.Range("C9").Text, "links": [os.path.basename(x) for x in (links or [])],
                             "widths": {c: nws.Columns(c).ColumnWidth for c in WIDTHS}}
    nb.Close(False)
    # 2b) 설정 시트와 같이(두 시트를 함께 골라) 복사
    wb.Activate()
    ws.Select()
    wb.Worksheets(S_SET).Select(False)  # Ctrl+클릭으로 두 시트 탭을 같이 고른 것과 같음
    before = [b.Name for b in xl.Workbooks]
    xl.ActiveWindow.SelectedSheets.Copy()
    nb = new_book(before)
    links = nb.LinkSources(1)
    R["copy_with_settings"] = {"sheets": [s.Name for s in nb.Worksheets], "c3": nb.Worksheets(S_MONTH).Range("C3").Formula,
                               "links": [os.path.basename(x) for x in (links or [])]}
    nb.Close(False)
    wb.Activate()
    ws.Select()

    # 3) 범위 복사 → 새 시트에 붙여넣기: 열 너비
    dst = wb.Worksheets.Add(M, wb.Worksheets(wb.Worksheets.Count))
    ws.Range("A1:D9").Copy(dst.Range("A1"))
    R["range_paste"] = {"widths": {c: dst.Columns(c).ColumnWidth for c in WIDTHS}, "c3": dst.Range("C3").Formula,
                        "standard_width": dst.StandardWidth}
    dst.Delete()

    # 4) 시트 보호
    P = {}
    P["default_locked"] = {"B3": ws.Range("B3").Locked, "C3": ws.Range("C3").Locked}
    ws.Protect()
    P["protect_no_pw"] = {"ProtectContents": ws.ProtectContents,
                          "write_B3": err(lambda: setattr(ws.Range("B3"), "Value", "X")),
                          "allow_format_cells": ws.Protection.AllowFormattingCells,
                          "allow_insert_rows": ws.Protection.AllowInsertingRows,
                          "allow_sort": ws.Protection.AllowSorting}
    P["unprotect_no_pw"] = err(lambda: ws.Unprotect())
    ws.Range("B3:B8").Locked = False
    ws.Range("C3:C9").FormulaHidden = True
    ws.Protect()
    P["unlocked_inputs"] = {"write_B5": err(lambda: setattr(ws.Range("B5"), "Value", "O")),
                            "c9_after": ws.Range("C9").Text,
                            "write_C3": err(lambda: setattr(ws.Range("C3"), "Value", 0)),
                            "write_A3": err(lambda: setattr(ws.Range("A3"), "Value", "?")),
                            "c3_formula_via_com": err(lambda: ws.Range("C3").Formula), "c3_text": ws.Range("C3").Text}
    ws.Range("B5").Value = "X"  # 되돌림(잠금 푼 칸이라 됨)
    # 보호된 시트를 복사하면?
    ws.Copy(M, wb.Worksheets(wb.Worksheets.Count))
    pc = wb.Worksheets(wb.Worksheets.Count)
    P["copy_of_protected"] = {"name": pc.Name, "ProtectContents": pc.ProtectContents,
                              "B3_locked": pc.Range("B3").Locked, "write_B3": err(lambda: setattr(pc.Range("B3"), "Value", "O")),
                              "write_C3": err(lambda: setattr(pc.Range("C3"), "Value", 0))}
    # 보호된 시트 이름 바꾸기 · 지우기 — 시트 보호만으로는?
    P["rename_protected"] = err(lambda: setattr(pc, "Name", "tmp"))
    P["delete_protected"] = err(lambda: pc.Delete())
    P["sheets_after_delete"] = [s.Name for s in wb.Worksheets]
    # 통합 문서 구조 보호
    wb.Protect("", True)  # Password 없음(빈 글자), Structure — None 은 E_INVALIDARG
    last = wb.Worksheets(wb.Worksheets.Count)
    P["structure"] = {"ProtectStructure": wb.ProtectStructure,
                      "delete": err(lambda: last.Delete()),
                      "rename": err(lambda: setattr(last, "Name", "tmp2")),
                      "add": err(lambda: wb.Worksheets.Add()),
                      "copy_within": err(lambda: ws.Copy(M, ws)),
                      "sheets": [s.Name for s in wb.Worksheets]}
    wb.Unprotect()
    # 암호
    ws.Unprotect()
    ws.Protect("1234")
    # 암호 없이 Unprotect() 를 부르면 엑셀이 암호 입력 창을 띄워 자동화가 멈춘다(11/11 두 번 확인) — 화면 단계(screens.py)에서 그 창을 찍는다
    P["password"] = {"unprotect_without": "prompts for password (dialog) — not called",
                     "unprotect_wrong": err(lambda: ws.Unprotect("0000")),
                     "unprotect_right": err(lambda: ws.Unprotect("1234")),
                     "ProtectContents_after": ws.ProtectContents}
    R["protect"] = P
    wb.Close(False)
finally:
    xl.DisplayAlerts = False
    xl.Quit()
out = os.path.join(HERE, "measure_en.json" if EN else "measure.json")
json.dump(R, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
