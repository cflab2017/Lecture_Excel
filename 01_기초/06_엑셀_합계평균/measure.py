"""엑셀 기초 6강 — 합계 · 평균 실측. 이 PC 엑셀(Microsoft 365)을 따로 띄워(DispatchEx) 수식을 넣고 칸에 보이는 값(.Text)을 읽는다.
가상 '문구점 일주일 판매 수량'(월~일) — 수요일은 쉬는 날(빈칸), 일요일은 0개, 목요일 18 은 글자로 저장된 숫자.
- SUM · AVERAGE · COUNT · COUNTA · 더하기(+)로 이은 식 · AVERAGEIF(>0)
- 자동 합계(AutoSum) 가 고르는 범위: 아래 칸에서 누를 때 / 범위를 같이 골랐을 때
- 숨긴 행 · 필터: SUM 그대로, SUBTOTAL(109) · SUBTOTAL(9) · 필터 뒤 자동 합계가 넣는 식
- 글자 숫자 고친 뒤 다시
결과 measure.json (argv 'en' 이면 measure_en.json). 실행: python measure.py [en]  (pywin32 필요)"""
import datetime, json, os, sys
import time
import pythoncom
import win32com.client as win32
import win32con, win32gui

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}
clean = lambda s: (s or "").replace("​", "").strip()
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"] if EN else ["월", "화", "수", "목", "금", "토", "일"]
HEAD = ("Day", "Sold") if EN else ("요일", "판매 수량")
# 수요일 빈칸(쉬는 날) · 목요일 글자 '18' · 일요일 0
SOLD = [12, 15, None, "18", 20, 25, 0]


def put(ws, cell, f):
    try:
        ws.Range(cell).Formula2 = f
        return {"formula": f, "text": ws.Range(cell).Text}
    except Exception as e:  # noqa: BLE001
        return {"formula": f, "rejected": str(e)[:120]}


def commit_edit():
    """자동 합계를 칸 하나에서 누르면 엑셀이 '편집 중'(식 제안)으로 멈춘다 — 그 창(EXCEL7)에 Enter 를 보내 확정(사람이 Enter 누르는 것과 같음)."""
    hits = []

    def cb(h, _):
        if win32gui.GetClassName(h) == "EXCEL7":
            hits.append(h)
        return True
    win32gui.EnumChildWindows(HWND, cb, 0)  # xl.Hwnd 는 편집 중엔 COM 이 막혀 미리 받아 둠
    for h in hits:
        win32gui.PostMessage(h, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(h, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    time.sleep(1.5)
    return len(hits)


def autosum_at(ws, cell):
    ws.Range(cell).Select()
    xl.CommandBars.ExecuteMso("AutoSum")
    n = commit_edit()
    for _ in range(20):
        try:
            return {"formula": ws.Range(cell).Formula, "text": ws.Range(cell).Text, "enter_sent_to": n}
        except Exception:  # noqa: BLE001
            time.sleep(0.5)
    return {"err": "still editing"}


def fill(ws):
    ws.Range("A1:B1").Value = (HEAD,)
    for i, (d, v) in enumerate(zip(DAYS, SOLD), start=2):
        ws.Range(f"A{i}").Value = d
        if v is None:
            continue
        if isinstance(v, str):
            ws.Range(f"B{i}").NumberFormat = "@"
        ws.Range(f"B{i}").Value = v


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = True  # 자동 합계 편집 상태에 Enter 를 보내려면 창이 있어야 한다
    xl.WindowState = -4143  # 보통(최소화하면 ExecuteMso 가 E_FAIL — 11/10)
    xl.Width, xl.Height = 760, 470
    HWND = xl.Hwnd
    xl.DisplayAlerts = False
    R["build"] = xl.Build
    R["labels"] = {}
    for k in ("AutoSum", "AutoSumMenu", "FunctionsSumInsertGallery", "Filter", "ConvertToNumber"):
        try:
            R["labels"][k] = clean(xl.CommandBars.GetLabelMso(k))
        except Exception as e:  # noqa: BLE001
            R["labels"][k] = "ERR " + str(e)[:60]
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    fill(ws)
    R["b5_is_text"] = put(ws, "D1", "=ISTEXT(B5)")
    R["b5_error_flag"] = bool(ws.Range("B5").Errors(3).Value)  # xlNumberAsText
    F = {}
    F["sum"] = put(ws, "D2", "=SUM(B2:B8)")
    F["plus"] = put(ws, "D3", "=B2+B3+B4+B5+B6+B7+B8")
    F["average"] = put(ws, "D4", "=AVERAGE(B2:B8)")
    F["count"] = put(ws, "D5", "=COUNT(B2:B8)")
    F["counta"] = put(ws, "D6", "=COUNTA(B2:B8)")
    F["averageif_pos"] = put(ws, "D7", '=AVERAGEIF(B2:B8,">0")')
    F["average_plus_text"] = put(ws, "D8", "=AVERAGE(B2,B3,B5,B6,B7,B8)")
    F["sum_cells_direct"] = put(ws, "D9", '=SUM(12,15,"18",20,25,0)')
    R["formulas"] = F
    # 상태 표시줄과 같은 값(선택 범위): 엑셀 함수로 대신 — 평균 · 개수(COUNTA) · 숫자 개수 · 합계
    # 자동 합계 ① B9 하나만 고르고 누르면(엑셀이 위로 숫자 칸을 찾아 범위를 제안)
    R["autosum_single"] = autosum_at(ws, "B9")
    ws.Range("B9").ClearContents()
    # 자동 합계 ② B2:B9 를 같이 고르고 누르면
    ws.Range("B2:B9").Select()
    try:
        xl.CommandBars.ExecuteMso("AutoSum")
        R["autosum_range"] = {"formula": ws.Range("B9").Formula, "text": ws.Range("B9").Text}
    except Exception as e:  # noqa: BLE001
        R["autosum_range"] = {"err": str(e)[:160]}
    ws.Range("B9").ClearContents()

    # 숨긴 행 · 필터 — 글자 숫자를 고친 표로(시트 2)
    ws2 = wb.Worksheets.Add(After=ws)
    fill(ws2)
    R["ready_before_fix"] = xl.Ready
    for _ in range(10):
        try:
            ws2.Range("B5").NumberFormat = "General"
            break
        except Exception as e:  # noqa: BLE001
            R.setdefault("fix_retry", []).append(str(e)[-80:]); time.sleep(1)
    ws2.Range("B5").Value = 18
    R["fixed"] = {"sum": put(ws2, "D2", "=SUM(B2:B8)"), "average": put(ws2, "D3", "=AVERAGE(B2:B8)"),
                  "averageif_pos": put(ws2, "D4", '=AVERAGEIF(B2:B8,">0")'), "count": put(ws2, "D5", "=COUNT(B2:B8)")}
    # 토요일(7행) 숨기기
    ws2.Rows(7).Hidden = True
    R["hidden"] = {"sum": put(ws2, "E2", "=SUM(B2:B8)"), "subtotal109": put(ws2, "E3", "=SUBTOTAL(109,B2:B8)"),
                   "subtotal9": put(ws2, "E4", "=SUBTOTAL(9,B2:B8)"), "aggregate": put(ws2, "E5", "=AGGREGATE(9,5,B2:B8)")}
    ws2.Rows(7).Hidden = False
    # 필터: 판매 수량 >= 15 만
    ws2.Range("A1:B8").AutoFilter(Field=2, Criteria1=">=15")
    vis = [i for i in range(2, 9) if not ws2.Rows(i).Hidden]
    R["filter"] = {"visible_rows": vis, "visible": [[ws2.Range(f"A{i}").Value, ws2.Range(f"B{i}").Value] for i in vis],
                   "sum": put(ws2, "F2", "=SUM(B2:B8)"), "subtotal9": put(ws2, "F3", "=SUBTOTAL(9,B2:B8)"),
                   "subtotal109": put(ws2, "F4", "=SUBTOTAL(109,B2:B8)")}
    R["filter"]["autosum"] = autosum_at(ws2, "B9")
    wb.Close(False)
finally:
    out = os.path.join(HERE, "measure_en.json" if EN else "measure.json")
    json.dump(R, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    xl.Quit()
json.dump(R, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
