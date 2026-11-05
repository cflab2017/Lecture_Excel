"""엑셀 기초 3강 — 틀 고정 실측. 이 PC 엑셀(Microsoft 365)을 따로 띄워(DispatchEx) 리본 명령(ExecuteMso)으로 틀 고정을 실행하고,
창이 몇 칸으로 나뉘었는지(SplitRow·SplitColumn·Panes), 내려 봤을 때 어느 행·열이 남는지(Panes(n).VisibleRange)를 잰다.
인쇄(PDF)에서 틀 고정·인쇄 제목이 어떻게 되는지도 본다. 결과 measure.json, 실습.xlsx · 정답.xlsx.
데이터는 모두 연습용 가상 문구점 판매 수량. 실행: python measure.py  (pywin32 · pymupdf 필요)
argv 'en' 이면 영어 머리글 데이터로 같은 것을 재서 measure_en.json · practice.xlsx · answer.xlsx."""
import datetime, json, os, sys, tempfile, time
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
from measure_util import fill

EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}
clean = lambda s: (s or "").replace("​", "").strip()


def state(xl):
    w = xl.ActiveWindow
    panes = []
    for k in range(1, w.Panes.Count + 1):
        p = w.Panes(k)
        panes.append({"scroll_row": p.ScrollRow, "scroll_col": p.ScrollColumn, "visible": p.VisibleRange.Address.replace("$", "")})
    return {"freeze": bool(w.FreezePanes), "split_row": w.SplitRow, "split_col": w.SplitColumn, "panes": panes}


def scroll_last(xl, row, col):
    """마지막(오른쪽 아래) 칸을 내린다 — 사람이 휠로 내리는 것과 같은 자리."""
    w = xl.ActiveWindow
    p = w.Panes(w.Panes.Count)
    p.ScrollRow, p.ScrollColumn = row, col
    time.sleep(0.3)


def run(xl, ws, idm, select=None, pre_scroll=None):
    ws.Activate()
    w = xl.ActiveWindow
    if w.FreezePanes:
        w.FreezePanes = False
    w.ScrollRow, w.ScrollColumn = (pre_scroll or (1, 1))
    if select:
        ws.Range(select).Select()
    xl.CommandBars.ExecuteMso(idm); time.sleep(0.6)
    out = {"cmd": idm, "select": select, "pre_scroll": pre_scroll, "after": state(xl)}
    scroll_last(xl, 40, 9)
    out["scrolled_to_row40_colI"] = state(xl)
    # 위로 끝까지 올려 본다(1행)
    scroll_last(xl, 1, 1)
    out["scroll_back_to_1"] = state(xl)
    return out


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    R["build"] = xl.Build
    R["labels"] = {idm: clean(xl.CommandBars.GetLabelMso(idm)) for idm in
                   ("ViewFreezePanesGallery", "FreezePanes", "FreezeTopRow", "FreezeFirstColumn", "ViewPageLayoutView", "ViewNormalViewExcel", "PrintTitles")}
    wb = xl.Workbooks.Add()
    xl.Visible = True                      # 창(ActiveWindow)이 있어야 틀 고정이 실제로 된다
    xl.WindowState = -4143
    xl.Width, xl.Height = 900, 520
    ws = wb.Worksheets(1); ws.Name = "Sales" if EN else "판매"
    last = fill(ws, EN)
    R["rows"], R["last_row"] = last - 1, last
    xl.ActiveWindow.Zoom = 100
    # 1) 첫 행 고정 · 2) 첫 열 고정 · 3) C2 를 고르고 틀 고정
    R["top_row"] = run(xl, ws, "FreezeTopRow", select="E10")
    R["first_col"] = run(xl, ws, "FreezeFirstColumn", select="E10")
    R["c2"] = run(xl, ws, "FreezePanes", select="C2")
    for cell in ("A3", "C1", "B2"):       # 도움말 규칙(고른 칸의 위쪽 행·왼쪽 열) 확인용
        R[cell.lower()] = run(xl, ws, "FreezePanes", select=cell)
    # 4) 함정 — 화면을 15행까지 내린 채 C20 을 고르고 틀 고정
    R["scrolled_trap"] = run(xl, ws, "FreezePanes", select="C20", pre_scroll=(15, 1))
    # 5) 함정 — A1(맨 위 왼쪽 칸)을 고른 채 틀 고정
    R["a1"] = run(xl, ws, "FreezePanes", select="A1")
    # 6) 보기 바꾸기 — 틀 고정(C2)을 켠 채 페이지 나누기 미리 보기(View=2) · 페이지 레이아웃(View=3)으로 바꾸면?
    #    (XlWindowView: 1 기본 · 2 페이지 나누기 미리 보기 · 3 페이지 레이아웃 — 11/5 처음에 2 를 페이지 레이아웃으로 잘못 썼다)
    R["views"] = {}
    for vname, v in (("page_break_preview", 2), ("page_layout", 3)):
        xl.ActiveWindow.View = 1; time.sleep(0.3)
        xl.ActiveWindow.FreezePanes = False
        ws.Range("C2").Select(); xl.CommandBars.ExecuteMso("FreezePanes"); time.sleep(0.5)
        before = bool(xl.ActiveWindow.FreezePanes)
        xl.ActiveWindow.View = v; time.sleep(0.8)
        d = {"view": xl.ActiveWindow.View, "freeze_before": before, "freeze_prop_in_view": bool(xl.ActiveWindow.FreezePanes),
             "enabled": {idm: bool(xl.CommandBars.GetEnabledMso(idm)) for idm in ("ViewFreezePanesGallery", "FreezePanes", "FreezeTopRow", "FreezeFirstColumn")}}
        xl.ActiveWindow.View = 1; time.sleep(0.5)
        d["freeze_back_in_normal"] = bool(xl.ActiveWindow.FreezePanes)
        R["views"][vname] = d
    R["views"]["normal_enabled"] = {idm: bool(xl.CommandBars.GetEnabledMso(idm)) for idm in ("ViewFreezePanesGallery", "FreezePanes", "FreezeTopRow", "FreezeFirstColumn")}
    # 7) 인쇄 — 틀 고정(C2) 그대로 PDF / 인쇄 제목 $1:$1 넣고 PDF
    import pymupdf
    # ⚠️ 리본 '틀 고정' 은 이미 고정돼 있으면 '틀 고정 취소' 로 동작한다(11/5 첫 실행에서 꺼졌다) — 먼저 끄고 다시 켠다
    xl.ActiveWindow.FreezePanes = False
    ws.Range("C2").Select(); xl.CommandBars.ExecuteMso("FreezePanes"); time.sleep(0.4)
    pr = {"freeze_on": bool(xl.ActiveWindow.FreezePanes), "toggle_note": "ExecuteMso('FreezePanes') 를 고정된 창에서 부르면 고정이 풀린다(첫 실행 freeze_on False)"}
    tmp = tempfile.mkdtemp(prefix="xl03_")
    for tag, titles, tcols in (("freeze_only", "", ""), ("print_titles", "$1:$1", ""), ("print_titles_rc", "$1:$1", "$A:$B")):
        ws.PageSetup.PrintTitleRows = titles
        ws.PageSetup.PrintTitleColumns = tcols
        pdf = os.path.join(tmp, tag + ".pdf")
        ws.ExportAsFixedFormat(0, pdf)
        doc = pymupdf.open(pdf)
        pages = []
        for pg in doc:
            words = sorted(pg.get_text("words"), key=lambda w: (round(w[1]), w[0]))
            first_line = [w[4] for w in words if abs(w[1] - words[0][1]) < 2] if words else []
            pages.append(" ".join(first_line))
        pr[tag] = {"pages": len(doc), "first_line_each_page": pages}
        doc.close()
    R["print"] = pr
    # 실습·정답 파일: 실습 = 틀 고정 없음, 정답 = C2 틀 고정 + 인쇄 제목 $1:$1
    xl.ActiveWindow.FreezePanes = False
    ws.PageSetup.PrintTitleRows = ""
    ws.PageSetup.PrintTitleColumns = ""
    xl.ActiveWindow.ScrollRow, xl.ActiveWindow.ScrollColumn = 1, 1
    ws.Range("A1").Select()
    names = ("practice.xlsx", "answer.xlsx") if EN else ("실습.xlsx", "정답.xlsx")
    wb.SaveAs(os.path.join(HERE, names[0]), 51)
    ws.Range("C2").Select(); xl.CommandBars.ExecuteMso("FreezePanes"); time.sleep(0.4)
    ws.PageSetup.PrintTitleRows = "$1:$1"
    ws.PageSetup.PrintTitleColumns = "$A:$B"
    xl.ActiveWindow.ScrollRow, xl.ActiveWindow.ScrollColumn = 1, 1
    wb.SaveAs(os.path.join(HERE, names[1]), 51)
    R["answer_state"] = state(xl)
    wb.Close(False)
finally:
    xl.Quit()

out = os.path.join(HERE, "measure_en.json" if EN else "measure.json")
json.dump(R, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
