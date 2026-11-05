"""엑셀 기초 3강 화면 — 보이는 엑셀(따로 띄운 것)에 판매표를 만들고 PrintWindow 로 창을 찍는다. 저장하지 않고 닫는다.
찍는 것: ① C2 틀 고정 뒤 40행·I열로 내린 화면(1행·A:B 남음) ② 15행까지 내린 채 C20 에서 틀 고정 → 맨 위로 올려도 15행부터
③ A1 을 고른 채 틀 고정(화면 가운데에 십자) ④ 페이지 레이아웃 보기에서 아래로 내린 화면(틀 고정이 보이는지).
결과: shots/*.png (원본 크기). 실행: python screens.py"""
import os, time
import win32com.client as win32
from measure_util import grab, fill

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)

xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    xl.Visible = True
    xl.WindowState = -4143
    xl.Width, xl.Height = 900, 520
    ws = wb.Worksheets(1); ws.Name = "판매"
    fill(ws)
    w = xl.ActiveWindow
    w.Zoom = 100

    def freeze(sel, pre=(1, 1)):
        if w.FreezePanes:
            w.FreezePanes = False
        w.ScrollRow, w.ScrollColumn = pre
        ws.Range(sel).Select()
        xl.CommandBars.ExecuteMso("FreezePanes"); time.sleep(0.6)

    def last_pane(row, col):
        p = w.Panes(w.Panes.Count)
        p.ScrollRow, p.ScrollColumn = row, col
        time.sleep(1.2)

    # ① C2 틀 고정 → 40행 · I열
    freeze("C2"); last_pane(40, 9); ws.Range("I40").Select(); time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, "s1_c2_scrolled.png"))
    # ② 15행까지 내린 채 C20 → 맨 위로 올려도 15행부터
    freeze("C20", pre=(15, 1)); last_pane(1, 1); ws.Range("C20").Select(); time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, "s2_trap_top.png"))
    # ③ A1 을 고른 채 틀 고정
    freeze("A1"); time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, "s3_a1_center.png"))
    # ④ 페이지 나누기 미리 보기(View=2) · ⑤ 페이지 레이아웃(View=3) — C2 틀 고정을 켠 채로 바꾸고 30행으로
    for name, v in (("s4_pagebreak.png", 2), ("s5_pagelayout.png", 3)):
        w.View = 1; time.sleep(0.5)
        freeze("C2")
        w.View = v; time.sleep(1.5)
        w.ScrollRow = 30; time.sleep(1.5)
        ws.Range("C30").Select(); time.sleep(1.0)
        grab(xl.Hwnd, os.path.join(OUT, name))
    w.View = 1; time.sleep(1.0)
    w.ScrollRow, w.ScrollColumn = 30, 1; time.sleep(1.2)
    grab(xl.Hwnd, os.path.join(OUT, "s6_back_normal.png"))
    wb.Close(False)
finally:
    xl.Quit()
