"""엑셀 기초 2강 화면 — 보이는 엑셀(따로 띄운 것)에 시험 시트를 만들고 PrintWindow 로 창을 찍는다. 저장하지 않고 닫는다.
찍는 것: ① 테두리로 그린 표 + 바로 아래 새 주문(테두리 없음·합계 그대로) ② 표 서식(Ctrl+T) + 요약 행(수식 입력줄에 SUBTOTAL)
③ 같은 표를 상품 '장미 꽃다발'로 거른 뒤(요약 행 210,000 · 표 밖 SUM 740,000).
결과: shots/*.png (원본 크기). 실행: python screens.py"""
import os, time
import win32com.client as win32
from measure_util import grab, HEAD, ROWS, NEW, fill

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)

xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    xl.Visible = True
    xl.WindowState = -4143
    xl.Width, xl.Height = 1000, 520
    # ① 테두리
    a = wb.Worksheets(1); a.Name = "테두리"; a.Activate()
    fill(a); a.Range("A1:F9").Borders.LineStyle = 1; a.Range("A1:F1").Font.Bold = True
    a.Range("A10:E10").Value = (NEW,); a.Range("F10").Formula = "=D10*E10"
    a.Range("H1").Value = "합계"; a.Range("I1").Formula = "=SUM(F2:F9)"; a.Range("I1").NumberFormat = "#,##0"; a.Columns("I").ColumnWidth = 10
    xl.ActiveWindow.Zoom = 120; xl.ActiveWindow.DisplayGridlines = True
    xl.ActiveWindow.ScrollRow, xl.ActiveWindow.ScrollColumn = 1, 1
    a.Range("I1").Select(); time.sleep(2.0)
    grab(xl.Hwnd, os.path.join(OUT, "s1_border.png"))
    # ② 표 서식 + 요약 행
    t = wb.Worksheets.Add(After=a); t.Name = "표"; t.Activate()
    fill(t, with_amount=False)
    lo = t.ListObjects.Add(1, t.Range("A1").CurrentRegion, None, 1)
    t.Range("F2").Formula = "=[@수량]*[@단가]"
    t.Range("A10:E10").Value = (NEW,)
    lo.ShowTotals = True
    t.Range("H1").Value = "SUM"; t.Range("I1").Formula = "=SUM(표1[금액])"; t.Range("I1").NumberFormat = "#,##0"; t.Columns("I").ColumnWidth = 10
    xl.ActiveWindow.Zoom = 120
    xl.ActiveWindow.ScrollRow, xl.ActiveWindow.ScrollColumn = 1, 1
    t.Range("F11").Select(); time.sleep(2.0)
    grab(xl.Hwnd, os.path.join(OUT, "s2_table.png"))
    # ③ 필터
    lo.Range.AutoFilter(Field=3, Criteria1="장미 꽃다발")
    t.Range("F11").Select(); time.sleep(2.0)
    grab(xl.Hwnd, os.path.join(OUT, "s3_filter.png"))
    wb.Close(False)
finally:
    xl.Quit()
