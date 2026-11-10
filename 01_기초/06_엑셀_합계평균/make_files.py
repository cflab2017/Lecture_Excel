"""엑셀 기초 6강(합계 · 평균) 실습 · 정답 파일 — 이 PC 엑셀을 따로 띄워(DispatchEx) 만든다. 내용은 연습용 가상 판매 수량.
실습: 문구점 일주일 판매 수량 — 수요일 쉬는 날(빈칸) · 일요일 0개 · 목요일 18 이 '글자로 저장된 숫자'. 합계 · 평균 칸은 비어 있음.
정답: 목요일을 숫자로 고침 → 합계 =SUM(B3:B9) 90 · 평균 =AVERAGE(B3:B9) 15 · 판 날만 평균 =AVERAGEIF(B3:B9,">0") 18 ·
      필터용 보이는 행 합계 =SUBTOTAL(109,B3:B9) · 확인용 =COUNT 6 / =COUNTA 6.
argv 'en' 이면 영어(practice.xlsx · answer.xlsx). 작성자 · 마지막 저장자는 CODINGNOW.
실행: python make_files.py [en]  (pywin32)"""
import os, sys
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
TITLE = "Stationery shop — items sold this week" if EN else "문구점 일주일 판매 수량"
H = ("Day", "Sold") if EN else ("요일", "판매 수량")
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"] if EN else ["월", "화", "수", "목", "금", "토", "일"]
SOLD = [12, 15, None, "18", 20, 25, 0]
L = (["Total", "Average", "Average (days with sales)", "Visible rows total", "Number cells", "Non-empty cells"] if EN else
     ["합계", "평균", "판 날만 평균", "보이는 행 합계", "숫자 칸 수", "빈칸 아닌 칸 수"])
NOTE_P = ("Wed: closed (blank) · Sun: open, 0 sold · Thu 18 was typed as text. Try: total and average in B11 · B12."
          if EN else "수: 쉬는 날(빈칸) · 일: 문은 열었지만 0개 · 목 18 은 글자로 들어가 있음. 해 보기: B11 합계 · B12 평균")
NOTE_A = ("Thu fixed to a number. Total 90 · Average 15 (blank Wed skipped, Sun 0 counted) · days with sales 18."
          if EN else "목을 숫자로 고침. 합계 90 · 평균 15(빈칸 수요일은 빼고 0 인 일요일은 셈) · 판 날만 평균 18")


def table(ws, fixed):
    ws.Range("A1").Value = TITLE
    ws.Range("A2:B2").Value = (H,)
    for i, (d, v) in enumerate(zip(DAYS, SOLD), start=3):
        ws.Range(f"A{i}").Value = d
        if v is None:
            continue
        if isinstance(v, str) and not fixed:
            ws.Range(f"B{i}").NumberFormat = "@"
            ws.Range(f"B{i}").Value = v
        else:
            ws.Range(f"B{i}").Value = int(v)
    ws.Range("A2:B2").Font.Bold = True
    ws.Range("A2:B2").Interior.Color = 0xEADBC8
    ws.Range("A2:B9").Borders.LineStyle = 1
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 24 if EN else 18
    ws.Columns("B").ColumnWidth = 12
    for k, lab in enumerate(L, start=11):
        ws.Range(f"A{k}").Value = lab
    ws.Range("A11:B16").Borders.LineStyle = 1
    ws.Range("A11:A16").Interior.Color = 0xF2F2F2


def author(wb):
    for k in ("Author", "Last Author"):
        wb.BuiltinDocumentProperties(k).Value = "CODINGNOW"


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Practice" if EN else "실습"
    table(ws, fixed=False)
    ws.Range("A18").Value = NOTE_P
    author(wb)
    p = os.path.join(HERE, "practice.xlsx" if EN else "실습.xlsx")
    if os.path.exists(p):
        os.remove(p)
    wb.SaveAs(p, 51)
    wb.Close(False)
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Answer" if EN else "정답"
    table(ws, fixed=True)
    for k, f in enumerate(["=SUM(B3:B9)", "=AVERAGE(B3:B9)", '=AVERAGEIF(B3:B9,">0")', "=SUBTOTAL(109,B3:B9)", "=COUNT(B3:B9)", "=COUNTA(B3:B9)"], start=11):
        ws.Range(f"B{k}").Formula = f
    ws.Range("A2:B9").AutoFilter()
    ws.Range("A18").Value = NOTE_A
    author(wb)
    p2 = os.path.join(HERE, "answer.xlsx" if EN else "정답.xlsx")
    if os.path.exists(p2):
        os.remove(p2)
    wb.SaveAs(p2, 51)
    vals = [ws.Range(f"B{k}").Text for k in range(11, 17)]
    wb.Close(False)
    print("ok", p, p2, vals)
    assert vals == ["90", "15", "18", "90", "6", "6"], vals
finally:
    xl.Quit()
