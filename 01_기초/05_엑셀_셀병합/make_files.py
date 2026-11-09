"""엑셀 기초 5강(셀 병합) 실습 · 정답 파일 — 이 PC 엑셀을 따로 띄워(DispatchEx) 만든다. 내용은 연습용 가상 점수표.
실습: 제목 A1:C1 '병합하고 가운데 맞춤' + 반(A열) 세로 병합 → 점수로 정렬하면 오류, '1반' 필터는 한 줄만.
정답: 제목은 '선택 영역의 가운데로'(병합 아님) + 반 병합 해제 뒤 빈칸을 위 칸 값으로 채움 → 점수 내림차순 정렬 · 필터 켬.
argv 'en' 이면 영어(practice.xlsx · answer.xlsx). 작성자 · 마지막 저장자는 CODINGNOW.
실행: python make_files.py [en]  (pywin32)"""
import os, sys
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
TITLE = "1st term scores" if EN else "1학기 점수표"
H = ("Class", "Name", "Score") if EN else ("반", "이름", "점수")
ROWS = ([("Class 1", "Kim", 82), ("", "Lee", 95), ("", "Park", 77), ("Class 2", "Choi", 88), ("", "Jung", 91), ("", "Kang", 70)] if EN else
        [("1반", "김민준", 82), ("", "이서연", 95), ("", "박지후", 77), ("2반", "최하은", 88), ("", "정도윤", 91), ("", "강수아", 70)])
NOTE_P = ("Try: sort by Score (largest first), then filter Class = Class 1." if EN
          else "해 보기: 점수 큰 순서로 정렬 → 반 = 1반 으로 필터")
NOTE_A = ("Title: Center Across Selection (not merged). Class: unmerged, blanks filled with the cell above."
          if EN else "제목: 선택 영역의 가운데로(병합 아님) · 반: 병합 해제 뒤 빈칸을 위 칸 값으로 채움")


def table(ws):
    ws.Range("A1").Value = TITLE
    ws.Range("A2:C2").Value = (H,)
    for i, r in enumerate(ROWS, start=3):
        ws.Range(f"A{i}:C{i}").Value = (r,)
    ws.Range("A2:C2").Font.Bold = True
    ws.Range("A2:C2").Interior.Color = 0xEADBC8  # 연한 파랑(BGR)
    ws.Range("A2:C8").Borders.LineStyle = 1
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Columns("A:C").ColumnWidth = 12


def author(wb):
    for k in ("Author", "Last Author"):
        wb.BuiltinDocumentProperties(k).Value = "CODINGNOW"


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False
    xl.DisplayAlerts = False
    # 실습
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Practice" if EN else "실습"
    table(ws)
    ws.Range("A1:C1").Merge()
    ws.Range("A1:C1").HorizontalAlignment = -4108
    for top, bot in ((3, 5), (6, 8)):
        ws.Range(f"A{top}:A{bot}").Merge()
        ws.Range(f"A{top}:A{bot}").VerticalAlignment = -4108
        ws.Range(f"A{top}:A{bot}").HorizontalAlignment = -4108
    ws.Range("A10").Value = NOTE_P
    author(wb)
    p = os.path.join(HERE, "practice.xlsx" if EN else "실습.xlsx")
    if os.path.exists(p):
        os.remove(p)
    wb.SaveAs(p, 51)
    wb.Close(False)
    # 정답
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Answer" if EN else "정답"
    table(ws)
    ws.Range("A1:C1").HorizontalAlignment = 7  # 선택 영역의 가운데로
    blanks = ws.Range("A3:A8").SpecialCells(4)
    blanks.FormulaR1C1 = "=R[-1]C"
    ws.Range("A3:A8").Value = ws.Range("A3:A8").Value
    ws.Range("A2:C8").Sort(Key1=ws.Range("C2"), Order1=2, Header=1)
    ws.Range("A2:C8").AutoFilter()
    ws.Range("A10").Value = NOTE_A
    author(wb)
    p2 = os.path.join(HERE, "answer.xlsx" if EN else "정답.xlsx")
    if os.path.exists(p2):
        os.remove(p2)
    wb.SaveAs(p2, 51)
    rows = [[ws.Range(f"{c}{i}").Value for c in "ABC"] for i in range(3, 9)]
    wb.Close(False)
    print("ok", p, p2)
    print(rows)
finally:
    xl.Quit()
