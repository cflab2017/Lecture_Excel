"""엑셀 기초 1강 영어판 실습·정답 파일(practice.xlsx · answer.xlsx) — 데이터만 영어, 만드는 법은 measure.py 와 같다. 실행: python make_en_files.py"""
import os, pythoncom, win32com.client as win32
HERE = os.path.dirname(os.path.abspath(__file__))
ROWS = [("Nov 10", "Room 3A", "Quarterly report", ["Laptop", "Projector", "10 handouts"]),
        ("Nov 11", "Room 2B", "New hire training", ["Training book", "Name tags"]),
        ("Nov 12", "Room 3A", "Client meeting", ["Quote", "Business cards", "3 samples", "Drinks"]),
        ("Nov 13", "Seminar room 5", "Team workshop", ["Whiteboard markers", "Sticky notes"])]
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    for name, joiner, wrap in (("practice.xlsx", ", ", False), ("answer.xlsx", "\n", True)):
        w = xl.Workbooks.Add(); s = w.Worksheets(1); s.Name = "Bookings"
        s.Range("A1:D1").Value = (("Date", "Room", "Meeting", "Items to bring"),); s.Range("A1:D1").Font.Bold = True
        for i, (d, room, what, items) in enumerate(ROWS, start=2):
            s.Range(f"A{i}:D{i}").Value = ((d, room, what, joiner.join(items)),)
        s.Columns("A").ColumnWidth = 9; s.Columns("B").ColumnWidth = 15; s.Columns("C").ColumnWidth = 18; s.Columns("D").ColumnWidth = 20
        s.Range("D2:D5").WrapText = wrap
        s.Range("A2:C5").VerticalAlignment = -4108 if wrap else -4107
        s.Rows("1:5").AutoFit()
        p = os.path.join(HERE, name)
        if os.path.exists(p): os.remove(p)
        w.SaveAs(p, 51); w.Close(False); print(name)
finally:
    xl.Quit()
