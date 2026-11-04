"""엑셀 기초 1강 영어판 실측 — 영어 데이터로 measure.py 의 핵심만 다시(LEN·코드·자동 줄 바꿈·줄 높이·수식·바꾸기). 결과 measure_en.json."""
import datetime, json, os, pythoncom, win32com.client as win32
HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds")}
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Columns("D").ColumnWidth = 20
    R["font"] = [ws.Range("D2").Font.Name, ws.Range("D2").Font.Size]
    R["h0"] = ws.Rows(2).RowHeight
    ws.Range("D2").Value = "Laptop\nProjector\n10 handouts"
    R["typed"] = {"len": int(ws.Evaluate("=LEN(D2)")), "code": int(ws.Evaluate("=CODE(MID(D2,7,1))")), "wrap": ws.Range("D2").WrapText, "h": ws.Rows(2).RowHeight}
    ws.Range("B3").Value = "Nov 10"; ws.Range("C3").Value = "Room 3A"
    ws.Range("D3").Formula = "=B3&CHAR(10)&C3"
    R["formula"] = {"value": ws.Range("D3").Value, "wrap": ws.Range("D3").WrapText, "h": ws.Rows(3).RowHeight}
    ws.Range("D3").WrapText = True
    R["formula"]["h_after_wrap"] = ws.Rows(3).RowHeight
    ws.Rows(4).RowHeight = ws.Rows(1).RowHeight
    ws.Range("D4").Value = "Laptop\nProjector"; ws.Range("D4").WrapText = True
    R["manual"] = {"h_set": ws.Rows(4).RowHeight}
    ws.Rows(4).AutoFit(); R["manual"]["h_autofit"] = ws.Rows(4).RowHeight
    ws.Range("D5").Value = "Quote, Business cards, 3 samples, Drinks"
    ws.Range("D5").Replace(What=", ", Replacement="\n", LookAt=2)
    R["replace"] = {"value": ws.Range("D5").Value, "wrap": ws.Range("D5").WrapText, "h": ws.Rows(5).RowHeight}
    ws.Rows(5).AutoFit(); R["replace"]["h_autofit"] = ws.Rows(5).RowHeight
    ws.Range("E2").Formula = '=SUBSTITUTE(D2,CHAR(10),", ")'; R["substitute"] = ws.Range("E2").Value
    wb.Close(False)
finally:
    xl.Quit()
json.dump(R, open(os.path.join(HERE, "measure_en.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False))
