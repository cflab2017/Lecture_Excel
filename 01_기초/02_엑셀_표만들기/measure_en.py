"""Excel Basics 2 (English) — same checks as measure.py with the English flower-shop order list (prices in dollars).
Writes measure_en.json, practice.xlsx and answer.xlsx. This PC's Excel runs with a Korean display language,
so default names it makes (표1, 요약, 열1) are Korean here — the English page uses Microsoft's English names (Table1, Total, Column1).
Run: python measure_en.py  (needs pywin32)"""
import datetime, json, os
import pythoncom
import win32com.client as win32
from measure_util import fill

HERE = os.path.dirname(os.path.abspath(__file__))
HEAD = ("Order date", "Customer", "Item", "Qty", "Price", "Amount")
ROWS = [
    ("Nov 3", "Brightline Co.", "Rose bouquet", 2, 35),
    ("Nov 3", "Sophie", "Wreath", 1, 89),
    ("Nov 4", "Daniel", "Tulip pot", 3, 18),
    ("Nov 4", "Brightline Co.", "Orchid pot", 1, 65),
    ("Nov 5", "Mia", "Rose bouquet", 1, 35),
    ("Nov 5", "Sunny Kids", "Tulip pot", 6, 18),
    ("Nov 6", "Leo", "Preserved flowers", 2, 42),
    ("Nov 6", "Sophie", "Rose bouquet", 3, 35),
]
NEW = ("Nov 7", "Oak Pharmacy", "Orchid pot", 2, 65)
R = {"run": datetime.datetime.now().isoformat(timespec="seconds")}


def en(ws, with_amount=True):
    fill(ws, with_amount, HEAD, ROWS)
    ws.Columns("B").ColumnWidth = 15; ws.Columns("C").ColumnWidth = 17
    ws.Range("E2:F20").NumberFormat = "$#,##0"


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False; xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    a = wb.Worksheets(1); en(a)
    a.Range("A1:F9").Borders.LineStyle = 1
    a.Range("I1").Formula = "=SUM(F2:F9)"; before = a.Range("I1").Value
    a.Range("A10:E10").Value = (NEW,); a.Range("F10").Formula = "=D10*E10"
    R["border"] = {"sum_before": before, "sum_after_new_row": a.Range("I1").Value, "new_row_border_bottom": a.Range("A10").Borders(9).LineStyle}
    s = wb.Worksheets.Add(); en(s)
    b = [(s.Range(f"B{i}").Value, s.Range(f"C{i}").Value) for i in range(2, 10)]
    s.Range("C2:C9").Sort(Key1=s.Range("C2"), Order1=1, Header=2)
    c = [(s.Range(f"B{i}").Value, s.Range(f"C{i}").Value) for i in range(2, 10)]
    R["sort_trap"] = {"after_items": [y for _, y in c], "moved_rows": sum(1 for x, y in zip(b, c) if x[1] != y[1])}
    t = wb.Worksheets.Add(); en(t, with_amount=False)
    lo = t.ListObjects.Add(1, t.Range("A1").CurrentRegion, None, 1)
    t.Range("F2").Formula = "=[@Qty]*[@Price]"
    t.Range("A10:E10").Value = (NEW,)
    lo.ShowTotals = True
    tot = lo.TotalsRowRange
    R["table"] = {"range": lo.Range.Address.replace("$", ""), "F10": t.Range("F10").Formula, "total": tot.Cells(1, 6).Value, "total_formula": tot.Cells(1, 6).Formula}
    lo.Range.AutoFilter(Field=3, Criteria1="Rose bouquet")
    R["filter"] = {"total": tot.Cells(1, 6).Value, "visible": sum(1 for r in range(2, 11) if not t.Rows(r).Hidden)}
    lo.Range.AutoFilter(Field=3)
    lo.Sort.SortFields.Clear(); lo.Sort.SortFields.Add(Key=lo.ListColumns("Item").DataBodyRange, Order=1); lo.Sort.Header = 1; lo.Sort.Apply()
    R["table_sort"] = [(t.Range(f"B{i}").Value, t.Range(f"C{i}").Value, t.Range(f"F{i}").Value) for i in range(2, 11)]
    wb.Close(False)

    p = xl.Workbooks.Add(); ws = p.Worksheets(1); ws.Name = "Orders"; en(ws, with_amount=False); ws.Range("A1:F1").Font.Bold = True
    p.SaveAs(os.path.join(HERE, "practice.xlsx"), 51); p.Close(False)
    q = xl.Workbooks.Add(); ws = q.Worksheets(1); ws.Name = "Orders"; en(ws, with_amount=False)
    lo = ws.ListObjects.Add(1, ws.Range("A1").CurrentRegion, None, 1); lo.Name = "Orders"
    ws.Range("F2").Formula = "=[@Qty]*[@Price]"
    lo.ShowTotals = True; lo.TotalsRowRange.Cells(1, 1).Value = "Total"
    q.SaveAs(os.path.join(HERE, "answer.xlsx"), 51)
    R["answer"] = {"total": lo.TotalsRowRange.Cells(1, 6).Value}
    q.Close(False)
finally:
    xl.Quit()

json.dump(R, open(os.path.join(HERE, "measure_en.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False))
