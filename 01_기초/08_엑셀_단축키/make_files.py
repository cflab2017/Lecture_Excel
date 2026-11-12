"""엑셀 기초 8강(단축키) 실습 · 정답 파일 — 가상 카페 11월 주문 기록(실제 가게 아님).
실습: 2~7행 · 8행 빈 줄 · 9~13행, D14 '합계'(E14 비움), F1 '비율'(비움) — Ctrl+방향키 · 자동 합계(Alt+=)가 빈 줄에서 멈추는 것을 직접 보게.
정답: 빈 줄(8행)을 지워 2~12행, E13 =SUM(E2:E12), F2:F12 =E2/$E$13 (0.0%).
실행: python make_files.py [en]  (pywin32)"""
import os, sys
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
EN = len(sys.argv) > 1 and sys.argv[1] == "en"

ORDERS = [  # (날짜, 메뉴 번호, 수량) — 메뉴 번호는 MENU 순서
    ("2026-11-02", 0, 2), ("2026-11-02", 1, 1), ("2026-11-02", 2, 3), ("2026-11-03", 3, 2), ("2026-11-03", 0, 4), ("2026-11-03", 1, 2),
    None,
    ("2026-11-04", 0, 3), ("2026-11-04", 2, 1), ("2026-11-04", 3, 2), ("2026-11-05", 1, 3), ("2026-11-05", 0, 1),
]
MENU_KO = [("아메리카노", 4500), ("카페라테", 5000), ("바닐라라테", 5500), ("녹차라테", 4800)]
MENU_EN = [("Americano", 4500), ("Latte", 5000), ("Vanilla latte", 5500), ("Green tea latte", 4800)]
HEAD_KO = ["날짜", "메뉴", "수량", "단가", "금액", "비율"]
HEAD_EN = ["Date", "Item", "Qty", "Price", "Amount", "Share"]


def build(wb, answer):
    ws = wb.Worksheets(1)
    ws.Name = "Orders" if EN else "주문"
    menu = MENU_EN if EN else MENU_KO
    ws.Range("A1:F1").Value = (tuple(HEAD_EN if EN else HEAD_KO),)
    ws.Range("A1:F1").Font.Bold = True
    rows = [o for o in ORDERS if o is not None] if answer else ORDERS
    r = 2
    for o in rows:
        if o is not None:
            d, m, q = o
            ws.Range(f"A{r}").Value = d
            ws.Range(f"A{r}").NumberFormat = "yyyy-mm-dd"
            ws.Range(f"B{r}").Value = menu[m][0]
            ws.Range(f"C{r}").Value = q
            ws.Range(f"D{r}").Value = menu[m][1]
            ws.Range(f"E{r}").Formula = f"=C{r}*D{r}"
        r += 1
    last = r - 1
    t = last + 1
    ws.Range(f"D{t}").Value = "Total" if EN else "합계"
    ws.Range(f"D{t}:E{t}").Font.Bold = True
    if answer:
        ws.Range(f"E{t}").Formula = f"=SUM(E2:E{last})"
        for k in range(2, last + 1):
            ws.Range(f"F{k}").Formula = f"=E{k}/$E${t}"
        ws.Range(f"F2:F{last}").NumberFormat = "0.0%"
    ws.Range(f"D2:E{t}").NumberFormat = "#,##0"
    for c, w in zip("ABCDEF", (12, 16, 7, 9, 11, 9)):
        ws.Columns(c).ColumnWidth = w
    ws.Range("A1").Select()
    return t


def fix_core(path):
    """docProps/core.xml 의 작성자 · 마지막 수정한 사람을 CODINGNOW 로(7강 make_files.fix_core 와 같음)."""
    import re, shutil, tempfile, zipfile
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "docProps/core.xml":
                s = data.decode("utf-8")
                for tag in ("dc:creator", "cp:lastModifiedBy"):
                    s = re.sub(f"<{tag}>[^<]*</{tag}>", f"<{tag}>CODINGNOW</{tag}>", s)
                data = s.encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, path)


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = True
    xl.DisplayAlerts = False
    for answer, name in ((False, "practice.xlsx" if EN else "실습.xlsx"), (True, "answer.xlsx" if EN else "정답.xlsx")):
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count > 1:
            wb.Worksheets(wb.Worksheets.Count).Delete()
        t = build(wb, answer)
        wb.BuiltinDocumentProperties("Author").Value = "CODINGNOW"
        path = os.path.join(HERE, name)
        if os.path.exists(path):
            os.remove(path)
        wb.SaveAs(path, 51)
        wb.Close(False)
        fix_core(path)
        print("saved", name, "total row", t)
finally:
    xl.Quit()
