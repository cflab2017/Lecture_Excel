"""엑셀 기초 7강(시트 복사 · 보호) 실습 · 정답 파일 — 이 PC 엑셀을 따로 띄워(DispatchEx) 만든다. 내용은 연습용 가상 독서 모임 회비.
실습: '11월' 시트(이름 · 납부 여부 O/X · 낸 금액 =IF(B3="O",설정!$B$1,0) · 합계) + '설정' 시트(월 회비 10,000). 보호 없음.
정답: '11월'을 복사해 '12월' 시트를 만들고(납부 여부는 비움), 두 달 시트 모두 납부 여부(B3:B8)만 잠금 해제 · 낸 금액(C3:C9) 수식 숨김 ·
      시트 보호(암호 없음 — 검토 → 시트 보호 해제로 바로 풀린다). 통합 문서 구조 보호는 하지 않음.
argv 'en' 이면 영어(practice.xlsx · answer.xlsx). 작성자 · 마지막 저장자는 CODINGNOW.
실행: python make_files.py [en]  (pywin32)"""
import os, sys
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
S_MONTH, S_NEXT, S_SET = ("Nov", "Dec", "Settings") if EN else ("11월", "12월", "설정")
NAMES = (["Ava", "Ben", "Chloe", "Dan", "Emma", "Finn"] if EN else ["김하늘", "이도윤", "박서연", "최지우", "정민준", "한예린"])
PAID = ["O", "O", "X", "O", "X", "O"]
HEAD = ("Name", "Paid?", "Amount", "Note") if EN else ("이름", "납부 여부", "낸 금액", "비고")
WIDTHS = {"A": 12, "B": 10, "C": 12, "D": 30}
NOTE_P = ("Try: copy this sheet as 'Dec' · protect it so only B3:B8 can be typed in."
          if EN else "해 보기: 이 시트를 복사해 '12월' 만들기 · B3:B8 만 입력되게 시트 보호")
NOTE_A = ("Dec = copy of Nov · only Paid? (B3:B8) unlocked · Amount formulas hidden · sheet protected (no password)."
          if EN else "12월 = 11월 복사본 · 납부 여부(B3:B8)만 잠금 해제 · 낸 금액 수식 숨김 · 시트 보호(암호 없음)")


def build(wb, answer):
    ws_set = wb.Worksheets(1)
    ws_set.Name = S_SET
    ws_set.Range("A1").Value = "Monthly fee" if EN else "월 회비"
    ws_set.Range("B1").Value = 10000
    ws_set.Range("B1").NumberFormat = "#,##0"
    ws_set.Columns("A").ColumnWidth = 12
    ws = wb.Worksheets.Add(ws_set)  # Before — 이름 붙인 인자는 늦은 바인딩에서 무시된다(measure.py 11/11)
    ws.Name = S_MONTH
    ws.Range("A1").Value = "Book club — November fees" if EN else "독서 모임 11월 회비"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2:D2").Value = (HEAD,)
    ws.Range("A2:D2").Font.Bold = True
    ws.Range("A2:D2").Interior.Color = 0xEFE6DA
    for i, (n, p) in enumerate(zip(NAMES, PAID), start=3):
        ws.Range(f"A{i}").Value = n
        ws.Range(f"B{i}").Value = p
        ws.Range(f"C{i}").Formula = f'=IF(B{i}="O",{S_SET}!$B$1,0)'
    ws.Range("A9").Value = "Total" if EN else "합계"
    ws.Range("A9").Font.Bold = True
    ws.Range("C9").Formula = "=SUM(C3:C8)"
    ws.Range("C3:C9").NumberFormat = "#,##0"
    ws.Range("B3:B8").HorizontalAlignment = -4108
    ws.Range("A2:D9").Borders.LineStyle = 1
    for c, w in WIDTHS.items():
        ws.Columns(c).ColumnWidth = w
    ws.Range("A11").Value = NOTE_A if answer else NOTE_P
    ws.Range("A11").Font.Color = 0x666666
    if answer:
        ws.Copy(None, ws)  # (Before 비움, After=ws)
        nx = wb.Worksheets(ws.Index + 1)
        nx.Name = S_NEXT
        nx.Range("A1").Value = "Book club — December fees" if EN else "독서 모임 12월 회비"
        nx.Range("B3:B8").ClearContents()
        for s in (ws, nx):
            s.Range("B3:B8").Locked = False
            s.Range("C3:C9").FormulaHidden = True
            s.Protect()
    ws.Activate()
    ws.Range("B3").Select()
    return ws


def author(wb):
    for k in ("Author", "Last Author"):
        wb.BuiltinDocumentProperties(k).Value = "CODINGNOW"


def fix_core(path):
    """창을 보이게 한 채 저장하면 엑셀이 '마지막 수정한 사람'에 Office 사용자 이름을 넣는다(11/11 — 6강은 창을 숨겨 괜찮았음).
    저장한 파일의 docProps/core.xml 을 직접 고쳐 작성자 · 마지막 수정한 사람을 CODINGNOW 로."""
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
    xl.Visible = True  # Activate · Select 는 창이 보여야 된다
    xl.DisplayAlerts = False
    for answer, name in ((False, "practice.xlsx" if EN else "실습.xlsx"), (True, "answer.xlsx" if EN else "정답.xlsx")):
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count > 1:
            wb.Worksheets(wb.Worksheets.Count).Delete()
        build(wb, answer)
        author(wb)
        path = os.path.join(HERE, name)
        if os.path.exists(path):
            os.remove(path)
        wb.SaveAs(path, 51)
        print(name, [s.Name for s in wb.Worksheets], [(s.Name, s.ProtectContents) for s in wb.Worksheets])
        wb.Close(False)
        fix_core(path)
finally:
    xl.DisplayAlerts = False
    xl.Quit()
