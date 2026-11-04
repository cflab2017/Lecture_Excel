"""엑셀 기초 1강 — 셀 안 줄바꿈 실측. 이 PC 엑셀(Microsoft 365)을 숨긴 채 따로 띄워(DispatchEx) 값·줄 높이·복사·CSV 를 잰다.
결과 measure.json, 실습.xlsx · 정답.xlsx 를 이 폴더에 만든다. 데이터는 모두 연습용 가상 회의실 예약표.
클립보드: 복사 시험 전에 글자 내용을 저장했다가 끝나면 되돌린다.
실행: python measure.py  (pywin32 필요)"""
import datetime, json, os
import pythoncom
import win32clipboard as wcb
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
ROWS = [  # 날짜, 회의실, 회의 내용, 준비물(실습 파일에서는 쉼표로 한 줄)
    ("11월 10일", "3층 대회의실", "분기 실적 보고", ["노트북", "빔프로젝터", "회의 자료 10부"]),
    ("11월 11일", "2층 소회의실", "신입 사원 교육", ["교육 교재", "이름표"]),
    ("11월 12일", "3층 대회의실", "거래처 미팅", ["견적서", "명함", "샘플 3종", "음료"]),
    ("11월 13일", "5층 세미나실", "팀 워크숍", ["화이트보드 펜", "포스트잇"]),
]
R = {"run": datetime.datetime.now().isoformat(timespec="seconds")}


def clip_get():
    wcb.OpenClipboard()
    try:
        return wcb.GetClipboardData(wcb.CF_UNICODETEXT) if wcb.IsClipboardFormatAvailable(wcb.CF_UNICODETEXT) else None
    finally:
        wcb.CloseClipboard()


def clip_set(s):
    wcb.OpenClipboard()
    try:
        wcb.EmptyClipboard(); wcb.SetClipboardData(wcb.CF_UNICODETEXT, s)
    finally:
        wcb.CloseClipboard()


def fill(ws, joiner):
    ws.Range("A1:D1").Value = (("날짜", "회의실", "회의 내용", "준비물"),)
    for i, (d, room, what, items) in enumerate(ROWS, start=2):
        ws.Range(f"A{i}:D{i}").Value = ((d, room, what, joiner.join(items)),)
    ws.Columns("A").ColumnWidth = 10; ws.Columns("B").ColumnWidth = 13; ws.Columns("C").ColumnWidth = 16; ws.Columns("D").ColumnWidth = 18
    ws.Range("A1:D1").Font.Bold = True


saved = clip_get()
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False; xl.DisplayAlerts = False
    R["build"] = xl.Build
    labels = {}
    for idm in ("WrapText", "AlignmentGroup", "GroupAlignmentExcel", "FindDialogExcel", "ReplaceDialog", "FormatCellsDialog", "RowHeight", "RowAutofitHeight"):
        try:
            labels[idm] = xl.CommandBars.GetLabelMso(idm).replace("​", "")
        except Exception:
            labels[idm] = None
    R["labels"] = labels
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1); ws.Name = "예약"
    fill(ws, ", ")
    R["row_h_default"] = ws.Rows(2).RowHeight
    # 1) Alt+Enter 와 같은 줄바꿈 문자(LF, 10)를 넣었을 때
    c = ws.Range("D2")
    c.Value = "\n".join(ROWS[0][3])
    R["lf"] = {"len": int(ws.Evaluate('=LEN(D2)')), "code_at_break": int(ws.Evaluate('=CODE(MID(D2,4,1))')),
               "wrap_after_value": c.WrapText, "row_h_after_value": ws.Rows(2).RowHeight}
    c.WrapText = True
    R["lf"]["row_h_after_wrap"] = ws.Rows(2).RowHeight
    # 2) 줄 높이를 직접 정해 둔 줄은 자동으로 안 늘어남
    c2 = ws.Range("D3"); ws.Rows(3).RowHeight = 16.5
    c2.Value = "\n".join(ROWS[1][3]); c2.WrapText = True
    R["manual"] = {"row_h_set": 16.5, "row_h_after_wrap": ws.Rows(3).RowHeight}
    ws.Rows(3).AutoFit()
    R["manual"]["row_h_after_autofit"] = ws.Rows(3).RowHeight
    # 3) 줄바꿈 없는 긴 글 + 자동 줄 바꿈: 열 너비에 따라 줄 수(=줄 높이)가 바뀜
    ws.Range("F1").Value = "다음 주 월요일 오전 10시까지 회의 자료를 팀 공유 폴더에 올려 주세요"
    ws.Range("F1").WrapText = True
    hs = {}
    for w in (12, 24, 48):
        ws.Columns("F").ColumnWidth = w; ws.Rows(1).AutoFit()
        hs[w] = ws.Rows(1).RowHeight
    R["wrap_long"] = hs
    ws.Range("F1").ClearContents(); ws.Columns("F").ColumnWidth = 8.43; ws.Rows(1).AutoFit()
    # 4) 수식으로 줄바꿈: CHAR(10)
    ws.Range("H2").Formula = '=A2&CHAR(10)&B2'
    R["formula"] = {"text": ws.Range("H2").Value, "wrap": ws.Range("H2").WrapText}
    ws.Range("I2").Formula2 = '=TEXTJOIN(CHAR(10),TRUE,"노트북","빔프로젝터")'
    R["textjoin"] = ws.Range("I2").Value
    # 5) 줄바꿈 없애기: SUBSTITUTE / 바꾸기(Ctrl+H 에서 Ctrl+J)
    ws.Range("J2").Formula = '=SUBSTITUTE(D2,CHAR(10),", ")'
    R["substitute"] = ws.Range("J2").Value
    ws.Range("K2").Value = "\n".join(ROWS[2][3]); ws.Range("K3").Value = "\n".join(ROWS[3][3])
    ok = ws.Range("K2:K3").Replace(What="\n", Replacement=" / ", LookAt=2)
    R["replace"] = {"ok": ok, "K2": ws.Range("K2").Value, "K3": ws.Range("K3").Value}
    # 6) 줄바꿈이 든 칸을 복사해 메모장 등에 붙이면 — 클립보드 글자
    ws.Range("D2").Copy()
    R["clip_one_cell"] = clip_get()
    ws.Range("C2:D2").Copy()
    R["clip_two_cells"] = clip_get()
    xl.CutCopyMode = False
    # 7) CSV 로 저장하면
    tmp = os.path.join(os.environ["TEMP"], "claude", "excel01"); os.makedirs(tmp, exist_ok=True)
    w2 = xl.Workbooks.Add(); s2 = w2.Worksheets(1)
    s2.Range("A1:B1").Value = (("회의실", "준비물"),); s2.Range("A2").Value = "3층 대회의실"; s2.Range("B2").Value = "노트북\n빔프로젝터"
    csv = os.path.join(tmp, "예약.csv")
    if os.path.exists(csv): os.remove(csv)
    w2.SaveAs(csv, 62)   # xlCSVUTF8
    w2.Close(False)
    raw = open(csv, "rb").read()
    R["csv"] = {"bytes": raw.decode("utf-8-sig"), "bom": raw.startswith(b"\xef\xbb\xbf")}
    wb.Close(False)
    # 실습.xlsx · 정답.xlsx
    for name, joiner, wrap in (("실습.xlsx", ", ", False), ("정답.xlsx", "\n", True)):
        w = xl.Workbooks.Add(); s = w.Worksheets(1); s.Name = "예약"
        fill(s, joiner)
        if wrap:
            s.Range("D2:D5").WrapText = True
        s.Range("A2:C5").VerticalAlignment = -4108 if wrap else -4107   # 가운데 / 아래
        s.Rows("1:5").AutoFit()
        R.setdefault("files", {})[name] = [s.Rows(i).RowHeight for i in range(1, 6)]
        p = os.path.join(HERE, name)
        if os.path.exists(p): os.remove(p)
        w.SaveAs(p, 51); w.Close(False)
finally:
    xl.Quit()
    if saved is not None:
        clip_set(saved)
R["clip_restored"] = saved is not None
json.dump(R, open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
