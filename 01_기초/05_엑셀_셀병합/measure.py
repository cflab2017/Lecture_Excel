"""엑셀 기초 5강 — 셀 병합 실측. 이 PC 엑셀(Microsoft 365)을 따로 띄워(DispatchEx) 병합 · 병합 해제 · 정렬 · 필터를 하고 칸 값을 읽는다.
- 병합: Range.Merge() (DisplayAlerts=False 면 묻지 않고 왼쪽 위 값만 남김) · 전체 병합: Range.Merge(True)
- 정렬: 크기가 다른 병합 칸이 섞인 범위를 Range.Sort → 엑셀이 내는 오류 문구를 그대로 기록
- 필터: 반(A열) 세로 병합 표에 AutoFilter '1반' → 보이는 행 수
- 선택 영역의 가운데로: HorizontalAlignment = 7(xlCenterAcrossSelection) — 칸은 따로, 보기만 가운데
- 병합 해제 뒤 빈칸 채우기: UnMerge → SpecialCells(빈 셀) → =위 칸 수식
결과 measure.json (argv 'en' 이면 measure_en.json). 실행: python measure.py [en]  (pywin32 필요)"""
import datetime, json, os, sys
import pythoncom
import pywintypes
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}
clean = lambda s: (s or "").replace("​", "").strip()

H = ("Class", "Name", "Score") if EN else ("반", "이름", "점수")
ROWS = [("1", "Kim", 82), ("", "Lee", 95), ("", "Park", 77), ("2", "Choi", 88), ("", "Jung", 91), ("", "Kang", 70)] if EN else \
       [("1반", "김민준", 82), ("", "이서연", 95), ("", "박지후", 77), ("2반", "최하은", 88), ("", "정도윤", 91), ("", "강수아", 70)]
if EN:
    ROWS = [("Class 1" if r[0] == "1" else ("Class 2" if r[0] == "2" else ""), r[1], r[2]) for r in ROWS]

pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False
    xl.DisplayAlerts = False
    R["build"] = xl.Build
    R["labels"] = {}
    for k in ("MergeCenterMenu", "MergeCenter", "MergeCellsAcross", "MergeCells", "UnmergeCells", "SortDialog", "Filter"):
        try:
            R["labels"][k] = clean(xl.CommandBars.GetLabelMso(k))
        except Exception as e:  # noqa: BLE001
            R["labels"][k] = "ERR " + str(e)[:60]
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)

    # ① 병합하면 왼쪽 위 값만 남는다
    ws.Range("A1:C1").Value = (("A", "B", "C"),)
    before = [ws.Range(a).Value for a in ("A1", "B1", "C1")]
    ws.Range("A1:C1").Merge()
    ws.Range("A1:C1").HorizontalAlignment = -4108  # 가운데
    after_merge = [ws.Range(a).Value for a in ("A1", "B1", "C1")]
    ws.Range("A1:C1").UnMerge()
    after_unmerge = [ws.Range(a).Value for a in ("A1", "B1", "C1")]
    R["merge_keeps_top_left"] = {"before": before, "after_merge": after_merge, "after_unmerge": after_unmerge}

    # ② 전체 병합(Merge Across): 줄마다 따로 합쳐진다
    ws.Range("E1:G2").Value = (("e1", "f1", "g1"), ("e2", "f2", "g2"))
    ws.Range("E1:G2").Merge(True)
    R["merge_across"] = {"E1_area": ws.Range("E1").MergeArea.Address, "E2_area": ws.Range("E2").MergeArea.Address,
                         "values": [ws.Range(a).Value for a in ("E1", "F1", "E2", "F2")]}

    # 표: 반(A열) 세로 병합
    ws2 = wb.Worksheets.Add()
    ws2.Range("A1:C1").Value = (H,)
    for i, r in enumerate(ROWS, start=2):
        ws2.Range(f"A{i}:C{i}").Value = (r,)
    for top, bot in ((2, 4), (5, 7)):
        ws2.Range(f"A{top}:A{bot}").Merge()
        ws2.Range(f"A{top}:A{bot}").VerticalAlignment = -4108
    R["table_merge_areas"] = [ws2.Range("A2").MergeArea.Address, ws2.Range("A5").MergeArea.Address]

    # ③ 정렬: 점수 내림차순 — 병합 칸(3칸)과 보통 칸(1칸)이 한 범위에
    try:
        ws2.Range("A1:C7").Sort(Key1=ws2.Range("C1"), Order1=2, Header=1)
        R["sort_with_merged"] = {"ok": True, "A": [ws2.Range(f"A{i}").Value for i in range(2, 8)]}
    except pywintypes.com_error as e:
        desc = e.excepinfo[2] if e.excepinfo else str(e)
        R["sort_with_merged"] = {"ok": False, "message": clean(desc)}
    # 정렬이 실패해 표는 그대로여야 한다 — 확인
    R["table_after_sort_try"] = [[ws2.Range(f"{c}{i}").Value for c in "ABC"] for i in range(2, 8)]

    # ④ 필터: 반 = 1반
    first = ROWS[0][0]
    ws2.Range("A1:C7").AutoFilter(Field=1, Criteria1=first)
    vis = [i for i in range(2, 8) if not ws2.Rows(i).Hidden]
    R["filter_merged"] = {"criteria": first, "visible_rows": vis, "visible_names": [ws2.Range(f"B{i}").Value for i in vis]}
    ws2.AutoFilterMode = False

    # ⑤ 병합 해제 → 빈칸 → 위 칸으로 채우기
    ws2.Range("A2:A7").UnMerge()
    after = [ws2.Range(f"A{i}").Value for i in range(2, 8)]
    blanks = ws2.Range("A2:A7").SpecialCells(4)  # xlCellTypeBlanks
    nblank = blanks.Count
    blanks.FormulaR1C1 = "=R[-1]C"
    filled = [ws2.Range(f"A{i}").Value for i in range(2, 8)]
    ws2.Range("A2:A7").Value = ws2.Range("A2:A7").Value  # 수식을 값으로
    R["unmerge_fill"] = {"after_unmerge": after, "blank_count": nblank, "filled": filled, "blank_address": blanks.Address}
    # 해제 뒤 필터와 정렬
    ws2.Range("A1:C7").AutoFilter(Field=1, Criteria1=first)
    vis = [i for i in range(2, 8) if not ws2.Rows(i).Hidden]
    R["filter_after_fill"] = {"visible_rows": vis, "visible_names": [ws2.Range(f"B{i}").Value for i in vis]}
    ws2.AutoFilterMode = False
    ws2.Range("A1:C7").Sort(Key1=ws2.Range("C1"), Order1=2, Header=1)
    R["sort_after_fill"] = [[ws2.Range(f"{c}{i}").Value for c in "ABC"] for i in range(2, 8)]

    # ⑥ 선택 영역의 가운데로: 제목 A1:C1 — 칸은 따로, 정렬·복사 그대로
    ws3 = wb.Worksheets.Add()
    ws3.Range("A1").Value = "Report" if EN else "성적표"
    ws3.Range("A1:C1").HorizontalAlignment = 7
    ws3.Range("A2:C3").Value = (("x", 2, 3), ("y", 1, 4))
    R["center_across"] = {"A1_merge": bool(ws3.Range("A1").MergeCells), "B1_value": ws3.Range("B1").Value,
                          "align_B1": ws3.Range("B1").HorizontalAlignment}
    try:
        ws3.Range("A1:C3").Sort(Key1=ws3.Range("B2"), Order1=1, Header=1)
        R["center_across"]["sort_ok"] = True
    except pywintypes.com_error as e:
        R["center_across"]["sort_ok"] = clean(e.excepinfo[2] if e.excepinfo else str(e))
    wb.Close(False)
finally:
    xl.Quit()
out = os.path.join(HERE, "measure_en.json" if EN else "measure.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)
print(json.dumps(R, ensure_ascii=False, indent=1, default=str))
