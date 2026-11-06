"""엑셀 기초 4강 — 자동 채우기 · 행 추가 실측. 이 PC 엑셀(Microsoft 365)을 따로 띄워(DispatchEx) 채우기 핸들을 끄는 것과 같은
Range.AutoFill(대상, 형식)을 부르고 칸에 나온 글자(.Text)를 읽는다. 행 추가는 Rows(..).Insert 뒤 합계 수식·값을 읽는다.
AutoFill 형식(XlAutoFillType): 0 기본(끌기 그대로) · 1 셀 복사 · 2 연속 데이터 채우기 · 6 평일 단위 · 7 월 단위.
결과 measure.json(argv 'en' 이면 영어 데이터로 measure_en.json) · 실습.xlsx · 정답.xlsx (영어: practice.xlsx · answer.xlsx).
실행: python measure.py [en]  (pywin32 필요)"""
import datetime, json, os, sys
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
from measure_util import fill, PEOPLE

EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}
clean = lambda s: (s or "").replace("​", "").strip()
NEW_ROW = (11, "2026-11-16", "Mon" if EN else "월", "Week 3" if EN else "3주차", "Korean" if EN else "국어", 20)

pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False
    xl.DisplayAlerts = False
    R["build"] = xl.Build
    R["labels"] = {}
    for k in ("FillMenu", "FillDown", "FillRight", "FillSeries", "SheetRowsInsert", "SheetRowsDelete", "CellsInsertDialog", "CellsInsertMenu"):
        try:
            R["labels"][k] = clean(xl.CommandBars.GetLabelMso(k))
        except Exception as e:  # noqa: BLE001
            R["labels"][k] = "ERR " + str(e)[:60]
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Test" if EN else "실험"
    ws.Range("A1:A6").Value = ((10,), (20,), (30,), (40,), (50,), (60,))
    ws.Range("B1").Value = 100
    col = [3]

    def fill_case(key, seeds, typ=0, n=6, fmt=None):
        c = col[0]
        col[0] += 1
        for k, v in enumerate(seeds):
            cell = ws.Cells(1 + k, c)
            if isinstance(v, str) and v.startswith("="):
                cell.Formula = v
            else:
                cell.Value = v
            if fmt:
                cell.NumberFormat = fmt
        src = ws.Range(ws.Cells(1, c), ws.Cells(len(seeds), c))
        dst = ws.Range(ws.Cells(1, c), ws.Cells(n, c))
        rec = {"seeds": [str(s) for s in seeds], "type": typ}
        try:
            src.AutoFill(dst, typ)
            rec["texts"] = [ws.Cells(r, c).Text for r in range(1, n + 1)]
            if any(isinstance(s, str) and s.startswith("=") for s in seeds):
                rec["formulas"] = [ws.Cells(r, c).Formula for r in range(1, n + 1)]
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)[:160]
        R.setdefault("fill", {})[key] = rec

    cases = [
        ("num1_default", [1], 0), ("num1_series", [1], 2), ("num12_default", [1, 2], 0), ("num2_4_default", [2, 4], 0),
        ("date_default", ["2026-11-06"], 0, "yyyy-mm-dd"), ("date_weekdays", ["2026-11-06"], 6, "yyyy-mm-dd"),
        ("date_months", ["2026-11-06"], 7, "yyyy-mm-dd"), ("text_plain", ["Math" if EN else "수학"], 0),
        ("formula_rel", ["=A1*2"], 0), ("formula_abs", ["=A1*$B$1"], 0),
    ]
    if EN:
        cases += [("day_short", ["Mon"], 0), ("day_long", ["Monday"], 0), ("month_short", ["Jan"], 0),
                  ("week_text", ["Week 1"], 0), ("room_text", ["Room 1"], 0), ("other_lang_day", ["월"], 0)]
    else:
        cases += [("day_short", ["월"], 0), ("day_long", ["월요일"], 0), ("month_ko", ["1월"], 0),
                  ("week_text", ["1주차"], 0), ("room_text", ["1반"], 0), ("other_lang_day", ["Mon"], 0)]
    for cse in cases:
        fill_case(cse[0], cse[1], cse[2], fmt=cse[3] if len(cse) > 3 else None)
    # Ctrl+D(아래로 채우기) = Range.FillDown — 맨 위 칸을 아래 칸들에 복사
    c = col[0]
    ws.Cells(1, c).Formula = "=A1+1"
    ws.Range(ws.Cells(1, c), ws.Cells(4, c)).FillDown()
    R["fill_down"] = {"texts": [ws.Cells(r, c).Text for r in range(1, 5)], "formulas": [ws.Cells(r, c).Formula for r in range(1, 5)]}

    # ── 행 추가 ──
    def table(name):
        s = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        s.Name = name
        return s, fill(s, EN)

    total = sum(PEOPLE)
    s1, t = table("RowsA" if EN else "행A")          # ① 합계 행(12행) 바로 위에 새 행
    s1.Rows(t).Insert()
    s1.Range(f"A{t}:F{t}").Value = (NEW_ROW,)
    R["insert_above_total"] = {"inserted_at": t, "total_row_after": t + 1, "formula": s1.Range(f"F{t+1}").Formula,
                               "value": s1.Range(f"F{t+1}").Text, "sum_before": total, "if_included": total + 20}
    # 엑셀 오류 검사(초록 삼각형) — Errors(5) = xlOmittedCells '수식이 인접한 셀을 생략'
    try:
        R["insert_above_total"]["omitted_cells_flag"] = bool(s1.Range(f"F{t+1}").Errors(5).Value)
    except Exception as e:  # noqa: BLE001
        R["insert_above_total"]["omitted_cells_flag"] = "ERR " + str(e)[:80]
    s2, t = table("RowsB" if EN else "행B")          # ② 범위 안(11행 = 마지막 데이터 행)에 새 행
    s2.Rows(t - 1).Insert()
    s2.Range(f"E{t-1}:F{t-1}").Value = (("Korean" if EN else "국어", 20),)
    R["insert_inside"] = {"inserted_at": t - 1, "total_row_after": t + 1, "formula": s2.Range(f"F{t+1}").Formula, "value": s2.Range(f"F{t+1}").Text}
    s3, t = table("RowsC" if EN else "행C")          # ③ 5:7 세 줄을 골라 한꺼번에
    s3.Rows("5:7").Insert()
    R["insert_three"] = {"A_col": [s3.Range(f"A{r}").Text for r in range(2, 15)], "total_formula": s3.Range(f"F{t+3}").Formula}
    s3.Range("A2:A3").AutoFill(s3.Range("A2:A14"), 0)  # 번호 다시: 1·2 를 골라 끝까지
    R["insert_three"]["A_after_refill"] = [s3.Range(f"A{r}").Text for r in range(2, 15)]

    # 날짜를 평일 단위로 채우면 따로 채운 요일 칸과 어긋남 → 요일은 날짜에서 TEXT 로
    s4, t = table("Weekday" if EN else "요일")
    s4.Range("B2").AutoFill(s4.Range(f"B2:B{t-1}"), 6)          # 날짜: 평일 단위
    s4.Range("C2").AutoFill(s4.Range(f"C2:C{t-1}"), 0)          # 요일: 따로 끌기
    R["weekday_mismatch"] = {"dates": [s4.Range(f"B{r}").Text for r in range(2, t)], "days_dragged": [s4.Range(f"C{r}").Text for r in range(2, t)]}
    fmt = "ddd" if EN else "aaa"
    s4.Range("C2").Formula = f'=TEXT(B2,"{fmt}")'
    s4.Range("C2").AutoFill(s4.Range(f"C2:C{t-1}"), 0)
    R["weekday_mismatch"]["days_text_formula"] = [s4.Range(f"C{r}").Text for r in range(2, t)]
    R["weekday_mismatch"]["formula"] = s4.Range("C2").Formula
    # 주차: '1주차' 를 끌면 줄마다 하나씩 늘어 날짜와 안 맞음(한 주 = 5줄)
    s4.Range("D2").AutoFill(s4.Range(f"D2:D{t-1}"), 0)
    R["week_dragged"] = [s4.Range(f"D{r}").Text for r in range(2, t)]

    # ── 실습·정답 파일 ──
    names = ("practice.xlsx", "answer.xlsx") if EN else ("실습.xlsx", "정답.xlsx")
    for kind, nm in zip(("practice", "answer"), names):
        w2 = xl.Workbooks.Add()
        s = w2.Worksheets(1)
        s.Name = "Classes" if EN else "수업표"
        t = fill(s, EN)
        if kind == "practice":
            s.Range(f"A3:D{t-1}").ClearContents()        # 번호·날짜·요일·주차 첫 줄만 두고 비움 → 채우기 연습
        else:
            s.Rows(t).Insert()
            s.Range(f"A{t}:F{t}").Value = (NEW_ROW,)
            s.Range(f"B{t}").NumberFormat = "yyyy-mm-dd"
            s.Range(f"F{t+1}").Formula = f"=SUM(F2:F{t})"   # 정답: 합계 범위를 새 행까지
        for p in ("Author", "Last Author"):
            w2.BuiltinDocumentProperties(p).Value = "CODINGNOW"
        w2.SaveAs(os.path.join(HERE, nm), 51)
        w2.Close(False)
    wb.Close(False)
finally:
    xl.Quit()

json.dump(R, open(os.path.join(HERE, "measure_en.json" if EN else "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k, v in R.get("fill", {}).items():
    print(k, v.get("texts") or v.get("error"), v.get("formulas") or "")
for k in ("fill_down", "insert_above_total", "insert_inside", "insert_three", "weekday_mismatch", "week_dragged", "labels"):
    print(k, json.dumps(R.get(k), ensure_ascii=False)[:400])
