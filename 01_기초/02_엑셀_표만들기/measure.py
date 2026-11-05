"""엑셀 기초 2강 — 표 만들기 실측. 이 PC 엑셀(Microsoft 365)을 숨긴 채 따로 띄워(DispatchEx)
테두리로 그린 표와 표 서식(Ctrl+T, ListObject)이 정렬·필터·합계·새 줄에서 어떻게 다른지 잰다.
결과 measure.json, 실습.xlsx · 정답.xlsx 를 이 폴더에 만든다. 데이터는 모두 연습용 가상 꽃집 주문 목록.
실행: python measure.py  (pywin32 필요)"""
import datetime, json, os
import pythoncom
import win32com.client as win32

from measure_util import HEAD, ROWS, NEW, fill

HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds")}


def col(ws, c, n=None):
    n = n or ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    return [ws.Range(f"{c}{i}").Value for i in range(2, n + 1)]


pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False; xl.DisplayAlerts = False
    R["build"] = xl.Build
    labels = {}
    for idm in ("BordersAll", "BordersGallery", "FormatAsTableGallery", "TableInsert", "TableInsertExcel", "TableStylesGallery",
                "TableTotalsRowShow", "TableTotalRowShow", "TableConvertToRange", "TableToolsConvertToRange", "ConvertTableToRange",
                "TableBandedRows", "TableHeaderRowShow", "TableFilterButton", "SortAscendingExcel", "SortDescendingExcel",
                "SortDialog", "Filter", "AutoFilterClassic", "FormatAsTable", "TableResize", "TableName", "TableToolsDesign"):
        try:
            labels[idm] = xl.CommandBars.GetLabelMso(idm).replace("​", "")
        except Exception:
            labels[idm] = None
    R["labels"] = labels
    R["autoexpand_setting"] = xl.AutoCorrect.AutoExpandListRange
    R["autofill_formulas_setting"] = xl.AutoCorrect.AutoFillFormulasInLists

    wb = xl.Workbooks.Add()
    # ── A. 테두리로 그린 표 ─────────────────────────────────────────
    a = wb.Worksheets(1); a.Name = "테두리"
    fill(a)
    a.Range("A1:F9").Borders.LineStyle = 1          # 모든 테두리
    a.Range("A1:F1").Font.Bold = True
    a.Range("H1").Value = "합계"; a.Range("I1").Formula = "=SUM(F2:F9)"
    before = a.Range("I1").Value
    a.Range("A10:E10").Value = (NEW,); a.Range("F10").Formula = "=D10*E10"   # 바로 아래 줄에 새 주문
    R["border"] = {
        "sum_before": before, "sum_after_new_row": a.Range("I1").Value, "sum_formula_after": a.Range("I1").Formula,
        "new_row_border_bottom": a.Range("A10").Borders(9).LineStyle,  # xlEdgeBottom — -4142 = 없음
        "new_row_border_left": a.Range("A10").Borders(7).LineStyle,
        "row9_border_bottom": a.Range("A9").Borders(9).LineStyle,
        "autofilter": a.AutoFilterMode,
    }
    # 정렬 함정: 상품 열(C)만 골라 정렬(정렬 경고에서 '현재 선택 영역으로 정렬'을 고른 것과 같다)
    s = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)); s.Name = "정렬함정"
    fill(s)
    pairs_before = list(zip(col(s, "B", 9), col(s, "C", 9)))
    s.Range("C2:C9").Sort(Key1=s.Range("C2"), Order1=1, Header=2)
    pairs_after = list(zip(col(s, "B", 9), col(s, "C", 9)))
    R["sort_trap"] = {
        "before": pairs_before, "after": pairs_after,
        "rows_mismatched": sum(1 for x, y in zip(pairs_before, pairs_after) if x != y and (x[0], y[1]) not in pairs_before),
        "moved_rows": sum(1 for x, y in zip(pairs_before, pairs_after) if x[1] != y[1]),
    }
    # 빈 줄이 끼면 Ctrl+T 가 고르는 범위(=현재 영역)가 거기서 끊긴다
    g = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)); g.Name = "빈줄"
    fill(g); g.Rows(6).Insert()
    R["blank_row_current_region"] = g.Range("B3").CurrentRegion.Address.replace("$", "")

    # ── B. 표 서식(Ctrl+T) ──────────────────────────────────────────
    t = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)); t.Name = "표"
    fill(t, with_amount=False)
    lo = t.ListObjects.Add(1, t.Range("A1").CurrentRegion, None, 1)   # xlSrcRange, 머리글 있음
    R["table"] = {"name": lo.Name, "style": lo.TableStyle.Name if lo.TableStyle else None, "style_local": lo.TableStyle.NameLocal if lo.TableStyle else None,
                  "range": lo.Range.Address.replace("$", ""), "show_autofilter": lo.ShowAutoFilter,
                  "banded_rows": lo.ShowTableStyleRowStripes, "header_row": lo.ShowHeaders,
                  "header_fill": t.Range("A1").DisplayFormat.Interior.Color, "row2_fill": t.Range("A2").DisplayFormat.Interior.Color,
                  "row3_fill": t.Range("A3").DisplayFormat.Interior.Color}
    # 계산 열: 금액 칸 하나에만 수식 → 열 전체에 저절로?
    t.Range("F2").Formula = "=[@수량]*[@단가]"
    R["table"]["calc_col_F3"] = t.Range("F3").Formula
    R["table"]["calc_col_F9_value"] = t.Range("F9").Value
    R["table"]["calc_col_formula_local"] = t.Range("F2").FormulaLocal
    # 바로 아래 줄에 값을 넣으면 표가 늘어나는지(COM 으로 값 넣기)
    t.Range("A10:E10").Value = (NEW,)
    R["table"]["range_after_new_row_com"] = lo.Range.Address.replace("$", "")
    R["table"]["F10_after_new_row_com"] = t.Range("F10").Formula
    if lo.Range.Rows.Count < 10:   # COM 값 넣기로 안 늘면 ListRows.Add 로 늘린 결과를 따로 적는다
        t.Range("A10:E10").ClearContents()
        nr = lo.ListRows.Add()
        nr.Range.Value = ((*NEW, None),)
        R["table"]["range_after_listrows_add"] = lo.Range.Address.replace("$", "")
        R["table"]["F10_after_listrows_add"] = t.Range("F10").Formula
    # 요약 행
    lo.ShowTotals = True
    tot = lo.TotalsRowRange
    R["table"]["totals_row"] = tot.Address.replace("$", "")
    R["table"]["totals_label"] = tot.Cells(1, 1).Value
    R["table"]["totals_formula"] = tot.Cells(1, 6).Formula
    R["table"]["totals_value"] = tot.Cells(1, 6).Value
    t.Range("H1").Value = "SUM"; t.Range("I1").Formula = "=SUM(표1[금액])"
    R["table"]["sum_structured"] = t.Range("I1").Value
    # 필터: 상품 = 장미 꽃다발
    lo.Range.AutoFilter(Field=3, Criteria1="장미 꽃다발")
    vis = [r for r in range(2, lo.DataBodyRange.Rows.Count + 2) if not t.Rows(r).Hidden]
    R["filter"] = {"visible_rows": len(vis), "totals_value": tot.Cells(1, 6).Value, "sum_outside": t.Range("I1").Value,
                   "subtotal_vs_sum": [tot.Cells(1, 6).Value, t.Range("I1").Value]}
    lo.Range.AutoFilter(Field=3)   # 필터 풀기
    # 표 정렬: 상품 오름차순 → 줄이 통째로 움직이는지
    lo.Sort.SortFields.Clear(); lo.Sort.SortFields.Add(Key=lo.ListColumns("상품").DataBodyRange, Order=1); lo.Sort.Header = 1; lo.Sort.Apply()
    n = lo.DataBodyRange.Rows.Count
    after = [(t.Range(f"B{i}").Value, t.Range(f"C{i}").Value, t.Range(f"F{i}").Value) for i in range(2, n + 2)]
    orig = {(r[1], r[2], r[3] * r[4]) for r in ROWS + [NEW]}
    R["table_sort"] = {"rows": after, "all_rows_intact": all(x in orig for x in after)}

    # 머리글 이름 — 빈 칸·같은 이름
    h = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)); h.Name = "머리글"
    h.Range("A1:D1").Value = (("상품", None, "금액", "금액"),)
    h.Range("A2:D3").Value = (("장미", 2, 1, 3), ("튤립", 3, 4, 5))
    lo2 = h.ListObjects.Add(1, h.Range("A1:D3"), None, 1)
    R["header_names"] = [h.Range(f"{c}1").Value for c in "ABCD"]
    # 머리글 없음으로 만들면
    h.Range("F2:G3").Value = (("장미", 2), ("튤립", 3))
    lo3 = h.ListObjects.Add(1, h.Range("F2:G3"), None, 2)   # xlNo
    R["no_header"] = {"range": lo3.Range.Address.replace("$", ""), "names": [h.Range("F2").Value, h.Range("G2").Value]}
    # 병합한 칸이 있으면
    m = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)); m.Name = "병합"
    fill(m); m.Range("A2:A3").Merge()
    try:
        m.ListObjects.Add(1, m.Range("A1:F9"), None, 1)
        R["merged"] = {"ok": True, "merged_after": m.Range("A2").MergeCells}
    except Exception as e:
        R["merged"] = {"ok": False, "error": str(e.excepinfo[2] if getattr(e, "excepinfo", None) else e)}
    # 범위로 변환: 서식·수식은?
    u = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count)); u.Name = "범위로변환"
    fill(u, with_amount=False)
    lo4 = u.ListObjects.Add(1, u.Range("A1").CurrentRegion, None, 1)
    u.Range("F2").Formula = "=[@수량]*[@단가]"
    lo4.ShowTotals = True
    fill_before = u.Range("A2").DisplayFormat.Interior.Color
    lo4.Unlist()
    R["unlist"] = {"tables_left": u.ListObjects.Count, "F2": u.Range("F2").Formula, "total_row": u.Range("F10").Formula,
                   "fill_kept": u.Range("A2").DisplayFormat.Interior.Color == fill_before, "autofilter": u.AutoFilterMode}
    wb.Close(False)

    # ── 실습·정답 파일 ──────────────────────────────────────────────
    p = xl.Workbooks.Add(); ws = p.Worksheets(1); ws.Name = "주문"
    fill(ws, with_amount=False)
    ws.Range("A1:F1").Font.Bold = True
    p.SaveAs(os.path.join(HERE, "실습.xlsx"), 51); p.Close(False)
    q = xl.Workbooks.Add(); ws = q.Worksheets(1); ws.Name = "주문"
    fill(ws, with_amount=False)
    lo = ws.ListObjects.Add(1, ws.Range("A1").CurrentRegion, None, 1); lo.Name = "주문표"
    ws.Range("F2").Formula = "=[@수량]*[@단가]"
    lo.ShowTotals = True
    q.SaveAs(os.path.join(HERE, "정답.xlsx"), 51)
    R["answer"] = {"name": lo.Name, "range": lo.Range.Address.replace("$", ""), "total": lo.TotalsRowRange.Cells(1, 6).Value,
                   "total_formula": lo.TotalsRowRange.Cells(1, 6).Formula}
    q.Close(False)
finally:
    xl.Quit()

json.dump(R, open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
