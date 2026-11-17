"""업무 자동화 1편 — 엑셀 파일 24개 한 장으로 합치기: 엑셀 파워 쿼리 '폴더에서' 실측(가이드 /guides/combine-excel-files).

예제(가상 회사 「예시문구」 6지점 × 7~10월, 지점마다 모양이 다름)를 임시 폴더에 풀고, 숨긴 엑셀(DispatchEx · Visible=False)에서
파워 쿼리 세 개를 만들어 시트에 불러온 뒤 줄 수 · 금액 합계 · 지점별 월 합계를 읽어 검산 표(검산_지점월별_정답.csv)와 맞춘다.
  Q1 첫파일기준 — 폴더의 첫 파일 · 첫 시트를 머리글 기준으로 모든 파일을 이어 붙임(예시 파일 하나를 기준으로 합치는 방식)
  Q2 합계만뺌   — 고친 쿼리에서 '합계' 줄만 빼고 '총계'는 남김(부산 함정)
  Q3 고친쿼리   — 파일마다 '매출' 시트(없으면 첫 시트) · '품목' 이 든 줄을 머리글로 · 열 이름 맞춤 · 합계 · 총계 줄 뺌 · 글자 숫자를 숫자로
더 재는 것: 인천 원본의 글자 숫자에 SUM 을 걸면 얼마가 나오나 · 표 이름으로 쓴 SUMIFS 검산식(부산 7월) ·
폴더에 11월 파일 하나를 더하고 Q3 를 새로 고치면 따라오나(서울 7월 파일을 '서울_11월_매출.xlsx' 로 복사해 넣음).
그림은 불러온 범위를 PDF 로 내보내 만든다(창 · 화면 캡처 안 씀). 결과: pq_measure.json · shots/*.png(저장소에 안 올림)
실행: python pq_measure.py   (Windows + 엑셀 · pywin32 · pymupdf · pillow)"""
import csv
import datetime
import json
import os
import shutil
import sys
import tempfile
import zipfile

import fitz
import pythoncom
import pywintypes
import win32com.client as win32
from PIL import Image

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
os.makedirs(SHOTS, exist_ok=True)
ZIP = os.path.join(HERE, "예제_지점별_매출_24개.zip")
ANS = os.path.join(HERE, "검산_지점월별_정답.csv")
TMP = tempfile.mkdtemp(prefix="xl_merge_")
INPUT = os.path.join(TMP, "input")

# zip 풀기(한글 이름 — UTF-8 표시가 없으면 cp949 로)
with zipfile.ZipFile(ZIP) as z:
    for info in z.infolist():
        n = info.filename
        if info.flag_bits & 0x800 == 0:
            try:
                n = n.encode("cp437").decode("cp949")
            except Exception:
                pass
        if n.endswith("/"):
            continue
        dst = os.path.join(TMP, n)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        with z.open(info) as f, open(dst, "wb") as g:
            g.write(f.read())
assert os.path.isdir(INPUT), os.listdir(TMP)

M_FIRST = '''let
    원본 = Folder.Files("{folder}"),
    엑셀만 = Table.SelectRows(원본, each [Extension] = ".xlsx"),
    시트붙임 = Table.AddColumn(엑셀만, "자료", each Table.PromoteHeaders(Excel.Workbook([Content]){{0}}[Data], [PromoteAllScalars = true])),
    기준열 = Table.ColumnNames(시트붙임{{0}}[자료]),
    펼침 = Table.ExpandTableColumn(Table.SelectColumns(시트붙임, {{"Name", "자료"}}), "자료", 기준열)
in
    펼침'''

M_FIXED = '''let
    폴더 = "{folder}",
    원본 = Folder.Files(폴더),
    엑셀만 = Table.SelectRows(원본, each [Extension] = ".xlsx"),
    숫자로 = (v) => if v = null then null else Number.From(Text.Remove(Text.From(v), {{",", " "}})),
    한파일 = (내용 as binary, 이름 as text) as table =>
        let
            책 = Excel.Workbook(내용, null, true),
            시트들 = Table.SelectRows(책, each [Kind] = "Sheet"),
            시트 = if Table.RowCount(시트들) > 1 and List.Contains(시트들[Name], "매출") then 시트들{{[Name = "매출"]}}[Data] else 시트들{{0}}[Data],
            줄들 = Table.ToRows(시트),
            머리줄 = List.PositionOf(List.Transform(줄들, each List.Contains(_, "품목") or List.Contains(_, "품목명")), true),
            머리글 = List.Transform(List.Positions(줄들{{머리줄}}), (i) => let h = 줄들{{머리줄}}{{i}} in if h = null then "빈열" & Text.From(i) else Text.Trim(Text.From(h))),
            몸통 = Table.FromRows(List.Skip(줄들, 머리줄 + 1), 머리글),
            이름맞춤 = Table.RenameColumns(몸통, {{{{"품목명", "품목"}}, {{"판매수량", "수량"}}, {{"매출액", "금액"}}}}, MissingField.Ignore),
            네열 = Table.SelectColumns(이름맞춤, {{"품목", "수량", "단가", "금액"}}),
            합계뺌 = Table.SelectRows(네열, each [품목] <> null and not List.Contains({excl}, Text.Trim(Text.From([품목])))),
            숫자 = Table.TransformColumns(합계뺌, {{{{"수량", 숫자로, type number}}, {{"단가", 숫자로, type number}}, {{"금액", 숫자로, type number}}}}),
            지점 = List.First(List.Select({{"서울", "부산", "대구", "인천", "광주", "대전"}}, (b) => Text.Contains(이름, b))),
            월 = Number.From(Text.Select(Text.Replace(이름, "2026", ""), {{"0".."9"}})),
            붙임 = Table.AddColumn(Table.AddColumn(숫자, "지점", each 지점, type text), "월", each 월, Int64.Type)
        in
            Table.SelectColumns(붙임, {{"지점", "월", "품목", "수량", "단가", "금액"}}),
    합침 = Table.Combine(Table.AddColumn(엑셀만, "표", each 한파일([Content], [Name]))[표])
in
    합침'''

QUERIES = [
    ("Q1_첫파일기준", M_FIRST.format(folder=INPUT.replace('"', '""'))),
    ("Q2_합계만뺌", M_FIXED.format(folder=INPUT.replace('"', '""'), excl='{"합계"}')),
    ("Q3_고친쿼리", M_FIXED.format(folder=INPUT.replace('"', '""'), excl='{"합계", "총계"}')),
]


def load_query(wb, name, formula):
    wb.Queries.Add(name, formula)
    ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name[:30]
    conn = f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="{name}";Extended Properties=""'
    lo = ws.ListObjects.Add(0, conn, None, 1, ws.Range("A1"))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = [f"SELECT * FROM [{name}]"]
    qt.RefreshStyle = 1
    qt.Refresh(False)
    return ws, lo


def read_table(lo):
    head = [c.Value for c in lo.HeaderRowRange]
    body = lo.DataBodyRange
    rows = [] if body is None else [list(r) for r in body.Value2]
    return head, rows


def export_png(ws, rng, name):
    ps = ws.PageSetup
    ps.PrintArea = rng
    ps.PrintHeadings = True
    ps.PrintGridlines = True
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    pdf = os.path.join(TMP, name + ".pdf")
    ws.ExportAsFixedFormat(0, pdf)
    doc = fitz.open(pdf)
    pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    # 흰 여백 자르기
    gray = im.convert("L").point(lambda x: 0 if x > 245 else 255)
    box = gray.getbbox()
    if box:
        im = im.crop((max(0, box[0] - 10), max(0, box[1] - 10), min(im.width, box[2] + 10), min(im.height, box[3] + 10)))
    out = os.path.join(SHOTS, name + ".png")
    im.save(out)
    doc.close()
    return {"png": "shots/" + name + ".png", "size": im.size}


ans = {}
with open(ANS, encoding="utf-8-sig") as f:
    for r in csv.DictReader(f):
        if r["지점"] == "합계":
            continue
        ans[r["지점"]] = {int(k.replace("월", "")): float(v) for k, v in r.items() if k.endswith("월")}

R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "queries": {}, "m_code_fixed": M_FIXED.format(folder="C:\\예제\\input", excl='{"합계", "총계"}')}
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    R["excel"] = {"version": xl.Version, "build": xl.Build}
    # 글자 숫자 실측 — 인천 7월 원본의 금액 칸(D4:D15)에 SUM 을 걸면(저장하지 않고 닫음)
    wi = xl.Workbooks.Open(os.path.join(INPUT, "인천(7월)매출.xlsx"), ReadOnly=True)
    wsi = wi.Worksheets(1)
    wsi.Range("F1").Formula = "=SUM(D4:D15)"
    wsi.Range("F2").Formula = "=COUNT(D4:D15)"
    R["incheon_text_numbers"] = {"D4": wsi.Range("D4").Value, "D4_type": type(wsi.Range("D4").Value).__name__,
                                 "SUM(D4:D15)": wsi.Range("F1").Value, "COUNT(D4:D15)": wsi.Range("F2").Value}
    wi.Close(False)
    wb = xl.Workbooks.Add()
    loaded = {}
    for name, formula in QUERIES:
        item = {}
        try:
            ws, lo = load_query(wb, name, formula)
            loaded[name] = (ws, lo)
            item["table_name"] = lo.Name
            head, rows = read_table(lo)
            item["columns"] = head
            item["rows"] = len(rows)
            item["first3"] = rows[:3]
            if "금액" in head:
                ci = head.index("금액")
                vals = [r[ci] for r in rows]
                item["amount_sum"] = sum(v for v in vals if isinstance(v, (int, float)))
                item["amount_text_cells"] = sum(1 for v in vals if isinstance(v, str))
            if "지점" in head and "월" in head and "금액" in head:
                bi, mi, ci = head.index("지점"), head.index("월"), head.index("금액")
                by = {}
                for r in rows:
                    if isinstance(r[ci], (int, float)):
                        by.setdefault(r[bi], {}).setdefault(int(r[mi]), 0.0)
                        by[r[bi]][int(r[mi])] += r[ci]
                item["by_branch_month"] = by
                item["diff_vs_answer"] = [[b, m, by.get(b, {}).get(m), ans[b][m]] for b in ans for m in ans[b] if abs(by.get(b, {}).get(m, -1) - ans[b][m]) > 0.5]
            last_col = chr(ord("A") + max(0, len(head) - 1))
            item["shot"] = export_png(ws, f"A1:{last_col}{min(13, len(rows) + 1)}", name)
            if "지점" in head:
                # 검산식 — 표 이름으로 부산 7월 금액 합계(정답 1,741,700)
                t = lo.Name
                f = f'=SUMIFS({t}[금액],{t}[지점],"부산",{t}[월],7)'
                ws.Range("H2").Formula = f
                item["sumifs_busan_7"] = {"formula": f, "value": ws.Range("H2").Value}
                ws.Range("H2").ClearContents()
        except pywintypes.com_error as e:
            item["error"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:300]
        R["queries"][name] = item
    # 새로 고침 실측 — 폴더에 11월 파일 하나를 더하고(서울 7월 파일을 '서울_11월_매출.xlsx' 로 복사) Q3 를 새로 고침
    if "Q3_고친쿼리" in loaded:
        ws3, lo3 = loaded["Q3_고친쿼리"]
        extra = os.path.join(INPUT, "서울_11월_매출.xlsx")
        shutil.copy(os.path.join(INPUT, "서울_7월_매출.xlsx"), extra)
        try:
            lo3.QueryTable.Refresh(False)
            head, rows = read_table(lo3)
            bi, mi, ci = head.index("지점"), head.index("월"), head.index("금액")
            R["refresh_test"] = {
                "added": os.path.basename(extra),
                "rows": len(rows),
                "amount_sum": sum(r[ci] for r in rows if isinstance(r[ci], (int, float))),
                "seoul_11": sum(r[ci] for r in rows if r[bi] == "서울" and r[mi] == 11),
            }
        except pywintypes.com_error as e:
            R["refresh_test"] = {"error": (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:300]}
    wb.Close(False)
finally:
    xl.Quit()
    shutil.rmtree(TMP, ignore_errors=True)

with open(os.path.join(HERE, "pq_measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)
for k, v in R["queries"].items():
    print(k, {x: v.get(x) for x in ("error", "table_name", "columns", "rows", "amount_sum", "amount_text_cells", "diff_vs_answer", "sumifs_busan_7")})
print("인천 글자 숫자", R.get("incheon_text_numbers"))
print("새로 고침", R.get("refresh_test"))
