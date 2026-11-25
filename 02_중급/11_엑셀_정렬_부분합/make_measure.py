"""엑셀 중급 11강 — 정렬 · 부분합(컴활 2급 실기 분석작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님). 이름 · 숫자는 지어낸 것.

하는 일
1. 실습.xlsx(문제 시트 + 분석작업-31 · 32 · 33 시트, 표만 있음 — 정렬 안 된 상태) · 정답.xlsx(정렬 + 부분합까지) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 정렬(Sort) · 부분합(Subtotal)을 실행해 결과 행(글자 · 수식 · 개요 수준)을 읽는다
3. 자주 틀리는 것(정렬 안 함 · 새로운 값으로 대치 · 순서 거꾸로 · 가나다순 · 모두 제거)도 같은 자리에서 — 저장하지 않는 사본에서
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 '정렬' · '부분합' 대화 상자 화면은 이 스크립트로 만들지 않는다.
결과: 실습.xlsx · 정답.xlsx · measure.json · shots/*.png (shots 는 저장소에 올리지 않음)
실행: python make_measure.py   (pywin32 · pymupdf · pillow)"""
import datetime
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

import fitz
import pythoncom
import pywintypes
import win32com.client as win32
import win32gui
from PIL import Image, ImageDraw

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
TMP = tempfile.mkdtemp(prefix="xl_mid11_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 숫자는 모두 가상입니다."
F = {"합계": -4157, "평균": -4106, "최대값": -4136, "개수": -4112}  # xlSum · xlAverage · xlMax · xlCount
ASC, DESC = 1, 2

# 표는 A2(머리글) ~ D12(데이터 10줄) — 정렬 안 된 차례. 부분합은 그 자리에서(줄이 끼어듦)
SALES = [["서울", "김하늘", "노트", 420000], ["부산", "이바다", "볼펜", 180000], ["대구", "박들판", "노트", 260000], ["서울", "최구름", "파일", 310000],
         ["부산", "정파도", "노트", 350000], ["대구", "한솔", "볼펜", 150000], ["서울", "윤별", "볼펜", 220000], ["부산", "강모래", "파일", 270000],
         ["대구", "오햇살", "파일", 190000], ["서울", "임가람", "노트", 380000]]
T31 = dict(
    key="q31", sheet="분석작업-31", level="기본", title="[표31] 지점별 판매 실적", head=["지점", "담당자", "제품", "판매액"], rows=SALES, money=[3],
    sort=[(1, ASC)], subs=[(1, "합계", [4], True)],
    text="[표31]에서 부분합 기능을 이용하여 '지점'별 '판매액'의 합계를 계산하시오. ▶ 정렬은 '지점'을 기준으로 오름차순으로 처리하시오.",
    use="데이터 → 정렬(지점 오름차순) → 데이터 → 부분합(그룹화할 항목 지점 · 합계 · 판매액)",
)
PAY = [["가온", "영업부", 3200000, 150000], ["나린", "총무부", 2900000, 120000], ["다인", "기획부", 3500000, 200000], ["라온", "영업부", 2800000, 180000],
       ["미르", "기획부", 3100000, 100000], ["바다", "총무부", 3300000, 160000], ["새벽", "영업부", 3600000, 220000], ["아름", "기획부", 2700000, 140000],
       ["온새", "총무부", 3000000, 90000], ["하랑", "영업부", 3000000, 130000]]
T32 = dict(
    key="q32", sheet="분석작업-32", level="응용", title="[표32] 부서별 급여", head=["이름", "부서", "기본급", "수당"], rows=PAY, money=[2, 3],
    sort=[(2, DESC)], subs=[(2, "최대값", [3], True), (2, "평균", [4], False)],
    text="[표32]에서 부분합 기능을 이용하여 '부서'별 '기본급'의 최대값을 계산한 후 '수당'의 평균을 계산하시오. ▶ 정렬은 '부서'를 기준으로 내림차순으로 처리하시오. ▶ 최대값과 평균은 위에 표시된 순서대로 처리하시오.",
    use="정렬(부서 내림차순) → 부분합 ① 최대값 · 기본급 → 부분합 ② 평균 · 수당, '새로운 값으로 대치' 해제",
)
EDU = [["강이슬", "대리", "영업부", 12], ["고은비", "사원", "총무부", 8], ["김도윤", "부장", "기획부", 4], ["노을", "과장", "영업부", 10],
       ["문지호", "사원", "기획부", 16], ["배서연", "대리", "총무부", 6], ["송하준", "과장", "기획부", 14], ["유채원", "사원", "영업부", 20],
       ["장민재", "부장", "총무부", 2], ["허윤아", "대리", "기획부", 9]]
ORDER = "부장,과장,대리,사원"
T33 = dict(
    key="q33", sheet="분석작업-33", level="함정", title="[표33] 직원 교육 이수", head=["이름", "직위", "부서", "교육 시간"], rows=EDU, money=[],
    sort=[(2, ORDER)], subs=[(2, "개수", [1], True)],
    text="[표33]에서 부분합 기능을 이용하여 '직위'별 '이름'의 개수를 계산하시오. ▶ 정렬은 '직위'를 '부장, 과장, 대리, 사원' 순서로 처리하시오.",
    use="정렬 → 정렬 기준 직위 · 정렬 '사용자 지정 목록'에 부장, 과장, 대리, 사원 → 부분합(직위 · 개수 · 이름)",
)
TABLES = [T31, T32, T33]
TRAPS = [
    dict(key="q31_nosort", table=T31, sort=[], why="정렬하지 않고 부분합 — 같은 지점이 여러 묶음으로 나뉨"),
    dict(key="q31_remove", table=T31, remove=True, why="부분합 → 모두 제거 — 원래 10줄로 돌아오나"),
    dict(key="q32_first", table=T32, subs=[(2, "최대값", [3], True)], why="(중간 상태) 첫 번째 부분합(최대값)만 한 모습 — 애니메이션용"),
    dict(key="q32_replace", table=T32, subs=[(2, "최대값", [3], True), (2, "평균", [4], True)], why="두 번째 부분합에서 '새로운 값으로 대치'를 그대로 둠 — 최대값 줄이 사라짐"),
    dict(key="q32_reverse", table=T32, subs=[(2, "평균", [4], True), (2, "최대값", [3], False)], why="평균을 먼저, 최대값을 나중에 — 줄 순서가 바뀜"),
    dict(key="q33_abc", table=T33, sort=[(2, ASC)], why="직위를 그냥 오름차순(가나다) — 과장 · 대리 · 부장 · 사원"),
]
WIDTHS = dict(zip("ABCD", (14, 12, 12, 13)))
FIRST_N = 31
LESSON = "엑셀 중급 11강 — 정렬 · 부분합 (컴활 2급 실기 분석작업 유형 연습)"
LESSON_NO = "엑셀 중급 11강"


def author(wb):
    for k in ("Author", "Last Author"):
        wb.BuiltinDocumentProperties(k).Value = "CODINGNOW"


def fix_core(path):
    """저장하면 엑셀이 '마지막 수정한 사람'을 이 PC 사용자 이름으로 다시 쓴다 → docProps/core.xml 을 직접 CODINGNOW 로."""
    tmp = path + ".tmp"
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


def core_people(path):
    with zipfile.ZipFile(path) as z:
        s = z.read("docProps/core.xml").decode("utf-8")
    return {tag: (re.search(f"<{tag}>([^<]*)</{tag}>", s) or [None, None])[1] for tag in ("dc:creator", "cp:lastModifiedBy")}


def last_row(ws):
    r = 2
    while ws.Range(f"A{r + 1}").Text != "" or ws.Range(f"B{r + 1}").Text != "" or ws.Range(f"D{r + 1}").Text != "":
        r += 1
    return r


def do_sort(ws, t, keys):
    n = 2 + len(t["rows"])
    s = ws.Sort
    s.SortFields.Clear()
    for col, how in keys:
        key = ws.Range(f"{chr(64 + col)}3:{chr(64 + col)}{n}")
        if isinstance(how, str):
            s.SortFields.Add(Key=key, SortOn=0, Order=1, CustomOrder=how, DataOption=0)
        else:
            s.SortFields.Add(Key=key, SortOn=0, Order=how, DataOption=0)
    s.SetRange(ws.Range(f"A2:D{n}"))
    s.Header = 1  # xlYes
    s.MatchCase = False
    s.Orientation = 1
    s.Apply()


def do_subs(ws, subs):
    for col, fn, total, replace in subs:
        rng = ws.Range(f"A2:D{last_row(ws)}")
        rng.Subtotal(col, F[fn], tuple(total), replace, False, 1)  # PageBreaks=False · SummaryBelowData=xlSummaryBelow


def read_rows(ws):
    out = []
    for r in range(2, last_row(ws) + 1):
        row = ws.Rows(r)
        out.append({"r": r, "text": [ws.Range(f"{chr(65 + j)}{r}").Text for j in range(4)],
                    "formula": [ws.Range(f"{chr(65 + j)}{r}").Formula for j in range(4) if str(ws.Range(f"{chr(65 + j)}{r}").Formula).startswith("=")],
                    "level": row.OutlineLevel})
    return out


def summary(rows):
    """요약 줄(수식이 있는 줄)만 — 글자 · 수식"""
    return [{"r": x["r"], "label": next((c for c in x["text"] if c and not c.replace(",", "").isdigit()), ""), "text": x["text"], "formula": x["formula"], "level": x["level"]}
            for x in rows if x["formula"]]


def build_sheet(ws, t, with_answer):
    ws.Range("A1").Value = t["title"]
    ws.Range("A1").Font.Bold = True
    hdr = ws.Range("A2:D2")
    hdr.Value = (tuple(t["head"]),)
    for i, row in enumerate(t["rows"]):
        for j, v in enumerate(row):
            ws.Range(f"{chr(65 + j)}{3 + i}").Value2 = v
    last = 2 + len(t["rows"])
    ws.Range(f"A2:D{last}").Borders.LineStyle = 1
    ws.Range(f"A3:C{last}").HorizontalAlignment = -4108
    for j in t["money"]:
        ws.Range(f"{chr(65 + j)}3:{chr(65 + j)}{last}").NumberFormat = "#,##0"
    hdr.Font.Bold = True
    hdr.Interior.Color = 0xEADBC8
    hdr.HorizontalAlignment = -4108
    for col, w in WIDTHS.items():
        ws.Columns(col).ColumnWidth = w
    if with_answer:
        do_sort(ws, t, t["sort"])
        do_subs(ws, t["subs"])


def build_problem_sheet(ws):
    ws.Range("A1").Value = LESSON
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = NOTICE
    ws.Range("A2").Font.Color = 0x3030C0
    line = 4
    for n, t in enumerate(TABLES, start=FIRST_N):
        ws.Range(f"A{line}").Value = f"{n}. ({t['level']}) '{t['sheet']}' 시트 — {t['text']}"
        ws.Range(f"A{line + 1}").Value = f"   ▶ {t['use']}"
        line += 3
    ws.Range(f"A{line}").Value = f"풀이 · 실측 값 · 자주 틀리는 곳: https://www.coding-now.com/blog/excel ({LESSON_NO})"
    ws.Columns("A").ColumnWidth = 120


# ── 그림: 범위 → PDF → PNG (8 · 9 · 10강과 같은 방식) ───────────
def export(ws, rng, name, boxes=()):
    fg_before = win32gui.GetForegroundWindow()
    ps = ws.PageSetup
    ps.PrintArea = rng
    ps.PrintHeadings = True
    ps.PrintGridlines = True
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin", "HeaderMargin", "FooterMargin"):
        setattr(ps, m, 14)
    pdf = os.path.join(TMP, name + ".pdf")
    ws.ExportAsFixedFormat(0, pdf)
    fg_after = win32gui.GetForegroundWindow()
    doc = fitz.open(pdf)
    page = doc[0]
    xs, ys = set(), set()
    for dr in page.get_drawings():
        for it in dr["items"]:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.x - p2.x) < 0.5:
                    xs.add(round(p1.x, 1))
                elif abs(p1.y - p2.y) < 0.5:
                    ys.add(round(p1.y, 1))
            elif it[0] == "re":
                r = it[1]
                xs.update({round(r.x0, 1), round(r.x1, 1)})
                ys.update({round(r.y0, 1), round(r.y1, 1)})
    xs, ys = merge(sorted(xs)), merge(sorted(ys))
    z = 2.0
    pix = page.get_pixmap(matrix=fitz.Matrix(z, z))
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    a0 = ws.Range(rng).Cells(1, 1)
    ncol, nrow = ws.Range(rng).Columns.Count, ws.Range(rng).Rows.Count
    if len(xs) < ncol + 2 or len(ys) < nrow + 2:
        raise RuntimeError(f"{name}: 선을 덜 찾음 xs={len(xs)} ys={len(ys)} need={ncol + 2},{nrow + 2}")
    colx, rowy = xs[-(ncol + 1):], ys[-(nrow + 1):]
    dw = ImageDraw.Draw(im)
    for b in boxes:
        br = ws.Range(b)
        c0, r0 = br.Column - a0.Column, br.Row - a0.Row
        c1, r1 = c0 + br.Columns.Count, r0 + br.Rows.Count
        dw.rectangle((colx[c0] * z - 3, rowy[r0] * z - 3, colx[c1] * z + 3, rowy[r1] * z + 3), outline=(220, 30, 30), width=4)
    im = im.crop((int(max(0, xs[0] * z - 12)), int(max(0, ys[0] * z - 12)), int(min(im.width, xs[-1] * z + 12)), int(min(im.height, ys[-1] * z + 12))))
    im.save(os.path.join(SHOTS, name + ".png"))
    doc.close()
    return {"png": "shots/" + name + ".png", "size": im.size, "foreground_changed": fg_before != fg_after,
            "fg_class": [cls(fg_before), cls(fg_after)] if fg_before != fg_after else None}


def cls(h):
    if not h:
        return None
    try:
        return win32gui.GetClassName(h)
    except Exception:  # noqa: BLE001
        return "(닫힘)"


def merge(v, tol=1.2):
    out = []
    for x in v:
        if not out or x - out[-1] > tol:
            out.append(x)
    return out


def boxes_for(rows):
    """요약 줄마다 A:D 빨간 테두리"""
    return [f"A{x['r']}:D{x['r']}" for x in rows if x["formula"]]


R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "notice": NOTICE, "problems": [], "traps": [], "shots": {}}
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    R["excel"] = {"version": xl.Version, "build": xl.Build}

    def new_book(with_answer):
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < 4:
            wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        wb.Worksheets(1).Name = "문제"
        build_problem_sheet(wb.Worksheets(1))
        for k, t in enumerate(TABLES):
            ws = wb.Worksheets(k + 2)
            ws.Name = t["sheet"]
            build_sheet(ws, t, with_answer)
        wb.Worksheets(1).Activate()
        author(wb)
        return wb

    paths = {}
    for with_answer, fname in ((False, "실습.xlsx"), (True, "정답.xlsx")):
        wb = new_book(with_answer)
        p = os.path.join(HERE, fname)
        if os.path.exists(p):
            os.remove(p)
        wb.SaveAs(p, 51)
        wb.Close(False)
        fix_core(p)
        paths[fname] = p

    tmp_ans, tmp_pr = os.path.join(TMP, "ans.xlsx"), os.path.join(TMP, "pr.xlsx")
    shutil.copy(paths["정답.xlsx"], tmp_ans)
    shutil.copy(paths["실습.xlsx"], tmp_pr)
    wb = xl.Workbooks.Open(tmp_ans)
    for n, t in enumerate(TABLES, start=FIRST_N):
        ws = wb.Worksheets(t["sheet"])
        rows = read_rows(ws)
        R["problems"].append({"n": n, "key": t["key"], "text": t["text"], "rows": rows, "summary": summary(rows), "last": rows[-1]["r"]})
        R["shots"][f"{t['key']}_result"] = export(ws, f"A1:D{rows[-1]['r']}", f"{t['key']}_result", boxes_for(rows))
    wb.Close(False)
    # 함정 — 실습 사본(정렬 안 된 표)에서 하나씩, 매번 새로 연다
    for tr in TRAPS:
        t = tr["table"]
        wb = xl.Workbooks.Open(tmp_pr)
        ws = wb.Worksheets(t["sheet"])
        item = {"key": tr["key"], "why": tr["why"]}
        try:
            do_sort(ws, t, tr.get("sort", t["sort"])) if tr.get("sort", t["sort"]) else None
            do_subs(ws, tr.get("subs", t["subs"]))
            if tr.get("remove"):
                item["before_remove_last"] = last_row(ws)
                ws.Range(f"A2:D{last_row(ws)}").RemoveSubtotal()
            rows = read_rows(ws)
            item["rows"] = rows
            item["summary"] = summary(rows)
            item["last"] = rows[-1]["r"]
            item["shot"] = export(ws, f"A1:D{rows[-1]['r']}", f"trap_{tr['key']}", boxes_for(rows))
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
        R["traps"].append(item)
        wb.Close(False)
    wb = xl.Workbooks.Open(tmp_pr)
    for t in TABLES:
        ws = wb.Worksheets(t["sheet"])
        R["shots"][f"{t['key']}_start"] = export(ws, f"A1:D{2 + len(t['rows'])}", f"{t['key']}_start", [])
    wb.Close(False)
    R["file_props"] = {k: core_people(p) for k, p in paths.items()}
finally:
    xl.Quit()
    shutil.rmtree(TMP, ignore_errors=True)

with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)


def brief(rows):
    return [(x["r"], x["level"], " | ".join(x["text"]), x["formula"]) for x in rows]


print(json.dumps({"excel": R.get("excel"), "props": R.get("file_props"),
                  "problems": [(p["n"], brief(p["rows"])) for p in R["problems"]],
                  "traps": [(t["key"], t.get("rejected") or (t.get("before_remove_last"), brief(t["rows"]))) for t in R["traps"]],
                  "fg": [(k, v.get("fg_class")) for k, v in R["shots"].items() if v.get("foreground_changed")] + [(t["key"], t["shot"].get("fg_class")) for t in R["traps"] if t.get("shot", {}).get("foreground_changed")]},
                 ensure_ascii=False, indent=1, default=str))
