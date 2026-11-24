"""엑셀 중급 10강 — 고급 필터, 조건 범위 AND · OR(컴활 2급 실기 기본작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님). 이름 · 숫자는 지어낸 것.

하는 일
1. 실습.xlsx(문제 시트 + 기본작업-28 · 29 · 30 시트, 조건 · 결과 칸 비어 있음) · 정답.xlsx(조건 입력 + 고급 필터 결과) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 고급 필터(다른 장소에 복사)를 실행해 결과 칸에 복사된 행을 읽는다
3. 자주 틀리는 조건(함정)도 같은 자리에 넣어 결과 읽기 — 저장하지 않는 사본에서
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 '고급 필터' 대화 상자 화면은 이 스크립트로 만들지 않는다.
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
TMP = tempfile.mkdtemp(prefix="xl_mid10_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 숫자는 모두 가상입니다."
X_COPY = 2  # xlFilterCopy

# 표는 A2(머리글) ~ D12(데이터 10줄), 조건은 A14 부터, 결과는 A18 부터 — 시험 기본작업 꼴
EDU = [["가람", "영업부", "사원", 86], ["나래", "총무부", "대리", 92], ["다솜", "영업부", "대리", 74], ["로운", "기획부", "사원", 88],
       ["마루", "영업부", "과장", 80], ["보미", "기획부", "사원", 95], ["새롬", "영업부", "사원", 79], ["아라", "총무부", "과장", 83],
       ["여름", "영업부", "대리", 91], ["윤슬", "기획부", "대리", 68]]
T28 = dict(
    key="q28", sheet="기본작업-28", level="기본", title="[표28] 사원 교육 점수", head=["이름", "부서", "직급", "점수"], rows=EDU,
    crit=[["부서", "점수"], ["영업부", ">=80"]], crit_rng="A14:B15", out="A18", out_head=None,
    expect=[r[0] for r in EDU if r[1] == "영업부" and r[3] >= 80],
    text="[표28]에서 부서가 '영업부'이고 점수가 80 이상인 데이터를 고급 필터를 사용하여 검색하시오. 조건은 [A14:B15] 범위 내에 알맞게 입력하고, 결과는 [A18] 셀부터 표시하시오.",
    use="데이터 → 고급 → 다른 장소에 복사 · 조건은 같은 줄(그리고)",
)
BOOK = [["별빛 정원", "소설", 12, 2021], ["숫자의 힘", "경제", 25, 2019], ["바람의 노래", "소설", 31, 2018], ["코딩 첫걸음", "IT", 8, 2023],
        ["여행 노트", "에세이", 22, 2020], ["작은 습관", "자기계발", 17, 2022], ["푸른 밤", "소설", 5, 2024], ["데이터 읽기", "IT", 20, 2021],
        ["마음 산책", "에세이", 9, 2019], ["돈의 흐름", "경제", 14, 2023]]
T29 = dict(
    key="q29", sheet="기본작업-29", level="응용", title="[표29] 도서 대출 현황", head=["도서명", "분류", "대출 횟수", "출판 연도"], rows=BOOK,
    crit=[["분류", "대출 횟수"], ["소설", ""], ["", ">=20"]], crit_rng="A14:B16", out="A18", out_head=["도서명", "대출 횟수"],
    expect=[r[0] for r in BOOK if r[1] == "소설" or r[2] >= 20],
    text="[표29]에서 분류가 '소설'이거나 대출 횟수가 20 이상인 데이터의 '도서명', '대출 횟수' 필드만 고급 필터를 사용하여 검색하시오. 조건은 [A14:B16] 범위 내에 알맞게 입력하고, 결과는 [A18] 셀부터 표시하시오.",
    use="조건은 다른 줄(또는) · 결과 자리에 보일 필드 이름을 먼저 적어 둠",
)
STOCK = [["볼펜", "문구", 120, 500], ["노트", "문구", 45, 1200], ["노트북", "전자", 8, 890000], ["노트 정리함", "생활", 30, 6500],
         ["형광펜", "문구", 15, 800], ["수첩", "문구", 60, 1500], ["노트", "문구", 200, 900], ["연필", "문구", 300, 300],
         ["노트패드", "문구", 70, 2500], ["지우개", "문구", 90, 400]]
T30 = dict(
    key="q30", sheet="기본작업-30", level="함정", title="[표30] 문구 재고", head=["품목", "분류", "재고", "단가"], rows=STOCK,
    crit=[["품목", "재고"], ['="=노트"', "<100"]], crit_rng="A14:B15", out="A18", out_head=None,
    expect=[r[0] for r in STOCK if r[0] == "노트" and r[2] < 100],
    text="[표30]에서 품목이 정확히 '노트'이고 재고가 100 미만인 데이터를 고급 필터를 사용하여 검색하시오. 조건은 [A14:B15] 범위 내에 알맞게 입력하고, 결과는 [A18] 셀부터 표시하시오. ('노트'로 시작하는 다른 품목은 나오면 안 됨)",
    use="정확히 같은 글자만: 조건 칸에 =\"=노트\" (칸에는 =노트 로 보임)",
)
TABLES = [T28, T29, T30]
# 함정 — crit(조건 칸) · crit_rng · out_head 를 바꿔 같은 자리에서 실행
TRAPS = [
    dict(key="q28_two_rows", table=T28, crit=[["부서", "점수"], ["영업부", ""], ["", ">=80"]], crit_rng="A14:B16", why="조건을 다른 줄에 — '또는'이 됨"),
    dict(key="q28_gt", table=T28, crit=[["부서", "점수"], ["영업부", ">80"]], why=">80 — 딱 80 인 마루가 빠짐"),
    dict(key="q28_blank_row", table=T28, crit=[["부서", "점수"], ["영업부", ">=80"], ["", ""]], crit_rng="A14:B16", why="조건 범위에 빈 줄까지 — 모든 줄이 맞음"),
    dict(key="q28_head_space", table=T28, crit=[["부서", "점수 "], ["영업부", ">=80"]], why="조건 머리글 '점수 '(끝에 공백) — 표의 '점수'와 다름"),
    dict(key="q28_value_80", table=T28, crit=[["부서", "점수"], ["영업부", "80"]], why=">= 없이 80 — 80 인 것만"),
    dict(key="q29_same_row", table=T29, crit=[["분류", "대출 횟수"], ["소설", ">=20"]], crit_rng="A14:B15", why="같은 줄에 — '그리고'가 됨"),
    dict(key="q29_no_head", table=T29, out_head=None, why="결과 자리에 필드 이름을 안 적음 — 모든 필드가 복사됨"),
    dict(key="q29_bad_head", table=T29, out_head=["도서명", "대출횟수"], why="결과 필드 이름 띄어쓰기가 다름('대출횟수')"),
    dict(key="q30_plain", table=T30, crit=[["품목", "재고"], ["노트", "<100"]], why="그냥 '노트' — '노트'로 시작하는 것까지"),
    dict(key="q30_typed_formula", table=T30, crit=[["품목", "재고"], ["=노트", "<100"]], why="칸에 =노트 를 그대로 침 — 수식으로 읽혀 #NAME?"),
    dict(key="q30_wild", table=T30, crit=[["품목", "재고"], ["*노트*", "<100"]], why="*노트* — '노트'가 들어간 것 모두"),
]
WIDTHS = dict(zip("ABCD", (13, 11, 11, 11)))
FIRST_N = 28
LESSON = "엑셀 중급 10강 — 고급 필터, 조건 범위 AND · OR (컴활 2급 실기 기본작업 유형 연습)"
LESSON_NO = "엑셀 중급 10강"


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


def put_crit(ws, crit):
    ws.Range("A14:D17").Clear()
    for i, row in enumerate(crit):
        for j, v in enumerate(row):
            c = ws.Range(f"{chr(65 + j)}{14 + i}")
            if v == "":
                continue
            if v.startswith("=") and v.startswith('="'):
                c.Formula = v  # ="=노트" — 칸에는 =노트 로 보이는 글자
            elif v.startswith("="):
                c.Formula = v  # 그대로 친 =노트 → 수식(#NAME?)
            else:
                c.Value = v
    ws.Range(f"A14:{chr(64 + len(crit[0]))}14").Font.Bold = True


def clear_out(ws):
    ws.Range("A18:F40").Clear()


def run_filter(ws, t, crit_rng, out_head):
    clear_out(ws)
    src = ws.Range(f"A2:D{2 + len(t['rows'])}")
    if out_head:
        for j, h in enumerate(out_head):
            ws.Range(f"{chr(65 + j)}18").Value = h
        dest = ws.Range(f"A18:{chr(64 + len(out_head))}18")
    else:
        dest = ws.Range("A18")
    src.AdvancedFilter(X_COPY, ws.Range(crit_rng), dest, False)
    ws.Range("A18:D18").Font.Bold = True
    return read_out(ws)


def read_out(ws):
    head = [ws.Range(f"{chr(65 + j)}18").Text for j in range(4)]
    rows = []
    r = 19
    while ws.Range(f"A{r}").Text != "" and r < 40:
        rows.append([ws.Range(f"{chr(65 + j)}{r}").Text for j in range(4)])
        r += 1
    return {"head": head, "rows": rows, "names": [x[0] for x in rows]}


def crit_texts(ws):
    return [[ws.Range(f"{chr(65 + j)}{14 + i}").Text for j in range(2)] for i in range(3)]


# ── 그림: 범위 → PDF → PNG (8 · 9강과 같은 방식) ───────────
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
    ws.Range(f"B3:D{last}").HorizontalAlignment = -4108
    hdr.Font.Bold = True
    hdr.Interior.Color = 0xEADBC8
    hdr.HorizontalAlignment = -4108
    ws.Range("A13").Value = "조건"
    ws.Range("A17").Value = "결과"
    ws.Range("A13,A17").Font.Color = 0x808080
    for col, w in WIDTHS.items():
        ws.Columns(col).ColumnWidth = w
    if with_answer:
        put_crit(ws, t["crit"])
        run_filter(ws, t, t["crit_rng"], t["out_head"])


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


def area(n):
    """표 · 조건 · 결과까지만(결과 줄 수 + 한 줄)."""
    return f"A1:D{18 + max(1, n) + 1}"


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
        out = read_out(ws)
        R["problems"].append({"n": n, "key": t["key"], "text": t["text"], "crit": crit_texts(ws), "crit_formula": [ws.Range(f"A{14 + i}").Formula for i in range(3)],
                              "out": out, "expect": t["expect"], "ok": out["names"] == t["expect"]})
        R["shots"][f"{t['key']}_result"] = export(ws, area(len(out["rows"])), f"{t['key']}_result", [t["crit_rng"], f"A18:{chr(64 + len(t['out_head'] or t['head']))}{18 + len(out['rows'])}"])
    for tr in TRAPS:
        t = tr["table"]
        ws = wb.Worksheets(t["sheet"])
        crit, crit_rng = tr.get("crit", t["crit"]), tr.get("crit_rng", t["crit_rng"])
        out_head = tr["out_head"] if "out_head" in tr else t["out_head"]
        item = {"key": tr["key"], "why": tr["why"], "crit": crit, "crit_rng": crit_rng, "out_head": out_head}
        put_crit(ws, crit)
        item["crit_shown"] = crit_texts(ws)
        try:
            out = run_filter(ws, t, crit_rng, out_head)
            item["out"] = out
            item["count"] = len(out["rows"])
            item["shot"] = export(ws, area(len(out["rows"])), f"trap_{tr['key']}", [crit_rng, f"A18:{chr(64 + len(out_head or t['head']))}{18 + max(1, len(out['rows']))}"])
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
            item["shot"] = export(ws, area(1), f"trap_{tr['key']}", [crit_rng])
        R["traps"].append(item)
        put_crit(ws, t["crit"])
        run_filter(ws, t, t["crit_rng"], t["out_head"])
    R["after_traps_ok"] = all(read_out(wb.Worksheets(t["sheet"]))["names"] == t["expect"] for t in TABLES)
    wb.Close(False)
    wb = xl.Workbooks.Open(tmp_pr)
    for t in TABLES:
        R["shots"][f"{t['key']}_start"] = export(wb.Worksheets(t["sheet"]), area(1), f"{t['key']}_start", [t["crit_rng"]])
    wb.Close(False)
    R["file_props"] = {k: core_people(p) for k, p in paths.items()}
finally:
    xl.Quit()
    shutil.rmtree(TMP, ignore_errors=True)

with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)
print(json.dumps({"excel": R.get("excel"),
                  "problems": [(p["n"], p["ok"], p["out"]["head"], p["out"]["names"], p["crit"]) for p in R["problems"]],
                  "props": R.get("file_props"), "after": R.get("after_traps_ok"),
                  "traps": [(t["key"], t.get("rejected") or (t["count"], t["out"]["names"], t["out"]["head"]), t["crit_shown"]) for t in R["traps"]],
                  "fg": [(k, v.get("fg_class")) for k, v in R["shots"].items() if v.get("foreground_changed")] + [(t["key"], t["shot"].get("fg_class")) for t in R["traps"] if t.get("shot", {}).get("foreground_changed")]},
                 ensure_ascii=False, indent=1, default=str))
