"""엑셀 중급 4강 — 반올림 ROUND · ROUNDUP · ROUNDDOWN · INT · MOD (컴활 2급 실기 계산작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님). 이름은 지어낸 이름.

하는 일
1. 실습.xlsx(문제 시트 + 계산작업 시트, 답 칸 비어 있음) · 정답.xlsx(답 칸에 수식) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 수식을 넣고(첫 칸에 넣은 뒤 아래로 채우기 = FillDown) 칸에 보이는 값(.Text)을 셀 주소별로 읽기
3. 자주 틀리는 식(함정)도 같은 칸에 넣어 결과 읽기 — 저장하지 않는 사본에서. 한 칸짜리 사실(음수 INT · .5 반올림 등)도 빈 칸에서 잼
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 수식 입력줄 · 함수 인수 창 화면은 이 스크립트로 만들지 않는다.
결과: 실습.xlsx · 정답.xlsx · measure.json · shots/*.png (shots 는 저장소에 올리지 않음)
실행: python make_measure.py   (pywin32 · pymupdf · pillow)"""
import datetime
import json
import os
import shutil
import sys
import tempfile

import fitz
import pythoncom
import pywintypes
import win32com.client as win32
import win32gui
from PIL import Image, ImageDraw

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
TMP = tempfile.mkdtemp(prefix="xl_mid04_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 코드 · 숫자는 모두 가상입니다."
WON = "#,##0"
PCT = "0%"


def grid(top_left_col, top_row, rows, head=True):
    """옆 표(side) 칸들 — rows 는 행마다 값 목록. 첫 행을 머리글(굵게)로."""
    out = {}
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            out[f"{chr(ord(top_left_col) + j)}{top_row + i}"] = (v, head and i == 0)
    return out


def hgrid(left_col, top_row, rows):
    """가로 표 — 첫 열을 머리글(굵게)로."""
    out = {}
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            out[f"{chr(ord(left_col) + j)}{top_row + i}"] = (v, j == 0)
    return out


# ── 문제 정의 ─────────────────────────────────────────────
T10 = dict(
    key="q10", level="기본", r=1, title="[표10] 사내 교육 평가",
    head=["이름", "1차", "2차", "3차", "평균"],
    rows=[["가온솔", 88, 92, 79], ["나린결", 75, 80, 82], ["다솜빛", 91, 85, 88], ["라온새", 66, 72, 70],
          ["마루별", 95, 97, 92], ["바다슬", 58, 63, 61], ["새봄이", 82, 79, 84], ["아라찬", 70, 71, 73]],
    side={},
    answers=[("E3", "E3:E10", "=ROUND(AVERAGE(B3:D3),1)", ["86.3", "79", "88", "69.3", "94.7", "60.7", "81.7", "71.3"])],
    area="A1:E10", boxes=["E3:E10"], fmt={}, border_side=[],
    text="[표10]에서 1차[B3:B10] · 2차[C3:C10] · 3차[D3:D10] 점수의 평균을 반올림하여 소수 첫째 자리까지 평균[E3:E10]에 표시하시오.",
    use="ROUND, AVERAGE 함수 사용",
)
T11 = dict(
    key="q11", level="응용", r=12, title="[표11] 문구 주문 포장",
    head=["상품", "수량", "단가", "상자 수", "할인 금액"],
    rows=[["볼펜", 50, 650], ["노트", 24, 1800], ["형광펜", 37, 900], ["지우개", 12, 400],
          ["파일", 80, 1300], ["점착 메모", 15, 2500], ["수정 테이프", 9, 1700], ["클립", 100, 300]],
    side={},
    answers=[("D14", "D14:D21", "=ROUNDUP(B14/12,0)", ["5", "2", "4", "1", "7", "2", "1", "9"]),
             ("E14", "E14:E21", "=ROUNDDOWN(B14*C14*0.9,-2)", ["29,200", "38,800", "29,900", "4,300", "93,600", "33,700", "13,700", "27,000"])],
    area="A12:E21", boxes=["D14:E21"], fmt={"C14:C21": WON, "E14:E21": WON}, border_side=[],
    text=("[표11]에서 ① 한 상자에 12개씩 담을 때 필요한 상자 수를 상자 수[D14:D21]에, "
          "② 수량[B14:B21] × 단가[C14:C21]에서 10% 할인한 금액을 백 원 미만은 버려 할인 금액[E14:E21]에 표시하시오."),
    use="① ROUNDUP 함수 ② ROUNDDOWN 함수 사용",
)
T12 = dict(
    key="q12", level="함정", r=23, title="[표12] 작업 시간 기록",
    head=["작업", "작업 시간(분)", "시간", "분"],
    rows=[["포장", 95], ["검수", 130], ["청소", 45], ["재고 정리", 240], ["입고", 61], ["교육", 180], ["회의", 59], ["발송", 125]],
    side={},
    answers=[("C25", "C25:C32", "=INT(B25/60)", ["1", "2", "0", "4", "1", "3", "0", "2"]),
             ("D25", "D25:D32", "=MOD(B25,60)", ["35", "10", "45", "0", "1", "0", "59", "5"])],
    area="A23:D32", boxes=["C25:D32"], fmt={}, border_side=[],
    text="[표12]에서 작업 시간(분)[B25:B32]을 시간[C25:C32]과 분[D25:D32]으로 나누어 표시하시오. (예: 95분 → 1시간 35분)",
    use="시간은 INT 함수, 분은 MOD 함수 사용",
)
TABLES = [T10, T11, T12]
# 함정 — exp_idx 는 그 표의 answers 중 몇 번째와 견주는지
TRAPS = [
    dict(key="q10_round0", table=T10, top="E3", fill="E3:E10", exp_idx=0, why="자릿수를 0 으로 → 정수로 반올림",
         formula="=ROUND(AVERAGE(B3:D3),0)"),
    dict(key="q10_rounddown", table=T10, top="E3", fill="E3:E10", exp_idx=0, why="ROUND 대신 ROUNDDOWN → 소수 둘째 자리부터 버림",
         formula="=ROUNDDOWN(AVERAGE(B3:D3),1)"),
    dict(key="q10_noround", table=T10, top="E3", fill="E3:E10", exp_idx=0, why="ROUND 없이 AVERAGE 만",
         formula="=AVERAGE(B3:D3)"),
    dict(key="q11_round", table=T11, top="D14", fill="D14:D21", exp_idx=0, why="상자 수에 ROUND(반올림) → 남는 물건이 상자에 안 들어감",
         formula="=ROUND(B14/12,0)"),
    dict(key="q11_int", table=T11, top="D14", fill="D14:D21", exp_idx=0, why="상자 수에 INT(내림)",
         formula="=INT(B14/12)"),
    dict(key="q11_digits2", table=T11, top="E14", fill="E14:E21", exp_idx=1, why="'백 원 미만 버림'을 자릿수 2 로 → 소수 둘째 자리라 아무것도 안 버림",
         formula="=ROUNDDOWN(B14*C14*0.9,2)"),
    dict(key="q11_digits3", table=T11, top="E14", fill="E14:E21", exp_idx=1, why="자릿수 -3 → 천 원 미만까지 버림",
         formula="=ROUNDDOWN(B14*C14*0.9,-3)"),
    dict(key="q12_round", table=T12, top="C25", fill="C25:C32", exp_idx=0, why="시간에 ROUND → 30분 넘으면 1시간이 올라감",
         formula="=ROUND(B25/60,0)"),
    dict(key="q12_divide", table=T12, top="C25", fill="C25:C32", exp_idx=0, why="INT 없이 나누기만",
         formula="=B25/60"),
    dict(key="q12_mod_swap", table=T12, top="D25", fill="D25:D32", exp_idx=1, why="MOD 인수 순서를 거꾸로(60 을 작업 시간으로 나눈 나머지)",
         formula="=MOD(60,B25)"),
]
# 한 칸 사실 — 빈 칸(Z1)에 넣고 읽는다
FACTS = [
    ("=INT(-7.5)", "음수 INT 는 더 작은 정수로"), ("=ROUNDDOWN(-7.5,0)", "ROUNDDOWN 은 0 쪽으로"), ("=TRUNC(-7.5)", "TRUNC 도 0 쪽으로"),
    ("=ROUND(2.5,0)", ".5 반올림"), ("=ROUND(-2.5,0)", "음수 .5 반올림"), ("=ROUND(1234.567,-2)", "자릿수 -2 = 백의 자리"),
    ("=ROUNDUP(1201,-2)", "1201 백의 자리 올림"), ("=ROUNDDOWN(1299,-2)", "1299 백의 자리 내림"), ("=ROUNDUP(-1.2,0)", "음수 ROUNDUP"),
    ("=MOD(-7,3)", "음수 MOD"), ("=MOD(7,-3)", "나누는 수가 음수인 MOD"), ("=MOD(10,0)", "0 으로 나눈 MOD"),
]
WIDTHS = dict(zip("ABCDEFGHI", (11, 12, 9, 9, 11, 10, 10, 10, 10)))
FW = {"q10": {"E": 40}, "q11": {"D": 30, "E": 40}, "q12": {"C": 22, "D": 22}}
FIRST_N = 10
LESSON = "엑셀 중급 4강 — 반올림 ROUND · ROUNDUP · ROUNDDOWN · INT · MOD (컴활 2급 실기 계산작업 유형 연습)"
LESSON_NO = "엑셀 중급 4강"


def first(t):
    return t["r"] + 2


def last(t):
    return t["r"] + 1 + len(t["rows"])


def author(wb):
    for k in ("Author", "Last Author"):
        wb.BuiltinDocumentProperties(k).Value = "CODINGNOW"


def fix_core(path):
    """저장하면 엑셀이 '마지막 수정한 사람'을 이 PC 사용자 이름으로 다시 쓴다 → docProps/core.xml 을 직접 CODINGNOW 로(기초 7 · 8강과 같음)."""
    import re
    import zipfile
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
    import re
    import zipfile
    with zipfile.ZipFile(path) as z:
        s = z.read("docProps/core.xml").decode("utf-8")
    return {tag: (re.search(f"<{tag}>([^<]*)</{tag}>", s) or [None, None])[1] for tag in ("dc:creator", "cp:lastModifiedBy")}

def texts(ws, rng):
    return {c.Address.replace("$", ""): c.Text for c in ws.Range(rng)}


def formulas(ws, rng):
    return {c.Address.replace("$", ""): c.Formula for c in ws.Range(rng)}


# ── 그림: 범위 → PDF → PNG ───────────────────────────────
def export(ws, rng, name, boxes=(), show_formulas=False, widths=None):
    """rng 를 PDF 로 내보내 PNG 로. boxes = 빨간 테두리를 그릴 범위 주소들."""
    fg_before = win32gui.GetForegroundWindow()
    saved = {}
    if show_formulas:
        # 수식 보기에서는 엑셀이 열을 두 배 넓게 그린다(ColumnWidth 값은 그대로) → 반으로 줄여 원래 폭처럼 보이게
        widths = widths or {}
        for k in range(ws.Range(rng).Columns.Count):
            col = ws.Range(rng).Columns(k + 1).Address.split(":")[0].split("$")[1]  # '$A$2' → 'A'
            saved[col] = ws.Columns(col).ColumnWidth
            ws.Columns(col).ColumnWidth = widths.get(col, saved[col]) / 2
    ws.Parent.Windows(1).DisplayFormulas = show_formulas
    ps = ws.PageSetup
    ps.PrintArea = rng
    ps.PrintHeadings = True
    ps.PrintGridlines = True
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    for m in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin", "HeaderMargin", "FooterMargin"):
        setattr(ps, m, 14)
    ps.CenterHeader = ""
    ps.CenterFooter = ""
    pdf = os.path.join(TMP, name + ".pdf")
    ws.ExportAsFixedFormat(0, pdf)
    ws.Parent.Windows(1).DisplayFormulas = False
    for col, w in saved.items():
        ws.Columns(col).ColumnWidth = w
    fg_after = win32gui.GetForegroundWindow()

    doc = fitz.open(pdf)
    page = doc[0]
    xs, ys = set(), set()
    for d in page.get_drawings():
        for it in d["items"]:
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
    xs = merge(sorted(xs))
    ys = merge(sorted(ys))
    z = 2.0
    pix = page.get_pixmap(matrix=fitz.Matrix(z, z))
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    # 범위 → 쪽 좌표: 인쇄 영역 왼쪽 위 칸의 엑셀 좌표를 기준으로, 머리글 오른쪽 · 아래 선에 맞춘다
    a0 = ws.Range(rng).Cells(1, 1)
    ncol = ws.Range(rng).Columns.Count
    nrow = ws.Range(rng).Rows.Count
    if len(xs) < ncol + 2 or len(ys) < nrow + 2:
        raise RuntimeError(f"{name}: 선을 덜 찾음 xs={xs} ys={ys} pages={doc.page_count} need={ncol + 2},{nrow + 2}")
    colx = xs[-(ncol + 1):]
    rowy = ys[-(nrow + 1):]
    d = ImageDraw.Draw(im)
    for b in boxes:
        br = ws.Range(b)
        c0 = br.Column - a0.Column
        r0 = br.Row - a0.Row
        c1 = c0 + br.Columns.Count
        r1 = r0 + br.Rows.Count
        x0, y0, x1, y1 = colx[c0] * z, rowy[r0] * z, colx[c1] * z, rowy[r1] * z
        d.rectangle((x0 - 3, y0 - 3, x1 + 3, y1 + 3), outline=(220, 30, 30), width=4)
    left, top = xs[0] * z - 12, ys[0] * z - 12
    right, bottom = xs[-1] * z + 12, ys[-1] * z + 12
    im = im.crop((int(max(0, left)), int(max(0, top)), int(min(im.width, right)), int(min(im.height, bottom))))
    out = os.path.join(SHOTS, name + ".png")
    im.save(out)
    doc.close()
    return {"png": "shots/" + name + ".png", "size": im.size, "cols_found": len(xs), "rows_found": len(ys),
            "need": [ncol + 2, nrow + 2], "foreground_changed": fg_before != fg_after}


def merge(v, tol=1.2):
    out = []
    for x in v:
        if not out or x - out[-1] > tol:
            out.append(x)
    return out


def build_sheet(ws, with_answer):
    for t in TABLES:
        r = t["r"]
        ws.Range(f"A{r}").Value = t["title"]
        ws.Range(f"A{r}").Font.Bold = True
        c = [chr(ord("A") + i) for i in range(len(t["head"]))]
        hdr = ws.Range(f"A{r + 1}:{c[-1]}{r + 1}")
        hdr.Value = (tuple(t["head"]),)
        for i, row in enumerate(t["rows"]):
            for j, v in enumerate(row):
                ws.Range(f"{c[j]}{first(t) + i}").Value = v
        ws.Range(f"A{r + 1}:{c[-1]}{last(t)}").Borders.LineStyle = 1
        ws.Range(f"B{first(t)}:{c[-1]}{last(t)}").HorizontalAlignment = -4108
        heads = [hdr]
        for cell, (v, is_head) in t["side"].items():
            ws.Range(cell).Value = v
            ws.Range(cell).HorizontalAlignment = -4108
            if is_head:
                heads.append(ws.Range(cell))
        for rng in t["border_side"]:
            ws.Range(rng).Borders.LineStyle = 1
        for h in heads:
            h.Font.Bold = True
            h.Interior.Color = 0xEADBC8  # 연한 파랑(BGR)
            h.HorizontalAlignment = -4108
        for rng, f in t["fmt"].items():
            ws.Range(rng).NumberFormat = f
        if with_answer:
            for top, fill, formula, _ in t["answers"]:
                ws.Range(top).Formula = formula
                if fill != top:
                    ws.Range(fill).FillDown()
    for col, w in WIDTHS.items():
        ws.Columns(col).ColumnWidth = w


def build_problem_sheet(ws):
    ws.Range("A1").Value = LESSON
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = NOTICE
    ws.Range("A2").Font.Color = 0x3030C0
    line = 4
    for n, t in enumerate(TABLES, start=FIRST_N):
        ws.Range(f"A{line}").Value = f"{n}. ({t['level']}) {t['text']}"
        ws.Range(f"A{line + 1}").Value = f"   ▶ {t['use']}"
        line += 3
    ws.Range(f"A{line}").Value = f"풀이 · 실측 값 · 자주 틀리는 곳: https://www.coding-now.com/blog/excel ({LESSON_NO})"
    ws.Columns("A").ColumnWidth = 120


# ── 실행 ──────────────────────────────────────────────────
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
        while wb.Worksheets.Count < 2:
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        wp, wc = wb.Worksheets(1), wb.Worksheets(2)
        wp.Name, wc.Name = "문제", "계산작업"
        build_problem_sheet(wp)
        build_sheet(wc, with_answer)
        wc.Activate()
        wc.Range("A1").Select()
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

    tmp_ans = os.path.join(TMP, "ans.xlsx")
    tmp_pr = os.path.join(TMP, "pr.xlsx")
    shutil.copy(paths["정답.xlsx"], tmp_ans)
    shutil.copy(paths["실습.xlsx"], tmp_pr)
    wb = xl.Workbooks.Open(tmp_ans)
    ws = wb.Worksheets("계산작업")
    for n, t in enumerate(TABLES, start=FIRST_N):
        item = {"n": n, "key": t["key"], "level": t["level"], "text": t["text"], "use": t["use"], "answers": []}
        for top, fill, formula, expect in t["answers"]:
            got = texts(ws, fill)
            item["answers"].append({"first_cell": top, "first_cell_formula": ws.Range(top).Formula,
                                    "formulas": formulas(ws, fill), "values": got,
                                    "value2": {c.Address.replace("$", ""): c.Value2 for c in ws.Range(fill)},
                                    "ok": list(got.values()) == expect})
        R["problems"].append(item)
    # 그림 — 정답 결과 · 수식 보기
    for t in TABLES:
        R["shots"][f"{t['key']}_result"] = export(ws, t["area"], f"{t['key']}_result", t["boxes"])
        a = t["area"].replace(f"A{t['r']}:", f"A{t['r'] + 1}:")
        R["shots"][f"{t['key']}_formula"] = export(ws, a, f"{t['key']}_formula", t["boxes"], show_formulas=True, widths=FW[t["key"]])
    # 표시 형식으로만 소수 한 자리 — 보이는 값과 칸에 든 값(문제 10 함정): F3 에 =AVERAGE 를 넣고 0.0 형식
    ws.Range("F3").Formula = "=AVERAGE(B3:D3)"
    ws.Range("F3").NumberFormat = "0.0"
    ws.Range("G3").Formula = "=F3*3"
    ws.Range("H3").Formula = "=E3*3"
    R["format_only"] = {"F3_text": ws.Range("F3").Text, "F3_value": ws.Range("F3").Value2, "E3_text": ws.Range("E3").Text,
                        "E3_value": ws.Range("E3").Value2, "F3x3": ws.Range("G3").Text, "E3x3": ws.Range("H3").Text}
    ws.Range("F3:H3").Clear()
    R["facts"] = []
    for f, label in FACTS:
        try:
            ws.Range("Z1").Formula = f
            R["facts"].append({"formula": f, "label": label, "text": ws.Range("Z1").Text, "value": ws.Range("Z1").Value2})
        except pywintypes.com_error as e:
            R["facts"].append({"formula": f, "label": label, "rejected": str(e)[:120]})
    ws.Range("Z1").Clear()
    # 함정 — 같은 칸에 틀린 식
    for tr in TRAPS:
        t = tr["table"]
        item = {"key": tr["key"], "why": tr["why"], "formula": tr["formula"], "range": tr["fill"]}
        try:
            ws.Range(tr["top"]).Formula = tr["formula"]
            if tr["fill"] != tr["top"]:
                ws.Range(tr["fill"]).FillDown()
            item["formulas"] = formulas(ws, tr["fill"])
            item["values"] = texts(ws, tr["fill"])
            exp = t["answers"][tr.get("exp_idx", 0)][3]
            item["wrong_vs_answer"] = sum(1 for a_, b_ in zip(item["values"].values(), exp) if a_ != b_)
            item["shot"] = export(ws, t["area"], f"trap_{tr['key']}", [tr["fill"]])
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
        for top, fill, formula, _ in t["answers"]:  # 정답 식으로 되돌림
            ws.Range(top).Formula = formula
            if fill != top:
                ws.Range(fill).FillDown()
        R["traps"].append(item)
    wb.Close(False)
    wb = xl.Workbooks.Open(tmp_pr)
    ws = wb.Worksheets("계산작업")
    for t in TABLES:
        R["shots"][f"{t['key']}_start"] = export(ws, t["area"], f"{t['key']}_start", t["boxes"])
    wb.Close(False)
    R["file_props"] = {k: core_people(p) for k, p in paths.items()}
finally:
    xl.Quit()
    shutil.rmtree(TMP, ignore_errors=True)

with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1)
print(json.dumps({"ok": [[a["ok"] for a in p["answers"]] for p in R["problems"]], "props": R.get("file_props"),
                  "values": [list(p["answers"][0]["values"].values()) for p in R["problems"]],
                  "traps": [(t["key"], t.get("rejected") or list(t.get("values", {}).values())) for t in R["traps"]],
                  "fg": any(v["foreground_changed"] for v in R["shots"].values()),
                  "facts": [(x["formula"], x.get("text", x.get("rejected"))) for x in R.get("facts", [])], "format_only": R.get("format_only")},
                 ensure_ascii=False, indent=1))
