"""엑셀 중급 7강 — 데이터베이스 함수 DSUM · DAVERAGE · DCOUNT · DMAX · DMIN (컴활 2급 실기 계산작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님). 이름은 지어낸 이름.

하는 일
1. 실습.xlsx(문제 시트 + 계산작업 시트, 답 칸 비어 있음 — 문제 21 은 조건 칸도 비어 있음) · 정답.xlsx(답 칸에 수식 · 조건) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 수식을 넣고 칸에 보이는 값(.Text)과 칸에 든 값의 종류를 셀 주소별로 읽기
3. 자주 틀리는 식(함정)도 같은 칸에 넣어 결과 읽기 — 조건 칸을 잠깐 바꿔(빈 줄 · 다른 줄 · 머리글 공백 · '문구' 와 '=문구') 재고 되돌린다. 저장하지 않는 사본에서
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
TMP = tempfile.mkdtemp(prefix="xl_mid07_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 수량 · 금액 · 기록은 모두 가상입니다."
YMD = "yyyy-mm-dd"
GEN = "G/표준"  # 한국어 엑셀의 '일반' 형식 코드(NumberFormat = "General" 은 거부됨, 11/18 실측)
DOW = "월화수목금토일"


def d(s):
    return datetime.date.fromisoformat(s)


# ── 문제 정의 ─────────────────────────────────────────────
WON = "#,##0"
SALES = [["서울", "가람", "볼펜", 120, 96000], ["부산", "나래", "노트", 80, 96000], ["서울", "다솜", "형광펜", 60, 72000], ["대구", "로운", "볼펜", 150, 120000],
         ["서울", "마루", "파일철", 40, 100000], ["부산", "보미", "볼펜", 90, 72000], ["대구", "새롬", "노트", 70, 84000], ["서울", "아라", "가위", 30, 105000]]
SEOUL = [r for r in SALES if r[0] == "서울"]
T19 = dict(
    key="q19", level="기본", r=1, title="[표19] 지점별 판매 기록",
    head=["지점", "담당자", "품목", "수량", "금액"],
    rows=SALES,
    side={"G2": ("지점", True), "G3": ("서울", False), "H2": ("금액 합계", True), "I2": ("판매 건수", True)},
    answers=[("H3", "H3", "=DSUM(A2:E10,5,G2:G3)", [f"{sum(r[4] for r in SEOUL):,}"]),
             ("I3", "I3", "=DCOUNT(A2:E10,4,G2:G3)", [str(len(SEOUL))])],
    area="A1:I10", boxes=["H3:I3"], fmt={"E3:E10": WON, "H3": WON}, fmt_after={}, border_side=["G2:G3", "H2:I3"],
    text=("[표19]에서 지점이 '서울'인 판매의 금액 합계를 [H3]에, 판매 건수를 [I3]에 표시하시오. "
          "(조건은 [G2:G3] 영역 사용, 판매 건수는 수량 필드로 셈)"),
    use="DSUM, DCOUNT 함수 사용",
)
FIT = [["하늘", "여", 34, "수영", 42], ["바다", "남", 28, "요가", 35], ["별", "여", 41, "요가", 50], ["솔", "남", 37, "수영", 38],
       ["이랑", "여", 26, "수영", 45], ["누리", "남", 45, "걷기", 60], ["아름", "여", 30, "걷기", 55], ["한결", "여", 52, "요가", 47]]
W30 = [r[4] for r in FIT if r[1] == "여" and r[2] >= 30]
SY = [r[4] for r in FIT if r[3] in ("수영", "요가")]
T20 = dict(
    key="q20", level="응용", r=12, title="[표20] 회원 운동 기록",
    head=["이름", "성별", "나이", "종목", "기록(분)"],
    rows=FIT,
    side={"G13": ("성별", True), "H13": ("나이", True), "G14": ("여", False), "H14": (">=30", False),
          "G16": ("종목", True), "G17": ("수영", False), "G18": ("요가", False),
          "I13": ("여 · 30세 이상 평균", True), "I16": ("수영 · 요가 최장", True)},
    answers=[("I14", "I14", "=ROUND(DAVERAGE(A13:E21,5,G13:H14),1)", [f"{round(sum(W30) / len(W30), 1):g}"]),
             ("I17", "I17", "=DMAX(A13:E21,5,G16:G18)", [str(max(SY))])],
    area="A12:I21", boxes=["I14", "I17"], fmt={}, fmt_after={}, border_side=["G13:H14", "G16:G18", "I13:I14", "I16:I17"],
    text=("[표20]에서 성별이 '여'이고 나이가 30 이상인 회원의 기록(분) 평균을 반올림하여 소수 첫째 자리까지 [I14]에, "
          "종목이 '수영' 또는 '요가'인 회원의 가장 긴 기록(분)을 [I17]에 표시하시오. (조건은 [G13:H14] · [G16:G18] 영역 사용)"),
    use="① ROUND, DAVERAGE 함수 ② DMAX 함수 사용",
)
STOCK = [["볼펜", "문구", 120, 500], ["노트", "문구", 45, 1200], ["파일함", "문구류", 10, 3000], ["형광펜", "문구", 15, 800],
         ["건전지", "생활", 30, 2500], ["수첩", "문구", 60, 1500], ["휴지", "생활", 200, 700], ["연필", "문구", 300, 300]]
MUNGU = [r[2] for r in STOCK if r[1] == "문구"]
T21 = dict(
    key="q21", level="함정", r=23, title="[표21] 상품 재고",
    head=["품목", "분류", "재고", "단가"],
    rows=STOCK,
    side={"F24": ("분류", True), "G24": ("최소 재고", True), "H24": ("재고 합계", True)},
    side_answer={"F25": '="=문구"'},  # 실습 파일에서는 비워 둠 — 학생이 조건을 직접 입력
    answers=[("F25", "F25", '="=문구"', ["=문구"]),
             ("G25", "G25", "=DMIN(A24:D32,3,F24:F25)", [str(min(MUNGU))]),
             ("H25", "H25", "=DSUM(A24:D32,3,F24:F25)", [str(sum(MUNGU))])],
    area="A23:H32", boxes=["F24:F25", "G25:H25"], fmt={"D25:D32": WON}, fmt_after={}, border_side=["F24:F25", "G24:H25"],
    text=("[표21]에서 분류가 정확히 '문구'인 상품(문구류는 제외)의 최소 재고를 [G25]에, 재고 합계를 [H25]에 표시하시오. "
          "(조건은 [F24:F25] 영역에 입력하여 사용)"),
    use="DMIN, DSUM 함수 사용",
)
TABLES = [T19, T20, T21]
# 함정 — set 이 있으면 그 칸들을 잠깐 바꿔 넣고 잰 뒤 되돌린다
TRAPS = [
    dict(key="q19_no_header", table=T19, top="H3", fill="H3", exp_idx=0, why="데이터베이스 범위에서 머리글 줄을 뺌(A3:E10)",
         formula="=DSUM(A3:E10,5,G2:G3)"),
    dict(key="q19_field4", table=T19, top="H3", fill="H3", exp_idx=0, why="필드 번호를 4(수량)로",
         formula="=DSUM(A2:E10,4,G2:G3)"),
    dict(key="q19_dcount_text", table=T19, top="I3", fill="I3", exp_idx=1, why="DCOUNT 를 글자 필드(1 = 지점)로 → 숫자만 세서 0",
         formula="=DCOUNT(A2:E10,1,G2:G3)"),
    dict(key="q19_crit_nohead", table=T19, top="H3", fill="H3", exp_idx=0, why="조건 범위에서 머리글(G2)을 뺌",
         formula="=DSUM(A2:E10,5,G3)"),
    dict(key="q20_blank_row", table=T20, top="I17", fill="I17", exp_idx=1, why="조건 범위에 빈 줄(G19)까지 넣음 → 빈 조건은 모두와 맞음",
         formula="=DMAX(A13:E21,5,G16:G19)"),
    dict(key="q20_gt", table=T20, top="I14", fill="I14", exp_idx=0, why="'이상'을 >30 으로(30세 아름이 빠짐)",
         formula="=ROUND(DAVERAGE(A13:E21,5,G13:H14),1)", set={"H14": ">30"}),
    dict(key="q20_or_rows", table=T20, top="I14", fill="I14", exp_idx=0, why="두 조건을 다른 줄에(여 / >=30) → '그리고'가 아니라 '또는'",
         formula="=ROUND(DAVERAGE(A13:E21,5,G13:H15),1)", set={"H14": "", "H15": ">=30"}),
    dict(key="q20_field_text", table=T20, top="I14", fill="I14", exp_idx=0, why="필드 번호를 4(종목 — 글자)로",
         formula="=ROUND(DAVERAGE(A13:E21,4,G13:H14),1)"),
    dict(key="q21_plain", table=T21, top="G25", fill="G25:H25", exp_idx=1, expect=[str(min(MUNGU)), str(sum(MUNGU))], why="조건을 그냥 '문구'로 → '문구'로 시작하는 문구류까지",
         formula="=DMIN(A24:D32,3,F24:F25)", set={"F25": "문구"}, also={"H25": "=DSUM(A24:D32,3,F24:F25)"}),
    dict(key="q21_head_space", table=T21, top="G25", fill="G25:H25", exp_idx=1, expect=[str(min(MUNGU)), str(sum(MUNGU))], why="조건 머리글 끝에 공백('분류 ')",
         formula="=DMIN(A24:D32,3,F24:F25)", set={"F24": "분류 "}, also={"H25": "=DSUM(A24:D32,3,F24:F25)"}),
]
# 한 칸 사실 — 빈 칸(Z1)에 넣고 읽는다(표를 참조하는 식도 있음)
FACTS = [
    ('=DSUM(A2:E10,"금액",G2:G3)', "필드 번호 대신 머리글 이름"), ("=DCOUNTA(A2:E10,1,G2:G3)", "DCOUNTA 는 글자 칸도 셈"),
    ("=DCOUNT(A2:E10,,G2:G3)", "필드를 비우면 조건에 맞는 줄 수"), ('=SUMIF(A3:A10,"서울",E3:E10)', "SUMIF 로 같은 합계"),
    ('=DAVERAGE(A13:E21,"기록(분)",G13:H14)', "ROUND 없이 평균"), ("=DMIN(A13:E21,5,G16:G18)", "수영 · 요가 최단"),
    ('=COUNTIFS(B14:B21,"여",C14:C21,">=30")', "COUNTIFS 로 여 · 30 이상 사람 수"),
]
WIDTHS = dict(zip("ABCDEFGHI", (10, 10, 9, 11, 11, 10, 11, 18, 18)))
FW = {"q19": {"H": 46, "I": 46}, "q20": {"I": 70}, "q21": {"F": 16, "G": 44, "H": 44}}
FIRST_N = 19
LESSON = "엑셀 중급 7강 — 데이터베이스 함수 DSUM · DAVERAGE · DCOUNT · DMAX · DMIN (컴활 2급 실기 계산작업 유형 연습)"
LESSON_NO = "엑셀 중급 7강"
AUTO_FMT = {}


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
    xs = merge(sorted(xs))
    ys = merge(sorted(ys))
    z = 2.0
    pix = page.get_pixmap(matrix=fitz.Matrix(z, z))
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    a0 = ws.Range(rng).Cells(1, 1)
    ncol = ws.Range(rng).Columns.Count
    nrow = ws.Range(rng).Rows.Count
    if len(xs) < ncol + 2 or len(ys) < nrow + 2:
        raise RuntimeError(f"{name}: 선을 덜 찾음 xs={xs} ys={ys} pages={doc.page_count} need={ncol + 2},{nrow + 2}")
    colx = xs[-(ncol + 1):]
    rowy = ys[-(nrow + 1):]
    dw = ImageDraw.Draw(im)
    for b in boxes:
        br = ws.Range(b)
        c0 = br.Column - a0.Column
        r0 = br.Row - a0.Row
        c1 = c0 + br.Columns.Count
        r1 = r0 + br.Rows.Count
        x0, y0, x1, y1 = colx[c0] * z, rowy[r0] * z, colx[c1] * z, rowy[r1] * z
        dw.rectangle((x0 - 3, y0 - 3, x1 + 3, y1 + 3), outline=(220, 30, 30), width=4)
    left, top = xs[0] * z - 12, ys[0] * z - 12
    right, bottom = xs[-1] * z + 12, ys[-1] * z + 12
    im = im.crop((int(max(0, left)), int(max(0, top)), int(min(im.width, right)), int(min(im.height, bottom))))
    out = os.path.join(SHOTS, name + ".png")
    im.save(out)
    doc.close()
    return {"png": "shots/" + name + ".png", "size": im.size, "cols_found": len(xs), "rows_found": len(ys),
            "need": [ncol + 2, nrow + 2], "foreground_changed": fg_before != fg_after,
            # 바뀌었으면 어떤 종류의 창이 앞으로 왔는지(창 제목은 사용자 문서 이름일 수 있어 적지 않고 클래스 이름만 — 엑셀이면 XLMAIN)
            "fg_class": [cls(fg_before), cls(fg_after)] if fg_before != fg_after else None}


def cls(h):
    """창 클래스 이름(창이 이미 닫혔으면 '(닫힘)')."""
    if not h:
        return None
    try:
        return win32gui.GetClassName(h)
    except Exception:
        return "(닫힘)"


def merge(v, tol=1.2):
    out = []
    for x in v:
        if not out or x - out[-1] > tol:
            out.append(x)
    return out


def put_answers(ws, t, record=False):
    for top, fill, formula, _ in t["answers"]:
        ws.Range(fill).NumberFormat = GEN if t["fmt_after"].get(fill) == GEN or fill not in t["fmt"] else t["fmt"][fill]
        ws.Range(top).Formula = formula
        if fill != top:
            ws.Range(fill).FillDown()
        if record:
            AUTO_FMT[top] = ws.Range(top).NumberFormat  # 수식을 넣은 순간 엑셀이 붙인 표시 형식
    for rng, f in t["fmt_after"].items():
        ws.Range(rng).NumberFormat = f


def build_sheet(ws, with_answer):
    for t in TABLES:
        r = t["r"]
        ws.Range(f"A{r}").Value = t["title"]
        ws.Range(f"A{r}").Font.Bold = True
        c = [chr(ord("A") + i) for i in range(len(t["head"]))]
        hdr = ws.Range(f"A{r + 1}:{c[-1]}{r + 1}")
        hdr.Value = (tuple(t["head"]),)
        for rng, f in t["fmt"].items():
            ws.Range(rng).NumberFormat = f
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
        if with_answer:
            for cell, f in t.get("side_answer", {}).items():
                ws.Range(cell).Formula = f
                ws.Range(cell).HorizontalAlignment = -4108
        for rng in t["border_side"]:
            ws.Range(rng).Borders.LineStyle = 1
        for h in heads:
            h.Font.Bold = True
            h.Interior.Color = 0xEADBC8  # 연한 파랑(BGR)
            h.HorizontalAlignment = -4108
        for rng, f in t["fmt"].items():
            ws.Range(rng).NumberFormat = f
        if with_answer:
            put_answers(ws, t, record=True)
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
        if with_answer:
            R["auto_format_on_entry"] = dict(AUTO_FMT)
        ws0 = wb.Worksheets("계산작업")
        R.setdefault("input_check", {})[fname] = {"E3": ws0.Range("E3").Text, "H14": ws0.Range("H14").Text, "G17": ws0.Range("G17").Text,
                                                  "F25_formula": ws0.Range("F25").Formula, "F25_text": ws0.Range("F25").Text}
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
                                    "formulas": formulas(ws, fill), "values": got, "expect": expect,
                                    "value2": {c.Address.replace("$", ""): c.Value2 for c in ws.Range(fill)},
                                    "text_cells": sum(1 for c in ws.Range(fill) if isinstance(c.Value2, str)),
                                    "ok": list(got.values()) == expect and (fill not in t.get("num", set()) or all(not isinstance(c.Value2, str) for c in ws.Range(fill)))})
        R["problems"].append(item)
    # 그림 — 정답 결과 · 수식 보기
    for t in TABLES:
        R["shots"][f"{t['key']}_result"] = export(ws, t["area"], f"{t['key']}_result", t["boxes"])
        a = t["area"].replace(f"A{t['r']}:", f"A{t['r'] + 1}:")
        R["shots"][f"{t['key']}_formula"] = export(ws, a, f"{t['key']}_formula", t["boxes"], show_formulas=True, widths=FW[t["key"]])
    # 한 칸 사실 — 매번 General 로 되돌린 Z1 에 넣고, 보이는 값과 엑셀이 붙인 형식을 읽는다
    R["facts"] = []
    ws.Columns("Z").ColumnWidth = 24  # 긴 글자 · 날짜가 잘리지 않게
    for f, label in FACTS:
        try:
            ws.Range("Z1").Clear()
            ws.Range("Z1").NumberFormat = GEN
            ws.Range("Z1").Formula = f
            R["facts"].append({"formula": f, "label": label, "text": ws.Range("Z1").Text, "value": ws.Range("Z1").Value2,
                               "is_text": isinstance(ws.Range("Z1").Value2, str), "auto_format": ws.Range("Z1").NumberFormat})
        except pywintypes.com_error as e:
            R["facts"].append({"formula": f, "label": label, "rejected": str(e)[:120]})
    ws.Range("Z1").Clear()
    # 함정 — 같은 칸에 틀린 식(General 로 되돌린 뒤 넣어 엑셀이 붙이는 형식도 기록)
    for tr in TRAPS:
        t = tr["table"]
        item = {"key": tr["key"], "why": tr["why"], "formula": tr["formula"], "range": tr["fill"]}
        saved = {c: ws.Range(c).Formula for c in tr.get("set", {})}
        try:
            for c, v in tr.get("set", {}).items():
                ws.Range(c).Formula = v
            item["set"] = tr.get("set")
            ws.Range(tr["fill"]).NumberFormat = GEN
            ws.Range(tr["top"]).Formula = tr["formula"]
            for c, f in tr.get("also", {}).items():
                ws.Range(c).Formula = f
            if tr["fill"] != tr["top"] and ws.Range(tr["fill"]).Rows.Count > 1:  # 가로 한 줄은 채우지 않음(FillDown 은 위 줄을 복사함)
                ws.Range(tr["fill"]).FillDown()
            item["auto_format"] = ws.Range(tr["top"]).NumberFormat
            if tr.get("fmt"):
                ws.Range(tr["fill"]).NumberFormat = tr["fmt"]
                item["fmt_applied"] = tr["fmt"]
            item["formulas"] = formulas(ws, tr["fill"])
            item["values"] = texts(ws, tr["fill"])
            exp = tr.get("expect") or t["answers"][tr.get("exp_idx", 0)][3]
            item["wrong_vs_answer"] = sum(1 for a_, b_ in zip(item["values"].values(), exp) if a_ != b_)
            item["text_cells"] = sum(1 for c in ws.Range(tr["fill"]) if isinstance(c.Value2, str))
            if tr.get("sum"):
                ws.Range(tr["sum"]).Formula = f"=SUM({tr['fill']})"
                item["sum"] = {"cell": tr["sum"], "formula": ws.Range(tr["sum"]).Formula, "text": ws.Range(tr["sum"]).Text}
                ws.Range(tr["sum"]).Clear()
            item["shot"] = export(ws, t["area"], f"trap_{tr['key']}", [tr["fill"]])
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
        for c, f in saved.items():
            ws.Range(c).Formula = f
        put_answers(ws, t)  # 정답 식 · 형식으로 되돌림
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
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)
print(json.dumps({"ok": [[a["ok"] for a in p["answers"]] for p in R["problems"]], "props": R.get("file_props"),
                  "values": [[list(a["values"].values()) for a in p["answers"]] for p in R["problems"]],
                  "auto_fmt": R.get("auto_format_on_entry"), "input": R.get("input_check"),
                  "text_cells": [[a["text_cells"] for a in p["answers"]] for p in R["problems"]],
                  "traps": [(t["key"], t.get("rejected") or list(t.get("values", {}).values()), t.get("wrong_vs_answer"), t.get("text_cells"), t.get("sum")) for t in R["traps"]],
                  "fg": [(k, v.get("fg_class")) for k, v in list(R["shots"].items()) + [(t["key"], t.get("shot", {})) for t in R["traps"]] if v.get("foreground_changed")],
                  "facts": [(x["formula"], x.get("text", x.get("rejected")), x.get("is_text")) for x in R.get("facts", [])]},
                 ensure_ascii=False, indent=1, default=str))
