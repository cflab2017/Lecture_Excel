"""엑셀 중급 5강 — 날짜 함수 DATE · YEAR · MONTH · DAY · WEEKDAY (컴활 2급 실기 계산작업 유형 연습).
문제 · 데이터 · 문장은 모두 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제다. 이름은 지어낸 이름.

하는 일
1. 실습.xlsx(문제 시트 + 계산작업 시트, 답 칸 비어 있음) · 정답.xlsx(답 칸에 수식) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 수식을 넣고(첫 칸에 넣은 뒤 아래로 채우기 = FillDown) 칸에 보이는 값(.Text)을 셀 주소별로 읽기
   — 수식을 넣은 순간 엑셀이 스스로 붙인 표시 형식(자동 서식)도 칸마다 기록한다(날짜 함수는 결과가 날짜로 보이는 일이 잦다)
3. 자주 틀리는 식(함정)도 같은 칸에 넣어 결과 읽기 — 저장하지 않는 사본에서. 한 칸짜리 사실(DATE 13월 · 두 자리 연도 등)도 빈 칸에서 잼
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 수식 입력줄 · 함수 인수 창 화면은 이 스크립트로 만들지 않는다.
날짜는 TODAY() 대신 기준일 칸을 쓴다 — TODAY 는 여는 날마다 값이 바뀌어 실측 값을 적어 둘 수 없다.
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
TMP = tempfile.mkdtemp(prefix="xl_mid05_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제입니다. 이름 · 날짜 · 숫자는 모두 가상입니다."
YMD = "yyyy-mm-dd"
GEN = "G/표준"  # 한국어 엑셀의 '일반' 형식 코드(NumberFormat = "General" 은 거부됨, 11/18 실측)
DOW = "월화수목금토일"


def d(s):
    return datetime.date.fromisoformat(s)


# ── 문제 정의 ─────────────────────────────────────────────
BASE = "2026-12-31"
HIRE = [["고은별", "2015-03-02"], ["노을빛", "2019-11-18"], ["도담이", "2021-07-01"], ["로아린", "2008-01-14"],
        ["모아름", "2024-02-26"], ["보라솔", "2012-09-03"], ["서하늘", "2026-04-01"], ["예슬기", "2017-12-11"]]
T13 = dict(
    key="q13", level="기본", r=1, title="[표13] 직원 입사 현황",
    head=["이름", "입사일", "입사 연도", "입사 월", "근속 연수"],
    rows=HIRE,
    side={"G2": ("기준일", True), "G3": (BASE, False)},
    answers=[("C3", "C3:C10", "=YEAR(B3)", [str(d(s).year) for _, s in HIRE]),
             ("D3", "D3:D10", "=MONTH(B3)", [str(d(s).month) for _, s in HIRE]),
             ("E3", "E3:E10", "=YEAR($G$3)-YEAR(B3)", [str(d(BASE).year - d(s).year) for _, s in HIRE])],
    area="A1:G10", boxes=["C3:E10"], fmt={"B3:B10": YMD, "G3": YMD}, fmt_after={"C3:E10": GEN}, border_side=["G2:G3"],
    text=("[표13]에서 입사일[B3:B10]을 이용하여 입사 연도[C3:C10]와 입사 월[D3:D10]을 표시하고, "
          "기준일[G3]의 연도에서 입사 연도를 빼서 근속 연수[E3:E10]를 구하시오."),
    use="YEAR, MONTH 함수 사용",
)
EVENTS = [["신입 환영회", 2026, 3, 7], ["봄 소풍", 2026, 4, 18], ["체육 대회", 2026, 5, 23], ["여름 캠프", 2026, 7, 30],
          ["가을 공연", 2026, 10, 9], ["바자회", 2026, 11, 21], ["송년회", 2026, 12, 26], ["새해 모임", 2027, 1, 2]]
T14 = dict(
    key="q14", level="응용", r=12, title="[표14] 동아리 행사 일정",
    head=["행사", "년", "월", "일", "날짜", "요일"],
    rows=EVENTS,
    side={},
    answers=[("E14", "E14:E21", "=DATE(B14,C14,D14)", [datetime.date(y, m, dd).isoformat() for _, y, m, dd in EVENTS]),
             ("F14", "F14:F21", '=CHOOSE(WEEKDAY(E14,2),"월","화","수","목","금","토","일")',
              [DOW[datetime.date(y, m, dd).weekday()] for _, y, m, dd in EVENTS])],
    area="A12:F21", boxes=["E14:F21"], fmt={"E14:E21": YMD}, fmt_after={"E14:E21": YMD}, border_side=[],
    text=("[표14]에서 년[B14:B21] · 월[C14:C21] · 일[D14:D21]을 이용하여 날짜[E14:E21]를 구하고, "
          "그 날짜의 요일을 '월'~'일'로 요일[F14:F21]에 표시하시오. (단, 요일은 월요일이 1 인 유형 사용)"),
    use="① DATE 함수 ② CHOOSE, WEEKDAY 함수 사용",
)
STAYS = [["하람결", "2026-12-04", 2], ["이든솔", "2026-12-07", 1], ["다올이", "2026-12-12", 3], ["소리빛", "2026-12-15", 7],
         ["하예린", "2026-12-19", 14], ["윤슬아", "2027-01-04", 31], ["미르샘", "2027-01-10", 32], ["온새미", "2027-02-01", 45]]
STAY_ROWS = [[n, s, (d(s) + datetime.timedelta(days=k)).isoformat()] for n, s, k in STAYS]
T15 = dict(
    key="q15", level="함정", r=23, title="[표15] 숙소 예약",
    head=["예약자", "체크인", "체크아웃", "숙박 일수", "요금 구분"],
    rows=STAY_ROWS,
    side={},
    answers=[("D25", "D25:D32", "=C25-B25", [str(k) for _, _, k in STAYS]),
             ("E25", "E25:E32", '=IF(WEEKDAY(B25,2)>=6,"주말","평일")', ["주말" if d(s).weekday() >= 5 else "평일" for _, s, _ in STAYS])],
    area="A23:E32", boxes=["D25:E32"], fmt={"B25:C32": YMD}, fmt_after={"D25:D32": GEN}, border_side=[],
    text=("[표15]에서 체크아웃[C25:C32]에서 체크인[B25:B32]을 빼서 숙박 일수[D25:D32]를 숫자로 구하고, "
          "체크인이 토요일 또는 일요일이면 '주말', 아니면 '평일'을 요금 구분[E25:E32]에 표시하시오."),
    use="① 날짜 빼기(숫자로 표시) ② IF, WEEKDAY 함수 사용",
)
TABLES = [T13, T14, T15]
# 함정 — exp_idx 는 그 표의 answers 중 몇 번째와 견주는지. fmt 가 있으면 수식을 넣은 뒤 그 표시 형식을 입힌다
TRAPS = [
    dict(key="q13_no_dollar", table=T13, top="E3", fill="E3:E10", exp_idx=2, why="기준일에 $ 를 빼고 채움 → 아래 칸은 빈 칸(G4…)의 연도",
         formula="=YEAR(G3)-YEAR(B3)"),
    dict(key="q13_days", table=T13, top="E3", fill="E3:E10", exp_idx=2, why="YEAR 없이 날짜끼리 빼기 → 연수가 아니라 날짜 수",
         formula="=$G$3-B3"),
    dict(key="q13_date_format", table=T13, top="C3", fill="C3:C10", exp_idx=0, why="입사 연도 칸이 날짜 형식(입사일 열 서식을 복사해 옴)",
         formula="=YEAR(B3)", fmt=YMD),
    dict(key="q14_weekday_type1", table=T14, top="F14", fill="F14:F21", exp_idx=1, why="WEEKDAY 유형을 빼서 일요일이 1 → 한 칸씩 밀림",
         formula='=CHOOSE(WEEKDAY(E14),"월","화","수","목","금","토","일")'),
    dict(key="q14_weekday_only", table=T14, top="F14", fill="F14:F21", exp_idx=1, why="CHOOSE 없이 WEEKDAY 만 → 요일 이름 대신 숫자",
         formula="=WEEKDAY(E14,2)"),
    dict(key="q14_text_join", table=T14, top="E14", fill="E14:E21", exp_idx=0, why="DATE 대신 & 로 이어 붙임 → 날짜가 아니라 글자",
         formula='=B14&"-"&C14&"-"&D14'),
    dict(key="q15_day", table=T15, top="D25", fill="D25:D32", exp_idx=0, why="숙박 일수를 DAY(체크아웃-체크인)로 → 31일 넘으면 틀림",
         formula="=DAY(C25-B25)"),
    dict(key="q15_date_format", table=T15, top="D25", fill="D25:D32", exp_idx=0, why="숙박 일수 칸이 날짜 형식으로 남음",
         formula="=C25-B25", fmt=YMD),
    dict(key="q15_reverse", table=T15, top="D25", fill="D25:D32", exp_idx=0, why="빼는 순서를 거꾸로(체크인-체크아웃)",
         formula="=B25-C25"),
    dict(key="q15_weekday_type1", table=T15, top="E25", fill="E25:E32", exp_idx=1, why="WEEKDAY 유형을 빼서 >=6 → 금 · 토가 주말, 일요일은 평일",
         formula='=IF(WEEKDAY(B25)>=6,"주말","평일")'),
]
# 한 칸 사실 — 빈 칸(Z1)에 넣고 읽는다(General 칸에 넣었을 때 보이는 값 + 엑셀이 붙인 표시 형식)
FACTS = [
    ("=DATE(2026,13,1)", "13월 → 다음 해 1월"), ("=DATE(2026,2,30)", "2월 30일 → 3월로 넘어감"), ("=DATE(26,5,5)", "두 자리 연도 26"),
    ("=DATE(2026,3,0)", "0일 = 앞 달 마지막 날"), ("=DAY(DATE(2026,3,0))", "그 달이 며칠까지인지"),
    ("=WEEKDAY(DATE(2026,11,18))", "유형 생략(일=1)"), ("=WEEKDAY(DATE(2026,11,18),2)", "유형 2(월=1)"), ("=WEEKDAY(DATE(2026,11,18),3)", "유형 3(월=0)"),
    ("=YEAR(0)", "빈 칸 · 0 의 연도"), ("=MONTH(0)", "빈 칸 · 0 의 월"), ("=DAY(0)", "빈 칸 · 0 의 일"),
    ('=YEAR("2026-11-18")', "글자 날짜도 읽음"), ('="2026-11-20"-"2026-11-18"', "글자 날짜끼리 빼기"),
    ("=DATE(2026,11,18)+30", "날짜 + 30"), ("=DATE(2026,11,18)-DATE(2026,1,1)", "날짜 - 날짜"),
]
WIDTHS = dict(zip("ABCDEFG", (11, 12, 12, 10, 12, 9, 12)))
FW = {"q13": {"C": 24, "D": 24, "E": 40}, "q14": {"E": 40, "F": 86}, "q15": {"D": 24, "E": 64}}
FIRST_N = 13
LESSON = "엑셀 중급 5강 — 날짜 함수 DATE · YEAR · MONTH · DAY · WEEKDAY (컴활 2급 실기 계산작업 유형 연습)"
LESSON_NO = "엑셀 중급 5강"
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
            "fg_class": [win32gui.GetClassName(fg_before) if fg_before else None, win32gui.GetClassName(fg_after) if fg_after else None] if fg_before != fg_after else None}


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
        R.setdefault("input_check", {})[fname] = {"B3_value2": ws0.Range("B3").Value2, "B3_text": ws0.Range("B3").Text,
                                                  "G3_text": ws0.Range("G3").Text, "B25_text": ws0.Range("B25").Text}
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
                                    "ok": list(got.values()) == expect})
        R["problems"].append(item)
    # 그림 — 정답 결과 · 수식 보기
    for t in TABLES:
        R["shots"][f"{t['key']}_result"] = export(ws, t["area"], f"{t['key']}_result", t["boxes"])
        a = t["area"].replace(f"A{t['r']}:", f"A{t['r'] + 1}:")
        R["shots"][f"{t['key']}_formula"] = export(ws, a, f"{t['key']}_formula", t["boxes"], show_formulas=True, widths=FW[t["key"]])
    # 한 칸 사실 — 매번 General 로 되돌린 Z1 에 넣고, 보이는 값과 엑셀이 붙인 형식을 읽는다
    R["facts"] = []
    ws.Columns("Z").ColumnWidth = 14  # 날짜가 ####### 로 안 보이게
    for f, label in FACTS:
        try:
            ws.Range("Z1").Clear()
            ws.Range("Z1").NumberFormat = GEN
            ws.Range("Z1").Formula = f
            R["facts"].append({"formula": f, "label": label, "text": ws.Range("Z1").Text, "value": ws.Range("Z1").Value2,
                               "auto_format": ws.Range("Z1").NumberFormat})
        except pywintypes.com_error as e:
            R["facts"].append({"formula": f, "label": label, "rejected": str(e)[:120]})
    ws.Range("Z1").Clear()
    # 함정 — 같은 칸에 틀린 식(General 로 되돌린 뒤 넣어 엑셀이 붙이는 형식도 기록)
    for tr in TRAPS:
        t = tr["table"]
        item = {"key": tr["key"], "why": tr["why"], "formula": tr["formula"], "range": tr["fill"]}
        try:
            ws.Range(tr["fill"]).NumberFormat = GEN
            ws.Range(tr["top"]).Formula = tr["formula"]
            if tr["fill"] != tr["top"]:
                ws.Range(tr["fill"]).FillDown()
            item["auto_format"] = ws.Range(tr["top"]).NumberFormat
            if tr.get("fmt"):
                ws.Range(tr["fill"]).NumberFormat = tr["fmt"]
                item["fmt_applied"] = tr["fmt"]
            item["formulas"] = formulas(ws, tr["fill"])
            item["values"] = texts(ws, tr["fill"])
            exp = t["answers"][tr.get("exp_idx", 0)][3]
            item["wrong_vs_answer"] = sum(1 for a_, b_ in zip(item["values"].values(), exp) if a_ != b_)
            item["shot"] = export(ws, t["area"], f"trap_{tr['key']}", [tr["fill"]])
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
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
                  "traps": [(t["key"], t.get("auto_format"), t.get("rejected") or list(t.get("values", {}).values()), t.get("wrong_vs_answer")) for t in R["traps"]],
                  "fg": [(k, v.get("fg_class")) for k, v in list(R["shots"].items()) + [(t["key"], t.get("shot", {})) for t in R["traps"]] if v.get("foreground_changed")],
                  "facts": [(x["formula"], x.get("text", x.get("rejected")), x.get("auto_format")) for x in R.get("facts", [])]},
                 ensure_ascii=False, indent=1, default=str))
