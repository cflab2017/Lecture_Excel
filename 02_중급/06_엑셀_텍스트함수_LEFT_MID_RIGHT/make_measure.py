"""엑셀 중급 6강 — 텍스트 함수 LEFT · MID · RIGHT · LEN · & (컴활 2급 실기 계산작업 유형 연습).
문제 · 데이터 · 문장은 모두 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제다. 이름은 지어낸 이름, 전화번호는 가상(010-0000-…).

하는 일
1. 실습.xlsx(문제 시트 + 계산작업 시트, 답 칸 비어 있음) · 정답.xlsx(답 칸에 수식) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 수식을 넣고(첫 칸에 넣은 뒤 아래로 채우기 = FillDown) 칸에 보이는 값(.Text)과 칸에 든 값의 종류(글자 · 숫자)를 셀 주소별로 읽기
   — 텍스트 함수의 결과는 숫자처럼 보여도 글자다. 보이는 값만 같고 종류가 다르면 틀린 것으로 센다
3. 자주 틀리는 식(함정)도 같은 칸에 넣어 결과 읽기 — 저장하지 않는 사본에서. 한 칸짜리 사실(글자 수 · 공백 · & 와 날짜 등)도 빈 칸에서 잼
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
TMP = tempfile.mkdtemp(prefix="xl_mid06_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제입니다. 이름 · 코드 · 전화번호(010-0000-…) · 금액은 모두 가상입니다."
YMD = "yyyy-mm-dd"
GEN = "G/표준"  # 한국어 엑셀의 '일반' 형식 코드(NumberFormat = "General" 은 거부됨, 11/18 실측)
DOW = "월화수목금토일"


def d(s):
    return datetime.date.fromisoformat(s)


# ── 문제 정의 ─────────────────────────────────────────────
WON = "#,##0"
CODES = [["KR-2024-A17", "무선 마우스"], ["JP-2023-B05", "키보드"], ["US-2025-A02", "모니터 받침"], ["CN-2022-C11", "USB 허브"],
         ["KR-2025-B08", "노트북 거치대"], ["VN-2024-A20", "웹캠"], ["DE-2021-C03", "헤드셋"], ["KR-2023-A14", "마우스 패드"]]
T16 = dict(
    key="q16", level="기본", r=1, title="[표16] 제품 코드",
    head=["제품코드", "제품명", "국가", "연도", "등급"],
    rows=CODES, side={},
    answers=[("C3", "C3:C10", "=LEFT(A3,2)", [c[:2] for c, _ in CODES]),
             ("D3", "D3:D10", "=MID(A3,4,4)", [c[3:7] for c, _ in CODES]),
             ("E3", "E3:E10", "=RIGHT(A3,3)", [c[-3:] for c, _ in CODES])],
    area="A1:E10", boxes=["C3:E10"], fmt={}, fmt_after={}, border_side=[],
    text=("[표16]에서 제품코드[A3:A10]의 앞 두 글자를 국가[C3:C10]에, 넷째 글자부터 네 글자를 연도[D3:D10]에, "
          "마지막 세 글자를 등급[E3:E10]에 표시하시오."),
    use="LEFT, MID, RIGHT 함수 사용",
)
PEOPLE = [["한별", "010-0000-5830"], ["가온솔", "010-0000-0476"], ["남궁하늘", "010-0000-7215"], ["이슬", "010-0000-1943"],
          ["다온빛", "010-0000-3361"], ["선우람", "010-0000-8829"], ["미르", "010-0000-4407"], ["해밀아", "010-0000-2150"]]
T17 = dict(
    key="q17", level="응용", r=12, title="[표17] 행사 참가자 명단",
    head=["이름", "전화번호", "가린 이름", "가린 번호"],
    rows=PEOPLE, side={},
    answers=[("C14", "C14:C21", '=LEFT(A14,1)&REPT("*",LEN(A14)-1)', [n[0] + "*" * (len(n) - 1) for n, _ in PEOPLE]),
             ("D14", "D14:D21", '=LEFT(B14,4)&"****"&RIGHT(B14,5)', [p[:4] + "****" + p[-5:] for _, p in PEOPLE])],
    area="A12:D21", boxes=["C14:D21"], fmt={}, fmt_after={}, border_side=[],
    text=("[표17]에서 이름[A14:A21]의 첫 글자만 남기고 나머지 글자 수만큼 '*'로 바꾸어 가린 이름[C14:C21]에 표시하고, "
          "전화번호[B14:B21]의 가운데 네 자리를 '****'로 바꾸어 가린 번호[D14:D21]에 표시하시오. (예: 가온솔 → 가**, 010-0000-0476 → 010-****-0476)"),
    use="① LEFT, REPT, LEN 함수와 & 연산자 ② LEFT, RIGHT 함수와 & 연산자 사용",
)
ORDERS = [["ORD-0903-01", "볼펜", 1500], ["ORD-0915-07", "노트", 3200], ["ORD-1002-02", "형광펜", 1200], ["ORD-1018-11", "파일철", 4500],
          ["ORD-1105-04", "가위", 3800], ["ORD-1121-09", "스테이플러", 12500], ["ORD-1201-05", "풀", 900], ["ORD-1214-10", "클립", 1000]]
T18 = dict(
    key="q18", level="함정", r=23, title="[표18] 주문 내역",
    head=["주문번호", "상품", "단가", "주문 월", "안내 문구"],
    rows=ORDERS, side={},
    answers=[("D25", "D25:D32", "=VALUE(MID(A25,5,2))", [str(int(o[4:6])) for o, _, _ in ORDERS]),
             ("E25", "E25:E32", '=B25&" "&TEXT(C25,"#,##0")&"원"', [f"{n} {p:,}원" for _, n, p in ORDERS])],
    num={"D25:D32"},  # 숫자로 들어 있어야 맞는 칸(보이는 값이 같아도 글자면 틀림)
    area="A23:E32", boxes=["D25:E32"], fmt={"C25:C32": WON}, fmt_after={}, border_side=[],
    text=("[표18]에서 주문번호[A25:A32]의 다섯째 글자부터 두 글자를 숫자로 바꾸어 주문 월[D25:D32]에 표시하고(예: 09 → 9), "
          "상품[B25:B32] · 공백 · 단가[C25:C32](천 단위 쉼표) · '원'을 이어 안내 문구[E25:E32]에 표시하시오. (예: 볼펜 1,500원)"),
    use="① MID, VALUE 함수 ② & 연산자와 TEXT 함수 사용",
)
TABLES = [T16, T17, T18]
# 함정 — exp_idx 는 그 표의 answers 중 몇 번째와 견주는지. sum 이 있으면 그 칸에 SUM 을 넣어 읽는다
TRAPS = [
    dict(key="q16_left_default", table=T16, top="C3", fill="C3:C10", exp_idx=0, why="LEFT 의 글자 수를 빼면 한 글자만",
         formula="=LEFT(A3)"),
    dict(key="q16_mid_start", table=T16, top="D3", fill="D3:D10", exp_idx=1, why="MID 시작 위치를 3 으로 → 하이픈부터",
         formula="=MID(A3,3,4)"),
    dict(key="q16_right_count", table=T16, top="E3", fill="E3:E10", exp_idx=2, why="RIGHT 글자 수를 2 로",
         formula="=RIGHT(A3,2)"),
    dict(key="q17_fixed_stars", table=T17, top="C14", fill="C14:C21", exp_idx=0, why="'*' 를 두 개로 고정 → 두 글자 · 네 글자 이름이 틀림",
         formula='=LEFT(A14,1)&"**"'),
    dict(key="q17_no_hyphen", table=T17, top="D14", fill="D14:D21", exp_idx=1, why="LEFT 3 · RIGHT 4 → 하이픈이 빠짐",
         formula='=LEFT(B14,3)&"****"&RIGHT(B14,4)'),
    dict(key="q17_plus", table=T17, top="D14", fill="D14:D21", exp_idx=1, why="& 대신 + 로 이음",
         formula='=LEFT(B14,4)+"****"+RIGHT(B14,5)'),
    dict(key="q18_no_value", table=T18, top="D25", fill="D25:D32", exp_idx=0, why="VALUE 없이 MID 만 → 숫자처럼 보이는 글자",
         formula="=MID(A25,5,2)", sum="G25"),
    dict(key="q18_value_ok_sum", table=T18, top="D25", fill="D25:D32", exp_idx=0, why="(비교용) 정답 식의 합계",
         formula="=VALUE(MID(A25,5,2))", sum="G25"),
    dict(key="q18_no_text", table=T18, top="E25", fill="E25:E32", exp_idx=1, why="TEXT 없이 단가를 그대로 이음 → 쉼표가 사라짐",
         formula='=B25&" "&C25&"원"'),
    dict(key="q18_no_space", table=T18, top="E25", fill="E25:E32", exp_idx=1, why="공백(\" \")을 빠뜨림",
         formula='=B25&TEXT(C25,"#,##0")&"원"'),
]
# 한 칸 사실 — 빈 칸(Z1)에 넣고 읽는다(보이는 값 + 칸에 든 값의 종류)
FACTS = [
    ('=LEN("KR-2024-A17")', "하이픈까지 센 글자 수"), ('=LEN("가온솔")', "한글 한 글자 = 1"), ('=LEN(" 가온솔")', "앞 공백도 한 글자"),
    ('=LEFT("가온솔")', "글자 수를 빼면 1"), ('=MID("KR-2024-A17",4,100)', "글자 수가 넘치면 끝까지"), ('=RIGHT("KR",5)', "글자보다 많이 달라고 하면"),
    ('=MID("KR-2024-A17",0,2)', "시작 위치 0"), ('=ISTEXT(MID("KR-2024-A17",4,4))', "MID 결과는 글자인가"),
    ('=MID("KR-2024-A17",4,4)+1', "글자 숫자에 + 1"), ('=VALUE("09")', "VALUE 는 앞 0 을 떼고 숫자로"),
    ('=TRIM("  가온   솔 ")', "TRIM 은 앞뒤 공백을 지우고 가운데는 하나로"), ('=12500&"원"', "숫자를 & 로 이으면 쉼표 없음"),
    ('=TEXT(12500,"#,##0")&"원"', "TEXT 로 쉼표"), ('=DATE(2026,11,19)&""', "날짜를 & 로 이으면"), ('=TEXT(DATE(2026,11,19),"yyyy-mm-dd")&"까지"', "TEXT 로 날짜 모양"),
    ('=REPT("*",3)', "REPT 는 같은 글자를 반복"),
]
WIDTHS = dict(zip("ABCDEFG", (13, 13, 11, 15, 20, 9, 9)))
FW = {"q16": {"C": 24, "D": 26, "E": 26}, "q17": {"C": 64, "D": 64}, "q18": {"D": 40, "E": 64}}
FIRST_N = 16
LESSON = "엑셀 중급 6강 — 텍스트 함수 LEFT · MID · RIGHT · LEN · & (컴활 2급 실기 계산작업 유형 연습)"
LESSON_NO = "엑셀 중급 6강"
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
        R.setdefault("input_check", {})[fname] = {"A3": ws0.Range("A3").Text, "B14": ws0.Range("B14").Text,
                                                  "C25_text": ws0.Range("C25").Text, "C25_value2": ws0.Range("C25").Value2}
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
            item["text_cells"] = sum(1 for c in ws.Range(tr["fill"]) if isinstance(c.Value2, str))
            if tr.get("sum"):
                ws.Range(tr["sum"]).Formula = f"=SUM({tr['fill']})"
                item["sum"] = {"cell": tr["sum"], "formula": ws.Range(tr["sum"]).Formula, "text": ws.Range(tr["sum"]).Text}
                ws.Range(tr["sum"]).Clear()
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
                  "text_cells": [[a["text_cells"] for a in p["answers"]] for p in R["problems"]],
                  "traps": [(t["key"], t.get("rejected") or list(t.get("values", {}).values()), t.get("wrong_vs_answer"), t.get("text_cells"), t.get("sum")) for t in R["traps"]],
                  "fg": [(k, v.get("fg_class")) for k, v in list(R["shots"].items()) + [(t["key"], t.get("shot", {})) for t in R["traps"]] if v.get("foreground_changed")],
                  "facts": [(x["formula"], x.get("text", x.get("rejected")), x.get("is_text")) for x in R.get("facts", [])]},
                 ensure_ascii=False, indent=1, default=str))
