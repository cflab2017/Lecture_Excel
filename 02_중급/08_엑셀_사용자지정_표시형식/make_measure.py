"""엑셀 중급 8강 — 사용자 지정 표시 형식(셀 서식 · 컴활 2급 실기 기본작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님). 이름 · 금액 · 날짜는 지어낸 것.

하는 일
1. 실습.xlsx(문제 시트 + 기본작업 시트, 답 범위는 처음 형식 그대로) · 정답.xlsx(답 범위에 사용자 지정 표시 형식) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 형식 코드를 넣고(NumberFormatLocal — 한국어 엑셀의 셀 서식 '형식' 칸에 치는 것과 같은 글자)
   칸에 보이는 값(.Text) · 칸에 든 값(Value2 — 형식을 바꿔도 그대로인지) · 화면 글자색(DisplayFormat.Font.Color) · 엑셀이 저장한 형식 코드를 셀 주소별로 읽기
3. 자주 틀리는 형식 코드(함정)도 같은 범위에 넣어 결과 읽기 — 값을 글자로 직접 쳐 넣은 경우(평균이 어떻게 되나)도. 저장하지 않는 사본에서
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png
   '값 보기' 그림은 수식 보기(DisplayFormulas)로 내보내 형식 없이 칸에 든 값이 보이게 한다

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 셀 서식 대화 상자 화면은 이 스크립트로 만들지 않는다.
결과: 실습.xlsx · 정답.xlsx · measure.json · shots/*.png (shots 는 저장소에 올리지 않음)
실행: python make_measure.py   (pywin32 · pymupdf · pillow)"""
import datetime
import decimal
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
TMP = tempfile.mkdtemp(prefix="xl_mid08_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 나이 · 금액 · 날짜는 모두 가상입니다."
GEN = "G/표준"  # 한국어 엑셀의 '일반' 형식 코드(NumberFormat = "General" 은 거부됨, 11/18 실측)
DOW = "월화수목금토일"
RED = 255  # RGB(255,0,0) — 엑셀 Color 값(BGR 정수)


def serial(s):
    return (datetime.date.fromisoformat(s) - datetime.date(1899, 12, 30)).days


def xl_round(x):
    return int(decimal.Decimal(str(x)).quantize(decimal.Decimal("1"), rounding=decimal.ROUND_HALF_UP))


# ── 문제 정의 ─────────────────────────────────────────────
MEM = [["가람", 34, 30000], ["나래", 28, 30000], ["다솜", 41, 0], ["로운", 37, 45000],
       ["마루", 26, 30000], ["보미", 45, 0], ["새롬", 30, 45000], ["아라", 52, 30000]]
F_AGE = '0"세"'
F_WON = '#,##0"원"'
T22 = dict(
    key="q22", level="기본", r=1, title="[표22] 동호회 회원 회비",
    head=["이름", "나이", "회비"], rows=MEM,
    start_fmt={},
    answers=[("B3:B10", F_AGE, [f"{r[1]}세" for r in MEM]),
             ("C3:C10", F_WON, [f"{r[2]:,}원" for r in MEM])],
    side={"E2": ("평균 나이(확인)", True)}, side_formula={"E3": ("=AVERAGE(B3:B10)", "0.0")},
    area="A1:E10", boxes=["B3:C10"], border_side=["E2:E3"],
    text=("[표22]에서 나이[B3:B10]는 숫자 뒤에 '세'를 붙여 '34세'처럼, 회비[C3:C10]는 천 단위 구분 기호와 '원'을 붙여 '30,000원'처럼 표시하시오. "
          "(회비가 0 이면 '0원'으로 표시)"),
    use="셀 서식(Ctrl+1)의 사용자 지정 표시 형식 사용",
)
EVENTS = [["송년회", "2026-12-18", 1250000], ["워크숍", "2026-12-04", 860000], ["봉사활동", "2026-12-12", 320000],
          ["신년회", "2027-01-08", 2480000], ["세미나", "2027-01-22", 540000], ["체육대회", "2027-02-20", 1975500],
          ["시상식", "2027-02-26", 1234560], ["견학", "2027-03-05", 698400]]
F_DATE = "yyyy-mm-dd(aaa)"
F_K = '#,##0,"천원"'
T23 = dict(
    key="q23", level="응용", r=12, title="[표23] 모임 행사 일정",
    head=["행사", "날짜", "예산"], rows=[[e[0], serial(e[1]), e[2]] for e in EVENTS],
    start_fmt={"B14:B21": "yyyy-mm-dd"},
    answers=[("B14:B21", F_DATE, [f"{e[1]}({DOW[datetime.date.fromisoformat(e[1]).weekday()]})" for e in EVENTS]),
             ("C14:C21", F_K, [f"{xl_round(e[2] / 1000):,}천원" for e in EVENTS])],
    side={"E13": ("예산 합계(확인)", True)}, side_formula={"E14": ("=SUM(C14:C21)", "#,##0")},
    area="A12:E21", boxes=["B14:C21"], border_side=["E13:E14"],
    text=("[표23]에서 날짜[B14:B21]는 '2026-12-18(금)'처럼 연-월-일 뒤 괄호 안에 요일을, "
          "예산[C14:C21]은 천 원 단위로 줄여 '1,250천원'처럼 표시하시오."),
    use="사용자 지정 표시 형식 사용(요일은 aaa)",
)
BR = [["서울", 12520, 1120], ["부산", 9800, -450], ["대구", 7600, 0], ["인천", 11050, 305],
      ["광주", 6400, -1080], ["대전", 8200, 0], ["울산", 4550, -15], ["수원", 13300, 2100]]
F_SHOP = '@"점"'
F_DELTA = '#,##0;[빨강]-#,##0;"-"'


def delta_txt(n):
    return f"{n:,}" if n > 0 else ("-" if n == 0 else f"-{abs(n):,}")


T24 = dict(
    key="q24", level="함정", r=23, title="[표24] 지점별 이번 달 판매",
    head=["지점", "판매", "전월 대비"], rows=BR,
    start_fmt={"B25:B32": "#,##0"},
    answers=[("A25:A32", F_SHOP, [f"{r[0]}점" for r in BR]),
             ("C25:C32", F_DELTA, [delta_txt(r[2]) for r in BR])],
    red={"C25:C32": [r[2] < 0 for r in BR]},  # 이 칸들은 빨강이어야 함
    side={}, side_formula={},
    area="A23:C32", boxes=["A25:A32", "C25:C32"], border_side=[],
    text=("[표24]에서 지점[A25:A32]은 이름 뒤에 '점'을 붙여 '서울점'처럼, 전월 대비[C25:C32]는 양수는 '1,120'처럼, "
          "음수는 빨강으로 '-450'처럼, 0 은 '-'로 표시하시오."),
    use="사용자 지정 표시 형식 사용(양수;음수;0 구역)",
)
TABLES = [T22, T23, T24]
# 함정 — fmt 를 그 범위에 넣거나(NumberFormatLocal), set 으로 칸 값을 잠깐 바꿔 재고 되돌린다. watch = 같이 읽을 칸
TRAPS = [
    dict(key="q22_typed_text", table=T22, range="B3:B10", exp_idx=0, watch=["E3"],
         set={f"B{3 + i}": f"{r[1]}세" for i, r in enumerate(MEM)}, why="형식 대신 '34세'를 글자로 직접 입력 — 보이는 값은 같지만 숫자가 아님"),
    dict(key="q22_hash_zero", table=T22, range="C3:C10", exp_idx=1, fmt='#,###"원"', why="0 자리에 # — 값이 0 이면 숫자가 안 보임"),
    dict(key="q23_ddd", table=T23, range="B14:B21", exp_idx=0, fmt="yyyy-mm-dd(ddd)", why="요일을 ddd 로 — 영어 요일"),
    dict(key="q23_aaaa", table=T23, range="B14:B21", exp_idx=0, fmt="yyyy-mm-dd(aaaa)", why="요일을 aaaa 로 — '금요일'까지"),
    dict(key="q23_no_comma", table=T23, range="C14:C21", exp_idx=1, fmt='#,##0"천원"', why="끝 쉼표를 빼먹음 — 줄지 않고 '천원'만 붙음"),
    dict(key="q23_two_comma", table=T23, range="C14:C21", exp_idx=1, fmt='#,##0,,"천원"', why="쉼표 둘 — 백만 단위로 줄어듦"),
    dict(key="q24_text_zero", table=T24, range="A25:A32", exp_idx=0, fmt='0"점"', why="글자 칸에 숫자 형식(0) — 글자에는 안 붙음"),
    dict(key="q24_order", table=T24, range="C25:C32", exp_idx=1, fmt='[빨강]-#,##0;#,##0;"-"', why="구역 순서를 바꿈 — 첫 구역은 양수"),
    dict(key="q24_two_sections", table=T24, range="C25:C32", exp_idx=1, fmt="#,##0;[빨강]-#,##0", why="0 구역을 빼먹음 — 0 이 0 으로"),
    dict(key="q24_no_minus", table=T24, range="C25:C32", exp_idx=1, fmt='#,##0;[빨강]#,##0;"-"', why="음수 구역에 - 를 안 붙임 — 빨강만, 부호 없음"),
    dict(key="q24_red_en", table=T24, range="C25:C32", exp_idx=1, fmt='#,##0;[Red]-#,##0;"-"', why="한국어 엑셀 형식 칸에 영어 [Red]"),
]
# 한 칸 사실 — 빈 칸(Z1)에 형식을 넣고 값(또는 식)을 넣어 읽는다
FACTS = [
    dict(fmt='#,##0,"천원"', value=1975500, label="천 단위로 줄이면 반올림해서 보임"),
    dict(fmt='#,##0,"천원"', formula="=SUM(C14:C21)", label="예산 합계를 천원으로(실제 값으로 더함)"),
    dict(fmt="#,##0,,", value=1234567, label="쉼표 둘 = 백만 단위"),
    dict(fmt="aaa", value=serial("2026-12-18"), label="aaa = 요일 한 글자"),
    dict(fmt="aaaa", value=serial("2026-12-18"), label="aaaa = 요일 전체"),
    dict(fmt="ddd", value=serial("2026-12-18"), label="ddd = 영어 요일 줄임"),
    dict(fmt='yyyy"년" m"월" d"일"', value=serial("2026-12-04"), label="한글 날짜 · m d 하나는 앞 0 없음"),
    dict(fmt="yyyy-mm-dd", value=serial("2026-12-04"), label="mm dd 는 앞 0"),
    dict(fmt="0세", value=45, label="따옴표 없이 세"),
    dict(fmt="#,##0원", value=30000, label="따옴표 없이 원"),
    dict(fmt="000", value=45, label="0 의 개수만큼 자리(앞 0)"),
    dict(fmt="0", value=12.5, label="소수가 반올림돼 보임(값은 12.5)"),
    dict(fmt='@"님"', value="가람", label="글자 뒤에 붙이기 @"),
    dict(fmt="#,##0;-#,##0;", value=0, label="0 구역을 비우면 0 이 안 보임"),
    dict(fmt="#,##0;(#,##0)", value=-450, label="음수를 괄호로"),
    dict(fmt='0"세"', formula='="34세"+1', label="글자 '34세'에 더하기"),
    dict(fmt="#,##0;[빨간색]-#,##0", value=-450, label="색 이름 [빨간색](MS 도움말 예시의 이름)"),
    dict(fmt="[파랑]#,##0;[빨강]-#,##0", value=1120, label="양수 파랑"),
]
WIDTHS = dict(zip("ABCDE", (11, 17, 17, 3, 16)))
FIRST_N = 22
LESSON = "엑셀 중급 8강 — 사용자 지정 표시 형식 (컴활 2급 실기 기본작업 유형 연습)"
LESSON_NO = "엑셀 중급 8강"


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


def addr(c):
    return c.Address.replace("$", "")


def texts(ws, rng):
    return {addr(c): c.Text for c in ws.Range(rng)}


def colors(ws, rng):
    out = {}
    for c in ws.Range(rng):
        try:
            out[addr(c)] = int(c.DisplayFormat.Font.Color)
        except Exception as e:  # noqa: BLE001
            out[addr(c)] = f"?{str(e)[:40]}"
    return out


def values(ws, rng):
    return {addr(c): c.Value2 for c in ws.Range(rng)}


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


def put_formats(ws, t, answer):
    """처음 형식(start_fmt · 나머지는 G/표준)으로, answer 면 그 위에 정답 형식을."""
    for rng, _, _ in t["answers"]:
        ws.Range(rng).NumberFormatLocal = t["start_fmt"].get(rng, GEN)
    if answer:
        for rng, f, _ in t["answers"]:
            ws.Range(rng).NumberFormatLocal = f


def put_rows(ws, t):
    c = [chr(ord("A") + i) for i in range(len(t["head"]))]
    for i, row in enumerate(t["rows"]):
        for j, v in enumerate(row):
            ws.Range(f"{c[j]}{first(t) + i}").Value2 = v


def build_sheet(ws, with_answer):
    for t in TABLES:
        r = t["r"]
        ws.Range(f"A{r}").Value = t["title"]
        ws.Range(f"A{r}").Font.Bold = True
        c = [chr(ord("A") + i) for i in range(len(t["head"]))]
        hdr = ws.Range(f"A{r + 1}:{c[-1]}{r + 1}")
        hdr.Value = (tuple(t["head"]),)
        for rng, f in t["start_fmt"].items():
            ws.Range(rng).NumberFormatLocal = f
        put_rows(ws, t)
        ws.Range(f"A{r + 1}:{c[-1]}{last(t)}").Borders.LineStyle = 1
        ws.Range(f"A{first(t)}:A{last(t)}").HorizontalAlignment = -4108
        heads = [hdr]
        for cell, (v, is_head) in t["side"].items():
            ws.Range(cell).Value = v
            ws.Range(cell).HorizontalAlignment = -4108
            if is_head:
                heads.append(ws.Range(cell))
        for cell, (f, nf) in t["side_formula"].items():
            ws.Range(cell).NumberFormatLocal = nf
            ws.Range(cell).Formula = f
        for rng in t["border_side"]:
            ws.Range(rng).Borders.LineStyle = 1
        for h in heads:
            h.Font.Bold = True
            h.Interior.Color = 0xEADBC8  # 연한 파랑(BGR)
            h.HorizontalAlignment = -4108
        put_formats(ws, t, with_answer)
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
        wp.Name, wc.Name = "문제", "기본작업"
        build_problem_sheet(wp)
        build_sheet(wc, with_answer)
        wc.Activate()
        wc.Range("A1").Select()
        author(wb)
        return wb

    paths = {}
    for with_answer, fname in ((False, "실습.xlsx"), (True, "정답.xlsx")):
        wb = new_book(with_answer)
        ws0 = wb.Worksheets("기본작업")
        R.setdefault("input_check", {})[fname] = {k: ws0.Range(k).Text for k in ("B3", "C5", "B14", "C14", "A25", "C26", "E3", "E14")}
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
    ws = wb.Worksheets("기본작업")
    for n, t in enumerate(TABLES, start=FIRST_N):
        item = {"n": n, "key": t["key"], "level": t["level"], "text": t["text"], "use": t["use"], "answers": [],
                "side": {c: {"formula": ws.Range(c).Formula, "text": ws.Range(c).Text} for c in t["side_formula"]}}
        raw = {}
        cl = [chr(ord("A") + i) for i in range(len(t["head"]))]
        for i, row in enumerate(t["rows"]):
            for j, v in enumerate(row):
                raw[f"{cl[j]}{first(t) + i}"] = v
        for rng, f, expect in t["answers"]:
            got = texts(ws, rng)
            v2 = values(ws, rng)
            col = colors(ws, rng)
            want_red = t.get("red", {}).get(rng)
            red_ok = True if want_red is None else all((col[k] == RED) == w for k, w in zip(col, want_red))
            item["answers"].append({"range": rng, "fmt_typed": f, "fmt_local": ws.Range(rng).Cells(1, 1).NumberFormatLocal,
                                    "fmt_en": ws.Range(rng).Cells(1, 1).NumberFormat,
                                    "values": got, "expect": expect, "value2": v2, "colors": col,
                                    "value_unchanged": all(v2[k] == raw[k] for k in v2),
                                    "ok": list(got.values()) == expect and red_ok, "red_ok": red_ok})
        R["problems"].append(item)
    # 그림 — 정답 결과 · 값 보기(수식 보기 = 형식 없이 칸에 든 값)
    for t in TABLES:
        R["shots"][f"{t['key']}_result"] = export(ws, t["area"], f"{t['key']}_result", t["boxes"])
        a = t["area"].replace(f"A{t['r']}:", f"A{t['r'] + 1}:")
        R["shots"][f"{t['key']}_values"] = export(ws, a, f"{t['key']}_values", t["boxes"], show_formulas=True)
    # 한 칸 사실 — Z1 을 지우고 형식(NumberFormatLocal)을 먼저 넣은 뒤 값 · 식
    R["facts"] = []
    ws.Columns("Z").ColumnWidth = 24
    for fa in FACTS:
        item = dict(fa)
        try:
            ws.Range("Z1").Clear()
            ws.Range("Z1").NumberFormatLocal = fa["fmt"]
            if "formula" in fa:
                ws.Range("Z1").Formula = fa["formula"]
            else:
                ws.Range("Z1").Value2 = fa["value"]
            c = ws.Range("Z1")
            item.update({"text": c.Text, "value2": c.Value2, "fmt_local": c.NumberFormatLocal, "fmt_en": c.NumberFormat,
                         "color": int(c.DisplayFormat.Font.Color)})
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
        R["facts"].append(item)
    ws.Range("Z1").Clear()
    # 함정
    for tr in TRAPS:
        t = tr["table"]
        item = {"key": tr["key"], "why": tr["why"], "range": tr["range"], "fmt": tr.get("fmt"), "set": tr.get("set")}
        try:
            for c, v in tr.get("set", {}).items():
                ws.Range(c).Value = v
            if tr.get("fmt"):
                ws.Range(tr["range"]).NumberFormatLocal = tr["fmt"]
                item["fmt_local"] = ws.Range(tr["range"]).Cells(1, 1).NumberFormatLocal
                item["fmt_en"] = ws.Range(tr["range"]).Cells(1, 1).NumberFormat
            item["values"] = texts(ws, tr["range"])
            item["value2"] = values(ws, tr["range"])
            item["colors"] = colors(ws, tr["range"])
            item["text_cells"] = sum(1 for c in ws.Range(tr["range"]) if isinstance(c.Value2, str))
            exp = t["answers"][tr["exp_idx"]][2]
            item["wrong_vs_answer"] = sum(1 for a_, b_ in zip(item["values"].values(), exp) if a_ != b_)
            item["watch"] = {c: {"formula": ws.Range(c).Formula, "text": ws.Range(c).Text} for c in tr.get("watch", [])}
            item["shot"] = export(ws, t["area"], f"trap_{tr['key']}", [tr["range"]] + tr.get("watch", []))
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
        put_rows(ws, t)  # 값 되돌림
        put_formats(ws, t, True)  # 정답 형식으로 되돌림
        R["traps"].append(item)
    # 되돌린 뒤 정답이 그대로인지 한 번 더
    R["after_traps_ok"] = all(list(texts(ws, rng).values()) == exp for t in TABLES for rng, _, exp in t["answers"])
    wb.Close(False)
    wb = xl.Workbooks.Open(tmp_pr)
    ws = wb.Worksheets("기본작업")
    for t in TABLES:
        R["shots"][f"{t['key']}_start"] = export(ws, t["area"], f"{t['key']}_start", t["boxes"])
    wb.Close(False)
    R["file_props"] = {k: core_people(p) for k, p in paths.items()}
finally:
    xl.Quit()
    shutil.rmtree(TMP, ignore_errors=True)

with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)
print(json.dumps({"ok": [[a["ok"] for a in p["answers"]] for p in R["problems"]],
                  "unchanged": [[a["value_unchanged"] for a in p["answers"]] for p in R["problems"]],
                  "fmt": [[(a["fmt_typed"], a["fmt_local"], a["fmt_en"]) for a in p["answers"]] for p in R["problems"]],
                  "values": [[list(a["values"].values()) for a in p["answers"]] for p in R["problems"]],
                  "colors": [[sorted(set(a["colors"].values()), key=str) for a in p["answers"]] for p in R["problems"]],
                  "side": [p["side"] for p in R["problems"]], "after_traps_ok": R.get("after_traps_ok"),
                  "props": R.get("file_props"), "input": R.get("input_check"),
                  "traps": [(t["key"], t.get("rejected") or list(t.get("values", {}).values()), t.get("wrong_vs_answer"), t.get("fmt_local"),
                             sorted(set(t.get("colors", {}).values()), key=str), t.get("watch")) for t in R["traps"]],
                  "fg": [(k, v.get("fg_class")) for k, v in list(R["shots"].items()) + [(t["key"], t.get("shot", {})) for t in R["traps"]] if v.get("foreground_changed")],
                  "facts": [(x["fmt"], x.get("text", x.get("rejected")), x.get("fmt_local"), x.get("color")) for x in R.get("facts", [])]},
                 ensure_ascii=False, indent=1, default=str))
