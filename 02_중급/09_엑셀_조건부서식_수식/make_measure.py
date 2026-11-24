"""엑셀 중급 9강 — 조건부 서식, 수식으로 행 전체 칠하기(컴활 2급 실기 기본작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님). 이름 · 숫자는 지어낸 것.

하는 일
1. 실습.xlsx(문제 시트 + 기본작업 시트, 조건부 서식 없음) · 정답.xlsx(답 범위에 '수식을 사용하여 서식을 지정할 셀 결정' 규칙) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 규칙을 넣고(범위를 고르고 첫 칸을 활성 셀로 둔 채 수식 입력 — 대화 상자에서 하는 것과 같음)
   칸마다 화면에 보이는 서식(DisplayFormat — 채우기 색 · 글꼴 색 · 굵게)을 읽어 칠해진 행을 센다
3. 자주 틀리는 규칙(함정)도 같은 범위에 넣어 결과 읽기 — 저장하지 않는 사본에서
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 '새 서식 규칙' 대화 상자 화면은 이 스크립트로 만들지 않는다.
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
TMP = tempfile.mkdtemp(prefix="xl_mid09_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 숫자는 모두 가상입니다."
X_EXPR, X_CELL, X_GE = 2, 1, 7  # xlExpression · xlCellValue · xlGreaterEqual
RED, BLUE, YELLOW = 0x0000FF, 0xFF0000, 0x00FFFF  # 엑셀 Color 는 BGR 정수 — RGB(255,0,0) · RGB(0,0,255) · RGB(255,255,0)
WHITE, BLACK = 0xFFFFFF, 0

# ── 문제 정의 ─────────────────────────────────────────────
CLUB = [["가온", 1, 11, 1], ["누리", 2, 9, 3], ["다온", 1, 12, 0], ["라희", 3, 8, 4],
        ["미르", 2, 10, 2], ["비주", 4, 7, 5], ["세찬", 3, 9, 3], ["온유", 4, 11, 1]]
T25 = dict(
    key="q25", level="기본", r=1, title="[표25] 동아리 출석", head=["이름", "학년", "출석", "결석"], rows=CLUB,
    range="A3:D10", first="A3", formula="=$D3>=3", style=dict(font=RED, bold=True),
    expect_rows=[i for i, r in enumerate(CLUB) if r[3] >= 3],
    area="A1:D10", boxes=["A3:D10"],
    text=("[표25]의 [A3:D10] 영역에서 결석이 3 이상인 행 전체에 글꼴 스타일 '굵게', 글꼴 색 '표준 색-빨강'을 적용하시오. "
          "(규칙 유형은 '수식을 사용하여 서식을 지정할 셀 결정', 한 개의 규칙으로)"),
    use="조건부 서식 → 새 규칙 → 수식 =$D3>=3",
)
STOCK = [["생수", "음료", 5], ["콜라", "음료", 24], ["라면", "식품", 3], ["주스", "음료", 8],
         ["과자", "식품", 12], ["커피", "음료", 10], ["우유", "음료", 2], ["휴지", "생활", 6]]
T26 = dict(
    key="q26", level="응용", r=12, title="[표26] 매장 재고", head=["상품", "분류", "재고"], rows=STOCK,
    range="A14:C21", first="A14", formula='=AND($B14="음료",$C14<10)', style=dict(fill=YELLOW),
    expect_rows=[i for i, r in enumerate(STOCK) if r[1] == "음료" and r[2] < 10],
    area="A12:C21", boxes=["A14:C21"],
    text=("[표26]의 [A14:C21] 영역에서 분류가 '음료'이면서 재고가 10 미만인 행 전체에 채우기 색 '표준 색-노랑'을 적용하시오. "
          "(규칙 유형은 '수식을 사용하여 서식을 지정할 셀 결정', AND 함수 사용)"),
    use="수식 =AND($B14=\"음료\",$C14<10)",
)
SALES = [["강민", "사원", 820], ["김솔", "대리", 1240], ["나윤", "과장", 960], ["도현", "사원", 1510],
         ["류진", "대리", 730], ["문결", "과장", 1105], ["박하", "사원", 990], ["서율", "대리", 1380]]
AVG27 = sum(r[2] for r in SALES) / len(SALES)
T27 = dict(
    key="q27", level="함정", r=23, title="[표27] 영업 실적", head=["이름", "직급", "실적"], rows=SALES,
    range="A25:C32", first="A25", formula="=$C25>=AVERAGE($C$25:$C$32)", style=dict(font=BLUE, bold=True),
    expect_rows=[i for i, r in enumerate(SALES) if r[2] >= AVG27],
    area="A23:E32", boxes=["A25:C32"], side={"E24": "실적 평균(확인)"}, side_formula={"E25": ("=AVERAGE(C25:C32)", "#,##0.0")},
    text=("[표27]의 [A25:C32] 영역에서 실적이 실적 전체의 평균 이상인 행 전체에 글꼴 스타일 '굵게', 글꼴 색 '표준 색-파랑'을 적용하시오. "
          "(규칙 유형은 '수식을 사용하여 서식을 지정할 셀 결정', AVERAGE 함수 사용)"),
    use="수식 =$C25>=AVERAGE($C$25:$C$32)",
)
TABLES = [T25, T26, T27]
# 함정 — 같은 범위(또는 range)에 규칙 하나만 넣고 칸마다 읽는다. first = 범위를 고를 때 활성 셀(규칙의 상대 주소 기준)
TRAPS = [
    dict(key="q25_no_dollar", table=T25, formula="=D3>=3", why="$ 없이 =D3>=3 — 열도 따라 움직여 A열만 칠해짐(B는 E, C는 F 를 봄)"),
    dict(key="q25_both_dollar", table=T25, formula="=$D$3>=3", why="$D$3 — 모든 칸이 3행만 봄(D3=1 이라 아무 행도 안 칠해짐)"),
    dict(key="q25_only_d", table=T25, range="D3:D10", first="D3", formula="=$D3>=3", why="범위를 D3:D10 만 고름 — 결석 칸만 칠해짐"),
    dict(key="q25_cell_value", table=T25, kind="cell", formula="=3", why="'셀 값 >= 3' 규칙(수식 아님) — 칸마다 자기 값으로 판단, 학년 · 출석 칸도 칠해짐"),
    dict(key="q25_header_start", table=T25, range="A2:D10", first="A2", formula="=$D3>=3",
         why="머리글(2행)부터 골랐는데 수식은 3행 기준 — 한 행씩 밀림"),
    dict(key="q26_or", table=T26, formula='=OR($B14="음료",$C14<10)', why="AND 대신 OR — 둘 중 하나만 맞아도 칠해짐"),
    dict(key="q26_no_quote", table=T26, formula="=AND($B14=음료,$C14<10)", why="음료에 큰따옴표 없음 — 이름으로 읽혀 오류"),
    dict(key="q26_relative", table=T26, formula='=AND(B14="음료",C14<10)', why="$ 없음 — 열이 밀려 A열만 맞음"),
    dict(key="q26_le", table=T26, formula='=AND($B14="음료",$C14<=10)', why="<= 10 — 재고가 딱 10 인 커피도 칠해짐"),
    dict(key="q27_rel_avg", table=T27, formula="=$C25>=AVERAGE(C25:C32)", why="평균 범위에 $ 없음 — 행마다 평균 범위가 아래로 밀림"),
    dict(key="q27_header_start", table=T27, range="A24:C32", first="A24", formula="=$C25>=AVERAGE($C$25:$C$32)",
         why="머리글(24행)부터 골랐는데 수식은 25행 기준 — 한 행씩 밀림"),
    dict(key="q27_gt", table=T27, formula="=$C25>AVERAGE($C$25:$C$32)", why="> 평균 — '이상'이 아니라 '초과'(이 데이터에서는 같은 줄)"),
]
WIDTHS = dict(zip("ABCDE", (11, 9, 9, 9, 16)))
FIRST_N = 25
LESSON = "엑셀 중급 9강 — 조건부 서식, 수식으로 행 전체 칠하기 (컴활 2급 실기 기본작업 유형 연습)"
LESSON_NO = "엑셀 중급 9강"


def first_row(t):
    return t["r"] + 2


def last_row(t):
    return t["r"] + 1 + len(t["rows"])


def author(wb):
    for k in ("Author", "Last Author"):
        wb.BuiltinDocumentProperties(k).Value = "CODINGNOW"


def fix_core(path):
    """저장하면 엑셀이 '마지막 수정한 사람'을 이 PC 사용자 이름으로 다시 쓴다 → docProps/core.xml 을 직접 CODINGNOW 로(기초 7 · 8강과 같음)."""
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


def addr(c):
    return c.Address.replace("$", "")


def look(ws, rng):
    """칸마다 화면에 보이는 서식 — 채우기 색 · 글꼴 색 · 굵게(조건부 서식 결과 포함)."""
    out = {}
    for c in ws.Range(rng):
        d = c.DisplayFormat
        out[addr(c)] = {"fill": int(d.Interior.Color), "font": int(d.Font.Color), "bold": bool(d.Font.Bold)}
    return out


def styled(cell, style):
    ok = True
    if "fill" in style:
        ok &= cell["fill"] == style["fill"]
    if "font" in style:
        ok &= cell["font"] == style["font"]
    if style.get("bold"):
        ok &= cell["bold"]
    return ok


def rows_hit(t, cells, style):
    """표 데이터 행(0부터)마다 — 칸이 모두 칠해졌나 · 일부만 칠해졌나."""
    ncol = len(t["head"])
    full, part = [], []
    for i in range(len(t["rows"])):
        r = first_row(t) + i
        hits = [styled(cells[f"{chr(65 + j)}{r}"], style) for j in range(ncol) if f"{chr(65 + j)}{r}" in cells]
        if hits and all(hits):
            full.append(i)
        elif any(hits):
            part.append(i)
    return full, part


def add_rule(ws, rng, first, formula, style, kind="expr"):
    """범위를 고르고 첫 칸을 활성 셀로 둔 채 규칙 추가(대화 상자와 같은 기준). 늦은 바인딩이라 인수는 자리 순서로."""
    ws.Activate()
    ws.Range(rng).Select()
    ws.Range(first).Activate()
    if kind == "cell":
        fc = ws.Range(rng).FormatConditions.Add(X_CELL, X_GE, formula)
    else:
        fc = ws.Range(rng).FormatConditions.Add(X_EXPR, None, formula)
    if "fill" in style:
        fc.Interior.Color = style["fill"]
    if "font" in style:
        fc.Font.Color = style["font"]
    if style.get("bold"):
        fc.Font.Bold = True
    return {"formula1_read": fc.Formula1, "applies_to": fc.AppliesTo.Address.replace("$", ""),
            "active": addr(ws.Application.ActiveCell)}


# ── 그림: 범위 → PDF → PNG (8강과 같은 방식) ───────────────
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
        raise RuntimeError(f"{name}: 선을 덜 찾음 xs={xs} ys={ys} need={ncol + 2},{nrow + 2}")
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
            # 바뀌었으면 어떤 종류의 창인지(제목은 사용자 문서 이름일 수 있어 클래스 이름만 — 엑셀이면 XLMAIN)
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


def build_sheet(ws, with_answer):
    for t in TABLES:
        r = t["r"]
        ws.Range(f"A{r}").Value = t["title"]
        ws.Range(f"A{r}").Font.Bold = True
        c = [chr(65 + i) for i in range(len(t["head"]))]
        hdr = ws.Range(f"A{r + 1}:{c[-1]}{r + 1}")
        hdr.Value = (tuple(t["head"]),)
        for i, row in enumerate(t["rows"]):
            for j, v in enumerate(row):
                ws.Range(f"{c[j]}{first_row(t) + i}").Value2 = v
        ws.Range(f"A{r + 1}:{c[-1]}{last_row(t)}").Borders.LineStyle = 1
        ws.Range(f"A{first_row(t)}:{c[-1]}{last_row(t)}").HorizontalAlignment = -4108
        heads = [hdr]
        for cell, v in t.get("side", {}).items():
            ws.Range(cell).Value = v
            heads.append(ws.Range(cell))
        for cell, (f, nf) in t.get("side_formula", {}).items():
            ws.Range(cell).NumberFormatLocal = nf
            ws.Range(cell).Formula = f
            ws.Range(f"{cell[0]}{int(cell[1:]) - 1}:{cell}").Borders.LineStyle = 1
            ws.Range(cell).HorizontalAlignment = -4108
        for h in heads:
            h.Font.Bold = True
            h.Interior.Color = 0xEADBC8
            h.HorizontalAlignment = -4108
    for col, w in WIDTHS.items():
        ws.Columns(col).ColumnWidth = w
    if with_answer:
        for t in TABLES:
            add_rule(ws, t["range"], t["first"], t["formula"], t["style"])


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


def clear_rules(ws):
    ws.Cells.FormatConditions.Delete()


# ── 실행 ──────────────────────────────────────────────────
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "notice": NOTICE, "avg27": AVG27, "problems": [], "traps": [], "shots": {}}
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
            wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
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
        p = os.path.join(HERE, fname)
        if os.path.exists(p):
            os.remove(p)
        wb.SaveAs(p, 51)
        R.setdefault("saved_rules", {})[fname] = [
            {"applies_to": fc.AppliesTo.Address.replace("$", ""), "formula1": fc.Formula1, "type": fc.Type}
            for fc in wb.Worksheets("기본작업").Cells.FormatConditions]
        wb.Close(False)
        fix_core(p)
        paths[fname] = p

    tmp_ans, tmp_pr = os.path.join(TMP, "ans.xlsx"), os.path.join(TMP, "pr.xlsx")
    shutil.copy(paths["정답.xlsx"], tmp_ans)
    shutil.copy(paths["실습.xlsx"], tmp_pr)
    wb = xl.Workbooks.Open(tmp_ans)
    ws = wb.Worksheets("기본작업")
    ws.Activate()
    for n, t in enumerate(TABLES, start=FIRST_N):
        cells = look(ws, t["range"])
        full, part = rows_hit(t, cells, t["style"])
        R["problems"].append({"n": n, "key": t["key"], "text": t["text"], "formula": t["formula"],
                              "rows_full": [first_row(t) + i for i in full], "rows_part": [first_row(t) + i for i in part],
                              "expect": [first_row(t) + i for i in t["expect_rows"]],
                              "ok": full == t["expect_rows"] and not part,
                              "side": {c: ws.Range(c).Text for c in t.get("side_formula", {})}})
        R["shots"][f"{t['key']}_result"] = export(ws, t["area"], f"{t['key']}_result", t["boxes"])
    # 사실 — 값을 바꾸면 바로 따라 바뀌나(결석 D5 0 → 3)
    ws.Range("D5").Value2 = 3
    full, part = rows_hit(T25, look(ws, T25["range"]), T25["style"])
    R["fact_live"] = {"D5": 3, "rows_full": [first_row(T25) + i for i in full]}
    ws.Range("D5").Value2 = CLUB[2][3]
    # 함정 — 규칙을 모두 지우고 하나만 넣어 읽기
    for tr in TRAPS:
        t = tr["table"]
        rng, first = tr.get("range", t["range"]), tr.get("first", t["first"])
        item = {"key": tr["key"], "why": tr["why"], "range": rng, "first": first, "formula": tr["formula"], "kind": tr.get("kind", "expr")}
        clear_rules(ws)
        try:
            item["rule"] = add_rule(ws, rng, first, tr["formula"], t["style"], tr.get("kind", "expr"))
            cells = look(ws, t["range"])
            full, part = rows_hit(t, cells, t["style"])
            item["rows_full"] = [first_row(t) + i for i in full]
            item["rows_part"] = [first_row(t) + i for i in part]
            item["cells_hit"] = sorted(k for k, v in cells.items() if styled(v, t["style"]))
            item["wrong_rows"] = sorted(set(full) ^ set(t["expect_rows"]) | set(part))
            item["shot"] = export(ws, t["area"], f"trap_{tr['key']}", [rng])
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
        R["traps"].append(item)
    clear_rules(ws)
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
print(json.dumps({"excel": R.get("excel"), "avg27": AVG27,
                  "problems": [(p["n"], p["ok"], p["rows_full"], p["rows_part"], p["expect"], p["side"]) for p in R["problems"]],
                  "saved": R.get("saved_rules"), "live": R.get("fact_live"), "props": R.get("file_props"),
                  "traps": [(t["key"], t.get("rejected") or (t["rows_full"], t["rows_part"], t["cells_hit"][:12]), t.get("rule")) for t in R["traps"]],
                  "fg": [(k, v.get("fg_class")) for k, v in R["shots"].items() if v.get("foreground_changed")] + [(t["key"], t["shot"].get("fg_class")) for t in R["traps"] if t.get("shot", {}).get("foreground_changed")]},
                 ensure_ascii=False, indent=1, default=str))
