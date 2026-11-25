"""엑셀 중급 3강 — VLOOKUP · HLOOKUP (컴활 2급 실기 계산작업 유형 연습).
문제 · 데이터 · 문장은 모두 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제다. 이름은 지어낸 이름.

하는 일
1. 실습.xlsx(문제 시트 + 계산작업 시트, 답 칸 비어 있음) · 정답.xlsx(답 칸에 수식) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 수식을 넣고(첫 칸에 넣은 뒤 아래로 채우기 = FillDown) 칸에 보이는 값(.Text)을 셀 주소별로 읽기
3. 자주 틀리는 식(함정)도 같은 칸에 넣어 결과 읽기 — 저장하지 않는 사본에서
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
TMP = tempfile.mkdtemp(prefix="xl_mid03_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제입니다. 이름 · 코드 · 숫자는 모두 가상입니다."
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
T7 = dict(
    key="q7", level="기본", r=1, title="[표7] 문구 주문 내역",
    head=["제품코드", "수량", "단가"],
    rows=[["P-103", 12], ["P-101", 30], ["P-105", 4], ["P-102", 10], ["P-104", 25], ["P-101", 18], ["P-103", 6], ["P-105", 9]],
    side=grid("E", 2, [["제품코드", "제품명", "단가"], ["P-101", "볼펜", 500], ["P-102", "노트", 1200], ["P-103", "형광펜", 800],
                       ["P-104", "지우개", 300], ["P-105", "파일철", 1500]]),
    answers=[("C3", "C3:C10", "=VLOOKUP(A3,$E$3:$G$7,3,FALSE)", ["800", "500", "1,500", "1,200", "300", "500", "800", "1,500"])],
    area="A1:G10", boxes=["C3:C10"], fmt={"C3:C10": WON, "G3:G7": WON}, border_side=["E2:G7"],
    text="[표7]에서 제품코드[A3:A10]를 이용하여 제품표[E3:G7]에서 단가를 찾아 단가[C3:C10]에 표시하시오.",
    use="VLOOKUP 함수 사용",
)
T8 = dict(
    key="q8", level="응용", r=12, title="[표8] 고객 구매 금액",
    head=["고객", "구매 금액", "할인율"],
    rows=[["계바다", 85000], ["곽이음", 100000], ["구나래", 250000], ["기솔찬", 300000],
          ["남은솜", 499000], ["노해밀", 500000], ["도아름", 720000], ["라온결", 45000]],
    side=hgrid("E", 13, [["구매 금액", 0, 100000, 300000, 500000], ["할인율", 0, 0.05, 0.1, 0.15]]),
    answers=[("C14", "C14:C21", "=HLOOKUP(B14,$F$13:$I$14,2,TRUE)", ["0%", "5%", "5%", "10%", "10%", "15%", "15%", "0%"])],
    area="A12:I21", boxes=["C14:C21"], fmt={"B14:B21": WON, "F13:I13": WON, "F14:I14": PCT, "C14:C21": PCT}, border_side=["E13:I14"],
    text=("[표8]에서 구매 금액[B14:B21]에 따른 할인율을 할인율표[F13:I14]를 참조하여 할인율[C14:C21]에 표시하시오. "
          "할인율표의 구매 금액은 그 금액 '이상'부터 적용된다."),
    use="HLOOKUP 함수 사용",
)
T9 = dict(
    key="q9", level="함정", r=23, title="[표9] 사원 부서 배정",
    head=["부서코드", "이름", "부서명"],
    rows=[["D2", "갈하솔"], ["D6", "견새온"], ["D1", "낭다래"], ["D3", "단해나"], ["D5", "동라온"], ["D4", "맹솔빛"], ["D2", "묵여름"], ["D6", "사빛결"]],
    side=grid("E", 24, [["부서코드", "부서명"], ["D4", "인사팀"], ["D1", "총무팀"], ["D6", "개발팀"], ["D2", "영업팀"], ["D5", "회계팀"], ["D3", "기획팀"]]),
    answers=[("C25", "C25:C32", "=VLOOKUP(A25,$E$25:$F$30,2,FALSE)", ["영업팀", "개발팀", "총무팀", "기획팀", "회계팀", "인사팀", "영업팀", "개발팀"])],
    area="A23:F32", boxes=["C25:C32"], fmt={}, border_side=["E24:F30"],
    text="[표9]에서 부서코드[A25:A32]를 이용하여 부서표[E25:F30]에서 부서명을 찾아 부서명[C25:C32]에 표시하시오.",
    use="VLOOKUP 함수 사용 ▶ 부서표는 부서코드 순으로 정렬되어 있지 않음",
)
TABLES = [T7, T8, T9]
TRAPS = [
    dict(key="q7_col2", table=T7, top="C3", fill="C3:C10", why="열 번호를 2 로 → 제품표 둘째 열(제품명)이 나옴",
         formula="=VLOOKUP(A3,$E$3:$G$7,2,FALSE)"),
    dict(key="q7_relative", table=T7, top="C3", fill="C3:C10", why="표 범위에 $ 를 안 붙이고 채움 → 아래 칸일수록 표가 내려가 위쪽 코드를 못 찾음",
         formula="=VLOOKUP(A3,E3:G7,3,FALSE)"),
    dict(key="q7_sheetcol", table=T7, top="C3", fill="C3:C10", why="열 번호를 시트 열 번호(G = 7)로 → 표 범위가 3열뿐이라 오류",
         formula="=VLOOKUP(A3,$E$3:$G$7,7,FALSE)"),
    dict(key="q8_false", table=T8, top="C14", fill="C14:C21", why="구간 찾기에 FALSE(정확히 일치) → 표에 딱 있는 금액(100,000 · 300,000 · 500,000)만 찾음",
         formula="=HLOOKUP(B14,$F$13:$I$14,2,FALSE)"),
    dict(key="q8_omit", table=T8, top="C14", fill="C14:C21", why="넷째 인수 생략 → TRUE 와 같음(정답과 같은 결과)",
         formula="=HLOOKUP(B14,$F$13:$I$14,2)"),
    dict(key="q8_vlookup", table=T8, top="C14", fill="C14:C21", why="가로 표에 VLOOKUP → 첫 열(F13:F14)에서 세로로 찾음",
         formula="=VLOOKUP(B14,$F$13:$I$14,2,TRUE)"),
    dict(key="q9_omit", table=T9, top="C25", fill="C25:C32", why="넷째 인수(FALSE) 생략 → 근삿값 찾기라 정렬 안 된 부서표에서 엉뚱한 부서",
         formula="=VLOOKUP(A25,$E$25:$F$30,2)"),
    dict(key="q9_true", table=T9, top="C25", fill="C25:C32", why="TRUE 를 적음(생략과 같음)",
         formula="=VLOOKUP(A25,$E$25:$F$30,2,TRUE)"),
    dict(key="q9_relative", table=T9, top="C25", fill="C25:C32", why="표 범위에 $ 없이 채움",
         formula="=VLOOKUP(A25,E25:F30,2,FALSE)"),
]
WIDTHS = dict(zip("ABCDEFGHI", (11, 11, 10, 3, 11, 10, 10, 10, 10)))
FW = {"q7": {"C": 40}, "q8": {"C": 44}, "q9": {"C": 44}}
FIRST_N = 7
LESSON = "엑셀 중급 3강 — VLOOKUP · HLOOKUP (컴활 2급 실기 계산작업 유형 연습)"
LESSON_NO = "엑셀 중급 3강"


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
            exp = t["answers"][0][3]
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
                  "fg": any(v["foreground_changed"] for v in R["shots"].values())},
                 ensure_ascii=False, indent=1))
