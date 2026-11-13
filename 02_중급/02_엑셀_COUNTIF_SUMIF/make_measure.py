"""엑셀 중급 2강 — COUNTIF · SUMIF · AVERAGEIF (컴활 2급 실기 계산작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님).

하는 일
1. 실습.xlsx(문제 시트 + 계산작업 시트, 답 칸 비어 있음) · 정답.xlsx(답 칸에 수식) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 수식을 넣고(첫 칸에 넣은 뒤 아래로 채우기 = FillDown) 칸에 보이는 값(.Text)을 셀 주소별로 읽기
3. 자주 틀리는 식(함정)도 같은 칸에 넣어 결과 읽기 — 저장하지 않는 사본에서. 기준 칸(E25)을 바꿨을 때 답이 따라 바뀌는지도
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 수식 입력줄 · 함수 인수 창 화면은 이 스크립트로 만들지 않는다.
결과: 실습.xlsx · 정답.xlsx · measure.json · shots/*.png
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
TMP = tempfile.mkdtemp(prefix="xl_mid02_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 부서 · 숫자는 모두 가상입니다."
WON = "#,##0"

# ── 문제 정의 ─────────────────────────────────────────────
# 표마다: 제목 행 r, 머리글(A열부터), 데이터 8줄, 옆 칸(side: 셀 → 값, 머리글이면 굵게),
# answers: [(첫 칸, 채울 범위, 정답 식, 기대 .Text 목록)], area = 그림 범위
T4 = dict(
    key="q4", level="기본", r=1, title="[표4] 사내 엑셀 평가 점수",
    head=["이름", "부서", "점수"],
    rows=[["판솔길", "영업", 92], ["함슬기", "총무", 78], ["형여울", "개발", 80], ["호온새", "영업", 65],
          ["화우람", "개발", 88], ["경이든", "총무", 79], ["도찬솔", "개발", 100], ["예하랑", "영업", 70]],
    side={"E2": ("80점 이상 인원", True)},
    answers=[("E3", "E3", '=COUNTIF(C3:C10,">=80")&"명"', ["4명"])],
    area="A1:E10", boxes=["E3"], fmt={},
    text="[표4]에서 점수[C3:C10]가 80 이상인 인원 수를 [E3]에 표시하시오.",
    use="COUNTIF 함수와 & 연산자 사용 ▶ 표시 예: 3명",
)
T5 = dict(
    key="q5", level="응용", r=12, title="[표5] 영업팀 판매 실적",
    head=["담당", "부서", "판매액"],
    rows=[["마해든", "영업1팀", 1250000], ["반가온", "영업3팀", 980000], ["방늘봄", "영업2팀", 640000], ["변단비", "영업1팀", 1120000],
          ["석라미", "영업2팀", 760000], ["설모아", "영업3팀", 455000], ["소빛찬", "영업1팀", 1030000], ["염새롬", "영업2팀", 1315000]],
    side={"E13": ("부서", True), "F13": ("합계", True), "G13": ("평균", True),
          "E14": ("영업1팀", False), "E15": ("영업2팀", False), "E16": ("영업3팀", False)},
    answers=[("F14", "F14:F16", "=SUMIF($B$14:$B$21,E14,$C$14:$C$21)", ["3,400,000", "2,715,000", "1,435,000"]),
             ("G14", "G14:G16", "=AVERAGEIF($B$14:$B$21,E14,$C$14:$C$21)", ["1,133,333", "905,000", "717,500"])],
    area="A12:G21", boxes=["F14:G16"], fmt={"C14:C21": WON, "F14:G16": WON},
    text="[표5]에서 부서[B14:B21]별 판매액[C14:C21]의 합계와 평균을 [F14:G16]에 표시하시오.",
    use="SUMIF, AVERAGEIF 함수 사용 ▶ [F14] · [G14]에 넣은 수식을 아래로 채워 완성할 것",
)
T6 = dict(
    key="q6", level="함정", r=23, title="[표6] 문구 판매 수량",
    head=["품목", "판매 수량"],
    rows=[["젤 볼펜(검정)", 48], ["젤 볼펜(빨강)", 35], ["볼펜 리필", 60], ["형광펜", 72],
          ["연필", 50], ["수정 테이프", 18], ["노트 A5", 95], ["볼펜 세트", 22]],
    side={"E24": ("기준 수량", True), "E25": (50, False), "F24": ("기준 이상 품목 수", True), "G24": ("볼펜 들어간 품목 합계", True)},
    answers=[("F25", "F25", '=COUNTIF(B25:B32,">="&E25)', ["4"]),
             ("G25", "G25", '=SUMIF(A25:A32,"*볼펜*",B25:B32)', ["165"])],
    area="A23:G32", boxes=["F25:G25"], fmt={},
    text=("[표6]에서 ① 판매 수량[B25:B32]이 기준 수량[E25] 이상인 품목 수를 [F25]에, "
          "② 품목[A25:A32]에 '볼펜'이 들어간 품목의 판매 수량 합계를 [G25]에 표시하시오."),
    use="COUNTIF, SUMIF 함수 사용 ▶ ①의 기준은 [E25]를 참조할 것",
)
TABLES = [T4, T5, T6]
# 함정 — 같은 칸에 넣어 결과를 읽는 '틀린 식' (top, fill, formula)
TRAPS = [
    dict(key="q4_no_unit", table=T4, top="E3", fill="E3", why="& \"명\" 을 빼먹음 → 숫자 4(오른쪽 정렬), 표시 예와 다름",
         formula='=COUNTIF(C3:C10,">=80")'),
    dict(key="q4_no_quote", table=T4, top="E3", fill="E3", why="조건 >=80 을 따옴표 없이 씀",
         formula="=COUNTIF(C3:C10,>=80)"),
    dict(key="q4_reverse_op", table=T4, top="E3", fill="E3", why="'이상'을 => 로 거꾸로 씀",
         formula='=COUNTIF(C3:C10,"=>80")&"명"'),
    dict(key="q5_relative", table=T5, top="F14", fill="F14:F16", why="범위에 $ 를 안 붙이고 아래로 채움 → 아래 칸으로 갈수록 범위가 한 줄씩 내려감",
         formula="=SUMIF(B14:B21,E14,C14:C21)"),
    dict(key="q5_no_sum_range", table=T5, top="F14", fill="F14:F16", why="합계 범위(세 번째 인수)를 빼먹음 → 글자 칸(부서)을 더함",
         formula="=SUMIF($B$14:$B$21,E14)"),
    dict(key="q5_avg_relative", table=T5, top="G14", fill="G14:G16", why="AVERAGEIF 도 $ 없이 채움",
         formula="=AVERAGEIF(B14:B21,E14,C14:C21)"),
    dict(key="q6_cell_in_quote", table=T6, top="F25", fill="F25", why="셀 주소를 따옴표 안에 넣음 → 'E25' 라는 글자와 비교",
         formula='=COUNTIF(B25:B32,">=E25")'),
    dict(key="q6_no_quote", table=T6, top="F25", fill="F25", why="비교 연산자를 따옴표 없이 씀",
         formula="=COUNTIF(B25:B32,>=E25)"),
    dict(key="q6_typed_50", table=T6, top="F25", fill="F25", why="기준 50 을 식에 직접 씀(값은 맞지만 [E25] 참조 조건 위반)",
         formula='=COUNTIF(B25:B32,">=50")'),
    dict(key="q6_exact", table=T6, top="G25", fill="G25", why="만능 문자 없이 \"볼펜\" → 품목 이름이 딱 '볼펜'인 것만",
         formula='=SUMIF(A25:A32,"볼펜",B25:B32)'),
    dict(key="q6_starts", table=T6, top="G25", fill="G25", why="\"볼펜*\" → '볼펜'으로 시작하는 것만(젤 볼펜 빠짐)",
         formula='=SUMIF(A25:A32,"볼펜*",B25:B32)'),
]


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
        side_cells = []
        for cell, (v, is_head) in t["side"].items():
            ws.Range(cell).Value = v
            side_cells.append(cell)
            if is_head:
                heads.append(ws.Range(cell))
        for top, fill, _, _ in t["answers"]:
            side_cells.append(fill)
        for cell in side_cells:
            ws.Range(cell).Borders.LineStyle = 1
            ws.Range(cell).HorizontalAlignment = -4108
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
    for col, w in zip("ABCDEFG", (14, 10, 12, 3, 15, 17, 21)):
        ws.Columns(col).ColumnWidth = w


def build_problem_sheet(ws):
    ws.Range("A1").Value = "엑셀 중급 2강 — COUNTIF · SUMIF · AVERAGEIF (컴활 2급 실기 계산작업 유형 연습)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = NOTICE
    ws.Range("A2").Font.Color = 0x3030C0
    line = 4
    for n, t in enumerate(TABLES, start=4):
        ws.Range(f"A{line}").Value = f"{n}. ({t['level']}) {t['text']}"
        ws.Range(f"A{line + 1}").Value = f"   ▶ {t['use']}"
        line += 3
    ws.Range(f"A{line}").Value = "풀이 · 실측 값 · 자주 틀리는 곳: https://www.coding-now.com/blog/excel (엑셀 중급 2강)"
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
    for n, t in enumerate(TABLES, start=4):
        item = {"n": n, "key": t["key"], "level": t["level"], "text": t["text"], "use": t["use"], "answers": []}
        for top, fill, formula, expect in t["answers"]:
            got = texts(ws, fill)
            item["answers"].append({"first_cell": top, "first_cell_formula": ws.Range(top).Formula,
                                    "formulas": formulas(ws, fill), "values": got,
                                    "value2": {c.Address.replace("$", ""): c.Value2 for c in ws.Range(fill)},
                                    "align_left_or_right": {c.Address.replace("$", ""): type(c.Value2).__name__ for c in ws.Range(fill)},
                                    "ok": list(got.values()) == expect})
        R["problems"].append(item)
    # 기준 칸을 바꾸면 따라 바뀌는지(문제 6 ①)
    ws.Range("E25").Value = 60
    R["q6_criteria_change"] = {"E25": 60, "F25_correct": ws.Range("F25").Text}
    ws.Range("F25").Formula = '=COUNTIF(B25:B32,">=50")'
    R["q6_criteria_change"]["F25_typed_50"] = ws.Range("F25").Text
    ws.Range("F25").Formula = '=COUNTIF(B25:B32,">="&E25)'
    ws.Range("E25").Value = 50
    R["q6_criteria_change"]["back_to_50"] = ws.Range("F25").Text
    # 그림 — 정답 결과 · 수식 보기
    fw = {"q4": {"E": 36}, "q5": {"F": 40, "G": 42}, "q6": {"F": 34, "G": 36}}
    for t in TABLES:
        R["shots"][f"{t['key']}_result"] = export(ws, t["area"], f"{t['key']}_result", t["boxes"])
        a = t["area"].replace(f"A{t['r']}:", f"A{t['r'] + 1}:")
        R["shots"][f"{t['key']}_formula"] = export(ws, a, f"{t['key']}_formula", t["boxes"], show_formulas=True, widths=fw[t["key"]])
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
                  "q6_change": R.get("q6_criteria_change"),
                  "traps": [(t["key"], t.get("rejected") or list(t.get("values", {}).values())) for t in R["traps"]],
                  "shots": {k: (v["size"], v["foreground_changed"]) for k, v in R["shots"].items()}},
                 ensure_ascii=False, indent=1))
