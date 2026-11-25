"""엑셀 중급 1강 — IF · AND · OR · 중첩 IF (컴활 2급 실기 계산작업 유형 연습).
문제 · 데이터 · 문장은 모두 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제다.

하는 일
1. 실습.xlsx(문제 시트 + 계산작업 시트, 답 칸 비어 있음) · 정답.xlsx(답 칸에 수식) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 수식을 넣고(첫 칸에 넣은 뒤 아래로 채우기 = FillDown) 칸에 보이는 값(.Text)을 셀 주소별로 읽기
3. 자주 틀리는 식(함정)도 같은 칸에 넣어 결과 읽기 — 저장하지 않는 사본에서
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
TMP = tempfile.mkdtemp(prefix="xl_mid01_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 컴활 2급 실기 출제기준에 맞춰 새로 만든 실전 연습 문제입니다. 이름 · 점수는 모두 가상입니다."

# ── 문제 정의 ─────────────────────────────────────────────
# r = 표 제목 행. 머리글은 r+1, 데이터는 r+2 부터 8줄.
T1 = dict(
    key="q1", level="기본", r=1, title="[표1] 사내 엑셀 교육 수료 현황",
    head=["사번", "이름", "출석(회)", "수료 여부"],
    rows=[["E-101", "탁하늬", 10], ["E-102", "편예솔", 7], ["E-103", "봉은결", 8], ["E-104", "국다솜", 9],
          ["E-105", "피시온", 6], ["E-106", "어도담", 8], ["E-107", "감해람", 10], ["E-108", "견윤슬", 5]],
    ans="D", formula='=IF(C{r}>=8,"수료","미수료")',
    text="[표1]에서 출석(회)[C3:C10]이 8 이상이면 '수료', 그 외에는 '미수료'를 수료 여부[D3:D10]에 표시하시오.",
    use="IF 함수 사용",
    expect=["수료", "미수료", "수료", "수료", "미수료", "수료", "수료", "미수료"],
)
T2 = dict(
    key="q2", level="응용", r=12, title="[표2] 사내 자격 시험 결과",
    head=["이름", "필기", "실기", "출석(회)", "보충 수강", "결과"],
    rows=[["금가람", 72, 65, 9, None], ["기누리", 58, 80, 10, None], ["모라온", 60, 60, 7, "O"], ["범미르", 85, 59, 8, None],
          ["빈별하", 90, 88, 6, None], ["승새벽", 66, 71, 5, "O"], ["시아라", 61, 92, 8, None], ["옥이슬", 45, 40, 10, "O"]],
    ans="F", formula='=IF(AND(B{r}>=60,C{r}>=60,OR(D{r}>=8,E{r}="O")),"합격","불합격")',
    text=("[표2]에서 필기[B14:B21]와 실기[C14:C21]가 모두 60 이상이고, 출석(회)[D14:D21]이 8 이상이거나 "
          "보충 수강[E14:E21]이 'O'이면 '합격', 그 외에는 '불합격'을 결과[F14:F21]에 표시하시오."),
    use="IF, AND, OR 함수 사용",
    expect=["합격", "불합격", "합격", "불합격", "불합격", "합격", "합격", "불합격"],
)
T3 = dict(
    key="q3", level="함정", r=23, title="[표3] 사내 엑셀 평가 등급",
    head=["이름", "점수", "등급"],
    rows=[["온초롱", 95], ["용한빛", 90], ["운해솔", 88], ["위나린", 80], ["인다온", 76], ["제로아", 70], ["좌마루", 69], ["추보람", 52]],
    ans="C", formula='=IF(B{r}>=90,"A",IF(B{r}>=80,"B",IF(B{r}>=70,"C","D")))',
    text=("[표3]에서 점수[B25:B32]가 90 이상이면 'A', 80 이상이면 'B', 70 이상이면 'C', "
          "그 외에는 'D'를 등급[C25:C32]에 표시하시오."),
    use="IF 함수 사용",
    expect=["A", "A", "B", "B", "C", "C", "D", "D"],
)
TABLES = [T1, T2, T3]
# 함정 — 같은 칸에 넣어 결과를 읽는 '틀린 식'
TRAPS = [
    dict(key="q2_no_quote", table=T2, why="'O' 를 따옴표 없이 씀 → 엑셀이 O 를 이름으로 읽음",
         formula='=IF(AND(B{r}>=60,C{r}>=60,OR(D{r}>=8,E{r}=O)),"합격","불합격")'),
    dict(key="q2_and_only", table=T2, why="OR 부분을 빼먹음(출석 · 보충 수강 조건 없음)",
         formula='=IF(AND(B{r}>=60,C{r}>=60),"합격","불합격")'),
    dict(key="q3_small_first", table=T3, why="작은 기준(70)부터 물음 → 70 이상이면 모두 첫 IF 에서 'C'",
         formula='=IF(B{r}>=70,"C",IF(B{r}>=80,"B",IF(B{r}>=90,"A","D")))'),
    dict(key="q3_greater", table=T3, why="'이상'을 > 로 씀 → 90 · 80 · 70 딱 맞는 점수가 한 등급 내려감",
         formula='=IF(B{r}>90,"A",IF(B{r}>80,"B",IF(B{r}>70,"C","D")))'),
    dict(key="q1_no_quote_text", table=T1, why="'수료' 를 따옴표 없이 씀",
         formula="=IF(C{r}>=8,수료,미수료)"),
]


def cols(t):
    return [chr(ord("A") + i) for i in range(len(t["head"]))]


def first(t):
    return t["r"] + 2


def last(t):
    return t["r"] + 1 + len(t["rows"])


def area(t, title=True):
    c = cols(t)
    return f"A{t['r'] if title else t['r'] + 1}:{c[-1]}{last(t)}"


def ans_range(t):
    return f"{t['ans']}{first(t)}:{t['ans']}{last(t)}"


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


def build_sheet(ws, with_answer):
    for t in TABLES:
        r = t["r"]
        ws.Range(f"A{r}").Value = t["title"]
        ws.Range(f"A{r}").Font.Bold = True
        c = cols(t)
        ws.Range(f"A{r + 1}:{c[-1]}{r + 1}").Value = (tuple(t["head"]),)
        hdr = ws.Range(f"A{r + 1}:{c[-1]}{r + 1}")
        hdr.Font.Bold = True
        hdr.Interior.Color = 0xEADBC8  # 연한 파랑(BGR)
        hdr.HorizontalAlignment = -4108
        for i, row in enumerate(t["rows"]):
            for j, v in enumerate(row):
                if v is not None:
                    ws.Range(f"{c[j]}{first(t) + i}").Value = v
        ws.Range(area(t, title=False)).Borders.LineStyle = 1
        ws.Range(f"B{first(t)}:{c[-1]}{last(t)}").HorizontalAlignment = -4108
        if with_answer:
            top = f"{t['ans']}{first(t)}"
            ws.Range(top).Formula = t["formula"].format(r=first(t))
            ws.Range(ans_range(t)).FillDown()
    for col, w in zip("ABCDEF", (10, 10, 9, 10, 10, 9)):
        ws.Columns(col).ColumnWidth = w


def build_problem_sheet(ws):
    ws.Range("A1").Value = "엑셀 중급 1강 — IF · AND · OR · 중첩 IF (컴활 2급 실기 계산작업 유형 연습)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = NOTICE
    ws.Range("A2").Font.Color = 0x3030C0
    line = 4
    for n, t in enumerate(TABLES, start=1):
        ws.Range(f"A{line}").Value = f"{n}. ({t['level']}) {t['text']}"
        ws.Range(f"A{line + 1}").Value = f"   ▶ {t['use']}"
        line += 3
    ws.Range(f"A{line}").Value = "풀이 · 실측 값 · 자주 틀리는 곳: https://www.coding-now.com/blog/excel (엑셀 중급 1강)"
    ws.Columns("A").ColumnWidth = 120


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

    # 정답 파일을 임시 사본으로 열어 실측 · 그림(원본은 건드리지 않음)
    tmp_ans = os.path.join(TMP, "ans.xlsx")
    tmp_pr = os.path.join(TMP, "pr.xlsx")
    shutil.copy(paths["정답.xlsx"], tmp_ans)
    shutil.copy(paths["실습.xlsx"], tmp_pr)
    wb = xl.Workbooks.Open(tmp_ans)
    ws = wb.Worksheets("계산작업")
    for n, t in enumerate(TABLES, start=1):
        got = texts(ws, ans_range(t))
        R["problems"].append({
            "n": n, "key": t["key"], "level": t["level"], "text": t["text"], "use": t["use"],
            "first_cell_formula": ws.Range(f"{t['ans']}{first(t)}").Formula,
            "formulas": formulas(ws, ans_range(t)), "values": got,
            "ok": list(got.values()) == t["expect"],
            "count": {v: list(got.values()).count(v) for v in sorted(set(got.values()))},
        })
    # 그림 — 정답 결과 · 수식 보기
    for n, t in enumerate(TABLES, start=1):
        R["shots"][f"{t['key']}_result"] = export(ws, area(t), f"{t['key']}_result", [ans_range(t)])
        w = {t["ans"]: 64 if n == 2 else (50 if n == 3 else 27)}
        R["shots"][f"{t['key']}_formula"] = export(ws, area(t, title=False), f"{t['key']}_formula", [ans_range(t)], show_formulas=True, widths=w)
    # 함정 — 같은 칸에 틀린 식
    for tr in TRAPS:
        t = tr["table"]
        top = f"{t['ans']}{first(t)}"
        f = tr["formula"].format(r=first(t))
        item = {"key": tr["key"], "why": tr["why"], "formula": f, "range": ans_range(t)}
        try:
            ws.Range(top).Formula = f
            ws.Range(ans_range(t)).FillDown()
            item["values"] = texts(ws, ans_range(t))
            item["wrong_vs_answer"] = sum(1 for a, b in zip(item["values"].values(), t["expect"]) if a != b)
            item["shot"] = export(ws, area(t), f"trap_{tr['key']}", [ans_range(t)])
        except pywintypes.com_error as e:
            item["rejected"] = (e.excepinfo[2] if e.excepinfo and e.excepinfo[2] else str(e))[:200]
        # 정답 식으로 되돌림
        ws.Range(top).Formula = t["formula"].format(r=first(t))
        ws.Range(ans_range(t)).FillDown()
        R["traps"].append(item)
    wb.Close(False)
    # 그림 — 실습(시작) 화면
    wb = xl.Workbooks.Open(tmp_pr)
    ws = wb.Worksheets("계산작업")
    for t in TABLES:
        R["shots"][f"{t['key']}_start"] = export(ws, area(t), f"{t['key']}_start", [ans_range(t)])
    wb.Close(False)
    # 저장한 파일 속성 확인(파일 안 core.xml 을 직접 읽음)
    R["file_props"] = {k: core_people(p) for k, p in paths.items()}
finally:
    xl.Quit()
    shutil.rmtree(TMP, ignore_errors=True)

with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1)
print(json.dumps({"ok": [p["ok"] for p in R["problems"]], "props": R.get("file_props"),
                  "traps": [(t["key"], t.get("rejected") or list(t.get("values", {}).values())) for t in R["traps"]],
                  "shots": {k: (v["size"], v["cols_found"], v["rows_found"], v["need"], v["foreground_changed"]) for k, v in R["shots"].items()}},
                 ensure_ascii=False, indent=1))
