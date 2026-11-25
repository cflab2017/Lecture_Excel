"""엑셀 중급 12강 — 피벗 테이블(컴활 2급 실기 분석작업 유형 연습).
문제 · 데이터 · 문장은 모두 새로 만든 연습용 자작 문제다(대한상공회의소 문제가 아님). 이름 · 숫자는 지어낸 것.

하는 일
1. 실습.xlsx(문제 시트 + 분석작업-34 · 35 · 36 시트, 표만 있음) · 정답.xlsx(같은 시트 A15 에 피벗 테이블까지) 만들기 — 작성자 · 마지막 저장자 CODINGNOW
2. 이 PC 엑셀에서 피벗 테이블을 만들어(행 · 열 · 값 · 레이아웃 · 빈 셀 · 그룹) 보고서 칸의 글자를 그대로 읽는다
3. 자주 틀리는 것(값의 기본 함수 · 새로 고침 · 줄 더하기 · 머리글 빈칸 · 압축 형식 · 빈 셀 · 그룹 안 함 · 글자 날짜)도 같은 자리에서 — 저장하지 않는 사본에서
4. 그림: 범위를 PDF 로 내보내(행 · 열 머리글 · 눈금선 포함) pymupdf 로 그림 → 볼 곳에 빨간 테두리 → shots/*.png

엑셀은 따로 띄워(DispatchEx) 숨긴 채(Visible=False) 쓴다. 창을 앞으로 가져오기 · 키 보내기 · 클립보드 · 화면 캡처를 쓰지 않는다
(content_hub 공용 규칙 '바탕화면 · 포커스', 2026-11-12). 그래서 '피벗 테이블 만들기' · 필드 목록 화면은 이 스크립트로 만들지 않는다.
피벗 테이블은 코드(PivotCaches.Create · CreatePivotTable · PivotFields.Orientation · AddDataField)로 만든다 — 화면에서 끌어 넣는 것과
다를 수 있는 곳(날짜 자동 묶기 등)은 measure.json 에 '코드로' 라고 적는다.
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
TMP = tempfile.mkdtemp(prefix="xl_mid12_")
os.makedirs(SHOTS, exist_ok=True)
NOTICE = "※ 대한상공회의소 문제가 아닌 연습용 자작 문제입니다. 이름 · 숫자는 모두 가상입니다."
FN = {"합계": -4157, "평균": -4106, "개수": -4112, "최대값": -4136}
FN_NAME = {v: k for k, v in FN.items()}
ROW, COL, DATA = 1, 2, 4  # xlRowField · xlColumnField · xlDataField
LAYOUT = {"압축": 0, "테이블": 1, "개요": 2}  # xlCompactRow · xlTabularRow · xlOutlineRow
PT_VER = 5  # xlPivotTableVersion15 — 엑셀 2013 이후 형식(압축 형식이 기본)
DEST = "A15"

D = datetime.date
T34 = dict(
    key="q34", sheet="분석작업-34", level="기본", title="[표34] 지점별 제품 판매", head=["지점", "담당자", "제품", "판매액"], money=[3], dates=[],
    rows=[["인천", "김바람", "모니터", 320000], ["광주", "이슬기", "키보드", 85000], ["대전", "박나래", "모니터", 290000], ["인천", "최우주", "키보드", 64000],
          ["광주", "정다운", "마우스", 42000], ["대전", "한가람", "키보드", 73000], ["인천", "윤소리", "마우스", 38000], ["광주", "강하늘", "모니터", 310000],
          ["대전", "오세움", "마우스", 51000], ["인천", "임보람", "모니터", 275000]],
    rowf=["지점"], colf=["제품"], data=[("판매액", "합계", "#,##0")], layout=None, null=None, group=None,
    text="[표34]를 이용하여 '지점'은 행, '제품'은 열로 처리하고, 값에 '판매액'의 합계를 계산하는 피벗 테이블을 작성하시오. "
         "▶ 피벗 테이블 보고서는 동일 시트의 A15 셀에서 시작하시오. ▶ 값 영역의 숫자는 1000 단위 구분 기호(,)를 표시하시오.",
    use="표 안 칸 하나 → 삽입 → 피벗 테이블 → 기존 워크시트 A15 → 지점은 행 · 제품은 열 · 판매액은 값(합계) → 값 필드 설정 → 표시 형식 → 숫자 · 1000 단위 구분 기호",
)
T35 = dict(
    key="q35", sheet="분석작업-35", level="응용", title="[표35] 직위별 기본급", head=["이름", "부서", "직위", "기본급"], money=[3], dates=[],
    rows=[["가람", "영업부", "과장", 3800000], ["나래", "기획부", "대리", 3200000], ["다솜", "영업부", "사원", 2700000], ["라희", "총무부", "과장", 3600000],
          ["마음", "기획부", "사원", 2600000], ["보라", "영업부", "대리", 3100000], ["새롬", "총무부", "사원", 2800000], ["아라", "기획부", "과장", 4000000],
          ["자람", "영업부", "사원", 2900000], ["하늘", "총무부", "사원", 2950000]],
    rowf=["직위"], colf=["부서"], data=[("기본급", "평균", "#,##0")], layout="개요", null="*", group=None,
    text="[표35]를 이용하여 '직위'는 행, '부서'는 열로 처리하고, 값에 '기본급'의 평균을 계산하는 피벗 테이블을 작성하시오. "
         "▶ 피벗 테이블 보고서는 동일 시트의 A15 셀에서 시작하시오. ▶ 보고서 레이아웃은 개요 형식으로 표시하시오. "
         "▶ 빈 셀은 '*' 기호로 표시하시오. ▶ 값 영역의 숫자는 1000 단위 구분 기호(,)를 표시하시오.",
    use="피벗 테이블(A15) → 직위 행 · 부서 열 · 기본급 값 → 값 필드 설정 → 평균 · 표시 형식 → 디자인 → 보고서 레이아웃 → 개요 형식으로 표시 → 피벗 테이블 옵션 → 빈 셀 표시 *",
)
T36 = dict(
    key="q36", sheet="분석작업-36", level="함정", title="[표36] 날짜별 판매 수량", head=["판매일", "제품", "담당", "수량"], money=[], dates=[0],
    rows=[[D(2026, 9, 3), "볼펜", "김솔", 30], [D(2026, 9, 17), "노트", "이담", 12], [D(2026, 9, 24), "파일", "박온", 8], [D(2026, 10, 2), "볼펜", "김솔", 25],
          [D(2026, 10, 15), "노트", "이담", "미정"], [D(2026, 10, 29), "파일", "박온", 14], [D(2026, 11, 5), "볼펜", "김솔", 40], [D(2026, 11, 12), "노트", "이담", 18],
          [D(2026, 11, 19), "파일", "박온", 6], [D(2026, 11, 26), "볼펜", "김솔", 22]],
    rowf=["판매일"], colf=[], data=[("수량", "합계", None)], layout=None, null=None, group="월",
    text="[표36]을 이용하여 '판매일'은 행으로 처리하고 '월' 단위로 그룹을 지정하시오. 값에는 '수량'의 합계를 계산하는 피벗 테이블을 작성하시오. "
         "▶ 피벗 테이블 보고서는 동일 시트의 A15 셀에서 시작하시오. ▶ 수량을 아직 못 적은 줄(D7 '미정')은 고치지 말고 그대로 두시오.",
    use="피벗 테이블(A15) → 판매일 행 · 수량 값 → 값 이름이 '개수 : 수량'이면 값 필드 설정에서 '합계'로 → 판매일 칸 하나 오른쪽 클릭 → 그룹 → 월",
)
TABLES = [T34, T35, T36]
WIDTHS = dict(zip("ABCDEF", (16, 13, 13, 13, 13, 13)))
FIRST_N = 34
LESSON = "엑셀 중급 12강 — 피벗 테이블 (컴활 2급 실기 분석작업 유형 연습)"
LESSON_NO = "엑셀 중급 12강"


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


def put_value(cell, v, is_date):
    if v is None:
        return
    if is_date:
        cell.NumberFormatLocal = "yyyy-mm-dd"
        cell.Formula = v.isoformat()  # 엑셀이 날짜로 받아들이는 꼴 — 아래에서 값 종류를 확인
    else:
        cell.Value2 = v


def build_sheet(ws, t, with_answer):
    ws.Range("A1").Value = t["title"]
    ws.Range("A1").Font.Bold = True
    hdr = ws.Range("A2:D2")
    hdr.Value = (tuple(t["head"]),)
    for i, row in enumerate(t["rows"]):
        for j, v in enumerate(row):
            put_value(ws.Range(f"{chr(65 + j)}{3 + i}"), v, j in t["dates"])
    last = 2 + len(t["rows"])
    ws.Range(f"A2:D{last}").Borders.LineStyle = 1
    ws.Range(f"A3:C{last}").HorizontalAlignment = -4108
    for j in t["money"]:
        ws.Range(f"{chr(65 + j)}3:{chr(65 + j)}{last}").NumberFormat = "#,##0"
    hdr.Font.Bold = True
    hdr.Interior.Color = 0xEADBC8
    hdr.HorizontalAlignment = -4108
    for col, w in WIDTHS.items():
        ws.Columns(col).ColumnWidth = w
    if with_answer:
        return make_pivot(ws, t, t)
    return None


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


def make_pivot(ws, t, spec, src=None, log=None):
    """spec: rowf · colf · data[(필드, 함수 이름 | None=엑셀이 고른 그대로, 표시 형식)] · layout · null · group"""
    log = log if log is not None else {}
    n = 2 + len(t["rows"])
    wb = ws.Parent
    pc = wb.PivotCaches().Create(1, ws.Range(src or f"A2:D{n}"), PT_VER)  # xlDatabase
    pt = pc.CreatePivotTable(ws.Range(DEST), f"피벗{t['key'][1:]}", True, PT_VER)
    ts = pt.TableStyle2
    log["table_style"] = ts if isinstance(ts, str) else getattr(ts, "Name", None)
    log["compact_layout_default"] = bool(pt.InGridDropZones) is False and pt.CompactLayoutRowHeader
    for k, f in enumerate(spec.get("rowf", [])):
        pf = pt.PivotFields(f)
        pf.Orientation = ROW
        pf.Position = k + 1
    for k, f in enumerate(spec.get("colf", [])):
        pf = pt.PivotFields(f)
        pf.Orientation = COL
        pf.Position = k + 1
    log["data"] = []
    for field, fn, fmt in spec.get("data", []):
        df = pt.AddDataField(pt.PivotFields(field))  # 함수 · 이름은 엑셀이 고르게 둔다 — 기본값을 잰다
        rec = {"field": field, "default_function": FN_NAME.get(df.Function, df.Function), "default_caption": df.Caption}
        if fn:
            df.Function = FN[fn]
        if fmt:
            df.NumberFormat = fmt
        rec.update({"function": FN_NAME.get(df.Function, df.Function), "caption": df.Caption, "number_format": df.NumberFormat})
        log["data"].append(rec)
    if spec.get("layout"):
        pt.RowAxisLayout(LAYOUT[spec["layout"]])
    if spec.get("null") is not None:
        pt.NullString = spec["null"]
        pt.DisplayNullString = True
    if spec.get("group") == "월":
        cell = pt.PivotFields(spec["rowf"][0]).DataRange.Cells(1)
        log["group_cell"] = cell.Address.replace("$", "")
        cell.Group(True, True, None, (False, False, False, False, True, False, False))  # 초 · 분 · 시 · 일 · 월 · 분기 · 연
    log["row_grand"], log["column_grand"] = pt.RowGrand, pt.ColumnGrand
    return pt


def read_pivot(ws, pt):
    tr = pt.TableRange2
    r0, c0 = tr.Row, tr.Column
    rows = []
    for i in range(tr.Rows.Count):
        rows.append({"r": r0 + i, "text": [ws.Cells(r0 + i, c0 + j).Text for j in range(tr.Columns.Count)]})
    return {"range": tr.Address.replace("$", ""), "body": pt.TableRange1.Address.replace("$", ""), "rows": rows,
            "last_row": r0 + tr.Rows.Count - 1, "last_col": chr(64 + c0 + tr.Columns.Count - 1)}


def find_cells(ws, rng, pred):
    out = []
    for c in ws.Range(rng).Cells:
        if pred(c.Text):
            out.append(c.Address.replace("$", ""))
    return out


# ── 그림: 범위 → PDF → PNG (8 · 9 · 10 · 11강과 같은 방식) ───────────
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


def err_text(e):
    return (e.excepinfo[2] if getattr(e, "excepinfo", None) and e.excepinfo[2] else str(e))[:300]


def shot_range(pv):
    return f"A1:{max(pv['last_col'], 'D')}{pv['last_row']}"


R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "notice": NOTICE, "pivot_version": PT_VER, "dest": DEST,
     "how": "코드로 만듦(PivotCaches.Create · CreatePivotTable · Orientation · AddDataField) — 화면에서 끌어 넣기 아님",
     "problems": [], "traps": [], "shots": {}}
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
try:
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    R["excel"] = {"version": xl.Version, "build": xl.Build}

    def new_book(with_answer, logs):
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count < 4:
            wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        wb.Worksheets(1).Name = "문제"
        build_problem_sheet(wb.Worksheets(1))
        for k, t in enumerate(TABLES):
            ws = wb.Worksheets(k + 2)
            ws.Name = t["sheet"]
            log = {}
            if with_answer:
                n = 2 + len(t["rows"])
                build_sheet(ws, t, False)
                make_pivot(ws, t, t, log=log)
                logs[t["key"]] = log
            else:
                build_sheet(ws, t, False)
        wb.Worksheets(1).Activate()
        author(wb)
        return wb

    paths, logs = {}, {}
    for with_answer, fname in ((False, "실습.xlsx"), (True, "정답.xlsx")):
        wb = new_book(with_answer, logs)
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

    # 원본 값 종류(날짜가 날짜로 들어갔나)
    wb = xl.Workbooks.Open(tmp_pr)
    ws = wb.Worksheets(T36["sheet"])
    R["q36_source"] = [{"cell": f"A{r}", "text": ws.Range(f"A{r}").Text, "value2": ws.Range(f"A{r}").Value2,
                        "d_text": ws.Range(f"D{r}").Text} for r in range(3, 13)]
    wb.Close(False)

    # 정답 — 저장한 정답 파일을 다시 열어 읽는다
    wb = xl.Workbooks.Open(tmp_ans)
    for n, t in enumerate(TABLES, start=FIRST_N):
        ws = wb.Worksheets(t["sheet"])
        pt = ws.PivotTables(1)
        pv = read_pivot(ws, pt)
        boxes = []
        if t is T34:
            boxes = find_cells(ws, pv["range"], lambda s: s == "595,000") + find_cells(ws, pv["range"], lambda s: s in ("합계 : 판매액",))
        elif t is T35:
            boxes = find_cells(ws, pv["range"], lambda s: s == "*") + find_cells(ws, pv["range"], lambda s: s in ("직위", "부서"))
        elif t is T36:
            boxes = find_cells(ws, pv["range"], lambda s: s.endswith("월")) + find_cells(ws, pv["range"], lambda s: s.startswith("합계 :"))
        R["problems"].append({"n": n, "key": t["key"], "text": t["text"], "build": logs.get(t["key"]), "pivot": pv, "boxes": boxes,
                              "fields": [{"name": pf.Name, "orientation": pf.Orientation} for pf in pt.PivotFields()]})
        R["shots"][f"{t['key']}_result"] = export(ws, shot_range(pv), f"{t['key']}_result", boxes)
    wb.Close(False)

    def trap(key, table, why, fn):
        """실습 사본(표만 있음)을 매번 새로 열어 fn(ws, item) — 저장하지 않음"""
        wb = xl.Workbooks.Open(tmp_pr)
        ws = wb.Worksheets(table["sheet"])
        item = {"key": key, "why": why}
        try:
            fn(ws, item)
        except pywintypes.com_error as e:
            item["rejected"] = err_text(e)
        R["traps"].append(item)
        wb.Close(False)

    # 34: 원본 값을 바꾸고 새로 고침 전 · 뒤
    def t_refresh(ws, item):
        log = {}
        pt = make_pivot(ws, T34, T34, log=log)
        cell = [c for c in find_cells(ws, pt.TableRange2.Address, lambda s: s == "595,000")]
        item["cell"] = cell
        ws.Range("D3").Value2 = 420000  # 인천 김바람 모니터 320,000 → 420,000
        item["source_changed"] = "D3 320,000 → 420,000"
        item["before_refresh"] = ws.Range(cell[0]).Text
        item["grand_before"] = read_pivot(ws, pt)["rows"][-1]["text"]
        item["after_refresh_only"] = read_pivot(ws, pt)["rows"]  # (이름은 요약 출력용) 새로 고침 전 피벗 전체
        item["shot_before"] = export(ws, shot_range(read_pivot(ws, pt)), "trap_q34_norefresh", cell + ["D3"])
        pt.RefreshTable()
        item["after_refresh"] = ws.Range(cell[0]).Text
        item["grand_after"] = read_pivot(ws, pt)["rows"][-1]["text"]
        item["after_change_source"] = read_pivot(ws, pt)["rows"]  # 새로 고침 뒤 피벗 전체
    trap("q34_refresh", T34, "원본 D3 을 바꾼 뒤 새로 고침 전 · 뒤", t_refresh)

    # 34: 원본 아래에 줄을 더하고 새로 고침만 / 데이터 원본 변경
    def t_addrow(ws, item):
        pt = make_pivot(ws, T34, T34)
        for j, v in enumerate(["광주", "문하람", "모니터", 100000]):
            ws.Range(f"{chr(65 + j)}13").Value2 = v
        item["added"] = "13행 광주 · 문하람 · 모니터 · 100,000"
        item["source"] = pt.SourceData
        pt.RefreshTable()
        item["after_refresh_only"] = read_pivot(ws, pt)["rows"]
        pt.ChangePivotCache(ws.Parent.PivotCaches().Create(1, ws.Range("A2:D13"), PT_VER))
        item["source_after_change"] = pt.SourceData
        item["after_change_source"] = read_pivot(ws, pt)["rows"]
    trap("q34_addrow", T34, "원본 아래 13행에 한 줄 더함 → 새로 고침만 / 데이터 원본 변경(A2:D13)", t_addrow)

    # 34: 머리글 한 칸이 비어 있으면
    def t_blankhead(ws, item):
        ws.Range("C2").Value2 = None
        item["changed"] = "C2(제품) 머리글을 지움"
        pc = ws.Parent.PivotCaches().Create(1, ws.Range("A2:D12"), PT_VER)
        pt = pc.CreatePivotTable(ws.Range(DEST), "피벗빈칸", True, PT_VER)
        item["made"] = True
        item["field_names"] = [pf.Name for pf in pt.PivotFields()]
    trap("q34_blankhead", T34, "머리글 C2 를 비운 채 피벗 테이블 만들기(코드) — 필드 이름", t_blankhead)

    # 34: 행 · 열을 바꿔 넣음
    def t_swap(ws, item):
        pt = make_pivot(ws, T34, dict(T34, rowf=["제품"], colf=["지점"]))
        item["pivot"] = read_pivot(ws, pt)
    trap("q34_swap", T34, "지점을 열, 제품을 행으로 바꿔 넣음", t_swap)

    # 35: 레이아웃을 그대로(압축 형식) · 빈 셀 그대로
    def t_compact(ws, item):
        log = {}
        pt = make_pivot(ws, T35, dict(T35, layout=None, null=None), log=log)
        item["build"] = log
        pv = read_pivot(ws, pt)
        item["pivot"] = pv
        boxes = find_cells(ws, pv["range"], lambda s: s in ("행 레이블", "열 레이블"))
        item["boxes"] = boxes
        item["shot"] = export(ws, shot_range(pv), "trap_q35_compact", boxes)
    trap("q35_compact", T35, "레이아웃 · 빈 셀을 건드리지 않음(엑셀이 고른 그대로) — 압축 형식 · 빈칸", t_compact)

    # 35: 테이블 형식(개요와 견줌)
    def t_tabular(ws, item):
        pt = make_pivot(ws, T35, dict(T35, layout="테이블"))
        item["pivot"] = read_pivot(ws, pt)
    trap("q35_tabular", T35, "보고서 레이아웃을 테이블 형식으로(개요 형식과 견줌)", t_tabular)

    # 35: 평균으로 안 바꿈(합계 그대로)
    def t_sum(ws, item):
        log = {}
        pt = make_pivot(ws, T35, dict(T35, data=[("기본급", None, "#,##0")]), log=log)
        item["build"] = log
        item["pivot"] = read_pivot(ws, pt)
    trap("q35_sum", T35, "값 필드를 평균으로 바꾸지 않음 — 엑셀이 고른 함수 그대로", t_sum)

    # 36: 값 함수를 엑셀이 고른 그대로(빈칸이 있는 수량)
    def t_count(ws, item):
        log = {}
        pt = make_pivot(ws, T36, dict(T36, data=[("수량", None, None)]), log=log)
        item["build"] = log
        pv = read_pivot(ws, pt)
        item["pivot"] = pv
        boxes = find_cells(ws, pv["range"], lambda s: s.startswith("개수 :") or s.startswith("합계 :")) + ["D7"]
        item["boxes"] = boxes
        item["shot"] = export(ws, shot_range(pv), "trap_q36_count", boxes)
    trap("q36_count", T36, "수량을 값에 넣고 함수를 고르지 않음 — 글자 '미정'(D7)이 있는 열", t_count)

    # 36: '미정'을 0 으로 바꾸면 엑셀이 고르는 함수
    def t_zero(ws, item):
        ws.Range("D7").Value2 = 0
        log = {}
        make_pivot(ws, T36, dict(T36, data=[("수량", None, None)], group=None), log=log)
        item["build"] = log
    trap("q36_zero", T36, "D7 '미정'을 0 으로 바꾸고 수량을 값에 넣음 — 엑셀이 고른 함수", t_zero)

    # 36: Orientation = xlDataField 로 넣었을 때(빈칸 있는 수량)
    def t_orient(ws, item):
        pt = make_pivot(ws, T36, dict(T36, data=[], group=None))
        pf = pt.PivotFields("수량")
        pf.Orientation = DATA
        df = pt.DataFields(1)
        item["function"], item["caption"] = FN_NAME.get(df.Function, df.Function), df.Caption
    trap("q36_orient", T36, "수량을 Orientation=값 으로 넣음(AddDataField 대신) — 글자 '미정'(D7) 있는 열", t_orient)

    # 36: '미정'을 지워 빈칸으로 두면
    def t_textqty(ws, item):
        ws.Range("D7").Value2 = None
        item["changed"] = "D7 '미정'을 지워 빈칸"
        log = {}
        pt = make_pivot(ws, T36, dict(T36, data=[("수량", None, None)]), log=log)
        item["build"] = log
        item["pivot"] = read_pivot(ws, pt)
    trap("q36_blankqty", T36, "D7 '미정'을 지워 빈칸으로 두고 수량을 값에 넣음 — 엑셀이 고른 함수", t_textqty)

    # 36: 그룹을 안 함
    def t_nogroup(ws, item):
        pt = make_pivot(ws, T36, dict(T36, group=None))
        pv = read_pivot(ws, pt)
        item["pivot"] = pv
        item["shot"] = export(ws, shot_range(pv), "trap_q36_nogroup", [])
    trap("q36_nogroup", T36, "판매일을 행에 넣고 그룹을 하지 않음(코드로 넣었을 때)", t_nogroup)

    # 36: 날짜 하나가 글자이면 그룹
    def t_textdate(ws, item):
        ws.Range("A8").NumberFormat = "@"
        ws.Range("A8").Value2 = "2026-10-29"
        item["changed"] = "A8 을 글자 '2026-10-29' 로(텍스트 서식)"
        item["a8"] = {"text": ws.Range("A8").Text, "value2": ws.Range("A8").Value2}
        log = {}
        pt = make_pivot(ws, T36, dict(T36, group=None), log=log)
        cell = pt.PivotFields("판매일").DataRange.Cells(1)
        item["group_cell"] = cell.Address.replace("$", "")
        try:
            cell.Group(True, True, None, (False, False, False, False, True, False, False))
            item["grouped"] = read_pivot(ws, pt)
        except pywintypes.com_error as e:
            item["group_rejected"] = err_text(e)
            item["pivot_ungrouped"] = read_pivot(ws, pt)
    trap("q36_textdate", T36, "판매일 한 칸(A8)이 글자일 때 '월' 그룹", t_textdate)

    # 36: 월과 함께 연도까지(견줌)
    def t_years(ws, item):
        pt = make_pivot(ws, T36, dict(T36, group=None))
        cell = pt.PivotFields("판매일").DataRange.Cells(1)
        cell.Group(True, True, None, (False, False, False, False, True, False, True))
        item["pivot"] = read_pivot(ws, pt)
    trap("q36_years", T36, "그룹에서 월과 연을 함께 고름(견줌)", t_years)

    wb = xl.Workbooks.Open(tmp_pr)
    for t in TABLES:
        ws = wb.Worksheets(t["sheet"])
        R["shots"][f"{t['key']}_start"] = export(ws, f"A1:D{2 + len(t['rows'])}", f"{t['key']}_start", [])
    wb.Close(False)
    R["file_props"] = {k: core_people(p) for k, p in paths.items()}
finally:
    xl.Quit()
    shutil.rmtree(TMP, ignore_errors=True)

with open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)


def brief(rows):
    return [(x["r"], " | ".join(x["text"])) for x in rows]


out = {"excel": R.get("excel"), "props": R.get("file_props"),
       "problems": [(p["n"], p["build"], p["pivot"]["range"], brief(p["pivot"]["rows"]), p["boxes"]) for p in R["problems"]],
       "traps": [], "fg": [(k, v.get("fg_class")) for k, v in R["shots"].items() if v.get("foreground_changed")]}
for t in R["traps"]:
    d = {k: v for k, v in t.items() if k not in ("pivot", "after_refresh_only", "after_change_source", "grouped", "pivot_ungrouped", "shot", "shot_before")}
    for k in ("pivot", "grouped", "pivot_ungrouped"):
        if k in t:
            d[k] = brief(t[k]["rows"])
    for k in ("after_refresh_only", "after_change_source"):
        if k in t:
            d[k] = brief(t[k])
    for k in ("shot", "shot_before"):
        if k in t and t[k].get("foreground_changed"):
            out["fg"].append((t["key"], t[k].get("fg_class")))
    out["traps"].append(d)
print(json.dumps(out, ensure_ascii=False, indent=1, default=str))
