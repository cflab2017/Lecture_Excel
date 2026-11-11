"""엑셀 기초 7강(시트 복사 · 보호) 화면 — 보이는 엑셀(따로 띄운 것)에서 PrintWindow 로 찍는다. 실습.xlsx 는 읽기 전용으로 열고 저장하지 않는다.
대화상자는 엑셀 창 그림 위 제자리에 겹쳐 붙인다(11/6 규칙: 창 전체가 보이게).
s1 이동/복사 대화상자(시트 탭 오른쪽 클릭 → 이동/복사 와 같은 창, xlDialogWorkbookCopy)
s2 같은 파일 안 복사 뒤 시트 탭(11월 · 11월 (2) · 설정)
s3 11월 시트만 새 통합 문서로 복사 — C3 수식 입력줄에 원본 파일 이름([실습.xlsx]설정!$B$1)
s4 셀 서식 → 보호 탭(B3:B8 고른 채, xlDialogCellProtection)
s5 시트 보호 대화상자(xlDialogProtectDocument)
s6 보호된 시트의 잠긴 칸(C3)에 글자를 치면 뜨는 경고
결과 shots/*.png, 읽은 값 screens.json. 실행: python screens.py (pywin32 · pillow)"""
import json, os, sys, threading, time
import win32com.client as win32, win32con, win32gui, win32process
from PIL import Image

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "04_엑셀_자동채우기_행추가"))
from measure_util import grab  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)
R = {}
DLG_CLASSES = ("bosa_sdm_XL9", "#32770", "NUIDialog")


def find_dialog(pid, exclude=()):
    found = []

    def cb(h, _):
        if (win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid
                and win32gui.GetClassName(h) in DLG_CLASSES and h not in exclude):
            found.append(h)
        return True
    win32gui.EnumWindows(cb, 0)
    return found[0] if found else None


def composite(main_hwnd, dlg_hwnd, out):
    """엑셀 창 그림 위에 대화상자 그림을 화면 위치 그대로 겹쳐 붙인다."""
    a, b = out + ".main.png", out + ".dlg.png"
    mr, dr = win32gui.GetWindowRect(main_hwnd), win32gui.GetWindowRect(dlg_hwnd)
    mw, mh, dw, dh = mr[2] - mr[0], mr[3] - mr[1], dr[2] - dr[0], dr[3] - dr[1]
    # 대화상자가 엑셀 창 밖으로 나가면 잘리므로 창 가운데로 옮긴다(크기는 그대로)
    win32gui.SetWindowPos(dlg_hwnd, 0, mr[0] + max(0, (mw - dw) // 2), mr[1] + max(40, (mh - dh) // 2), 0, 0,
                          win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
    time.sleep(0.8)
    grab(main_hwnd, a)
    grab(dlg_hwnd, b)
    m, d = Image.open(a), Image.open(b)
    mr, dr = win32gui.GetWindowRect(main_hwnd), win32gui.GetWindowRect(dlg_hwnd)
    m.paste(d, (dr[0] - mr[0], dr[1] - mr[1]))
    m.save(out)
    os.remove(a); os.remove(b)
    return {"title": win32gui.GetWindowText(dlg_hwnd), "class": win32gui.GetClassName(dlg_hwnd),
            "offset": (dr[0] - mr[0], dr[1] - mr[1]), "size": d.size}


def show_and_shoot(dialog_no, out):
    """xl.Dialogs(n).Show() 는 창이 닫힐 때까지 멈추므로 다른 스레드에서 찾아 찍고 닫는다(취소)."""
    res = {}

    def shoot():
        for _ in range(60):
            time.sleep(0.3)
            h = find_dialog(PID)
            if h:
                time.sleep(1.2)
                res.update(composite(HWND, h, os.path.join(OUT, out)))
                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                return
        res["err"] = "dialog not found"
    th = threading.Thread(target=shoot)
    th.start()
    try:
        res["show"] = xl.Dialogs(dialog_no).Show()
    except Exception as e:  # noqa: BLE001
        res["show_err"] = str(e)[:120]
    th.join()
    return res


def excel7():
    hits = []
    win32gui.EnumChildWindows(HWND, lambda h, _: hits.append(h) if win32gui.GetClassName(h) == "EXCEL7" else None or True, 0)
    return hits


xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    xl.Visible = True
    xl.WindowState = -4143
    xl.Width, xl.Height = 760, 610  # 셀 서식 창(약 560px)이 들어가게
    HWND = xl.Hwnd
    PID = win32process.GetWindowThreadProcessId(HWND)[1]
    wb = xl.Workbooks.Open(os.path.join(HERE, "실습.xlsx"), 0, True)  # UpdateLinks 0, ReadOnly
    xl.ActiveWindow.Zoom = 100
    ws = wb.Worksheets("11월")
    ws.Activate()
    ws.Range("A1").Select()
    time.sleep(1)
    # s1 이동/복사
    R["s1"] = show_and_shoot(283, "s1_move_copy.png")
    # s2 같은 파일 안 복사
    ws.Copy(None, ws)  # (Before 비움, After=ws) — 이름 붙인 인자는 새 통합 문서로 복사된다(measure.py)
    R["s2_tabs"] = [s.Name for s in wb.Worksheets]
    time.sleep(1.2)
    grab(HWND, os.path.join(OUT, "s2_tabs.png"))
    wb.Worksheets("11월 (2)").Delete()
    ws.Activate()
    # s3 새 통합 문서로(시트 하나만)
    before = [b.Name for b in xl.Workbooks]
    ws.Copy()
    nb = [b for b in xl.Workbooks if b.Name not in before][0]
    nb.Activate()
    xl.ActiveWindow.Zoom = 100
    nb.Worksheets(1).Range("C3").Select()
    time.sleep(1.5)
    R["s3"] = {"book": nb.Name, "c3": nb.Worksheets(1).Range("C3").Formula, "links": list(nb.LinkSources(1) or [])}
    R["s3"]["links"] = [os.path.basename(x) for x in R["s3"]["links"]]
    grab(xl.Hwnd, os.path.join(OUT, "s3_external_link.png"))
    nb.Close(False)
    wb.Activate()
    ws.Activate()
    # s4 셀 서식 → 보호 탭
    ws.Range("B3:B8").Select()
    time.sleep(0.8)
    R["s4"] = show_and_shoot(46, "s4_cell_protection.png")
    # s5 시트 보호 대화상자
    ws.Range("A1").Select()
    R["s5"] = show_and_shoot(28, "s5_protect_sheet.png")
    # s6 잠긴 칸에 입력 → 경고
    ws.Range("B3:B8").Locked = False
    ws.Protect()
    ws.Range("C3").Select()
    time.sleep(1)
    for h in excel7():  # 사람이 5 를 치는 것과 같게 — 누름 · 글자 · 뗌
        win32gui.PostMessage(h, win32con.WM_KEYDOWN, 0x35, 0x00060001)
        win32gui.PostMessage(h, win32con.WM_CHAR, ord("5"), 0x00060001)
        win32gui.PostMessage(h, win32con.WM_KEYUP, 0x35, 0xC0060001)
    dlg = None
    for _ in range(30):
        time.sleep(0.3)
        dlg = find_dialog(PID)
        if dlg:
            break
    if dlg:
        time.sleep(1)
        R["s6"] = composite(HWND, dlg, os.path.join(OUT, "s6_locked_warning.png"))
        win32gui.PostMessage(dlg, win32con.WM_CLOSE, 0, 0)
        time.sleep(1)
    else:
        R["s6"] = {"err": "warning not found"}
    R["s6"]["c3_text"] = ws.Range("C3").Text
    wb.Close(False)
finally:
    xl.DisplayAlerts = False
    xl.Quit()
json.dump(R, open(os.path.join(HERE, "screens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
