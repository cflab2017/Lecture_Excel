"""엑셀 '이동 옵션'(홈 → 찾기 및 선택 → 이동 옵션…, xlDialogSelectSpecial) 대화상자를, 실습 표의 반 병합을 풀고 A3:A8 을 고른 채 띄워
찍고 닫는다 + 메뉴 글자(GetLabelMso). 저장하지 않는다. 결과 shots/s7_goto_special.png · 글자는 화면에 출력.
실행: python goto_special.py (pywin32 · pillow)"""
import os, sys, threading, time
import win32com.client as win32, win32gui, win32process, win32con

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "04_엑셀_자동채우기_행추가"))
from measure_util import grab  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")


def dialog_shot(pid, main, name):
    for _ in range(60):
        time.sleep(0.3)
        found = []

        def cb(h, _):
            if h != main and win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid \
                    and win32gui.GetClassName(h) in ("bosa_sdm_XL9", "#32770", "NUIDialog"):
                found.append(h)
            return True
        win32gui.EnumWindows(cb, 0)
        if found:
            time.sleep(1.2)
            print(name, "|", win32gui.GetWindowText(found[0]))
            grab(found[0], os.path.join(OUT, name))
            win32gui.PostMessage(found[0], win32con.WM_CLOSE, 0, 0)
            return
    print(name, "dialog not found")


xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    for k in ("FindDialogExcel", "FindAndSelectMenu", "GoToSpecial", "GoTo", "SelectSpecialDialog", "FindReplaceMenu"):
        try:
            print(k, "=", xl.CommandBars.GetLabelMso(k))
        except Exception as e:  # noqa: BLE001
            print(k, "ERR", str(e)[:40])
    wb = xl.Workbooks.Open(os.path.join(HERE, "실습.xlsx"), ReadOnly=True)
    xl.Visible = True
    ws = wb.Worksheets(1)
    ws.Range("A3:A8").UnMerge()
    ws.Range("A3:A8").Select()
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    th = threading.Thread(target=dialog_shot, args=(pid, xl.Hwnd, "s7_goto_special.png")); th.start()
    try:
        xl.Dialogs(132).Show()  # xlDialogSelectSpecial
    except Exception as e:  # noqa: BLE001
        print("dialog:", str(e)[:100])
    th.join()
    wb.Close(False)
finally:
    xl.Quit()
