"""엑셀 '계열' 대화상자(홈 → 채우기 → 계열…, xlDialogDataSeries)를 B2:B11 을 고른 채 띄워 PrintWindow 로 찍고 취소로 닫는다. 저장하지 않는다.
결과 shots/s4_series.png. 실행: python series_dialog.py"""
import os, threading, time
import win32com.client as win32, win32gui, win32process, win32con
from measure_util import grab, fill

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    xl.Visible = True
    ws = wb.Worksheets(1)
    t = fill(ws)
    ws.Range(f"B3:B{t-1}").ClearContents()
    ws.Range(f"B2:B{t-1}").Select()
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)

    def shoot():
        for _ in range(60):
            time.sleep(0.3)
            found = []
            def cb(h, _):
                if win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid and win32gui.GetClassName(h) == "bosa_sdm_XL9":
                    found.append(h)
                return True
            win32gui.EnumWindows(cb, 0)
            if found:
                time.sleep(1.2)
                print("dialog:", win32gui.GetWindowText(found[0]))
                grab(found[0], os.path.join(OUT, "s4_series.png"))
                win32gui.PostMessage(found[0], win32con.WM_CLOSE, 0, 0)
                return
        print("dialog not found")
    th = threading.Thread(target=shoot); th.start()
    r = xl.Dialogs(40).Show()   # xlDialogDataSeries
    th.join()
    print("Show returned", r, "B3 =", ws.Range("B3").Text)
    wb.Close(False)
finally:
    xl.Quit()
