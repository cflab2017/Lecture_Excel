"""엑셀 기초 6강(합계 · 평균) 화면 — 보이는 엑셀(따로 띄운 것)에서 PrintWindow 로 창을 찍는다. 저장하지 않고 닫는다.
s1 실습 파일 B3:B9 선택 — 상태 표시줄 평균 · 개수 · 합계, 목요일 칸의 초록 삼각형
s2 실습 파일 B10 에서 자동 합계 — 엑셀이 제안하는 범위(점선)를 Enter 전에 찍고, Enter 뒤 식을 읽음
s3 정답 파일 — 합계 · 평균 · 판 날만 평균 · 보이는 행 합계
s4 정답 파일에 판매 수량 >= 15 필터 — 합계 90 그대로 · 보이는 행 합계 78
결과 shots/*.png, 읽은 값 screens.json. 실행: python screens.py (pywin32 · pillow)"""
import json, os, sys, time
import win32com.client as win32, win32con, win32gui

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "04_엑셀_자동채우기_행추가"))
from measure_util import grab  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)
R = {}


def key(hwnd, vk):
    hits = []
    win32gui.EnumChildWindows(hwnd, lambda h, _: hits.append(h) if win32gui.GetClassName(h) == "EXCEL7" else None or True, 0)
    for h in hits:
        win32gui.PostMessage(h, win32con.WM_KEYDOWN, vk, 0)
        win32gui.PostMessage(h, win32con.WM_KEYUP, vk, 0)
    time.sleep(1.5)


def wait_ready(f):
    for _ in range(20):
        try:
            return f()
        except Exception:  # noqa: BLE001
            time.sleep(0.5)
    return None


xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    xl.Visible = True
    xl.WindowState = -4143
    xl.Width, xl.Height = 760, 470
    HWND = xl.Hwnd
    wb = xl.Workbooks.Open(os.path.join(HERE, "실습.xlsx"), ReadOnly=True)
    xl.ActiveWindow.Zoom = 100
    ws = wb.Worksheets(1)
    ws.Range("B3:B9").Select()
    time.sleep(1.5)
    grab(HWND, os.path.join(OUT, "s1_status_bar.png"))
    R["b6_text_flag"] = bool(ws.Range("B6").Errors(3).Value)
    # s2 자동 합계 제안(Enter 전) → 찍고 Enter
    ws.Range("B10").Select()
    time.sleep(0.8)
    xl.CommandBars.ExecuteMso("AutoSum")
    time.sleep(1.5)
    grab(HWND, os.path.join(OUT, "s2_autosum_suggest.png"))
    key(HWND, win32con.VK_RETURN)
    R["autosum_b10"] = wait_ready(lambda: {"formula": ws.Range("B10").Formula, "text": ws.Range("B10").Text})
    wb.Close(False)
    # s3 정답
    wb = xl.Workbooks.Open(os.path.join(HERE, "정답.xlsx"), ReadOnly=True)
    xl.ActiveWindow.Zoom = 100
    ws = wb.Worksheets(1)
    ws.Range("B11").Select()
    time.sleep(1.5)
    grab(HWND, os.path.join(OUT, "s3_answer.png"))
    R["answer"] = [[ws.Range(f"A{k}").Text, ws.Range(f"B{k}").Formula, ws.Range(f"B{k}").Text] for k in range(11, 17)]
    # s4 필터 >= 15
    ws.Range("A2:B9").AutoFilter(Field=2, Criteria1=">=15")
    ws.Range("B14").Select()
    time.sleep(1.5)
    grab(HWND, os.path.join(OUT, "s4_filter.png"))
    R["filter"] = {"visible": [ws.Range(f"A{i}").Text for i in range(3, 10) if not ws.Rows(i).Hidden],
                   "b11": ws.Range("B11").Text, "b14": ws.Range("B14").Text, "b12": ws.Range("B12").Text}
    wb.Close(False)
finally:
    xl.DisplayAlerts = False
    xl.Quit()
json.dump(R, open(os.path.join(HERE, "screens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
