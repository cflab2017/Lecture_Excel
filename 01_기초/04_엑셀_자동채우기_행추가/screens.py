"""엑셀 기초 4강 화면 — 보이는 엑셀(따로 띄운 것)에 수업표를 만들고 PrintWindow 로 창을 찍는다. 저장하지 않고 닫는다.
찍는 것: ① 번호·날짜·요일·주차를 첫 줄에서 채운 수업표(채운 범위 선택) ② 합계 행 바로 위에 끼운 새 행 — 합계 칸 수식 =SUM(F2:F11)·120
③ 범위 안(11행)에 끼운 새 행 — =SUM(F2:F12)·140. argv 'en' 이면 영어 데이터(파일 이름 끝 _en). 결과 shots/*.png. 실행: python screens.py [en]"""
import os, sys, time
import win32com.client as win32
from measure_util import grab, fill

EN = "en" in sys.argv[1:]
SUF = "_en" if EN else ""
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)
NEW = (11, "2026-11-16", "Mon" if EN else "월", "Week 3" if EN else "3주차", "Korean" if EN else "국어", 20)

xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    xl.Visible = True
    xl.WindowState = -4143
    xl.Width, xl.Height = 900, 520
    ws = wb.Worksheets(1)
    ws.Name = "Classes" if EN else "수업표"
    t = fill(ws, EN)
    xl.ActiveWindow.Zoom = 100
    # ① 첫 줄(2행)만 둔 상태에서 A2:D2 를 골라 11행까지 채운 결과 — 실제 채우기 핸들 끌기와 같은 AutoFill(기본)
    ws.Range(f"A3:D{t-1}").ClearContents()
    ws.Range("A2:A3").Value = ((1,), (2,))
    ws.Range("A2:A3").AutoFill(ws.Range(f"A2:A{t-1}"), 0)
    ws.Range("B2:D2").AutoFill(ws.Range(f"B2:D{t-1}"), 0)
    ws.Range(f"A2:D{t-1}").Select()
    time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, f"s1_filled{SUF}.png"))
    # 날짜를 평일만으로 다시(채우기 → 계열 → 평일 과 같은 형식 6)
    ws.Range("B2").AutoFill(ws.Range(f"B2:B{t-1}"), 6)
    ws.Range(f"B2:B{t-1}").Select()
    time.sleep(1.2)
    grab(xl.Hwnd, os.path.join(OUT, f"s1b_weekdays{SUF}.png"))
    # ② 새 시트(깨끗한 표) — 합계 바로 위에 새 행 → 합계 칸 선택(수식 입력줄에 수식이 보이게)
    ws1 = wb.Worksheets.Add(After=ws)
    ws1.Name = "AboveTotal" if EN else "합계위"
    t1 = fill(ws1, EN)
    ws1.Rows(t1).Insert()
    ws1.Range(f"A{t1}:F{t1}").Value = (NEW,)
    ws1.Range(f"B{t1}").NumberFormat = "yyyy-mm-dd"
    ws1.Range(f"F{t1+1}").Select()
    time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, f"s2_trap{SUF}.png"))
    # ③ 새 시트 — 범위 안(마지막 데이터 행)에 끼우기
    ws2 = wb.Worksheets.Add(After=ws1)
    ws2.Name = "Inside" if EN else "범위안"
    t2 = fill(ws2, EN)
    ws2.Rows(t2 - 1).Insert()
    ws2.Range(f"A{t2-1}:F{t2-1}").Value = (NEW,)
    ws2.Range(f"B{t2-1}").NumberFormat = "yyyy-mm-dd"
    ws2.Range(f"F{t2+1}").Select()
    time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, f"s3_inside{SUF}.png"))
    wb.Close(False)
finally:
    xl.Quit()
