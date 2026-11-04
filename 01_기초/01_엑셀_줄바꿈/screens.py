"""엑셀 기초 1강 화면 — 보이는 엑셀(따로 띄운 것)에 시험 시트를 만들고 PrintWindow 로 창을 찍는다. 저장하지 않고 닫는다.
찍는 것: ① 정답 표(준비물 칸 줄바꿈) ② 홈 탭 맞춤 그룹의 '자동 줄 바꿈'이 눌린 리본 ③ 함정 네 줄(수식 줄바꿈 꺼짐/켬, 줄 높이 고정 잘림/자동 맞춤).
결과: shots/*.png (원본 크기). 실행: python screens.py"""
import ctypes, ctypes.wintypes as W, os, sys, time
import win32com.client as win32
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)
u32, g32 = ctypes.windll.user32, ctypes.windll.gdi32
u32.SetProcessDPIAware()


def grab(hwnd, out):
    r = W.RECT(); u32.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    hdc = u32.GetWindowDC(hwnd); mdc = g32.CreateCompatibleDC(hdc); bmp = g32.CreateCompatibleBitmap(hdc, w, h)
    g32.SelectObject(mdc, bmp); u32.PrintWindow(hwnd, mdc, 2)
    class BMI(ctypes.Structure):
        _fields_ = [("a", W.DWORD), ("w", W.LONG), ("h", W.LONG), ("p", W.WORD), ("b", W.WORD), ("c", W.DWORD), ("s", W.DWORD),
                    ("x", W.LONG), ("y", W.LONG), ("u", W.DWORD), ("i", W.DWORD)]
    bi = BMI(ctypes.sizeof(BMI), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    g32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bi), 0)
    Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1).convert("RGB").save(out)
    g32.DeleteObject(bmp); g32.DeleteDC(mdc); u32.ReleaseDC(hwnd, hdc)
    print(os.path.basename(out), w, h)


xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(os.path.join(HERE, "정답.xlsx"))
    xl.Visible = True
    xl.WindowState = -4143
    xl.Width, xl.Height = 1000, 470
    ws = wb.Worksheets(1); ws.Activate()
    xl.ActiveWindow.Zoom = 130
    xl.ActiveWindow.ScrollRow, xl.ActiveWindow.ScrollColumn = 1, 1
    ws.Range("D2").Select(); time.sleep(2.0)
    grab(xl.Hwnd, os.path.join(OUT, "s1_answer.png"))
    # 함정 시트
    t = wb.Worksheets.Add(); t.Name = "함정"; t.Activate()
    t.Range("A1:B1").Value = (("상황", "결과"),); t.Range("A1:B1").Font.Bold = True
    t.Columns("A").ColumnWidth = 24; t.Columns("B").ColumnWidth = 17
    t.Range("A2").Value = "수식 · 자동 줄 바꿈 꺼짐"; t.Range("B2").Formula = '="11월 10일"&CHAR(10)&"3층 대회의실"'; t.Range("B2").WrapText = False
    t.Range("A3").Value = "수식 · 자동 줄 바꿈 켬"; t.Range("B3").Formula = '="11월 10일"&CHAR(10)&"3층 대회의실"'; t.Range("B3").WrapText = True
    t.Rows(4).RowHeight = 16.5
    t.Range("A4").Value = "줄 높이를 직접 정한 줄"; t.Range("B4").Value = "노트북\n빔프로젝터"; t.Range("B4").WrapText = True; t.Rows(4).RowHeight = 16.5
    t.Range("A5").Value = "행 높이 자동 맞춤 뒤"; t.Range("B5").Value = "노트북\n빔프로젝터"; t.Range("B5").WrapText = True; t.Rows(5).AutoFit()
    t.Rows(3).AutoFit()
    t.Range("A2:A5").VerticalAlignment = -4108
    xl.ActiveWindow.Zoom = 130
    xl.ActiveWindow.ScrollRow, xl.ActiveWindow.ScrollColumn = 1, 1
    t.Range("B3").Select(); time.sleep(2.0)
    grab(xl.Hwnd, os.path.join(OUT, "s2_traps.png"))
    wb.Close(False)
finally:
    xl.Quit()
