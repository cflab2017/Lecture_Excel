"""기초 4강 공통 — 가상 학원 수업표(번호·날짜·요일·주차·과목·수강생) 10줄 + 합계, 창 찍기(PrintWindow: 3강 것과 같음).
데이터는 모두 연습용 가상 값이다(실제 학원 아님)."""
import ctypes, ctypes.wintypes as W, datetime

SUBJ = ["국어", "영어", "수학", "과학", "사회"]
SUBJ_EN = ["Korean", "English", "Math", "Science", "Social"]
PEOPLE = [12, 15, 9, 14, 11, 13, 10, 16, 8, 12]
START = datetime.date(2026, 11, 2)   # 월요일


def head(en=False):
    return ["No.", "Date", "Day", "Week", "Class", "Students"] if en else ["번호", "날짜", "요일", "주차", "과목", "수강생"]


def fill(ws, en=False, rows=10):
    """수업표 A1:F11 + 합계 행 12. 날짜는 평일만(월~금) 2주."""
    ws.Range("A1:F1").Value = (tuple(head(en)),)
    days, d = [], START
    while len(days) < rows:
        if d.weekday() < 5:
            days.append(d)
        d += datetime.timedelta(days=1)
    dn = ["Mon", "Tue", "Wed", "Thu", "Fri"] if en else ["월", "화", "수", "목", "금"]
    for i, day in enumerate(days):
        r = i + 2
        wk = f"Week {1 + (day - START).days // 7}" if en else f"{1 + (day - START).days // 7}주차"
        ws.Range(f"A{r}").Value = i + 1
        ws.Range(f"B{r}").Value = day.isoformat()
        ws.Range(f"B{r}").NumberFormat = "yyyy-mm-dd"
        ws.Range(f"C{r}").Value = dn[day.weekday()]
        ws.Range(f"D{r}").Value = wk
        ws.Range(f"E{r}").Value = (SUBJ_EN if en else SUBJ)[i % 5]
        ws.Range(f"F{r}").Value = PEOPLE[i]
    t = rows + 2
    ws.Range(f"E{t}").Value = "Total" if en else "합계"
    ws.Range(f"F{t}").Formula = f"=SUM(F2:F{t-1})"
    ws.Range("A1:F1").Font.Bold = True
    ws.Range(f"E{t}:F{t}").Font.Bold = True
    for c, w in zip("ABCDEF", (6, 12, 6, 9, 10, 9)):
        ws.Columns(c).ColumnWidth = w
    return t


def grab(hwnd, out):
    """창을 PrintWindow 로 찍어 PNG 로 (3강 measure_util.grab 과 같음)."""
    from PIL import Image
    u32, g32 = ctypes.windll.user32, ctypes.windll.gdi32
    u32.SetProcessDPIAware()
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
