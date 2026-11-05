"""기초 2강 공통 — 가상 꽃집 주문 목록 데이터, 시트 채우기, 창 찍기(PrintWindow)."""
import ctypes, ctypes.wintypes as W, os

HEAD = ("주문일", "주문자", "상품", "수량", "단가", "금액")
ROWS = [  # 주문일, 주문자, 상품, 수량, 단가 (금액 = 수량×단가)
    ("11월 3일", "한빛상사", "장미 꽃다발", 2, 35000),
    ("11월 3일", "서연", "축하 화환", 1, 89000),
    ("11월 4일", "도윤", "튤립 화분", 3, 18000),
    ("11월 4일", "한빛상사", "난 화분", 1, 65000),
    ("11월 5일", "지아", "장미 꽃다발", 1, 35000),
    ("11월 5일", "새봄유치원", "튤립 화분", 6, 18000),
    ("11월 6일", "민준", "프리저브드 꽃", 2, 42000),
    ("11월 6일", "서연", "장미 꽃다발", 3, 35000),
]
NEW = ("11월 7일", "하늘약국", "난 화분", 2, 65000)


def fill(ws, with_amount=True, head=HEAD, rows=ROWS):
    ws.Range("A1:F1").Value = (head,)
    for i, r in enumerate(rows, start=2):
        ws.Range(f"A{i}:E{i}").Value = (r,)
        if with_amount:
            ws.Range(f"F{i}").Formula = f"=D{i}*E{i}"
    for col, w in zip("ABCDEF", (10, 12, 15, 7, 10, 11)):
        ws.Columns(col).ColumnWidth = w
    ws.Range("E2:F20").NumberFormat = "#,##0"


def grab(hwnd, out):
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
    print(os.path.basename(out), w, h)
