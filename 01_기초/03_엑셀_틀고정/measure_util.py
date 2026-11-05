"""기초 3강 공통 — 가상 문구점 품목 60개의 월별 판매 수량(개), 시트 채우기, 창 찍기(PrintWindow).
데이터는 모두 연습용 가상 값이다(실제 가게 아님). 같은 씨앗(seed)으로 늘 같은 값이 나온다."""
import ctypes, ctypes.wintypes as W, os, random

MONTHS = [f"{m}월" for m in range(1, 13)]
MONTHS_EN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_KINDS = ["볼펜", "연필", "샤프", "지우개", "형광펜", "네임펜", "공책", "스프링노트", "포스트잇", "클립", "스테이플러", "풀", "가위", "자", "수정테이프", "파일철"]
_COLORS = ["검정", "빨강", "파랑", "초록"]
_KINDS_EN = ["Ballpoint pen", "Pencil", "Mechanical pencil", "Eraser", "Highlighter", "Marker", "Notebook", "Spiral notebook", "Sticky notes", "Paper clips",
             "Stapler", "Glue stick", "Scissors", "Ruler", "Correction tape", "Binder"]
_COLORS_EN = ["black", "red", "blue", "green"]
N_ITEMS = 60


def items(en=False):
    kinds, colors = (_KINDS_EN, _COLORS_EN) if en else (_KINDS, _COLORS)
    out = []
    for i in range(N_ITEMS):
        k = kinds[i % len(kinds)]
        c = colors[(i // len(kinds)) % len(colors)]
        out.append(f"{k} ({c})" if en else f"{k}({c})")
    return out


def sales():
    rnd = random.Random(20261106)
    return [[rnd.randint(5, 90) for _ in range(12)] for _ in range(N_ITEMS)]


def fill(ws, en=False):
    head = (["No.", "Item"] + MONTHS_EN + ["Total"]) if en else (["번호", "품목"] + MONTHS + ["합계"])
    ws.Range("A1:O1").Value = (tuple(head),)
    S = sales()
    names = items(en)
    for i in range(N_ITEMS):
        r = i + 2
        ws.Range(f"A{r}:N{r}").Value = (tuple([i + 1, names[i]] + S[i]),)
        ws.Range(f"O{r}").Formula = f"=SUM(C{r}:N{r})"
    ws.Range("A1:O1").Font.Bold = True
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 18
    for c in "CDEFGHIJKLMN":
        ws.Columns(c).ColumnWidth = 6
    ws.Columns("O").ColumnWidth = 7
    return N_ITEMS + 1   # 마지막 행


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
