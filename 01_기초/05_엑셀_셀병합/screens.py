"""엑셀 기초 5강(셀 병합) 화면 — 보이는 엑셀(따로 띄운 것)에서 PrintWindow 로 창 · 대화상자를 찍는다. 저장하지 않고 닫는다.
s1 실습 표(제목 · 반 병합, A3:A5 선택) · s2 병합 경고(A1:C1 에 A·B·C 를 넣고 병합 → 경고 창, 취소로 닫음)
s3 병합 칸이 섞인 표에서 정렬 → 오류 창 · s4 '1반' 필터 → 한 줄만 · s5 정답 표(병합 해제 · 채움 · 정렬 · 필터)
s6 셀 서식 '맞춤' 대화상자(A1:C1 선택 영역의 가운데로). 결과 shots/*.png. 실행: python screens.py (pywin32 · pillow)"""
import os, sys, threading, time
import win32com.client as win32, win32gui, win32process, win32con

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "04_엑셀_자동채우기_행추가"))
from measure_util import grab  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)


def dialog_shot(pid, main, name, close=True):
    """main 창이 아닌 그 엑셀 창(대화상자)이 뜨면 찍고 닫는다."""
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
            h = found[0]
            texts = []
            win32gui.EnumChildWindows(h, lambda c, _: texts.append(win32gui.GetWindowText(c)) or True, 0)
            print(name, "|", win32gui.GetClassName(h), "|", win32gui.GetWindowText(h), "|", [t for t in texts if t][:8])
            grab(h, os.path.join(OUT, name))
            if close:
                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
            return
    print(name, "dialog not found")


def open_book(p):
    wb = xl.Workbooks.Open(os.path.join(HERE, p), ReadOnly=True)
    xl.ActiveWindow.Zoom = 100
    return wb


xl = win32.DispatchEx("Excel.Application")
try:
    xl.DisplayAlerts = False
    xl.Visible = True
    xl.WindowState = -4143
    xl.Width, xl.Height = 760, 470
    _, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
    # s1 실습 표
    wb = open_book("실습.xlsx")
    ws = wb.Worksheets(1)
    ws.Range("A3:A5").Select()
    time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, "s1_practice.png"))
    # s3 정렬 오류 — 표 범위 고르고 정렬 대화상자(데이터 → 정렬)
    # 점수 칸 하나를 누르고 [데이터 → 내림차순 정렬] 과 같은 명령(ExecuteMso) — 정렬 대화상자(Dialogs 39)는 확인 전이라 오류가 안 뜸(11/9)
    ws.Range("C3").Select()
    th = threading.Thread(target=dialog_shot, args=(pid, xl.Hwnd, "s3_sort_error.png")); th.start()
    xl.DisplayAlerts = True  # 끄면 오류 창도 안 뜨고 정렬만 조용히 안 됨(11/9)
    try:
        xl.CommandBars.ExecuteMso("SortDescendingExcel")
    except Exception as e:  # noqa: BLE001
        print("sort:", str(e)[:120])
    th.join()
    xl.DisplayAlerts = False
    print("C3 after sort try =", ws.Range("C3").Value, "B3 =", ws.Range("B3").Value)
    # s4 필터 1반
    ws.Range("A2:C8").AutoFilter(Field=1, Criteria1="1반")
    ws.Range("A2").Select()
    time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, "s4_filter.png"))
    wb.Close(False)
    # s2 병합 경고 — 새 통합 문서
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = (("성적표", "1학기", "2026"),)
    ws.Range("A1:C1").Select()
    xl.DisplayAlerts = True
    th = threading.Thread(target=dialog_shot, args=(pid, xl.Hwnd, "s2_merge_warning.png")); th.start()
    try:
        ws.Range("A1:C1").Merge()
    except Exception as e:  # noqa: BLE001
        print("merge:", str(e)[:120])
    th.join()
    xl.DisplayAlerts = False
    print("after cancel B1 =", ws.Range("B1").Value, "merged =", ws.Range("A1").MergeCells)
    wb.Close(False)
    # s5 정답 표 · s6 맞춤 대화상자
    wb = open_book("정답.xlsx")
    ws = wb.Worksheets(1)
    ws.Range("A3:A8").Select()
    time.sleep(1.5)
    grab(xl.Hwnd, os.path.join(OUT, "s5_answer.png"))
    ws.Range("A1:C1").Select()
    th = threading.Thread(target=dialog_shot, args=(pid, xl.Hwnd, "s6_center_across.png")); th.start()
    try:
        xl.Dialogs(43).Show()  # xlDialogAlignment
    except Exception as e:  # noqa: BLE001
        print("alignment dialog:", str(e)[:120])
    th.join()
    wb.Close(False)
finally:
    xl.DisplayAlerts = False
    xl.Quit()
