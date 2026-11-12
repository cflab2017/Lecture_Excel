"""엑셀 기초 8강(단축키) 실측 — 실습 파일을 읽기 전용으로 열고, 단축키를 엑셀 창(EXCEL7)에 직접 보내 결과를 읽는다. 저장하지 않는다.
키 보내기: AttachThreadInput + SetKeyboardState 로 Ctrl · Shift · Alt 를 눌린 상태로 만든 뒤 PostMessage — 창을 앞으로 가져오지 않고
마우스도 쓰지 않는다(다른 세션이 바탕화면을 쓰는 중이어도 겹치지 않게). 입력기(IME)를 거치지 않는 방식이라 한/영 상태에 따른 차이는 재지 않았다.
결과 measure.json(·measure_en.json), 화면 shots/*.png(KR 만). 실행: python measure.py [en]"""
import ctypes, datetime, json, os, shutil, sys, tempfile, time
import pythoncom, pywintypes
import win32api, win32con, win32gui, win32process
import win32com.client as win32

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "04_엑셀_자동채우기_행추가"))
from measure_util import grab  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
EN = len(sys.argv) > 1 and sys.argv[1] == "en"
SRC = os.path.join(HERE, "practice.xlsx" if EN else "실습.xlsx")
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)
u32, k32 = ctypes.windll.user32, ctypes.windll.kernel32
R = {"when": datetime.datetime.now().isoformat(timespec="seconds"), "file": os.path.basename(SRC)}
EXT = {win32con.VK_DOWN, win32con.VK_UP, win32con.VK_LEFT, win32con.VK_RIGHT, win32con.VK_HOME, win32con.VK_END, win32con.VK_NEXT, win32con.VK_PRIOR}


def com(f, tries=40):
    """엑셀이 편집 중이면 COM 호출이 거절된다 — 잠깐씩 다시."""
    for _ in range(tries):
        try:
            return f()
        except pywintypes.com_error as e:
            if e.args[0] in (-2147418111, -2146777998):
                time.sleep(0.25)
                continue
            raise
    return f()


class Keys:
    def __init__(self, xl):
        self.xl = xl
        main = xl.Hwnd
        kids = []
        win32gui.EnumChildWindows(main, lambda h, _: kids.append(h) if win32gui.GetClassName(h) == "EXCEL7" else None, 0)
        self.grid = kids[0]
        self.tid = win32process.GetWindowThreadProcessId(self.grid)[0]
        self.pid = win32process.GetWindowThreadProcessId(main)[1]
        self.me = k32.GetCurrentThreadId()

    def send(self, vk, ctrl=False, shift=False, alt=False, wait=0.7):
        u32.AttachThreadInput(self.me, self.tid, True)
        st = (ctypes.c_ubyte * 256)()
        u32.GetKeyboardState(st)
        old = bytes(st)
        for on, a, b in ((ctrl, win32con.VK_CONTROL, win32con.VK_LCONTROL), (shift, win32con.VK_SHIFT, win32con.VK_LSHIFT), (alt, win32con.VK_MENU, win32con.VK_LMENU)):
            if on:
                st[a] = 0x80; st[b] = 0x80
        u32.SetKeyboardState(st)
        sc = win32api.MapVirtualKey(vk, 0)
        lp = 1 | (sc << 16) | ((1 << 24) if vk in EXT else 0) | ((1 << 29) if alt else 0)
        down, up = (win32con.WM_SYSKEYDOWN, win32con.WM_SYSKEYUP) if alt else (win32con.WM_KEYDOWN, win32con.WM_KEYUP)
        win32gui.PostMessage(self.grid, down, vk, lp)
        time.sleep(0.35)
        win32gui.PostMessage(self.grid, up, vk, lp | 0xC0000000)
        time.sleep(0.25)
        u32.SetKeyboardState((ctypes.c_ubyte * 256).from_buffer_copy(old))
        u32.AttachThreadInput(self.me, self.tid, False)
        time.sleep(wait)

    def char(self, s):
        for ch in s:
            win32gui.PostMessage(self.grid, win32con.WM_CHAR, ord(ch), 1)
            time.sleep(0.08)
        time.sleep(0.3)

    def dialog(self):
        found = []

        def cb(h, _):
            if (win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == self.pid
                    and win32gui.GetClassName(h) in ("bosa_sdm_XL9", "#32770", "NUIDialog")):
                found.append(h)
            return True
        win32gui.EnumWindows(cb, 0)
        return found[0] if found else None


def main():
    pythoncom.CoInitialize()
    work = os.path.join(tempfile.mkdtemp(), os.path.basename(SRC))
    shutil.copy(SRC, work)
    fg0 = win32gui.GetForegroundWindow()
    xl = win32.DispatchEx("Excel.Application")
    try:
        xl.Visible = True
        xl.DisplayAlerts = False
        xl.WindowState = -4143
        xl.Width, xl.Height = 760, 577
        wb = xl.Workbooks.Open(work)
        ws = wb.Worksheets(1)
        xl.ActiveWindow.Zoom = 100
        time.sleep(1.5)
        K = Keys(xl)
        sel = lambda: com(lambda: xl.Selection.Address.replace("$", ""))
        act = lambda: com(lambda: xl.ActiveCell.Address.replace("$", ""))
        go = lambda a: (com(lambda: ws.Range(a).Select()), time.sleep(0.3))

        # 1 Ctrl+방향키 — 빈 줄(8행)에서 멈춤
        go("E2"); jumps = []
        for _ in range(4):
            K.send(win32con.VK_DOWN, ctrl=True); jumps.append(act())
        R["ctrl_down_from_E2"] = jumps
        go("E2"); K.send(win32con.VK_DOWN, ctrl=True, shift=True)
        R["ctrl_shift_down_from_E2"] = sel()
        if not EN:
            com(lambda: setattr(xl.ActiveWindow, "ScrollRow", 1)); com(lambda: setattr(xl.ActiveWindow, "ScrollColumn", 1)); time.sleep(0.6)
            grab(xl.Hwnd, os.path.join(OUT, "s1_ctrl_shift_down.png"))
        go("E9"); K.send(win32con.VK_UP, ctrl=True); R["ctrl_up_from_E9"] = act()
        # 2 Ctrl+Home · Ctrl+End
        go("E5"); K.send(win32con.VK_HOME, ctrl=True); R["ctrl_home"] = act()
        K.send(win32con.VK_END, ctrl=True); R["ctrl_end"] = act()
        com(lambda: ws.Range("H30").__setattr__("Value", "x")); com(lambda: ws.Range("H30").ClearContents())
        go("A1"); K.send(win32con.VK_END, ctrl=True); R["ctrl_end_after_clear_H30"] = act()
        cp = os.path.join(tempfile.mkdtemp(), "copy.xlsx")
        com(lambda: wb.SaveCopyAs(cp))
        wb2 = com(lambda: xl.Workbooks.Open(cp)); time.sleep(1)
        com(lambda: wb2.Worksheets(1).Range("A1").Select())
        K2 = Keys(xl)
        K2.send(win32con.VK_END, ctrl=True); R["ctrl_end_after_save_reopen"] = act()
        com(lambda: wb2.Close(False)); com(lambda: wb.Activate()); time.sleep(0.8)
        # 3 Alt+= 자동 합계 — 빈 줄 위까지만
        go("E14"); K.send(0xBB, alt=True, wait=1.0)
        K.send(win32con.VK_RETURN)
        R["alt_equal_E14"] = {"formula": com(lambda: ws.Range("E14").Formula), "value": com(lambda: ws.Range("E14").Value)}
        go("E14")
        if not EN:
            com(lambda: setattr(xl.ActiveWindow, "ScrollRow", 1)); com(lambda: setattr(xl.ActiveWindow, "ScrollColumn", 1)); time.sleep(0.6)
            grab(xl.Hwnd, os.path.join(OUT, "s2_alt_equal_trap.png"))
        com(lambda: ws.Range("E14").ClearContents())
        go("E2:E13"); K.send(0xBB, alt=True, wait=1.2)
        R["alt_equal_with_E2_E13_selected"] = {"formula": com(lambda: ws.Range("E14").Formula), "value": com(lambda: ws.Range("E14").Value)}
        # 4 F4 — 수식 편집 중이면 $ 바꾸기
        com(lambda: ws.Range("F2").__setattr__("Formula", "=E2/E14")); cyc = []
        for _ in range(4):
            go("F2"); K.send(win32con.VK_F2, wait=0.4); K.send(win32con.VK_F4, wait=0.4); K.send(win32con.VK_RETURN)
            cyc.append(com(lambda: ws.Range("F2").Formula))
        R["f4_cycle_from_=E2/E14"] = cyc
        com(lambda: ws.Range("F2").ClearContents())
        # 5 F4 — 편집 중이 아니면 마지막 동작 반복(굵게)
        go("B2"); K.send(ord("B"), ctrl=True)
        go("B9"); K.send(win32con.VK_F4)
        R["f4_repeat_bold"] = {"B2": com(lambda: ws.Range("B2").Font.Bold), "B9_after_F4": com(lambda: ws.Range("B9").Font.Bold), "B10_untouched": com(lambda: ws.Range("B10").Font.Bold)}
        # 6 Ctrl+; · Ctrl+Shift+;
        go("A16"); K.send(0xBA, ctrl=True, wait=0.5); K.send(win32con.VK_RETURN)
        go("B16"); K.send(0xBA, ctrl=True, shift=True, wait=0.5); K.send(win32con.VK_RETURN)
        R["ctrl_semicolon"] = {"A16_text": com(lambda: ws.Range("A16").Text), "A16_formula": com(lambda: ws.Range("A16").Formula), "A16_format": com(lambda: ws.Range("A16").NumberFormat),
                               "B16_text": com(lambda: ws.Range("B16").Text), "B16_formula": com(lambda: ws.Range("B16").Formula), "B16_format": com(lambda: ws.Range("B16").NumberFormat),
                               "today": datetime.date.today().isoformat()}
        # 7 Ctrl+Enter — 고른 칸 모두에 같은 값, Ctrl+Z 한 번
        go("H2:H4"); K.char("OK"); K.send(win32con.VK_RETURN, ctrl=True)
        R["ctrl_enter_H2_H4"] = [com(lambda c=c: ws.Range(c).Value) for c in ("H2", "H3", "H4")]
        R["ctrl_enter_selection_after"] = sel()
        K.send(ord("Z"), ctrl=True)
        R["ctrl_z_after_ctrl_enter"] = [com(lambda c=c: ws.Range(c).Value) for c in ("H2", "H3", "H4")]
        # 8 Ctrl+Space · Shift+Space
        go("E5"); K.send(win32con.VK_SPACE, ctrl=True); R["ctrl_space_from_E5"] = sel()
        go("E5"); K.send(win32con.VK_SPACE, shift=True); R["shift_space_from_E5"] = sel()
        # 9 행 지우기 · 넣기 — 행을 고르고 Ctrl+- / Ctrl+Shift+=, 칸 하나면 대화상자
        com(lambda: ws.Rows(8).Select()); time.sleep(0.3)
        K.send(0xBD, ctrl=True, wait=1.0)
        R["ctrl_minus_row8"] = {"A8_now": com(lambda: ws.Range("A8").Text), "D13_now": com(lambda: ws.Range("D13").Value), "E13_now": com(lambda: ws.Range("E13").Formula)}
        K.send(ord("Z"), ctrl=True, wait=1.0)
        R["ctrl_z_after_delete"] = {"A8": com(lambda: ws.Range("A8").Text), "A9": com(lambda: ws.Range("A9").Text)}
        go("C5"); K.send(0xBB, ctrl=True, shift=True, wait=1.2)
        d = K.dialog()
        R["ctrl_shift_equal_single_cell"] = {"dialog": bool(d), "title": win32gui.GetWindowText(d) if d else None}
        if d:
            win32gui.PostMessage(d, win32con.WM_CLOSE, 0, 0); time.sleep(0.8)
        go("C5"); K.send(0xBD, ctrl=True, wait=1.2)
        d = K.dialog()
        R["ctrl_minus_single_cell"] = {"dialog": bool(d), "title": win32gui.GetWindowText(d) if d else None}
        if d:
            win32gui.PostMessage(d, win32con.WM_CLOSE, 0, 0); time.sleep(0.8)
        # 10 Ctrl+1
        go("E2"); K.send(ord("1"), ctrl=True, wait=1.5)
        d = K.dialog()
        R["ctrl_1"] = {"dialog": bool(d), "title": win32gui.GetWindowText(d) if d else None}
        if d:
            win32gui.PostMessage(d, win32con.WM_CLOSE, 0, 0); time.sleep(0.8)
        # 11 Ctrl+` 수식 보기
        go("E14"); K.send(0xC0, ctrl=True, wait=1.0)
        R["ctrl_backtick"] = {"display_formulas": com(lambda: xl.ActiveWindow.DisplayFormulas), "E2_shown": com(lambda: ws.Range("E2").Text)}
        if not EN:
            com(lambda: setattr(xl.ActiveWindow, "ScrollRow", 1)); com(lambda: setattr(xl.ActiveWindow, "ScrollColumn", 1)); time.sleep(0.6)
            grab(xl.Hwnd, os.path.join(OUT, "s3_show_formulas.png"))
        K.send(0xC0, ctrl=True, wait=1.0)
        R["ctrl_backtick_again"] = com(lambda: xl.ActiveWindow.DisplayFormulas)
        # 12 Ctrl+Page Down — 다음 시트
        com(lambda: wb.Worksheets.Add(None, ws)); com(lambda: ws.Activate()); time.sleep(0.5)
        K.send(win32con.VK_NEXT, ctrl=True); R["ctrl_pagedown"] = {"from": ws.Name, "to": com(lambda: xl.ActiveSheet.Name)}
        K.send(win32con.VK_PRIOR, ctrl=True); R["ctrl_pageup"] = com(lambda: xl.ActiveSheet.Name)
        # 13 연습 문제 흐름 — 새로 연 실습 파일에서 빈 줄 지우기(Shift+Space → Ctrl+-) → 합계 칸에서 Alt+=
        com(lambda: wb.Close(False))
        wb = com(lambda: xl.Workbooks.Open(work)); ws = wb.Worksheets(1); time.sleep(1.2)
        K = Keys(xl)
        go("A8"); K.send(win32con.VK_SPACE, shift=True); K.send(0xBD, ctrl=True, wait=1.0)
        go("E13"); K.send(0xBB, alt=True, wait=1.0); K.send(win32con.VK_RETURN)
        go("E2"); K.send(win32con.VK_DOWN, ctrl=True)
        R["flow_delete_blank_then_alt_equal"] = {"D13": com(lambda: ws.Range("D13").Value), "E13_formula": com(lambda: ws.Range("E13").Formula),
                                                 "E13_value": com(lambda: ws.Range("E13").Value), "ctrl_down_from_E2": act()}
        R["foreground_unchanged_by_us"] = win32gui.GetForegroundWindow() != xl.Hwnd
        R["fg_before_was_excel"] = fg0 == xl.Hwnd
        com(lambda: wb.Close(False))
    finally:
        xl.Quit()
    name = "measure_en.json" if EN else "measure.json"
    with open(os.path.join(HERE, name), "w", encoding="utf-8") as f:
        json.dump(R, f, ensure_ascii=False, indent=2, default=str)
    print(json.dumps(R, ensure_ascii=False, indent=1, default=str))


main()
