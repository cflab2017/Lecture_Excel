"""shots/ 원본에서 강의 페이지용 그림을 잘라 coding-now public/lectures/excel/ 에 넣는다(제목 표시줄·계정 사진은 잘라 냄).
영어판은 3강처럼 화면 그림을 쓰지 않는다(이 PC 엑셀 메뉴가 한국어)."""
import os
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\excel"
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
BOX = (8, 188, 1192, 545)   # 수식 입력줄 + 1~14행
out = {
    "basic-04-filled.png": S("s1_filled.png").crop(BOX),      # 번호 1·2 + 날짜·요일·주차 첫 줄에서 채운 결과
    "basic-04-weekdays.png": S("s1b_weekdays.png").crop((8, 188, 1192, 690)),  # 날짜만 평일 단위로 → 요일 칸 어긋남 + 상태 표시줄
    "basic-04-trap.png": S("s2_trap.png").crop(BOX),          # 합계 바로 위에 끼운 행 — =SUM(F2:F11) · 120 · 초록 삼각형
    "basic-04-inside.png": S("s3_inside.png").crop(BOX),      # 범위 안에 끼운 행 — =SUM(F2:F12) · 140
    "basic-04-series.png": S("s4_series.png"),                # 홈 → 채우기 → 계열… = 연속 데이터 창(날짜 단위 일·평일·월·년) — series_dialog.py
}
for n, im in out.items():
    im.save(os.path.join(DST, n), optimize=True)
    print(n, im.size)
