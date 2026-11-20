"""shots/ 원본에서 강의 페이지용 그림을 만든다 — 제목 표시줄(계정 사진 · 파일 이름)을 잘라 내고 빨간 테두리로 볼 곳 표시.
결과는 coding-now public/lectures/excel/basic-08-*.png. 영어판은 화면 그림을 쓰지 않는다(이 PC 엑셀 메뉴가 한국어)."""
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"F:\_GIT_Projects\05_content\content_hub\homepage\public\lectures\excel"
TOP = 44  # 제목 표시줄 높이


def win(src, dst, boxes):
    im = Image.open(os.path.join(HERE, "shots", src)).convert("RGB")
    d = ImageDraw.Draw(im)
    for b in boxes:
        d.rectangle(b, outline=(220, 30, 30), width=3)
    im = im.crop((0, TOP, im.width, im.height))
    im.save(os.path.join(DST, dst))
    print(dst, im.size)


# s1 E2 에서 Ctrl+Shift+↓ — E2:E7 까지만(8행 빈 줄) · 상태 표시줄 개수 6 · 합계 68,100
win("s1_ctrl_shift_down.png", "basic-08-ctrl-shift-down.png", [(410, 264, 510, 400), (40, 398, 510, 420), (466, 738, 690, 762)])
# s2 E14 에서 Alt+= — 수식 입력줄 =SUM(E9:E13) · 48,100 · 빈 8행
win("s2_alt_equal_trap.png", "basic-08-alt-equal-trap.png", [(228, 189, 520, 215), (410, 529, 510, 554), (40, 398, 510, 420)])
# s3 Ctrl+` 수식 보기 — E 열 수식 · 16행 Ctrl+; 날짜 · Ctrl+Shift+; 시간이 숫자로
win("s3_show_formulas.png", "basic-08-show-formulas.png", [(784, 264, 974, 555), (40, 574, 512, 598)])
