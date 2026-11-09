"""shots/ 원본에서 강의 페이지용 그림을 만든다 — 창 그림은 제목 표시줄(계정 사진)을 잘라 내고 빨간 테두리로 볼 곳 표시(11/6 규칙: 위치가 보이게).
결과는 coding-now public/lectures/excel/basic-05-*.png. 영어판은 화면 그림을 쓰지 않는다(이 PC 엑셀 메뉴가 한국어)."""
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\excel"
TOP = 44  # 제목 표시줄 높이


def win(src, dst, boxes):
    im = Image.open(os.path.join(HERE, "shots", src)).convert("RGB")
    d = ImageDraw.Draw(im)
    for b in boxes:
        d.rectangle(b, outline=(220, 30, 30), width=3)
    im = im.crop((0, TOP, im.width, im.height))
    im.save(os.path.join(DST, dst))
    print(dst, im.size)


def plain(src, dst):
    im = Image.open(os.path.join(HERE, "shots", src)).convert("RGB")
    im.save(os.path.join(DST, dst))
    print(dst, im.size)


win("s1_practice.png", "basic-05-practice.png", [(40, 243, 347, 271), (40, 291, 144, 360)])
plain("s2_merge_warning.png", "basic-05-merge-warning.png")
plain("s3_sort_error.png", "basic-05-sort-error.png")
win("s4_filter.png", "basic-05-filter.png", [(40, 291, 347, 316), (52, 594, 240, 618)])
win("s5_answer.png", "basic-05-answer.png", [(40, 291, 144, 428)])
plain("s6_center_across.png", "basic-05-center-across.png")
# 이동 옵션 — '빈 셀(K)' 에 빨간 테두리(goto_special.py)
im = Image.open(os.path.join(HERE, "shots", "s7_goto_special.png")).convert("RGB")
ImageDraw.Draw(im).rectangle((18, 224, 120, 248), outline=(220, 30, 30), width=3)
im.save(os.path.join(DST, "basic-05-goto-special.png"))
print("basic-05-goto-special.png", im.size)
