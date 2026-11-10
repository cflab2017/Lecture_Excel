"""shots/ 원본에서 강의 페이지용 그림을 만든다 — 제목 표시줄(계정 사진)을 잘라 내고 빨간 테두리로 볼 곳 표시(11/6 규칙: 위치가 보이게).
결과는 coding-now public/lectures/excel/basic-06-*.png. 영어판은 화면 그림을 쓰지 않는다(이 PC 엑셀 메뉴가 한국어)."""
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


# s1 상태 표시줄(평균 · 개수 · 합계) + 목요일 칸(초록 삼각형, 왼쪽 정렬 = 글자)
win("s1_status_bar.png", "basic-06-status-bar.png", [(500, 596, 682, 618), (189, 359, 293, 382)])
# s2 자동 합계가 제안한 범위 B7:B9 와 수식 입력줄
win("s2_autosum_suggest.png", "basic-06-autosum-suggest.png", [(189, 379, 293, 470), (236, 190, 352, 214)])
# s3 정답 합계 칸과 수식
win("s3_answer.png", "basic-06-answer.png", [(189, 469, 293, 493), (236, 190, 352, 214)])
# s4 필터 — 보이는 행 합계 78 · 상태 표시줄 '7개 중 4개'
win("s4_filter.png", "basic-06-filter.png", [(189, 469, 293, 493), (236, 190, 430, 214), (60, 596, 260, 618)])
