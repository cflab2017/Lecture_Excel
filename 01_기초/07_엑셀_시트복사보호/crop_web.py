"""shots/ 원본에서 강의 페이지용 그림을 만든다 — 제목 표시줄(계정 사진)을 잘라 내고 빨간 테두리로 볼 곳 표시(11/6 규칙: 위치가 보이게).
결과는 coding-now public/lectures/excel/basic-07-*.png. 영어판은 화면 그림을 쓰지 않는다(이 PC 엑셀 메뉴가 한국어)."""
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


# s1 이동/복사 창 — 다음 시트의 앞에 목록 · 복사본 만들기 체크 · 시트 탭 11월
win("s1_move_copy.png", "basic-07-move-copy.png", [(356, 338, 655, 534), (356, 538, 476, 560), (120, 752, 175, 778)])
# s2 같은 파일 안 복사 뒤 탭 — 11월 (2)
win("s2_tabs.png", "basic-07-tabs.png", [(180, 750, 252, 780)])
# s3 새 통합 문서로 복사 — 수식 입력줄 [실습.xlsx] · 탭은 11월 하나 · 제목 통합 문서1
win("s3_external_link.png", "basic-07-external-link.png", [(232, 190, 512, 214), (120, 752, 175, 778)])
# s4 셀 서식 보호 탭 — 잠금 · 숨김 · 안내 문구 / 고른 B3:B8
win("s4_cell_protection.png", "basic-07-cell-protection.png", [(244, 196, 772, 280), (142, 288, 229, 421)])
# s5 시트 보호 창 — 암호 칸 · 잠긴 셀 보호 · 허용 목록 위 두 줄
win("s5_protect_sheet.png", "basic-07-protect-sheet.png", [(368, 230, 642, 360)])
