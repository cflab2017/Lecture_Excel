"""shots/ 원본에서 강의 페이지용 그림을 잘라 coding-now public/lectures/excel/ 에 넣는다(제목 표시줄·계정 사진은 잘라 냄)."""
import os
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\excel"
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
out = {
    "basic-03-c2.png": S("s1_c2_scrolled.png").crop((8, 188, 1192, 640)),        # 수식 입력줄 + 1행·A:B 가 남은 40행 화면
    "basic-03-trap.png": S("s2_trap_top.png").crop((8, 188, 1192, 640)),         # 맨 위로 올려도 15행부터
    "basic-03-a1.png": S("s3_a1_center.png").crop((8, 188, 1192, 640)),          # A1 → 화면 가운데 십자
    "basic-03-pagelayout.png": S("s5_pagelayout.png").crop((8, 188, 1192, 685)), # 페이지 레이아웃: 1행이 안 남음 + 상태 표시줄
}
for n, im in out.items():
    im.save(os.path.join(DST, n), optimize=True); print(n, im.size)
