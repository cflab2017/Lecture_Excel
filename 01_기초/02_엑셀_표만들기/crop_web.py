"""shots/ 원본에서 강의 페이지용 그림을 잘라 coding-now public/lectures/excel/ 에 넣는다(제목 표시줄·계정 사진은 잘라 냄)."""
import os
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\excel"
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
out = {}
out["basic-02-border.png"] = S("s1_border.png").crop((8, 188, 1000, 525))      # 수식 입력줄 + 1~10행
out["basic-02-table.png"] = S("s2_table.png").crop((8, 48, 1000, 540))         # 테이블 디자인 리본 + 표 + 요약 행
f = S("s3_filter.png"); grid = f.crop((8, 188, 1000, 385)); bar = f.crop((8, 660, 430, 688))
im = Image.new("RGB", (992, grid.height + 8 + bar.height), "white"); im.paste(grid, (0, 0)); im.paste(bar, (0, grid.height + 8))
out["basic-02-filter.png"] = im                                               # 거른 표 + 상태 표시줄 '9개 중 3개'
for n, im in out.items():
    im.save(os.path.join(DST, n), optimize=True); print(n, im.size)
