# -*- coding: utf-8 -*-
"""업무 자동화 3편 — 워드 보고서로 만든 발표 PPT, 한 문장으로 고치기 전 · 뒤 실측 스크립트 (코딩나우 · 2026-11-26)

    py measure.py            → measure.json (같은 폴더)

필요: python-pptx (파워포인트 · 워드를 띄우지 않고 파일만 읽음)
재는 것
  1) 파일 속성 — 작성자 · 마지막 수정한 사람
  2) PPT 마다 장 수 · 장마다 글자 수(빈칸 · 그래프 · 자리 표시자[제목 · 쪽 머리 · 쪽 번호] 빼고, 표 안 글은 넣음)
     · 표 · 그래프 · 발표자 노트가 있는 장
  3) PPT 에 나온 숫자(글 · 표 · 그래프 값)가 보고서(예제 워드) 글에 있는가 — 쉼표는 빼고 비교
     · 쪽 머리의 「01」~「08」 같은 차례 번호는 「보고서에 없는 숫자」에서 따로 표시
"""
import json
import re
import statistics
import sys
import zipfile
from pathlib import Path

from pptx import Presentation

HERE = Path(__file__).resolve().parent
DOCX = HERE / "예제_2026년_3분기_영업보고서.docx"
PPTS = {"처음_한줄": HERE / "결과1_한줄로_처음_받은_PPT.pptx", "한문장으로_고친": HERE / "결과2_한문장으로_고친_PPT.pptx"}


def core_props(path):
    with zipfile.ZipFile(path) as z:
        core = z.read("docProps/core.xml").decode("utf-8")
    return dict(re.findall(r"<(?:dc|cp):(creator|lastModifiedBy)>([^<]*)<", core))


def report_text():
    with zipfile.ZipFile(DOCX) as z:
        x = z.read("word/document.xml").decode("utf-8")
    paras = ["".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)) for p in re.findall(r"<w:p[ >].*?</w:p>", x, re.S)]
    return "\n".join(p for p in paras if p.strip())


def nums(s):
    out = set()
    for m in re.finditer(r"\d[\d,]*(?:\.\d+)?", s):
        v = m.group().replace(",", "")
        out.add(v.rstrip("0").rstrip(".") if "." in v else v)
    return out


def measure_ppt(path, rep_nums):
    prs = Presentation(path)
    slides = []
    for i, s in enumerate(prs.slides, 1):
        chars, tables, charts, all_text, chart_vals = 0, 0, 0, [], set()
        for sh in s.shapes:
            if sh.has_chart:
                charts += 1
                for plot in sh.chart.plots:
                    for ser in plot.series:
                        for v in ser.values:
                            if v is not None:
                                chart_vals |= nums(f"{v:g}")
                continue
            if sh.has_table:
                tables += 1
                t = "\n".join(c.text for r in sh.table.rows for c in r.cells)
                chars += len(re.sub(r"\s", "", t))
                all_text.append(t)
                continue
            if sh.has_text_frame:
                t = sh.text_frame.text
                all_text.append(t)
                if not sh.is_placeholder:
                    chars += len(re.sub(r"\s", "", t))
        notes = s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip() != ""
        found = nums("\n".join(all_text)) | chart_vals
        missing = sorted(found - rep_nums, key=lambda v: (len(v), v))
        order = [v for v in missing if re.fullmatch(r"0\d", v) or v == str(i)]  # 쪽 머리 차례 번호 · 쪽 번호
        slides.append({"장": i, "글자": chars, "표": tables, "그래프": charts, "노트": notes,
                       "보고서에_없는_숫자": [v for v in missing if v not in order], "차례_쪽_번호": order})
    per = [x["글자"] for x in slides]
    return {"작성자": core_props(path), "장_수": len(slides), "장마다_글자": per,
            "장마다_글자_가운데값": statistics.median(per), "표": sum(x["표"] for x in slides),
            "그래프": sum(x["그래프"] for x in slides), "노트_있는_장": sum(x["노트"] for x in slides),
            "보고서에_없는_숫자": sorted({v for x in slides for v in x["보고서에_없는_숫자"]}), "장별": slides}


def main():
    rep = report_text()
    rep_nums = nums(rep)
    out = {"보고서": {"작성자": core_props(DOCX), "글자": len(re.sub(r"\s", "", rep)), "숫자_가짓수": len(rep_nums)}}
    for key, path in PPTS.items():
        out[key] = measure_ppt(path, rep_nums)
    (HERE / "measure.json").write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    for key in PPTS:
        m = out[key]
        print(f"{key}: {m['장_수']}장 · 글자 가운데 {m['장마다_글자_가운데값']} · 표 {m['표']} · 그래프 {m['그래프']} · 노트 {m['노트_있는_장']} · 보고서에 없는 숫자 {m['보고서에_없는_숫자'] or '없음'}")


main()
