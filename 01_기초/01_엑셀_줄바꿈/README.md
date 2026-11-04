# 기초 1강 — 엑셀 줄바꿈 (셀 안에서 줄 바꾸기 · 자동 줄 바꿈 · 줄바꿈 없애기)

- 강의: https://www.coding-now.com/blog/excel/excel-basic-01-line-break
- English: https://www.coding-now.com/en/blog/excel/excel-basic-01-line-break

## 파일

| 파일 | 내용 |
|---|---|
| `실습.xlsx` | 따라 할 파일 — 준비물 칸(D열)이 쉼표로 한 줄에 이어져 있다 |
| `정답.xlsx` | 준비물을 한 줄에 하나씩 나누고 줄 높이를 맞춘 결과 |
| `practice.xlsx` · `answer.xlsx` | 영어판 강의용 같은 연습 파일 |
| `measure.py` → `measure.json` | 강의에 적은 값을 엑셀로 다시 재는 스크립트(LEN·문자 코드·자동 줄 바꿈·줄 높이·바꾸기·CSV) |
| `measure_en.py` → `measure_en.json` | 영어판 자료로 같은 측정 |
| `make_en_files.py` | 영어판 연습 파일을 만드는 스크립트 |
| `screens.py` | 강의 화면을 찍은 스크립트(`shots/` 에 저장, 저장소에는 올리지 않음) |

스크립트는 Windows + 엑셀 + `pip install pywin32 pillow` 에서 돈다. 엑셀을 따로 띄워 쓰고, 파일은 저장하지 않고 닫는다.

## 핵심

- 셀 안 줄바꿈 = **Alt+Enter** (편집 상태에서). 들어가는 글자는 코드 10(CHAR(10)).
- 손으로 넣으면 자동 줄 바꿈이 저절로 켜지지만, 수식(`&CHAR(10)&`)으로 만든 줄바꿈은 자동 줄 바꿈을 직접 켜야 보인다.
- 줄 높이가 고정돼 있으면 잘린다 → 행 높이 자동 맞춤.
- 한꺼번에 없애기: 바꾸기(Ctrl+H)에서 찾을 내용에 **Ctrl+J**, 또는 `=SUBSTITUTE(D2,CHAR(10),", ")`.
