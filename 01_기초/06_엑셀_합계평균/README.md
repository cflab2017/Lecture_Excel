# 기초 6강 — 엑셀 합계 · 평균 (자동 합계가 일부만 잡을 때, 빈칸과 0, 글자로 된 숫자, 필터 뒤 합계)

- 강의: https://www.coding-now.com/blog/excel/excel-basic-06-sum-average
- English: https://www.coding-now.com/en/blog/excel/excel-basic-06-sum-average

## 파일

| 파일 | 내용 |
|---|---|
| `실습.xlsx` | 따라 할 파일 — 가상 문구점 일주일 판매 수량(월~일). 수요일은 쉬는 날(빈칸), 일요일은 0개, 목요일 18 은 글자로 들어가 있음 |
| `정답.xlsx` | 목요일을 숫자로 고친 뒤 합계 · 평균 · 판 날만 평균 · 보이는 행 합계(SUBTOTAL 109) · 숫자 칸 수 · 빈칸 아닌 칸 수, 필터 켬 |
| `practice.xlsx` · `answer.xlsx` | 영어판 강의용 같은 연습 파일 |
| `measure.py` → `measure.json` | 강의에 적은 값을 엑셀로 다시 재는 스크립트 — SUM · + · AVERAGE · COUNT · COUNTA · AVERAGEIF, 자동 합계가 고르는 범위, 숨긴 행 · 필터와 SUBTOTAL |
| `measure.py en` → `measure_en.json` | 영어판 자료로 같은 측정 |
| `make_files.py` | 실습 · 정답 파일 만들기(`en` 이면 영어판) |
| `screens.py` → `screens.json` · `crop_web.py` | 강의 화면(`shots/`, 저장소에는 올리지 않음) — 상태 표시줄 · 자동 합계 제안 · 정답 · 필터 |

스크립트는 Windows + 엑셀 + `pip install pywin32 pillow` 에서 돈다. 엑셀을 따로 띄워 쓰고, 시험 통합 문서는 저장하지 않고 닫는다.
자동 합계는 칸 하나에서 누르면 엑셀이 '식 제안' 상태로 멈추므로 엑셀 창에 Enter 를 보내 확정한다(사람이 Enter 누르는 것과 같음).

## 핵심

- 글자로 된 18 이 섞인 B2:B8: =SUM 72 · 더하기(=B2+…+B8) 90 · =AVERAGE 14.4(숫자 5칸) · =COUNT 5 · =COUNTA 6 · =AVERAGEIF(…,">0") 18.
- 상태 표시줄: 평균 14.4 · 개수 6 · 합계 72(실습 파일 B3:B9 선택).
- 표 바로 아래 칸 하나에서 자동 합계 → =SUM(B7:B9) 45 — 글자 칸(목)에서 멈춤. 범위를 같이 고르고 누르면 =SUM(B3:B9).
- 목요일을 숫자로 고친 뒤: 합계 90 · 평균 15(빈칸은 빼고 0 은 셈) · 판 날만 평균 18.
- 행 숨기기(토 25): SUM 90 그대로 · SUBTOTAL(9) 90 · SUBTOTAL(109) 65 · AGGREGATE(9,5) 65.
- 판매 수량 >= 15 필터: SUM 90 · SUBTOTAL(9)/(109) 78 · 표 아래에서 자동 합계 → =SUBTOTAL(9, 목~일) 63(빈칸 수요일 위는 빠짐).
