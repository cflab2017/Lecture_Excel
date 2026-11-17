# Lecture_Excel — 엑셀 강의 (기초 · 중급 · 고급)

사무에서 바로 쓰는 **엑셀** 를 기초 → 중급 → 고급 순서로 배우는 한국어 강의입니다. 2026년 11월부터 하루 한 편씩 연재합니다.

- 강의 페이지: https://www.coding-now.com/blog/excel · English: https://www.coding-now.com/en/blog/excel
- 확인 환경: 엑셀(Microsoft 365) — 강의에 적은 결과는 모두 이 환경에서 직접 해 보고 옮긴 것입니다.

## 폴더

| 폴더 | 내용 |
|---|---|
| `01_기초/` | 처음 쓰는 분 — 화면, 입력, 저장, 기본 서식 |
| `02_중급/` | 일에 바로 쓰는 기능 |
| `03_고급/` | 자동화·큰 자료 다루기 |
| `업무자동화/` | 반복 업무를 엑셀 · AI 로 줄이기 — 예제 파일과 요청 문장 |

강마다 폴더 하나(`01_기초/01_강의이름/`)를 두고 다음 파일을 넣습니다.

- `실습.xlsx` — 강의를 따라 할 연습 파일(내용은 모두 연습용 가상 데이터)
- `정답.xlsx` — 따라 한 뒤 결과
- `README.md` — 강의 요약과 강의 페이지 주소
- 결과를 다시 확인할 때 쓴 스크립트가 있으면 함께

## 커리큘럼

강이 올라갈 때마다 아래 표에 더합니다. 강 제목은 검색어를 앞에 둡니다(예: 'VLOOKUP 함수 — 엑셀 중급 2강').

| 단계 | # | 강의 | 핵심 |
|---|---|---|---|
| 기초 | 1 | [엑셀 줄바꿈 — 셀 안에서 줄 바꾸기](01_기초/01_엑셀_줄바꿈/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-01-line-break) | Alt+Enter · 자동 줄 바꿈 · Ctrl+J 로 없애기 |
| 기초 | 2 | [엑셀 표 만들기 — 테두리 표와 표 서식(Ctrl+T)](01_기초/02_엑셀_표만들기/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-02-table) | 새 줄·계산 열·요약 행(SUBTOTAL 109)·한 열 정렬 함정 |
| 기초 | 3 | [엑셀 틀 고정 — 첫 행 · 첫 열 · 원하는 칸 기준](01_기초/03_엑셀_틀고정/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-03-freeze-panes) | 고른 칸 위·왼쪽 고정 · 내린 채 고정 함정 · 페이지 레이아웃에서 풀림 · 인쇄 제목 |
| 기초 | 4 | [엑셀 자동 채우기 · 행 추가 — 채우기 핸들 · 평일 · 합계가 안 늘 때](01_기초/04_엑셀_자동채우기_행추가/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-04-autofill-insert-row) | 1 한 칸은 복사 · 날짜는 주말까지 · 계열… 평일 · 요일은 TEXT · 합계 바로 위 행은 SUM 밖 |
| 기초 | 5 | [엑셀 셀 병합 — 병합하고 가운데 맞춤, 정렬 오류 · 필터에 한 줄만 나올 때](01_기초/05_엑셀_셀병합/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-05-merge-cells) | 왼쪽 위 값만 남음 · 정렬 오류 · 필터 한 줄 · 선택 영역의 가운데로 |
| 기초 | 6 | [엑셀 합계 · 평균 — 자동 합계가 일부만 잡을 때, 빈칸과 0, 글자 숫자](01_기초/06_엑셀_합계평균/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-06-sum-average) | SUM 72 vs + 90 · 자동 합계 =SUM(B7:B9) · 평균 14.4→15 · SUBTOTAL 109 |
| 기초 | 7 | [엑셀 시트 복사 · 시트 보호 — 다른 파일로 복사하면 생기는 외부 연결, 일부 칸만 입력되게](01_기초/07_엑셀_시트복사보호/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-07-copy-protect-sheet) | 복사본 '11월 (2)' · 새 통합 문서로 [실습.xlsx]설정!$B$1 · 범위 복사 열 너비 8.38 · 잠금 해제 칸만 입력 · 시트 보호해도 삭제됨 → 통합 문서 보호 |
| 기초 | 8 | [엑셀 단축키 — 빈 칸에서 멈추는 Ctrl+방향키 · 자동 합계 Alt+= · F4 의 두 가지 뜻](01_기초/08_엑셀_단축키/) · [강의](https://www.coding-now.com/blog/excel/excel-basic-08-shortcuts) | Ctrl+↓ 빈 줄에서 멈춤 · Alt+= =SUM(E9:E13) · F4 $ 순서 · Ctrl+; 날짜 |
| 중급 | 1 | [엑셀 IF 함수 — AND · OR · 중첩 IF (컴활 2급 실기 계산작업 유형 연습)](02_중급/01_엑셀_IF함수/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-01-if-function) | 자작 3문제 · 따옴표 빠지면 #NAME? · OR 빼먹으면 8칸 중 1칸만 틀림 · 중첩 IF 는 큰 기준부터 · 이상은 >= |
| 중급 | 2 | [엑셀 COUNTIF · SUMIF · AVERAGEIF (컴활 2급 실기 계산작업 유형 연습)](02_중급/02_엑셀_COUNTIF_SUMIF/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-02-countif-sumif) | 자작 3문제 · & "명" · $ 없이 채우면 셋째 칸만 틀림 · ">="&E25 vs ">=E25" 0 · "*볼펜*" 165 vs "볼펜*" 82 |
| 중급 | 3 | [엑셀 VLOOKUP · HLOOKUP (컴활 2급 실기 계산작업 유형 연습)](02_중급/03_엑셀_VLOOKUP_HLOOKUP/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-03-vlookup-hlookup) | 자작 3문제 · 열 번호는 표 범위 안에서 · $ 없이 채우면 #N/A · 구간은 TRUE(FALSE 면 #N/A 5칸) · 정렬 안 된 표에 넷째 인수 생략하면 5칸 틀림 |
| 중급 | 4 | [엑셀 반올림 ROUND · ROUNDUP · ROUNDDOWN · INT · MOD (컴활 2급 실기 계산작업 유형 연습)](02_중급/04_엑셀_반올림_ROUND_INT_MOD/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-04-round-int-mod) | 자작 3문제 · 자릿수 1 · 0 · -2 · 상자 수는 ROUNDUP(ROUND 면 4칸 틀림) · 표시 형식만 줄이면 ×3 이 259 · INT(-7.5) = -8 · MOD 인수 거꾸로 8칸 틀림 |

중급의 '컴활 2급 실기 유형 연습' 문제는 모두 **대한상공회의소 문제가 아닌 연습용 자작 문제**입니다.

## 업무 자동화

| # | 예제 | 핵심 |
|---|---|---|
| 1 | [엑셀 파일 24개 한 장으로 합치기](업무자동화/01_엑셀_파일_합치기/) · [가이드](https://www.coding-now.com/guides/combine-excel-files) | 모양이 다른 지점별 월 매출 24개 · 파워 쿼리 '폴더에서' 고친 쿼리 43,665,600 일치 · 클로드 요청 문장 · '합계'만 빼면 부산 두 배 |

---
만든 곳: [코딩나우 CODINGNOW](https://www.coding-now.com)
