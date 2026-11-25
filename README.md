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
| 중급 | 5 | [엑셀 날짜 함수 DATE · YEAR · MONTH · DAY · WEEKDAY (컴활 2급 실기 계산작업 유형 연습)](02_중급/05_엑셀_날짜함수_DATE_WEEKDAY/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-05-date-functions) | 자작 3문제 · 기준일 $ 없으면 7칸 음수 · WEEKDAY 유형 2(생략하면 하루씩 밀림) · DAY(빼기)는 32박 → 1 · 날짜 형식 칸이면 2015 → 1905-07-07 · DATE(2026,13,1) = 2027-01-01 |
| 중급 | 6 | [엑셀 텍스트 함수 LEFT · MID · RIGHT · LEN · & (컴활 2급 실기 계산작업 유형 연습)](02_중급/06_엑셀_텍스트함수_LEFT_MID_RIGHT/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-06-text-functions) | 자작 3문제 · MID 시작 위치는 하이픈도 셈 · 이름 가리기 REPT("*",LEN-1)('**' 고정이면 4칸 틀림) · & 대신 + 는 #VALUE! · MID 로 꺼낸 '09' 는 글자라 SUM 0(VALUE 면 84) · & 는 쉼표 모양을 버림 → TEXT |
| 중급 | 7 | [엑셀 데이터베이스 함수 DSUM · DAVERAGE · DCOUNT · DMAX · DMIN (컴활 2급 실기 계산작업 유형 연습)](02_중급/07_엑셀_데이터베이스함수_DSUM_DAVERAGE/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-07-database-functions) | 자작 3문제 · 표 · 조건 범위 모두 머리글부터(빼면 0 · #VALUE!) · 같은 줄 그리고(48.5) · 다른 줄 또는(48.1) · 조건 범위 빈 줄은 전부 맞음(60) · '문구'는 문구류까지(10) → ="=문구"(15) |
| 중급 | 8 | [엑셀 사용자 지정 표시 형식 (컴활 2급 실기 기본작업 유형 연습)](02_중급/08_엑셀_사용자지정_표시형식/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-08-custom-number-format) | 자작 3문제 · 0"세" · #,##0"원" 은 값 그대로(평균 36.6) · '34세' 직접 입력은 글자(평균 #DIV/0!) · 요일 aaa(ddd 는 Fri) · 끝 쉼표 = 천 단위 · 양수;음수;0 — 0 구역 빠지면 0, 순서 바뀌면 양수가 빨강 · 글자는 @ · 한국어 엑셀은 [빨강]([Red] · [빨간색] 거부) |
| 중급 | 9 | [엑셀 조건부 서식 — 수식으로 행 전체 칠하기 (컴활 2급 실기 기본작업 유형 연습)](02_중급/09_엑셀_조건부서식_수식/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-09-conditional-formatting-formula) | 자작 3문제 · =$D3>=3 네 행 통째로 · $ 빼면 A열만 · $D$3 이면 안 칠해짐 · 머리글부터 고르면 한 행 밀림 · AND(따옴표 빠지면 안 칠해짐 · OR 이면 7행) · 평균 범위는 $C$25:$C$32 |
| 중급 | 10 | [엑셀 고급 필터 — 조건 범위 AND · OR (컴활 2급 실기 기본작업 유형 연습)](02_중급/10_엑셀_고급필터_AND_OR/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-10-advanced-filter) | 자작 3문제 · 같은 줄 그리고 3명 · 다른 줄 또는 9명 · 빈 줄까지 잡으면 10명 전부 · 머리글 끝 공백 0명 · 필드 두 개만 꺼내기(이름 다르면 오류) · '노트'는 노트북까지 4개 · ="=노트" 1개 · =노트 는 #NAME? |
| 중급 | 11 | [엑셀 정렬 · 부분합 (컴활 2급 실기 분석작업 유형 연습)](02_중급/11_엑셀_정렬_부분합/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-11-sort-subtotal) | 자작 3문제 · 지점별 합계(대구 · 부산 · 서울 요약 · 총합계) · 최대값 다음 평균(대치 해제 — 화면엔 평균 줄이 위) · 사용자 지정 목록 부장 · 과장 · 대리 · 사원 개수 · 정렬 없이 하면 요약 10줄 · 대치 그대로면 최대 줄 사라짐 |
| 중급 | 12 | [엑셀 피벗 테이블 (컴활 2급 실기 분석작업 유형 연습)](02_중급/12_엑셀_피벗테이블/) · [강의](https://www.coding-now.com/blog/excel/excel-intermediate-12-pivot-table) | 자작 3문제 · 지점 × 제품 합계(총합계 1,548,000) · 직위 × 부서 평균 · 개요 형식이면 머리글이 필드 이름 · 빈 셀 * · 날짜를 월로 그룹(9월 50 · 10월 39 · 11월 86) · 수량에 글자 '미정'이 있으면 엑셀이 개수를 고름 · 원본을 고치면 새로 고침, 줄을 더하면 데이터 원본 변경 · 글자 날짜는 그룹 거부 |

중급의 '컴활 2급 실기 유형 연습' 문제는 모두 **대한상공회의소 문제가 아닌 연습용 자작 문제**입니다.

## 업무 자동화

| # | 예제 | 핵심 |
|---|---|---|
| 1 | [엑셀 파일 24개 한 장으로 합치기](업무자동화/01_엑셀_파일_합치기/) · [가이드](https://www.coding-now.com/guides/combine-excel-files) | 모양이 다른 지점별 월 매출 24개 · 파워 쿼리 '폴더에서' 고친 쿼리 43,665,600 일치 · 클로드 요청 문장 · '합계'만 빼면 부산 두 배 |
| 2 | [파일 이름 30개 한꺼번에 정리](업무자동화/02_파일_이름_정리/) · [가이드](https://www.coding-now.com/guides/rename-files) | 이름이 제각각인 회의록 · 보고서 30개를 날짜_제목으로 · 엑셀 목록 「CSV UTF-8」 + PowerShell 한 줄 29개 · 그냥 CSV 는 한글이 깨져 29개 실패 · 클로드 요청 문장 · 수정 시각으로 날짜를 정하면 1/29 |

---
만든 곳: [코딩나우 CODINGNOW](https://www.coding-now.com)
