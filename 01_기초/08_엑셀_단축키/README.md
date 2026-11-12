# 기초 8강 — 엑셀 단축키 (빈 칸에서 멈추는 Ctrl+방향키 · 자동 합계 Alt+= · F4 의 두 가지 뜻)

- 강의: https://www.coding-now.com/blog/excel/excel-basic-08-shortcuts
- English: https://www.coding-now.com/en/blog/excel/excel-basic-08-shortcuts

## 파일

| 파일 | 내용 |
|---|---|
| `실습.xlsx` | 따라 할 파일 — 가상 카페 11월 주문표. 2~7행 · **8행 빈 줄** · 9~13행, D14 '합계'(E14 비움), F1 '비율'(비움). 금액 E = 수량 × 단가 |
| `정답.xlsx` | 빈 줄을 지운 표(2~12행), 합계 E13 =SUM(E2:E12) 116,200, 비율 F2:F12 =E2/$E$13 (0.0%) |
| `practice.xlsx` · `answer.xlsx` | 영어판 강의용 같은 연습 파일(Orders 시트) |
| `measure.py` → `measure.json` | 강의에 적은 값을 엑셀로 다시 재는 스크립트 — 단축키를 엑셀 창(EXCEL7)에 직접 보내고 셀 주소 · 수식 · 값을 읽음 |
| `measure.py en` → `measure_en.json` | 영어판 파일로 같은 측정 |
| `make_files.py` | 실습 · 정답 파일 만들기(`en` 이면 영어판) |
| `crop_web.py` | 강의 화면(`shots/`, 저장소에는 올리지 않음) — Ctrl+Shift+↓ 범위 · Alt+= 함정 · Ctrl+` 수식 보기 |

스크립트는 Windows + 엑셀 + `pip install pywin32 pillow` 에서 돈다. 실습 파일 사본을 %TEMP% 에서 열고 저장하지 않고 닫는다.
키는 `AttachThreadInput` + `SetKeyboardState` 로 Ctrl · Shift · Alt 를 눌린 상태로 만든 뒤 `PostMessage` 로 엑셀 창에 보낸다 — 창을 앞으로 가져오지 않고 마우스를 쓰지 않는다.
입력기(IME)를 거치지 않으므로 한글 입력 상태에서의 차이는 재지 않았다.

## 핵심(이 PC 엑셀, Microsoft 365)

- E2 에서 Ctrl+↓: E7 → E9 → E13 → E1048576 (빈 8행 위에서 멈춤). Ctrl+Shift+↓: E2:E7. E9 에서 Ctrl+↑: E7.
- Ctrl+Home A1, Ctrl+End F14. H30 에 썼다 지우면 Ctrl+End H30, 저장한 사본을 다시 열면 F14.
- E14 에서 Alt+=: =SUM(E9:E13) 48,100. E2:E13 을 고르고 Alt+=: E14 =SUM(E2:E13) 116,200. 빈 8행을 지우고(Shift+Space → Ctrl+-) E13 에서 Alt+=: =SUM(E2:E12) 116,200.
- F2 =E2/E14 에서 F2 → F4 를 거듭: =E2/$E$14 → =E2/E$14 → =E2/$E14 → =E2/E14. 편집 중이 아니면 F4 = 반복(B2 Ctrl+B 뒤 B9 F4 → 굵게).
- Ctrl+; → 2026-11-12(yyyy-mm-dd, 값 46338), Ctrl+Shift+; → h:mm AM/PM 시각(값). Ctrl+Enter → H2:H4 모두 OK, Ctrl+Z 한 번에 셋 다 비움.
- Ctrl+Space → E:E, Shift+Space → 5:5. 행을 고르고 Ctrl+- → 바로 지움, 칸 하나만 고르면 Ctrl+Shift+= '삽입' · Ctrl+- '삭제' 창. Ctrl+1 '셀 서식'. Ctrl+` 수식 보기 켜기 · 끄기. Ctrl+Page Down 다음 시트.
