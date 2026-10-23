# 거성호텔 나이트 근무 마감 자동화 툴 v1.2

오아시스 ERP 엑셀 → 호텔 마감 양식 자동 변환 (수입일보 / 카드 승인내역 / 대외후불 / 조식명단)

> 개발은 Claude Code로 진행합니다. 작업 지침은 `CLAUDE.md`를 참고하세요.

## 사용법 (근무자) — 자세한 내용은 `docs/사용법.txt`
1. `GeoseongClosingTool.exe` 더블클릭 (`config.json`, `templates` 폴더는 같은 폴더에 둘 것)
2. 영업일자 확인 (정오 이전에는 전날 날짜가 자동 선택, ◀ ▶로 변경) → 서류 카드의 [파일 선택] → ERP 엑셀 선택
3. 원본 파일이 있는 폴더의 `마감결과` 폴더에 저장됨 (카드의 "결과 열기" / "결과 폴더 열기")
4. **노란 칸은 직접 확인·입력**: 수입일보의 선/후·숙박수, 카드 승인내역의 객실·현금영수증

## 서류별 변환 규칙
- **수입일보**: 객실번호 순 정렬, 객실은 4자리(0302). 매출은 ERP의 NET(공급가) 열 기준. 누계는 매출과 같은 값(수식).
- **카드 승인내역**: 응답이 '성공'인 건만 표에 포함. 현금영수증은 표에서 빼고 '현금영수증' 줄에 합산, 계 = 카드 합계 + 현금영수증.
  ERP에는 객실번호가 없고 Acct#뿐이라 객실 칸은 비워 둠 (Acct#는 오른쪽에 참고용으로 표시, 인쇄 제외).
- **대외후불**: 입금액을 금액으로, 객실번호 순 정렬. 원본 합계와 다르면 경고.
- **조식명단**: IN HOUSE LIST의 비고에 '조식/조포'가 있는 객실만. '룸온리'는 제외. 인원은 비고의 'N인' 우선, 없으면 성인+소인.

## exe 만들기
- **방법 A (Windows PC)**: `build_exe.bat` 더블클릭 → `dist` 폴더의 exe + config.json을 호텔 PC로 복사
- **방법 B (GitHub, 권장)**: push하면 자동으로 테스트 → exe 빌드. Actions → "Build Windows EXE" → 맨 위 실행 → 아래 Artifacts의 `GeoseongClosingTool` 다운로드(압축 해제하면 exe + config.json + 사용법.txt)
- 호텔 PC가 Windows 7이면 Python 3.8로 빌드

## 양식 조정 (코드 수정 없이 config.json만 편집)
- `source_fields`: ERP 열 제목 별칭 / `output`: 결과 양식의 열 순서·이름·너비
- `sheet_title`, `file_pattern`: 제목·파일명 형식
- `business_day_rollover_hour`: 영업일자를 전날로 보는 기준 시각 (기본 12시)
- `mask_card_number`: 마스킹 안 된 카드번호 앞4/뒤4만 표시
- `template`: 호텔 빈 양식 파일 경로. 있으면 양식의 폰트·크기·열 너비·인쇄 설정을 그대로 사용 (`templates/README.md` 참고)
- `font`: 양식 파일이 없을 때 쓸 폰트 이름·크기(`title`/`header`/`body`)와 행 높이

## 호텔 양식 맞추기
`templates/`에 호텔 실제 양식에서 만든 빈 양식 4종이 들어 있습니다(수입일보·카드·대외후불: 굴림, 조식명단: 맑은 고딕). 호텔 양식이 바뀌면 새 파일로 `python tools/make_templates.py`를 실행해 다시 만듭니다.
exe 옆에 `templates` 폴더를 두고 빈 양식 4종(`수입일보.xlsx`, `카드승인내역.xlsx`, `대외후불리스트.xlsx`, `조식명단.xlsx`)을 넣으면 결과가 호텔 양식과 같은 모양으로 나옵니다.
