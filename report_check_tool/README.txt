보고서 생성 누락 점검 도구

용도
- 입력 엑셀 파일과 생성된 PDF 보고서 파일들을 비교합니다.
- 어떤 입력 행의 보고서가 생성됐고, 어떤 행이 누락됐는지 확인합니다.
- 기존 자동화 실행 코드는 건드리지 않는 별도 도구입니다.

가장 쉬운 사용법
1. 이 report_check_tool 폴더 안에 입력 엑셀 파일 1개를 넣습니다.
2. 같은 폴더 안에 생성된 PDF 보고서 파일들을 넣습니다.
3. check_reports.bat를 실행합니다.
4. report_check_result.xlsx 파일을 확인합니다.

결과 파일 시트
- summary: 전체 요약
- all: 입력 행 전체 점검 결과
- missing: 누락된 입력 행만 모음
- duplicates: 같은 입력 행에 PDF가 여러 개 매칭된 경우
- extra_pdfs: 입력 엑셀과 매칭되지 않는 PDF 파일

지원하는 PDF 파일명 예시
- 프랑스_자동차 엔진용 기타 부품(840999)_수출시장분석보고서.pdf
- 1946_840999_자동차엔진용기타부품_프랑스_20260526_100000.pdf

명령줄 옵션
- 다른 폴더를 검사:
  check_reports.bat --folder "C:\보고서점검"

- 엑셀 파일을 직접 지정:
  check_reports.bat --input "C:\보고서점검\보고서생성리스트.xlsx"

- PDF 폴더를 따로 지정:
  check_reports.bat --input "C:\보고서점검\보고서생성리스트.xlsx" --report-dir "C:\보고서점검\PDF"

- 하위 폴더 PDF까지 검사:
  check_reports.bat --recursive
