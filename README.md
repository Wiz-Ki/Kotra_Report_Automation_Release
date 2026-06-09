# KOTRA 수출시장 분석보고서 자동생성기

수출시장 분석보고서 생성 과정의 반복 입력, 보고서 생성 대기, PDF 다운로드를 자동화하는 내부 업무 보조 도구입니다.

- 대상 사이트: https://kotra.or.kr/mutugpt/export-assistant/report
- 실행 방식: Windows VM에서 ZIP 다운로드 후 `.bat` 파일 실행
- 브라우저: VM에 설치된 Microsoft Edge 우선 사용
- 배포 방식: PyInstaller exe 없이 Python 런타임을 폴더 내부에 자동 준비
- 백그라운드 실행: Playwright Chromium headless shell 우선 사용, 실패 시 Microsoft Edge headless로 대체

본 프로그램은 내부 업무 효율화를 위한 보조 도구이며, KOTRA의 공식 대외 서비스가 아닙니다.

## 빠른 사용법

자세한 현장용 안내는 `설명서.txt`와 `설명서(상세).txt`를 참고하세요.

1. GitHub 배포 저장소에서 `Code > Download ZIP`을 선택합니다.
2. Windows VM에서 ZIP 파일을 압축 해제합니다.
3. 압축 해제된 폴더 안에서 `install_vm.bat`를 실행합니다.
   - 처음 1회만 실행합니다.
   - 공식 Python을 `portable_python` 폴더에 설치합니다.
   - 필요한 Python 패키지를 자동 설치합니다.
4. 설치가 끝나면 `run_gui.bat`를 실행합니다.
5. GUI에서 입력 엑셀 파일과 다운로드 폴더를 선택한 뒤 실행합니다.

처음 한 번:

```bat
install_vm.bat
```

매번 실행:

```bat
run_gui.bat
```

콘솔 실행 또는 문제 확인:

```bat
run_cli.bat
```

### 📊 액셀 파일 업로드시 참고사항

> [!CAUTION]
> **중요:** 아래 양식에 맞게 액셀 파일을 수정하여 반출 후 프로그램에 입력해주세요!!   
> **선택항목은 행 자체는 유지하고, 값을 비워주세요!**
<img width="2284" height="354" alt="input_template_example" src="https://github.com/user-attachments/assets/a99390ac-a77b-46f5-a0e2-a1403e178a56" />



## 실행 조건

- Windows 원격 VM에서 실행합니다.
- VM은 인터넷 접속이 가능해야 합니다.
- VM에 Microsoft Edge가 설치되어 있어야 합니다.
- `.bat` 파일 실행이 허용되어야 합니다.
- ZIP 파일 안에서 바로 실행하지 말고 반드시 압축을 푼 뒤 실행해야 합니다.

## 입력 엑셀

기본 템플릿 파일은 `input_template.xlsx`입니다.

실제 실행용 엑셀은 `input_template.xlsx`를 복사해서 작성하거나, GUI에서 원하는 `.xlsx` 파일을 직접 선택하면 됩니다.

필수 입력값:

- 수출액 규모
- 해당 품목 수출 경험
- HS CODE
- 희망진출국가

입력 템플릿 컬럼:

| 컬럼명 | 설명 | 필수 |
| --- | --- | --- |
| `연번(선택)` | 원본 데이터 순번 | X |
| `회사명(선택)` | 회사명 | X |
| `사업자번호(선택)` | 사업자번호 | X |
| `수출액 규모(필수: 내수/초보/유망/성장/선도 중 1개)` | 내수/초보/유망/성장/선도 분류 | O |
| `해당 품목 수출 경험(필수: O/X)` | O = 수출 경험 있음, X = 처음입니다 | O |
| `HS CODE(필수: 6자리 숫자)` | 수출품 HS CODE | O |
| `수출품명(선택: 구체적으로 작성 권장)` | 수출품명 | X |
| `희망진출국가(필수: 국가 1개만 입력)` | 희망 진출 국가 | O |

기존 영문 컬럼(`hs_code`, `product_name`, `export_scale`, `export_experience`, `target_country`)도 지원합니다.

## 프로그램 동작

사용자가 입력 엑셀을 선택하면 프로그램은 각 행을 순서대로 처리합니다.

1. KOTRA 보고서 생성 페이지 접속
2. `희망 국가 직접 분석` 선택
3. 엑셀 한 행의 값 입력
4. 보고서 생성 버튼 클릭
5. 다운로드 버튼이 나타날 때까지 최대 15분 대기
6. 버튼이 나타나면 PDF 다운로드
7. `downloads` 폴더 또는 사용자가 선택한 폴더에 저장
8. 성공/실패 로그 기록
9. 다음 행 처리

기본 저장 파일명은 KOTRA 사이트가 제안한 파일명을 그대로 사용합니다.

GUI의 `저장 파일명 커스텀`을 켜면 항목과 문자를 조각으로 추가하고, 조각을 드래그해 순서를 바꾸거나 `×`로 삭제할 수 있습니다.

CLI에서 실행할 때는 `--filename-pattern`을 입력하면 같은 규칙으로 저장 파일명을 만들 수 있습니다. 사용할 수 있는 항목은 다음과 같습니다.

| 항목 | 패턴 |
| --- | --- |
| 연번 | `{row_index}` |
| HS CODE | `{hs_code}` |
| 수출품명 | `{product_name}` |
| 희망진출국가 | `{target_country}` |
| 생성날짜 | `{date}` |
| 생성시간 | `{time}` |
| 생성일시 | `{datetime}` |
| 생성연도 | `{year}` |
| 생성월 | `{month}` |
| 생성일 | `{day}` |
| 생성시 | `{hour}` |
| 생성분 | `{minute}` |
| 생성초 | `{second}` |
| 사이트 기본 파일명 | `{site_filename}` |

예시 패턴:

```text
{target_country}_{product_name}({hs_code})_수출시장분석보고서
```

날짜 구분자를 직접 넣고 싶다면 다음처럼 조합할 수 있습니다.

```text
{year}-{month}-{day}_{hour}.{minute}.{second}
```

예시 저장명:

```text
베트남_스킨케어(330499)_수출시장분석보고서.pdf
```

행 처리 중 실패하면 해당 항목은 재시도 대기열에 올라가며 기본 1회에 한하여 다시 시도합니다. 자동 재시도 후에도 실패한 항목만 `logs\failed_rows.xlsx`에 기록됩니다.

GUI의 `실패 항목 자동 재시도(1회)` 옵션은 실행 중에도 변경할 수 있으며, 변경 후 새로 실패하는 항목부터 반영됩니다. 이미 재시도 대기열에 올라간 항목은 그대로 1회 재시도됩니다.

다운로드 버튼이 30초 후 나타나면 30초 후 바로 다운로드합니다. 15분 안에 나타나지 않으면 해당 행만 실패 처리하고 다음 행으로 넘어갑니다.

## 설치 스크립트가 하는 일

`install_vm.bat`는 최초 1회 실행하는 준비 스크립트입니다.

- 공식 Python Windows 설치파일을 다운로드합니다.
- 현재 폴더 안에 `portable_python` 폴더를 만듭니다.
- `portable_python` 안에 Python을 설치합니다.
- `requirements-runtime.txt` 기준으로 필요한 패키지를 설치합니다.
- 백그라운드 실행용 Playwright Chromium headless shell 설치를 시도합니다.
- Microsoft Edge 설치 여부를 확인합니다.
- 프로그램 import 가능 여부를 확인합니다.

설치가 끝나면 `VM setup passed.` 메시지가 표시됩니다.

프로그램 실행 시 화면 표시 모드에서는 Microsoft Edge를 먼저 실행합니다. 백그라운드 실행 모드에서는 Windows 원격 VM에서 빈 브라우저 창이 뜨는 현상을 줄이기 위해 Playwright Chromium headless shell을 먼저 실행하고, 설치되어 있지 않으면 Microsoft Edge headless로 대체합니다. Edge 실행에 실패하면 macOS 등 로컬 테스트 편의를 위해 Playwright 기본 Chromium으로 한 번 더 시도합니다.

## 주요 파일

| 파일 | 용도 |
| --- | --- |
| `설명서.txt` | 가장 짧은 실행 안내 |
| `설명서(상세).txt` | 상세 실행 안내 |
| `install_vm.bat` | 최초 1회 설치/준비 |
| `run_gui.bat` | GUI 실행 |
| `run_cli.bat` | 콘솔 실행 |
| `requirements-runtime.txt` | 런타임 Python 패키지 목록 |
| `input_template.xlsx` | 입력 엑셀 템플릿 |
| `main.py` | 콘솔 진입점 |
| `gui_launcher.py` | GUI 진입점 |
| `automation.py` | 보고서 생성 자동화 로직 |

## 결과와 로그

기본 PDF 저장 폴더:

```text
downloads
```

기본 로그 폴더:

```text
logs
```

주요 로그 파일:

- `logs\success_log.xlsx`
- `logs\failed_rows.xlsx`
- `logs\processing_status.xlsx`
- `logs\startup_error.txt`
- `logs\diagnostics`

행 처리 중 오류가 나면 해당 시점의 화면 스크린샷과 최근 브라우저 콘솔/네트워크 오류가 `logs\diagnostics`에 저장됩니다.

## 실패 행 재시도

GUI에서 `실패 행만 재시도`를 선택하면 `logs\failed_rows.xlsx` 기준으로 실패한 행만 다시 실행합니다.

일반 실행 중 실패한 항목은 재시도 대기열에 올라가 1회 다시 처리되며, 자동 재시도 후에도 실패한 항목만 실패 행 재시도 대상에 남습니다.

콘솔에서 실행하려면:

```bat
run_cli.bat --retry-failed
```

## 콘솔 옵션

GUI 대신 콘솔에서 실행할 수 있습니다.

```bat
run_cli.bat
```

엑셀 파일 직접 지정:

```bat
run_cli.bat --input input.xlsx
```

브라우저 백그라운드 실행:

```bat
run_cli.bat --input input.xlsx --headless
```

자동 재시도 없이 실행:

```bat
run_cli.bat --input input.xlsx --no-auto-retry
```

입력 템플릿 생성:

```bat
run_cli.bat --create-template
```

## 자주 발생하는 문제

### `install_vm.bat`에서 Python 다운로드 실패

- VM 인터넷 연결을 확인하세요.
- `python.org` 접속이 보안 정책으로 막혀 있는지 확인하세요.

### pip 설치 실패

- VM에서 `pypi.org` 접속이 가능한지 확인하세요.
- 사내 프록시 설정이 필요한 환경인지 확인하세요.

### Microsoft Edge를 찾지 못함

- VM에 Microsoft Edge가 설치되어 있는지 확인하세요.
- Edge 실행이 보안 정책으로 차단되어 있는지 확인하세요.

macOS 등 로컬 테스트 환경에서는 Edge가 없어도 Playwright Chromium이 설치되어 있으면 fallback으로 실행될 수 있습니다.

### Edge는 열리지만 자동화가 실패함

- 사내 Edge 정책이 자동화 실행, 새 브라우저 프로필 생성, 다운로드를 막는지 확인하세요.
- `logs\startup_error.txt` 또는 `logs\diagnostics` 폴더를 확인하세요.

### GUI가 열리지 않음

- `install_vm.bat`를 먼저 실행했는지 확인하세요.
- `portable_python` 폴더가 생성되었는지 확인하세요.
- 아래 명령으로 Python 실행 여부를 확인하세요.

```bat
run_cli.bat --create-template
```

## 입력값 참고

- HS CODE는 6자리 숫자로 입력합니다.
- `수출품명`은 구체적으로 입력할수록 좋습니다.
- `수출액 규모`는 `내수`, `초보`, `유망`, `성장`, `선도`를 입력할 수 있습니다.
- 숫자 수출액이 들어오면 사이트의 달러 구간에 맞춰 자동 변환합니다.
- `해당 품목 수출 경험`은 `O` 또는 `X`만 입력합니다.
- `희망진출국가`는 한 행에 국가 1개만 입력하는 것을 권장합니다.

## 프로젝트 구조

```text
KOTRA_Report_Auto_Release/
 ├─ README.md
 ├─ 설명서.txt
 ├─ 설명서(상세).txt
 ├─ install_vm.bat
 ├─ run_gui.bat
 ├─ run_cli.bat
 ├─ requirements-runtime.txt
 ├─ input_template.xlsx
 ├─ main.py
 ├─ automation.py
 ├─ gui_launcher.py
 ├─ gui_v2.py
 ├─ config.py
 ├─ field_mapping.py
 ├─ logger.py
 ├─ site_selectors.py
 ├─ template.py
 ├─ compat.py
 ├─ assets/
 └─ scripts/
```

실행 후 생성되는 폴더:

```text
portable_python/
.setup_downloads/
downloads/
logs/
```
