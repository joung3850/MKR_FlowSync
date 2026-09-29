# Python 기반 MKR·Outlook 분석 프로그램 구현 계획

## 현재 구현 범위

`MKR_SYNC.exe` 하나를 실행하면 Edge WebView2 기반 HTML 화면이 네이티브 창으로 열린다. 별도 CMD와 외부 브라우저는 사용하지 않는다. Python은 Excel COM, Classic Outlook, PDF 분석, 수량 판정, 로그와 CSV 보고서를 담당한다.

1차 버전은 다음 기능을 제공한다.

- 대상 Excel 선택
- 일본어 설명문을 포함한 입력에서 MKR 번호 추출 및 중복 제거
- 외부 `MKR_TEMPLATE.xlsx`의 `MKR_TEMPLATE` 시트로 누락 MKR 시트 생성
- 대상 Excel 백업과 pending 파일 검증 후 원본 교체
- Classic Outlook 폴더, 발신인, 시작일, 최대 검색 건수 지정
- 주문 Excel, Sales Note PDF, 메일 본문, 백오더 자료 분석
- 주문 최초본·추가본·개정본의 시간순 상태 재구성
- 최신 유효 Sales Note와 백오더를 결합한 R·S·P·K·L 판정 미리보기
- `Reports/MKR_PREVIEW_날짜시간.csv`와 실행 로그 생성
- 성공한 MKR 미리보기 결과만 사용자 승인 후 부분 적용
- A:G, H:N, D3, G8, K4:K7 기록 전 백업, pending 저장, 값·구조 검증, 원본 교체
- 실행 중 진행률 표시와 처리 단위 경계의 취소

미리보기 단계에서는 Excel을 변경하지 않는다. 사용자가 `Excel에 적용`을 승인한 경우에만 검증된 미리보기 스냅샷을 기록한다.

## 실행 구조

```text
MKR_SYNC.exe
├─ pywebview + Edge WebView2
├─ mkr_sync/web/index.html
├─ WebApi JavaScript bridge
├─ JobManager
├─ 외부 템플릿 기반 시트 생성 서비스
├─ Classic Outlook 수집 서비스
├─ 주문서·Sales Note·백오더 파서
├─ STANDARD_V1 수량 판정 엔진
└─ 로그·CSV 보고서
```

로컬 웹 서버와 포트는 사용하지 않는다. HTML은 EXE 리소스로 포함되며 `window.pywebview.api`를 통해 Python만 호출한다.

## 외부 템플릿과 Excel 안전성

템플릿 검색 우선순위는 다음과 같다.

1. `MKR_SYNC.exe` 옆 `MKR_TEMPLATE.xlsx`
2. `settings.json`에 저장된 사용자 지정 경로
3. 네이티브 파일 선택창

템플릿은 읽기 전용으로 열며 수정하지 않는다. 누락 시트가 없으면 템플릿 파일 없이도 기존 시트를 검증하고 Outlook 분석으로 이동할 수 있다.

시트 생성 순서:

1. 대상 파일, 시트명, 기존 시트 D3와 헤더 검증
2. 템플릿 `MKR_TEMPLATE` 시트, 헤더, 병합 구조, 외부 참조 수식 검증
3. `Backup` 폴더에 원본 백업
4. 대상 파일의 pending 복사본 생성
5. 동일 Excel COM 인스턴스에서 템플릿 시트를 pending 파일로 복사
6. 시트명 변경 및 D3에 원래 MKR 번호 기록
7. 저장 후 COM 및 OOXML 시트 목록 재검증
8. 검증 성공 시에만 `os.replace`로 원본 교체

오류 시 pending 파일을 폐기하고 원본과 템플릿은 유지한다. 대상 Excel에는 `MKR_TEMPLATE` 시트가 남지 않는다.

## Outlook 분석

- 실행 중인 Classic Outlook만 지원한다.
- 사용자가 선택한 폴더를 EntryID와 StoreID로 다시 연다.
- 1~500건 범위에서 최신 메일부터 검색한다.
- 여러 발신인을 허용한다.
- 제목에서 MKR을 우선 찾고, 없을 때 현재 메일 본문을 보완 검색한다.
- 한 메일에 여러 대상 MKR이 있어도 각 대상별 자료로 분리한다.
- 주문 첨부를 시간순으로 처리한다. 개정본은 전체 상태를 교체하고, 명확한 추가본은 누적본 또는 증분본으로 병합한다.
- 같은 메일의 A(KRW)/B(THB) 파일을 합치고 품목별 Origin을 Japan/Thailand/Mixed/Unknown으로 보존한다.
- 최신 유효 Sales Note, 최신 백오더 스냅샷, 후속 일정 메일을 분석한다.
- Sales Note 미수신 시 Shipped/Shortage를 추정하지 않고 공란으로 유지한다.
- 주문서에는 없지만 Sales Note에 남은 출하 품목을 `SALES_NOTE_ONLY`로 보존한다.
- PDF는 `pypdf` 텍스트 추출을 우선하고 실패하면 Word COM을 보조 수단으로 사용한다.
- 한 MKR 분석이 실패해도 다음 MKR을 계속 분석하고 오류를 마지막에 모아 표시한다.

첨부파일은 실행별 `MKR_Attachments/Run_...` 폴더에 저장한다. 미리보기 단계는 대상 Excel 데이터 영역을 변경하지 않는다.

## Excel 데이터 적용

1. 완료된 미리보기 결과와 대상 파일 경로를 대조한다.
2. 미리보기 당시 SHA256과 현재 대상 Excel의 SHA256을 비교한다.
3. 파일이 변경되었으면 적용을 중단한다. MKR별 분석 오류는 해당 시트만 보류한다.
4. `Backup` 폴더에 원본을 백업한다.
5. pending Excel에 A10:G509, H10:N509, D3, G8, K4:K7을 기록한다.
6. 헤더, 병합 구조, 시트명, 품번, 수량, 숫자 서식을 재검증한다.
7. 검증에 성공한 pending 파일만 원본 위치로 교체한다.
8. `Reports/MKR_APPLIED_날짜시간.csv` 감사 보고서를 생성한다.

## 공개 JavaScript API

```python
get_startup_state()
choose_target_workbook()
choose_template_workbook()
extract_mkr_numbers(text)
create_mkr_sheets(request)
list_outlook_folders()
start_preview(request)
apply_preview(request)
get_job_status(job_id)
cancel_job(job_id)
open_report_folder()
```

대상 파일과 템플릿 파일은 네이티브 파일 선택창 또는 검증된 설정 경로에서만 받는다. 공개 API는 임의 명령 실행이나 파일 삭제 기능을 제공하지 않는다.

## 패키징

- Python 3.14 64비트
- `pywin32==312`
- `pywebview==6.2.1`
- `pypdf==6.18.1`
- Python 3.14 대응 PyInstaller
- `onefile`, `windowed`
- EdgeChromium, pywin32, 실행 HTML 포함
- Qt, CEF, GTK 제외
- 출력: 프로젝트 루트의 `MKR_SYNC.exe`

배포 폴더:

```text
MKR_SYNC.exe
MKR_TEMPLATE.xlsx
settings.json
Logs/
Reports/
Backup/
MKR_Attachments/
```

기존 Tkinter 화면은 회귀 확인용 `python -m mkr_sync tk-gui`로 남아 있으며, 기본 GUI와 EXE 진입점은 HTML 내장 앱을 실행한다.

## 완료 검증

- 일본어 설명문에서 MKR 8개 추출, 중복 제거, P/O 번호 제외
- JavaScript와 Python 추출 결과 일치
- 외부 템플릿 해시 불변
- 대상에 `MKR_TEMPLATE` 시트가 남지 않음
- 여러 누락 시트 일괄 생성
- 기존 시트 D3 또는 구조 불일치 시 전체 중단
- 실패 시 원본 대상 Excel 해시 불변
- 여러 발신인과 여러 MKR 제목 처리
- 수정 주문과 중복 첨부 처리
- PDF 텍스트 추출과 Word fallback
- 경고와 오류를 HTML에서 구분
- 일부 미리보기 오류가 있어도 성공한 MKR은 적용 가능하며 실패 MKR은 기존 값 유지
- 미리보기 후 대상 파일 변경 시 적용 중단
- 적용 실패 시 원본 해시 유지, 성공 시 A:G·H:N·D3·G8·K4:K7 값 검증
- 취소와 중복 실행 방지
- EXE 더블클릭 시 CMD와 외부 브라우저 없이 창 하나만 표시
- Windows 10/11, Excel 데스크톱, Classic Outlook, WebView2 환경 통합 테스트

## 다음 단계

실사용 메일 표본을 확대해 품목명 변형, 표 구조 예외, 일정 문구 예외를 회귀 테스트에 추가한다.
