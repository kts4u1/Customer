# 고객안내문(FSN) · FCO 통합 Dashboard — Claude Code 인계 문서

작업 시작 전에 `main.py`와 `dashboard_template.html`을 먼저 읽고, 아래 내용과 실제 코드가 일치하는지 확인하세요.
수정 후에는 `python -m py_compile main.py`와 `python tests/test_parsing.py`를 실행합니다.

## 사용자 / 작업 방식
- 사용자는 Philips 한국 현장 서비스 엔지니어이며 프로그래밍 경험이 없습니다. 코드는 전적으로 Claude가 작성·디버깅합니다.
- 대화는 한국어 존댓말로 합니다. 설명은 짧고 구체적으로, 실행할 명령어는 그대로 복사해서 쓸 수 있게 줍니다.
- 실행 환경은 회사 Windows PC, Python 3.14.x 입니다. `pip`가 PATH에 없어서 항상 `python -m pip`를 사용합니다.
- exe는 `--noconsole`로 빌드됩니다. 오류는 messagebox와 `~/FSN_Dashboard/debug_log.txt`로 보이게 합니다. "아무 반응 없음"이 가장 큰 디버깅 장애물이었습니다(FCO 때 경험).
- 큰 구조 변경은 먼저 사용자에게 확인을 받습니다. 사용자는 "일단 만들어 보고 수정 사항은 하나씩" 진행하는 방식을 원합니다.

## 프로그램 개요
SharePoint(docs.philips.com, OneDrive 개인 공유 링크)에 있는 고객안내문 전달 현황 엑셀을 받아서
HTML 대시보드(`~/FSN_Dashboard/FSN_FCO_Dashboard.html`)를 만들고 Edge로 엽니다.

- 스택: Python, Tkinter(진행 창), openpyxl, python-dateutil, Selenium(Edge), PyInstaller → `고객안내문.exe`
- 화면은 Tkinter가 아니라 HTML(외부 라이브러리 없는 순수 JS/CSS, 오프라인 동작)입니다. 데이터는 JSON으로 `__DASHBOARD_DATA__` 자리에 들어갑니다.
- 파일: `main.py`, `dashboard_template.html`, `FSN.ico`, `requirements.txt`, `build.bat`(CRLF, ASCII만), `build.py`, `README.txt`, `config.example.json`, `tools/make_sample_data.py`(가상 샘플), `tools/make_icon.py`, `tests/test_parsing.py`
- 개발용 CLI: `python main.py --build 고객안내문.xlsx --fco "MR=MR FCO Action Plan.xlsx" --out test.html --today 2026-10-23 --sample --no-open`

## 사용자 요구사항 (2026-10-23 최초 요청)
1. BU별, C&R별, 전달방법(등기/직접전달/이메일)별 FSN 전달 상황 시각화 + 아직 전달이 필요한 병원 리스트
2. 세일즈/FSE용 병원별 전달 현황
3. FCO + FSN 통합 Dashboard

## 대시보드 구성 (dashboard_template.html)
- 상단 공통 필터: 담당자(FSE·영업), BU, 검색. 모든 탭에 적용.
- 탭: 통합 개요 / 고객안내문(FSN) / FCO / 병원별 현황 / 데이터 정보(컬럼 인식 진단)
- 색상(검증 완료, 바꾸지 말 것): FSN 완료=파랑 `#2a78d6`, 미전달=빨강 `#d03b3b` (빨강-녹색은 적록색약 구분 불가로 제외).
  FCO 구간: 6개월 미만 `#d03b3b`, 1년 미만 `#fab219`, 1년 이상 `#0ca30c`, 기한 없음 회색. 라이트/다크 모드 모두 지원.
- 담당자 필터 규칙: 선택한 사람이 그 데이터(FSN 또는 FCO)의 담당자 컬럼에 있으면 행 단위로, 없으면(예: 영업 담당이 FCO 탭을 볼 때) 병원 단위로 거릅니다.
- CSV 저장은 UTF-8 BOM. (리눅스 headless 테스트에서 한글 파일명이 "download"로 나오는 건 로케일 문제, UTF-8 로케일/Windows에선 정상)

## 데이터 사양
### 고객안내문 파일 — 2026-10-23 사용자 캡처로 확인된 것
- 파일명: `고객안내문 발송현황_Latest.xlsx`
- 시트: IB_raw data, IGT DXR US HPM, MR, CT, (삭제예정)SMAX자료, IGT FCO release date, 자동연동 설정, (삭제예정)매칭 QA 점검, C&R, FCO 매칭, Closed, …
  → BU 컬럼이 없고 시트 이름이 BU 역할 → 시트가 여러 개면 시트 이름을 BU로 사용. 이름에 "삭제"/"closed"가 들어간 시트는 제외(`SKIP_SHEET_WORDS`, config `skip_sheets`).
- MR 시트 헤더(1행, 일부): C&R#, 장비 ID, 장비코드(6NC/12NC), 장비명, 일련번호(S/N), 고객코드(향후 업데이트), 영문병원명(업데이트예정), (가려짐…), 전달방법 값으로 보이는 칸(등기/직접전달/이메일+공문), 공문(방법 or N), 대리점명(or N), 지역, 비고, 최종증빙 확인, 수신확인서 (Y/N/NA), (최초…일)
- **병원 이름은 `국문병원명` / `고객명(국문)`** (사용자 지정). 우선순위: 국문병원명 → 고객명(국문) → 영문병원명. "-", "N", "0" 같은 값은 빈칸 취급(`row_hospital`, `PLACEHOLDER`).
- 아직 미확인: 어떤 컬럼이 "전달 완료"를 뜻하는지(발송일? 최종증빙? 수신확인서?). [데이터 정보] 탭 캡처나 1행 헤더 전체 캡처를 받아서 확정해야 함.

### 고객안내문 파일 (일반 규칙)
- URL은 `main.py`의 `FSN_SOURCE`. 회사 로그인이 필요해서 Claude 환경에서는 열어볼 수 없었습니다(403).
- 그래서 시트/헤더/컬럼을 자동 인식합니다: `FSN_FIELDS`(정확 일치 별칭 → 포함 단어 순), 위쪽 25행 안에서 가장 많이 맞는 헤더 행, 두 줄 헤더, 병합 셀(세로 병합 값 채움), 등기/직접전달/이메일이 칸으로 나뉜 형태(`WIDE_METHOD_HEADERS`) 지원.
- 같은 구조의 시트가 여러 개면 합치고(FSN 컬럼이 없으면 시트 이름을 FSN 구분으로), 구조가 다른 요약 시트는 제외합니다. 숨김 시트 제외.
- 완료 판정(`classify_status`): 전달여부 값 해석 → 비었으면 전달일/방법 칸 표시로 보충. 전달여부 컬럼이 없으면 전달일, 그것도 없으면 전달방법 입력 여부. "해당없음/제외/N/A/철거" 등은 집계 제외.
- 사용자가 `~/FSN_Dashboard/config.json`의 `"columns"`로 컬럼을 직접 지정할 수 있습니다.
- **실제 파일을 받으면 [데이터 정보] 탭 캡처로 매핑을 확인하고 별칭/판정 규칙을 조정하는 것이 다음 작업입니다.**

### FCO 파일
- 기존 FCO 프로그램과 동일: MR / XR / CT/AMI, 컬럼 FCO, Due Date, FCO Title, …, Account, Site, FSE, Action Plan, Status, 고객안내문 관련 …
- `Status`가 Closed인 행 제외. 구간: Due Date가 오늘+6개월 미만 빨강, +1년 미만 노랑, 그 외 녹색.
- FCO SharePoint URL은 아직 모릅니다(`FCO_SOURCES`의 url이 빈칸). 빈칸이면 기존 FCO 프로그램 캐시 `~/FCO_Action_Plan/{MR_FCO_Action_Plan,XR_Open_FCO_list,CT_AMI_FCO_Action_Plan}.xlsx`를 읽기 전용으로 사용합니다. 사용자에게 기존 FCO `main.py`의 `SOURCES` URL을 받으면 채우세요.
- 병원 연결: 병원명 정규화(`hospital_key`: 공백·(재)·의료법인 등 제거) + 장비 S/N이 같으면 FCO Account 쪽 병원으로 묶음.

## 자동 다운로드 (EdgeDownloader)
- Edge(Windows 계정 SSO) + Selenium, 프로필 `~/FSN_Dashboard/browser_profile`, 다운로드 폴더 `~/FSN_Dashboard/downloads`(+ `~/Downloads`, `~/Desktop` 감시).
- 순서: ⓪ 공유 링크 페이지를 연 뒤 페이지 안에서 `fetch(링크+download=1, credentials:'include')`로 파일 내용을 직접 받음(`FETCH_JS`, `_fetch_in_page`).
  브라우저 다운로드를 쓰지 않아서 Edge의 "이 파일을 열까요? [열기][취소]" 창이 안 뜸(1차 버전에서 사용자가 이 창을 눌러야 받아지던 문제). 실패 시 아래로.
  ① 공유 링크에 `download=1` 붙여 바로 받기(CDP `Browser.setDownloadBehavior allow`로 묻지 않게) → ② Excel Online `파일 → 복사본 만들기 → 복사본 다운로드`(영/한, iframe 재귀 탐색, 최대 3회) → ③ 사용자가 직접 받기 3분 대기.
- 로그인 페이지(login.microsoftonline 등)면 사용자에게 로그인 안내 후 최대 150초 대기.
- 브라우저 다운로드 경로에서는 고객안내문은 "시작 이후 새로 생긴 xlsx", FCO는 "FCO + MR/XR/CT"로 인식.
- 캐시: `~/FSN_Dashboard/data/FSN_customer_notice.xlsx`(+ `.meta.json`에 원래 파일명), 오늘 받은 파일이면 재사용.

### FCO 프로그램에서 해결했던 문제 (같은 방식 유지)
- `page_load_strategy="eager"` + 45초 타임아웃, `command_executor.set_timeout(20)`.
- "Please wait... loading the workbook" 배너가 iframe 안에 있음 → iframe 재귀 탐색으로 배너가 사라질 때까지 대기.
- 파일명 특수문자 제거, 엑셀 불허 제어문자 제거(`clean_text`).
- PyInstaller `--collect-submodules=selenium` 필수.

## 빌드 (Windows PC에서만)
`build.bat` 더블클릭 → `python -m pip install -r requirements.txt` → `python build.py`
→ PyInstaller로 `dist\FSN_Dashboard.exe`를 만든 뒤 `dist\고객안내문.exe`로 복사(배치 파일 한글 깨짐 방지).
`dashboard_template.html`, `FSN.ico`는 `--add-data`로 exe 안에 포함됩니다.

## 검증한 것 / 남은 것
- 검증: 샘플 2종(일반 표, 두 줄 헤더+병합+방법 칸 분리) 파싱, 대시보드 5개 탭 렌더링(데스크톱/모바일/다크), 필터·클릭·CSV, Tkinter 진행 창(캐시 사용/다운로드 실패), 리눅스 PyInstaller 빌드 실행.
- 미검증: 실제 회사 PC에서 Edge 자동 다운로드, 실제 고객안내문 파일의 컬럼 구조, Windows exe 빌드.
- FCO 단계의 "!" 표시는 실패가 아니라 "FCO 주소 미설정 → FCO 프로그램이 받아 둔 이전 날짜 파일 사용"이라는 뜻(진행 창에 이유 표시하도록 수정).
- 열린 항목: FCO URL 3개, 실제 파일 기준 컬럼/판정 규칙 조정, 사용자 피드백 반영.
