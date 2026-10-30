# 고객안내문(FSN) · FCO 통합 Dashboard — Claude Code 인계 문서

claude.ai / Claude Code에서 진행한 "고객안내문 전달 현황 대시보드" 작업을 이어서 하기 위한 문서입니다.
작업 시작 전에 이 폴더의 `main.py`와 `dashboard_template.html`을 먼저 끝까지 읽고, 아래 내용과 실제 코드가 일치하는지 확인해 주세요.
파일이 폴더에 없으면 사용자에게 달라고 요청하세요. (GitHub: `kts4u1/Customer`, 브랜치 `claude/peaceful-sagan-i5msvi`)

## 사용자 / 작업 방식
- 사용자는 Philips 한국 현장 서비스 엔지니어이며 프로그래밍 경험이 없습니다. 코드는 전적으로 Claude가 작성·디버깅합니다.
- 대화는 한국어 존댓말로 합니다. 설명은 짧고 구체적으로, 실행할 명령어는 그대로 복사해서 쓸 수 있게 줍니다.
- 실행 환경은 회사 Windows PC, Python 3.14.x 입니다. `pip`가 PATH에 없어서 항상 `python -m pip`를 사용합니다.
- 사용자는 GitHub 브랜치를 ZIP으로 받아 빌드합니다:
  `https://github.com/kts4u1/Customer/archive/refs/heads/claude/peaceful-sagan-i5msvi.zip`
  → 압축 풀기 → `build.bat` 더블클릭 → `dist\고객안내문.exe`. 회사 PC에서 직접 고친 코드는 다음 ZIP에서 사라지므로, 주소·설정은 코드에 넣거나 `config.json`을 씁니다.
- exe는 `--noconsole`로 빌드됩니다. 오류는 messagebox와 `~/FSN_Dashboard/debug_log.txt`로 보이게 합니다. "아무 반응 없음"이 가장 큰 디버깅 장애물이었습니다(FCO 때 경험).
- 수정 후에는 항상 `python -m py_compile main.py`와 `python tests/test_parsing.py`를 실행합니다.
- 큰 구조 변경은 먼저 사용자에게 확인을 받습니다. 사용자는 "일단 만들어 보고 수정 사항은 하나씩" 진행하는 방식을 원합니다.
- 실제 회사 데이터는 Claude 환경에서 볼 수 없습니다(SharePoint 403). 사용자 화면 캡처와 대시보드 [데이터 정보] 탭이 유일한 단서입니다.

## 프로그램 개요
SharePoint에 있는 고객안내문(FSN) 전달 현황 엑셀과 FCO Action Plan 엑셀 3개를 받아서
HTML 대시보드(`~/FSN_Dashboard/FSN_FCO_Dashboard.html`)를 만들고 Edge로 엽니다. 더블클릭하면 실행됩니다.

- 스택: Python, Tkinter(진행 창만), openpyxl, python-dateutil, Selenium(Edge), PyInstaller → `고객안내문.exe`
- 화면은 Tkinter가 아니라 HTML입니다(외부 라이브러리 없는 순수 JS/CSS, 오프라인 동작). 데이터는 JSON으로 템플릿의 `__DASHBOARD_DATA__` 자리에 들어갑니다. HTML 파일 하나라 메일로 보내도 열립니다.
- 파일 구성: `main.py`, `dashboard_template.html`, `FSN.ico`(청록→파랑 봉투 + 체크 아이콘), `requirements.txt`, `build.bat`(CRLF, ASCII만), `build.py`, `README.txt`, `config.example.json`, `tools/make_sample_data.py`(가상 샘플), `tools/make_icon.py`, `tests/test_parsing.py`, `preview/sample_dashboard.html`(샘플 데이터 미리보기)
- 개발용 CLI: `python main.py --build 고객안내문.xlsx --fco "MR=MR FCO Action Plan.xlsx" --out test.html --today 2026-10-23 --sample --no-open`

## 진행 창 (Tkinter) 흐름
1. 고객안내문 파일 → 2. FCO 파일 (MR · XR · CT/AMI) → 3. 대시보드 만들기. 단계별로 ✔ / ! / ✖ 표시.
- 버튼: 대시보드 열기 / 새로 받기(오늘 받은 파일이 있어도 다시 받기) / 파일 직접 선택(다운로드 실패 시) / 폴더 열기 / 닫기.
- 오늘 받은 파일이 있으면 받지 않고 재사용합니다. 받기 실패 시 이전 파일로 진행하고 "!"와 파일 날짜를 표시합니다.

## 사용자 요구사항 (2026-10-23 최초 요청)
1. BU별, C&R별, 전달방법(등기/직접전달/이메일)별 FSN 전달 상황 시각화 + 아직 전달이 필요한 병원 리스트
2. 세일즈/FSE용 병원별 전달 현황
3. FCO + FSN 통합 Dashboard

## 대시보드 구성 (dashboard_template.html)
- 탭: 통합 개요 / 고객안내문(FSN) / FCO / 병원별 현황 / 데이터 정보(시트·컬럼 인식 진단)
- 상단 공통 필터: 담당자(FSE·영업), 검색. 모든 탭에 적용. 막대·숫자 카드를 누르면 해당 필터가 걸림.
- 고객안내문 탭: C&R별, 담당 FSE별, 영업 담당별 전달 막대 + "아직 전달이 필요한 병원" 표(병원별/건별, CSV 저장).
- **BU와 전달방법은 화면에서 숨김** (2026-10-23 사용자 요청: 정보가 아직 정확하지 않음). `dashboard_template.html`의 `SHOW_BU`, `SHOW_METHOD` = false.
  사용자 설명: BU = 지역 묶음. BU별 지역 목록을 받으면 엑셀 '지역' 컬럼 → BU 매핑을 만들고 `SHOW_BU = true`로 다시 켤 것. (현재 파서는 시트 이름을 bu에 넣지만 화면엔 안 나옴)
- 색상(검증 완료, 바꾸지 말 것): FSN 완료=파랑 `#2a78d6`, 미전달=빨강 `#d03b3b` (빨강-녹색은 적록색약 구분 불가로 제외).
  FCO 구간: 6개월 미만 `#d03b3b`, 1년 미만 `#fab219`, 1년 이상 `#0ca30c`, 기한 없음 회색. 라이트/다크 모드 모두 지원.
- 담당자 필터 규칙: 선택한 사람이 그 데이터(FSN 또는 FCO)의 담당자 컬럼에 있으면 행 단위로, 없으면(예: 영업 담당이 FCO 탭을 볼 때) 병원 단위로 거릅니다.
- CSV 저장은 UTF-8 BOM. (리눅스 headless 테스트에서 한글 파일명이 "download"로 나오는 건 로케일 문제, Windows에선 정상)

## 데이터 사양
### 고객안내문 파일 (사용자 캡처로 확인)
- URL: `main.py`의 `FSN_SOURCE` (docs.philips.com OneDrive 개인 공유 링크, `:x:/g/personal/donghak_lee_philips_com/…`)
- 파일명: `고객안내문 발송현황_Latest.xlsx`
- 시트: IB_raw data, IGT DXR US HPM, MR, CT, (삭제예정)SMAX자료, IGT FCO release date, 자동연동 설정, (삭제예정)매칭 QA 점검, C&R, FCO 매칭, Closed, …
  → BU 컬럼이 없고 시트 이름이 모달리티 묶음. 같은 구조의 시트(MR/CT/IGT…)를 합쳐 읽음. 이름에 "삭제"/"closed"가 들어간 시트는 제외(`SKIP_SHEET_WORDS`, config `skip_sheets`).
- MR 시트 헤더(1행, 일부): C&R#, 장비 ID, 장비코드(6NC/12NC), 장비명, 일련번호(S/N), 고객코드(향후 업데이트), 영문병원명(업데이트예정), 국문병원명, …, 공문(방법 or N), 대리점명(or N), 지역, 비고, 최종증빙 확인, 수신확인서 (Y/N/NA), (최초…일)
- **병원 이름 = `국문병원명` → `고객명(국문)` → `영문병원명`** 순으로 처음 채워진 값(사용자 지정). "-", "N", "0" 같은 값은 빈칸 취급(`row_hospital`, `PLACEHOLDER`).
- **전달 완료 = `최종증빙 확인` 칸에 값이 있음** (사용자 확인). 비어 있으면 미전달(발송일이 있어도 미전달).
  N/X/0/-/미… 는 미전달, NA·N/A·해당없음은 집계 제외. 코드: `FSN_FIELDS`의 status 별칭 맨 앞 + `parse_fsn`의 `evidence` 처리.

### 고객안내문 파일 (일반 규칙 — 자동 인식)
- 시트/헤더/컬럼 자동 인식: `FSN_FIELDS`(정확 일치 별칭 → 포함 단어 순), 위쪽 25행 안에서 가장 많이 맞는 헤더 행, 두 줄 헤더, 병합 셀(세로 병합 값 채움), 등기/직접전달/이메일이 칸으로 나뉜 형태(`WIDE_METHOD_HEADERS`).
- 구조가 다른 요약 시트·숨김 시트는 제외.
- 사용자가 `~/FSN_Dashboard/config.json`으로 바꿀 수 있는 것: `"fsn_url"`, `"fco_urls"`, `"columns"`(필드 → 엑셀 헤더 직접 지정), `"skip_sheets"`.

### FCO 파일
- 기존 FCO 프로그램과 같은 파일: MR / XR / CT/AMI. 컬럼: FCO, Due Date, FCO Title, IP, System Code, System Code Description, SN, Account, Site, Case, WO, Region, FSE, SW Ver, Action Plan, Status, 고객안내문 관련, Remark, PM Cycle, Kit 1.
- URL 3개(`FCO_SOURCES`): `share.philips.com/:x:/r/sites/KRHSFCOstatus/_layouts/15/Doc.aspx?sourcedoc={GUID}&file=…`
  → `add_download_param`이 `…/sites/KRHSFCOstatus/_layouts/15/download.aspx?UniqueId=GUID`로 바꿔서 받음(Doc.aspx에 download=1은 안 됨).
- `Status`가 Closed인 행 제외. 구간: Due Date가 오늘+6개월 미만 빨강, +1년 미만 노랑, 그 외 녹색(기존 FCO 프로그램과 같은 기준).
- 받기 실패 시: 오늘 받은 파일 → 기존 FCO 프로그램 캐시 `~/FCO_Action_Plan/…xlsx`(읽기만) → 이전 파일 순으로 사용.
- 병원 연결: 병원명 정규화(`hospital_key`: 공백·(재)·의료법인 등 제거) + 장비 S/N이 같으면 같은 병원 + FCO Account가 영문이면 영문병원명으로도 연결.

## 자동 다운로드 (EdgeDownloader) — 회사 PC에서 동작 확인 완료
- Edge(Windows 계정 SSO) + Selenium. 프로필 `~/FSN_Dashboard/browser_profile`(잠겨 있으면 `browser_profile_2`로 재시도), 다운로드 폴더 `~/FSN_Dashboard/downloads`.
- 순서:
  ⓪ 파일 화면을 연 뒤 **페이지 안에서 `fetch(다운로드 주소, credentials:'include')`** 로 파일 내용을 직접 받음(`FETCH_JS`, `_fetch_in_page`). 첫 두 바이트가 `PK`인지 확인 후 저장.
     브라우저 다운로드를 쓰지 않아서 Edge의 "이 파일을 열까요? [열기][취소]" 창이 안 뜸(1차 버전에서 사용자가 이 창을 눌러야 받아지던 문제를 이렇게 해결). **2026-10-29 회사 PC에서 확인 완료.**
  ① 다운로드 주소로 브라우저 다운로드(CDP `Browser.setDownloadBehavior allow`로 묻지 않게)
  ② Excel Online `파일 → 복사본 만들기 → 복사본 다운로드`(영/한, iframe 재귀 탐색, 최대 3회)
  ③ 사용자가 직접 받기 3분 대기(고객안내문만)
- 로그인 페이지(login.microsoftonline 등)면 사용자에게 로그인 안내 후 최대 150초 대기.
- 브라우저 다운로드 경로(①~③)의 파일 인식: 고객안내문은 "시작 이후 새로 생긴 xlsx", FCO는 이름에 "FCO" + MR/XR/CT.
- 캐시: `~/FSN_Dashboard/data/FSN_customer_notice.xlsx`, `MR_FCO_Action_Plan.xlsx`, `XR_Open_FCO_list.xlsx`, `CT_AMI_FCO_Action_Plan.xlsx` (+ `.meta.json`에 원래 파일명). 오늘 받은 파일이면 재사용.

### 이미 해결한 문제 (되돌리지 말 것)
- `page_load_strategy="eager"` + 45초 타임아웃, `command_executor.set_timeout(20)` (FCO 때 해결).
- "Please wait... loading the workbook" 배너가 iframe 안에 있음 → iframe 재귀 탐색으로 배너가 사라질 때까지 대기 (FCO 때 해결).
- 파일명 특수문자 제거, 엑셀 불허 제어문자 제거(`clean_text`).
- PyInstaller `--collect-submodules=selenium` 필수.
- "이 파일을 열까요?" 창 → 페이지 안 fetch 방식(위 ⓪).
- 병원명이 "-"/영문으로 나오던 문제 → 국문병원명 우선(위 데이터 사양).
- 진행 창 완료 후 버튼 줄이 창 밖으로 밀리던 문제 → 버튼 줄을 아래에 고정, 창 크기 자동.

## 빌드 (Windows PC에서만)
`build.bat` 더블클릭 → `python -m pip install -r requirements.txt` → `python build.py`
→ PyInstaller로 `dist\FSN_Dashboard.exe`를 만든 뒤 `dist\고객안내문.exe`로 복사(배치 파일에 한글을 쓰면 깨질 수 있어서 파이썬에서 처리).
`dashboard_template.html`, `FSN.ico`는 `--add-data`로 exe 안에 포함됩니다. `build.bat`은 CRLF, ASCII만.

## 현재 상태 / 열린 항목
- 회사 PC 확인 완료(2026-10-29): Windows exe 빌드, Edge 자동 받기(고객안내문 + FCO 3개).
- 검증(Claude 환경): 샘플 파싱(일반 표, 두 줄 헤더+병합+방법 칸 분리, 시트별 FSN, 실제 파일과 비슷한 구조), 대시보드 5개 탭(데스크톱/모바일/다크), 필터·클릭·CSV, Tkinter 진행 창, 리눅스 PyInstaller 빌드.
- 열린 항목:
  - BU별 화면: 사용자에게 BU별 지역 목록을 받으면 '지역' 컬럼 → BU 매핑 후 `SHOW_BU = true`.
  - 전달방법별 화면: 등기/직접전달/이메일 정보가 확정되면 `SHOW_METHOD = true`(어느 컬럼인지 확인 필요).
  - 실제 파일 [데이터 정보] 탭으로 컬럼 매핑·판정 값 재확인.
- 부가 산출물: 발표용 5장 슬라이드(claude.ai 아티팩트 "고객안내문 · FCO 통합 Dashboard 소개" https://claude.ai/artifact/HTUvLJhYF14XQjUuVaRzZU). 5장 대시보드 캡처는 샘플 데이터 화면.

## 시작 방법
위 내용을 읽은 뒤 `main.py`, `dashboard_template.html`과 대조해 차이가 있으면 알려 주고, 사용자에게 오늘 이어서 할 작업이 무엇인지 물어보세요.
