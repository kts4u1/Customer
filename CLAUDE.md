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
1. 고객안내문 파일 → 2. FCO 파일 (MR · XR · CT/AMI) → 3. 담당자 파일 (PM Planning) → 4. 대시보드 만들기. 단계별로 ✔ / ! / ✖ 표시.
- 버튼: 대시보드 열기 / 다시 받기 / 파일 직접 선택(다운로드 실패 시) / 폴더 열기 / 닫기.
- **실행할 때마다 5개 파일을 모두 새로 받음**(2026-10-31 사용자 요청, "오늘 받은 파일 재사용" 없앰). 받기 실패 시 이전 파일로 진행하고 "!"와 파일 날짜를 표시합니다.
- 완료 문구에 "받기 n초 · 정리 n초" 표시(속도 확인용, debug_log에도 기록).

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
- **병원 이름 = 국문 우선**: 국문병원명 → 고객명(국문)/한글고객명 → 고객주소 (설치장소) 중 **한글이 들어간 첫 값**, 한글이 하나도 없을 때만 영문(`row_hospital`).
  2026-10-31 사용자 캡처: IGT DXR US HPM 시트는 '고객명(국문)'에 영문이, '고객주소 (설치장소)'에 진짜 국문 병원명이 있음 → `hospital_site` 필드 추가.
  영문 이름은 `hen`(FCO Account 연결용, `row_hospital_en`). "-", "N", "0", "폐기", "UNKN/SCRP"는 빈칸 취급(`PLACEHOLDER`).
  병원 표시 이름도 같은 병원의 국문 이름이 있으면 국문으로(FCO만 있는 병원의 영문 Account 대신).
- IGT 시트의 'Modality' 칸은 이메일/N 같은 엉뚱한 값 → BU로 읽지 않음(별칭에서 Modality 제거). 장비 구분은 시트 이름(`r["sheet"]`) 사용.
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
- 받기 실패 시: 이전에 받은 파일 → 기존 FCO 프로그램 캐시 `~/FCO_Action_Plan/…xlsx`(읽기만) 중 최신 파일 사용.
- 병원 연결: 병원명 정규화(`hospital_key`: 공백·(재)·의료법인 등 제거) + 장비 S/N이 같으면 같은 병원 + FCO Account가 영문이면 영문병원명으로도 연결.

### 담당자 파일 (PM Planning) — 2026-10-30 추가 (실제 파일은 못 봄, 사용자 캡처·결과로 맞춰 가는 중)
- 사용자 요청: "Hospital table"과 "FSE" 탭에 담당자 이름이 있으니 이걸로 고객안내문 대시보드의 담당자를 채울 것.
- URL(`PM_SOURCE`): `share.philips.com/:x:/r/sites/HS_CS_RPTeam/_layouts/15/Doc.aspx?sourcedoc={1BF6D0D9-2A63-41B5-B38A-D8E52C403F5E}&file=★MASTER★ 2025 CT MR XR PM Planning.xlsx`
  → FCO와 같이 download.aspx?UniqueId=…로 받음. 캐시 `~/FSN_Dashboard/data/PM_Planning_master.xlsx`. config `pm_url`.
- 파서 `parse_people` (read_only. 이름에 hospital/병원/fse/담당이 들어간 시트 + 앞 30행에 S/N·장비 ID 칸이 있는 장비 시트를 읽음):
  연결 칸(`PEOPLE_KEYS`: 병원명 국문/영문, S/N, 장비 ID, 고객코드, 모달리티, 지역) + 담당자 칸(헤더에 FSE/담당/엔지니어/Sales/영업/이름…,
  또는 헤더가 "MR"/"CT"/"XR"뿐이고 값이 사람 이름처럼 보이는 칸). 헤더의 MR/CT/XR → 그 모달리티 담당. Sales/영업 → 영업.
  FSE 시트(한글 이름 + 영문명/사번)는 별칭표로 써서 영문 이름을 한글 이름으로 바꿈.
- 적용 `apply_people`: 고객안내문 행마다 S/N → 장비 ID → 고객코드 → 병원명 순으로 찾고, 행의 시트(MR/CT/IGT…)와 모달리티가 맞는 FSE + 모달리티 없는 담당자(영업)를 넣음.
  (XR = DXR/IGT로 취급, `MOD_CANON`). 고객안내문 파일에 이미 담당자가 있으면 덮어쓰지 않음. `r["ppsrc"]="pm"` 표시.
- 진행 창 3단계 "담당자 파일 (PM Planning)" 추가 → 대시보드는 4단계. [데이터 정보] 탭에 연결 결과/못 찾은 병원/읽은 칸 표시.
- 2026-10-30 1차 실행 결과: 담당 FSE가 거의 비어 있음("(미지정)" 미전달 1,412건, 이름 붙은 행은 수십 건). "담당없음"이 사람 이름으로 나옴.
  → 보강(실제 원인은 아직 미확인): 이름이 안 맞는 시트도 S/N·장비 ID 칸이 있으면 장비 단위로 사용, 비슷한 병원명(`loose_key`, 담당자 파일에서
  한 병원만 가리킬 때만), 한 칸에 병원 여러 개(`split_hospitals`), S/N 앞자리 0 무시(`sn_norm`), 헤더 없이도 값이 한글 이름인 칸 인식
  (`looks_like_ko_name_column`), 정/부/Main/Sub 헤더, "담당없음" 등 제외(`PERSON_PLACEHOLDER`),
  그래도 없으면 FCO 파일(Closed 포함)의 S/N·병원별 FSE로 보충(`apply_fco_people`, `ppsrc="fco"`).
  [데이터 정보] 탭에 채워진 비율, 못 찾은 병원, 담당자 파일 병원명 예시, 각 시트의 칸 이름 전체를 표시.
- 2026-10-30 2차 실행 결과: 담당 FSE 채워짐 4,328/4,351(99%), 그중 S/N 연결 3,909. 그러나 "담당자 421명" — 장비 시트의
  Customer info(고객 쪽 담당자) 이름까지 FSE로 들어옴. 담당자 파일의 병원명 칸은 "분당서울대 31번 Ing Elition X 3.0T"처럼 장비 별칭이라 이름 연결은 거의 무의미, S/N이 핵심.
  → 수정: **FSE 명단(roster) = FSE 탭의 이름 + FCO 파일 FSE 이름**. FSE 칸 중 명단 이름 비율 50% 미만인 칸은 제외, 행 단위로도 명단 밖 이름은 제외.
  헤더 위 묶음 제목(예: "Customer info")을 오른쪽으로 채워서 고객 쪽 묶음(`CUSTOMER_WORDS`) 칸은 제외.
  [데이터 정보]에 FSE 명단 수, 제외한 칸과 이유, 명단 밖이라 뺀 이름(상위 20) 표시. FCO 파싱을 담당자 파일보다 먼저 함(명단에 FCO 이름 추가).
- **다음 할 일: 실행 후 담당자 수와 "명단에 없어 뺀 이름" 확인. FSE 탭이 명단이 아니면(비율 0% 등) 명단 기준을 다시 정해야 함.**

### 장비 구분 MR · CT · XR (2026-10-31 사용자 요청)
- 사용자: "PM Planning master file에서 MR, CT, XR쪽의 system 탭이 장비 탭" → 이 탭들을 장비 종류의 기준으로 씀.
- `parse_people`: 시트 이름에 system/시스템/장비 + MR/CT/XR(`title_mods`)이 있고 S/N·장비 ID 칸이 있으면 `res["equipment"]`에 S/N/장비 ID → 종류 저장.
  같은 탭의 담당자는 그 종류 담당으로 표시(헤더·모달리티 칸에 표시가 없을 때 시트 이름 사용).
- `assign_modality`: 고객안내문 행마다 `r["mod"]` ∈ MR/CT/XR/기타. 순서: PM 장비 탭 S/N → PM 장비 탭 장비 ID → FCO 파일 S/N →
  한 종류 시트(MR/CT) → 장비명 제품 이름(`SYSTEM_KEYWORDS`: Ingenia→MR, Incisive→CT, Azurion→XR, EPIQ→기타…) → 기타. CT/AMI는 CT, IGT/DXR은 XR.
- FCO 행 `r["mod"]` = 파일 종류. 담당자 고를 때(`row_mod_set`)도 이 종류를 씀.
- 화면: 상단 "장비" 필터(모든 탭), 개요·고객안내문·FCO 탭에 장비별 카드(고정 순서 MR, CT, XR, 기타), 표에 장비 칸, [데이터 정보]에 "장비 구분" 결과
  (장비별 건수, 어떻게 정했는지, PM 장비 탭 목록·대수). FCO 탭의 "소스" 선택칸은 "장비" 필터와 겹쳐 없앰.
- 다음 확인: 사용자 실행 후 [데이터 정보] "장비 구분"에서 PM 장비 탭이 인식됐는지(탭 이름/대수), "못 정함 → 기타"가 얼마나 되는지.

### S/N이 겹쳐 엉뚱한 담당자가 붙던 문제 (2026-10-31)
- 증상: 강릉동인병원에 관계없는 FSE(장재훈)가 붙음. 원인: IGT/XR S/N이 "115", "370"처럼 짧아 다른 병원 장비와 번호가 겹치는데,
  S/N만으로 연결하고 같은 번호의 담당자를 모두 합쳐 넣었음.
- 수정: S/N·장비 ID·고객코드 연결은 행 단위 기록(`res["recs"]`: 병원명·시트·담당자)으로 저장하고 `confirm_by_hospital`로 확인:
  번호 6자 이상(`SHORT_ID`) + 기록 내용이 하나면 인정 / 아니면 병원 이름이 같은 기록만(`same_hospital`) / 병원 칸 없는 기록이 하나뿐이고 장비 종류가 같으면 인정.
  장비 종류가 다른 기록은 먼저 제외. 장비 종류 판정(`assign_modality`), FCO S/N 보충, FCO S/N으로 병원 묶기에도 같은 기준.
- `same_hospital`: 정확/비슷한 이름/포함 + 한글만 비교해 3글자 이상 연속 일치 & 짧은 쪽 60% 이상(PM 장비 탭 병원 칸은 "인천성모 Ing CX 3.0T" 같은 별칭).
- 화면: 표에 "담당자 근거" 칸(`r["ppwhy"]`, 예: "담당자 파일 'XR System' · S/N 370", "FCO 파일 · S/N …", "담당자 파일 · 병원명").

## 자동 다운로드 (EdgeDownloader) — 회사 PC에서 동작 확인 완료
- Edge(Windows 계정 SSO) + Selenium. 프로필 `~/FSN_Dashboard/browser_profile`(잠겨 있으면 `browser_profile_2`로 재시도), 다운로드 폴더 `~/FSN_Dashboard/downloads`.
- **속도(2026-10-31)**: `fetch_many` — 사이트(host)마다 파일 화면을 한 번만 열고(고객안내문 docs.philips.com 1번, FCO 3개+PM share.philips.com 1번),
  그 사이트의 파일들을 페이지 안에서 **동시에** fetch(`FETCH_MANY_START_JS`) → 하나씩 base64로 꺼냄(`FETCH_MANY_GET_JS`). 못 받은 파일만 아래 ①②③으로 하나씩.
- 순서:
  ⓪ 파일 화면을 연 뒤 **페이지 안에서 `fetch(다운로드 주소, credentials:'include')`** 로 파일 내용을 직접 받음(`FETCH_JS`, `_fetch_in_page`). 첫 두 바이트가 `PK`인지 확인 후 저장.
     브라우저 다운로드를 쓰지 않아서 Edge의 "이 파일을 열까요? [열기][취소]" 창이 안 뜸(1차 버전에서 사용자가 이 창을 눌러야 받아지던 문제를 이렇게 해결). **2026-10-29 회사 PC에서 확인 완료.**
  ① 다운로드 주소로 브라우저 다운로드(CDP `Browser.setDownloadBehavior allow`로 묻지 않게)
  ② Excel Online `파일 → 복사본 만들기 → 복사본 다운로드`(영/한, iframe 재귀 탐색, 최대 3회)
  ③ 사용자가 직접 받기 3분 대기(고객안내문만)
- 로그인 페이지(login.microsoftonline 등)면 사용자에게 로그인 안내 후 최대 150초 대기.
- 브라우저 다운로드 경로(①~③)의 파일 인식: 고객안내문은 "시작 이후 새로 생긴 xlsx", FCO는 이름에 "FCO" + MR/XR/CT.
- 저장: `~/FSN_Dashboard/data/FSN_customer_notice.xlsx`, `MR_FCO_Action_Plan.xlsx`, `XR_Open_FCO_list.xlsx`, `CT_AMI_FCO_Action_Plan.xlsx`, `PM_Planning_master.xlsx` (+ `.meta.json`에 원래 파일명). 받기 실패 때만 이전 파일 사용.

## 엑셀 읽기 속도 (2026-10-31)
- 모든 파일을 openpyxl `read_only=True`로 읽음(`sheet_rows`: `reset_dimensions()` 후 값만). 고객안내문은 병합 셀 정보를 zip 안 시트 XML의 `<mergeCell>`에서 직접 읽음(`SheetMerges`, `fill_merges`).
- 시트마다 **앞 30행만 먼저 읽어** 헤더를 찾고, 쓸 시트가 아니면 나머지를 읽지 않음(IB_raw data 같은 큰 시트 건너뜀). 담당자 파일은 이름이 대상이 아닌 시트는 앞 30행에 S/N·장비 ID 칸이 있을 때만 끝까지 읽음.
- 측정(Claude 환경, 10MB·4,500행 + 4만 행 raw 시트 파일): 고객안내문 읽기 17.3초 → 3.4초.
- **주의: read_only 모드에선 행마다 길이가 다름**(마지막으로 채워진 칸까지만). 칸을 읽을 땐 항상 `c < len(row)` 확인.
  2026-10-31 회사 PC에서 `detect_header`의 두 줄 헤더 조합이 다음 행을 같은 길이로 가정해 "list index out of range" → 수정 + `test_ragged_rows`.
- 파일별 읽기는 `safe_parse`로 감쌈: 한 파일에서 오류가 나도 나머지로 대시보드를 만들고 오류는 로그·[데이터 정보]에 남김.

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

## 현재 상태 / 열린 항목 (2026-10-31 기준)
- 회사 PC 확인 완료: Windows exe 빌드, Edge 자동 받기(고객안내문 + FCO 3개 + PM Planning), 속도 개선 후 "속도는 이제 됐다"(사용자).
- 최근 수정 순서(10-31): 매번 새로 받기·동시 받기·빠른 읽기 → "list index out of range" 수정 → MR/CT/XR 장비 구분(PM system 탭)
  → IGT 시트 국문 병원명(고객주소 설치장소) → 짧은 S/N 겹침으로 엉뚱한 FSE가 붙던 문제 수정 + "담당자 근거" 칸.
- 테스트: `python tests/test_parsing.py` (17개, 모두 통과). 사용자 캡처로 확인된 구조(IGT 시트, 짧은 S/N 겹침 등)는 테스트로 재현해 둠.
- 열린 항목:
  - 최신 버전 실행 결과 확인: 강릉동인병원 등 담당자가 맞는지("담당자 근거" 칸), [데이터 정보] "장비 구분"에서 PM system 탭 3개(MR/CT/XR) 인식·대수,
    "못 정함 → 기타" 비율, 담당자 파일 "FSE 명단에 없어 뺀 이름".
  - BU별 화면: 사용자에게 BU별 지역 목록을 받으면 '지역' 컬럼 → BU 매핑 후 `SHOW_BU = true`.
  - 전달방법별 화면: 등기/직접전달/이메일 정보가 확정되면 `SHOW_METHOD = true`(어느 컬럼인지 확인 필요).
- 부가 산출물: 발표용 5장 슬라이드(claude.ai 아티팩트 "고객안내문 · FCO 통합 Dashboard 소개" https://claude.ai/artifact/HTUvLJhYF14XQjUuVaRzZU). 5장 대시보드 캡처는 샘플 데이터 화면.

## 시작 방법
위 내용을 읽은 뒤 `main.py`, `dashboard_template.html`과 대조해 차이가 있으면 알려 주고, 사용자에게 오늘 이어서 할 작업이 무엇인지 물어보세요.
