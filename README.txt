고객안내문 · FCO 통합 Dashboard
=====================================

SharePoint에 있는 "고객안내문(FSN) 전달 현황" 엑셀을 받아서
BU별 / C&R별 / 전달방법(등기·직접전달·이메일)별 전달 상황과
아직 전달이 필요한 병원 목록, 병원별 현황, FCO 현황을 한 화면(브라우저)에서 보여줍니다.


1. 파일 구성
-------------
  main.py                  프로그램 본체
  dashboard_template.html  대시보드 화면 (main.py가 데이터를 넣어서 HTML로 저장)
  FSN.ico                  아이콘
  requirements.txt         필요한 파이썬 패키지
  build.bat / build.py     exe 만들기
  config.example.json      (선택) 설정 예시
  tools/, tests/           개발·테스트용 (빌드에는 필요 없음)


2. exe 만들기 (회사 Windows PC)
--------------------------------
  위 파일들을 한 폴더에 두고 build.bat 을 더블클릭합니다.
  → dist\고객안내문.exe 가 생깁니다.

  직접 명령으로 할 때:
    python -m pip install -r requirements.txt
    python build.py

  수정 후 문법 확인:
    python -m py_compile main.py


3. 사용 방법
-------------
  고객안내문.exe 더블클릭
    ① 오늘 받은 파일이 없으면 Edge가 열리고 SharePoint에서 자동으로 받습니다.
       (공유 링크 바로 받기 → 안 되면 파일 → 복사본 만들기 → 복사본 다운로드 메뉴 클릭)
    ② 자동으로 안 되면 열린 Edge 창에서 직접 [복사본 다운로드]를 누르면 3분 동안 기다렸다가 인식합니다.
    ③ 대시보드가 Edge에서 열립니다.

  버튼
    대시보드 열기   마지막으로 만든 대시보드를 다시 엽니다.
    새로 받기       오늘 받은 파일이 있어도 SharePoint에서 다시 받습니다.
    파일 직접 선택  내려받아 둔 엑셀 파일을 골라서 대시보드를 만듭니다. (다운로드가 안 될 때)
    폴더 열기       데이터/로그 폴더(%USERPROFILE%\FSN_Dashboard)를 엽니다.


4. 대시보드 화면
-----------------
  맨 위 필터(담당자 FSE·영업 / BU / 검색)는 모든 탭에 적용됩니다.

  통합 개요        고객안내문 전달률, 미전달 건수, FCO 마감 구간별 건수, 우선 확인 병원
  고객안내문(FSN)  BU별 · C&R별 · 전달방법별 · 담당자별 전달 막대(파랑=완료, 빨강=미전달)
                   아래 표: 아직 전달이 필요한 병원 (병원별/건별, 엑셀(CSV) 저장)
  FCO              MR · XR · CT/AMI 남은 FCO (빨강 6개월 미만 / 노랑 1년 미만 / 녹색 1년 이상)
  병원별 현황      세일즈·FSE용. 병원을 고르면 그 병원의 고객안내문 + FCO를 한 번에 봅니다.
  데이터 정보      엑셀의 어떤 시트·컬럼을 읽었는지. 숫자가 이상하면 이 화면을 캡처해서 보내 주세요.

  막대나 숫자 카드를 누르면 해당 조건으로 필터가 걸립니다. (다시 누르면 해제)
  만들어진 HTML 파일(FSN_FCO_Dashboard.html)은 그대로 메일로 보내도 열립니다.


5. 담당자 데이터 (PM Planning)
-------------------------------
  ★MASTER★ 2025 CT MR XR PM Planning.xlsx 의 "Hospital table" / "FSE" 시트에서
  병원·장비별 담당 FSE / 영업 이름을 읽어 고객안내문 행에 넣습니다.
  연결 순서: S/N → 장비 ID → 고객코드 → 병원명(국문/영문). FSE 시트의 영문명·사번은 한글 이름으로 바꿉니다.
  결과는 대시보드 [데이터 정보] 탭 "담당자 파일" 부분에서 확인합니다.


6. FCO 데이터
--------------
  FCO 주소가 설정되지 않은 경우, 기존 FCO 프로그램(FCO.exe)이 받아 둔 파일
  (%USERPROFILE%\FCO_Action_Plan\*.xlsx)을 자동으로 같이 씁니다.
  오늘 받은 파일이 아니면 대시보드 위쪽에 "이전 파일" 안내가 나옵니다.


7. 설정 파일 (선택)
--------------------
  %USERPROFILE%\FSN_Dashboard\config.json 을 만들면 다시 빌드하지 않고 바꿀 수 있습니다.
  (config.example.json 참고)

    "fsn_url"   고객안내문 파일 주소
    "fco_urls"  FCO 파일 주소 (MR / XR / CT/AMI)
    "pm_url"    담당자 파일(PM Planning) 주소
    "skip_sheets" 읽지 않을 시트 이름에 들어간 단어 (기본: 삭제, closed)
    "columns"   컬럼 이름을 직접 지정 (자동 인식이 틀릴 때)
                예: {"hospital": "병원명", "status": "전달여부", "method": "전달방법"}
                항목 이름: hospital, bu, cr, method, status, date, fsn, fsn_title, fse, sales,
                           sn, system, site, region, reply, remark, fco


8. 문제 해결
-------------
  - 오류가 나면 메시지 창이 뜨고 %USERPROFILE%\FSN_Dashboard\debug_log.txt 에 기록됩니다.
  - 전달 완료/미전달 숫자가 이상하면 대시보드 [데이터 정보] 탭을 캡처해서 보내 주세요.
