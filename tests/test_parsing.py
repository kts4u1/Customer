# -*- coding: utf-8 -*-
"""
간단 회귀 테스트 (pytest 없이 실행 가능):  python tests/test_parsing.py
샘플 엑셀을 임시 폴더에 만들어 파싱/대시보드 생성까지 확인합니다.
"""
import datetime as dt
import json
import os
import sys
import tempfile
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

import main  # noqa: E402
import make_sample_data  # noqa: E402


def test_classify_status():
    cases = {
        "완료": "done", "전달완료": "done", "O": "done", "o": "done", "Y": "done", "발송": "done", "등기": "done",
        "이메일 발송": "done", "회신완료": "done", "2026-10-01": "done", dt.date(2026, 10, 1): "done",
        "미전달": "todo", "미발송": "todo", "X": "todo", "N": "todo", "진행중": "todo", "예정": "todo",
        "대기": "todo", "반송": "todo", "-": "",  # "-"는 빈칸 취급 → 미전달
        "해당없음": "excluded", "N/A": "excluded", "철거": "excluded",
        "": "", None: "",
        "???": "unknown",
    }
    for v, want in cases.items():
        got = main.classify_status(v)
        assert got == want, f"classify_status({v!r}) = {got!r}, 기대값 {want!r}"


def test_canon_method():
    cases = {
        "등기": "등기", "등기우편": "등기", "우편": "등기", "직접전달": "직접전달", "방문": "직접전달",
        "이메일": "이메일", "E-mail": "이메일", "메일": "이메일", "등기, 이메일": "등기+이메일",
        "Registered mail": "등기", "": "", "미정": "",
    }
    for v, want in cases.items():
        got = main.canon_method(v)
        assert got == want, f"canon_method({v!r}) = {got!r}, 기대값 {want!r}"


def test_to_date():
    assert main.to_date("2026-10-23") == dt.date(2026, 10, 23)
    assert main.to_date("2026.10.3") == dt.date(2026, 10, 3)
    assert main.to_date("2026년 10월 3일") == dt.date(2026, 10, 3)
    assert main.to_date("26.10.03") == dt.date(2026, 10, 3)
    assert main.to_date(46318) == dt.date(2026, 10, 23)  # 엑셀 일련번호
    assert main.to_date("완료") is None
    assert main.to_date(None) is None


def test_hospital_key():
    assert main.hospital_key("(재)가람 의료원") == main.hospital_key("가람의료원")
    assert main.hospital_key("의료법인 누리병원") == main.hospital_key("누리 병원")
    assert main.hospital_key("한빛대학교병원") != main.hospital_key("한빛병원")


def test_match_columns():
    hdr = [(0, "No"), (1, "병원명"), (2, "전달 방법"), (3, "전달여부"), (4, "전달일"), (5, "C&R"), (6, "BU"),
           (7, "담당 FSE"), (8, "영업담당"), (9, "S/N"), (10, "고객안내문 회신여부")]
    m = main.match_columns(hdr, main.FSN_FIELDS)
    want = {"hospital": 1, "method": 2, "status": 3, "date": 4, "cr": 5, "bu": 6, "fse": 7, "sales": 8, "sn": 9,
            "reply": 10}
    for k, v in want.items():
        assert m.get(k) == v, f"{k}: {m.get(k)} != {v}  (전체 {m})"


def _samples(tmp: Path):
    make_sample_data.sample_a(tmp / "A.xlsx")
    make_sample_data.sample_b(tmp / "B.xlsx")
    make_sample_data.fco_file(tmp / "MR FCO Action Plan.xlsx", "MR", ["MR"])


def test_parse_samples():
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        _samples(tmp)
        a = main.parse_fsn(tmp / "A.xlsx")
        assert a["ok"], a["error"]
        assert a["fields"]["hospital"] == "병원명" and a["fields"]["status"] == "전달여부"
        assert [s["name"] for s in a["sheets"] if s["used"]] == ["전달현황"], "요약 시트는 제외되어야 함"
        assert a["excluded"] > 0, "해당없음 행은 집계 제외"
        assert all(r["h"] for r in a["rows"])

        b = main.parse_fsn(tmp / "B.xlsx")
        assert b["ok"], b["error"]
        assert set(b["wide"].values()) == {"등기", "직접전달", "이메일"}, b["wide"]
        assert all(r["h"] for r in b["rows"]), "세로 병합된 병원명이 채워져야 함"
        assert all((r["s"] == 1) == bool(r["m"]) for r in b["rows"]), "전달방법 표시가 있으면 완료"

        today = dt.date(2026, 10, 23)
        f = main.parse_fco(tmp / "MR FCO Action Plan.xlsx", "MR", today)
        assert f["ok"], f["error"]
        assert f["closed"] > 0 and all(r["st"].lower() != "closed" for r in f["rows"])
        assert {r["b"] for r in f["rows"]} <= {"r", "o", "g", "n"}

        data = main.build_dashboard_data({"parse": a, "file": {}},
                                         [{"source": "MR", "parse": f, "file": {}}], today)
        assert data["link"]["matchedHospitals"] > 0
        out = main.render_dashboard(data, tmp / "out.html")
        html = out.read_text(encoding="utf-8")
        assert "__DASHBOARD_DATA__" not in html and "</script>" in html
        start = html.index('id="dashboard-data" type="application/json">') + len('id="dashboard-data" type="application/json">')
        payload = html[start:html.index("</script>", start)]
        assert json.loads(payload)["fsn"]["available"] is True


def test_sheet_per_fsn():
    """FSN별로 시트가 나뉜 경우: 시트 이름을 FSN으로, 요약 시트는 제외"""
    import openpyxl
    with tempfile.TemporaryDirectory() as d:
        path = Path(d) / "multi.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "FSN-001"
        ws.append(["병원명", "BU", "전달여부", "비고"])
        ws.append(["가람병원", "MR", "완료", ""])
        ws.append(["누리병원", "MR", "", "재방문"])
        ws2 = wb.create_sheet("FSN-002")
        ws2.append(["병원명", "BU", "전달여부"])
        ws2.append(["가람병원", "CT", "미전달"])
        ws3 = wb.create_sheet("요약")
        ws3.append(["병원명", "전달여부"])
        ws3.append(["가람병원", "1/2"])
        wb.save(path)
        r = main.parse_fsn(path)
        assert r["ok"], r["error"]
        used = [s["name"] for s in r["sheets"] if s["used"]]
        assert used == ["FSN-001", "FSN-002"], used
        assert sorted({x["fsn"] for x in r["rows"]}) == ["FSN-001", "FSN-002"]
        assert len(r["rows"]) == 3 and sum(x["s"] for x in r["rows"]) == 1


def test_pipeline_manual_file():
    """파일 직접 선택 경로: 다운로드 없이 대시보드 생성"""
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        _samples(tmp)
        old = (main.APP_DIR, main.DATA_DIR, main.DOWNLOAD_DIR, main.LOG_PATH, main.DASHBOARD_PATH,
               main.CONFIG_PATH, main.FCO_APP_DIR)
        try:
            main.APP_DIR = tmp / "app"
            main.DATA_DIR = main.APP_DIR / "data"
            main.DOWNLOAD_DIR = main.APP_DIR / "downloads"
            main.LOG_PATH = main.APP_DIR / "debug_log.txt"
            main.DASHBOARD_PATH = main.APP_DIR / "dash.html"
            main.CONFIG_PATH = main.APP_DIR / "config.json"
            main.FCO_APP_DIR = tmp / "fco_app"
            main.FCO_APP_DIR.mkdir()
            (tmp / "MR FCO Action Plan.xlsx").replace(main.FCO_APP_DIR / "MR_FCO_Action_Plan.xlsx")
            steps = []
            out = main.run_pipeline(False, tmp / "A.xlsx", lambda t: None, lambda i, s, t: steps.append((i, s)),
                                    threading.Event(), {})
            assert out.exists()
            assert (0, "run") in steps and (2, "ok") in steps, steps
            assert (main.DATA_DIR / main.FSN_SOURCE["cache"]).exists()
        finally:
            (main.APP_DIR, main.DATA_DIR, main.DOWNLOAD_DIR, main.LOG_PATH, main.DASHBOARD_PATH,
             main.CONFIG_PATH, main.FCO_APP_DIR) = old


if __name__ == "__main__":
    failed = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print(f"PASS  {name}")
            except Exception as e:  # noqa: BLE001
                failed += 1
                print(f"FAIL  {name}: {e}")
    print("모두 통과" if not failed else f"{failed}개 실패")
    sys.exit(1 if failed else 0)
