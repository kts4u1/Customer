# -*- coding: utf-8 -*-
"""
고객안내문.exe 빌드 (build.bat에서 호출).  python build.py

PyInstaller 이름은 영문(FSN_Dashboard)으로 빌드한 뒤 dist\\고객안내문.exe 로 복사합니다.
(배치 파일에 한글을 쓰면 코드페이지 문제로 깨질 수 있어서 파이썬에서 처리)
"""
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
NAME = "FSN_Dashboard"
FINAL = "고객안내문.exe"
SEP = ";" if sys.platform == "win32" else ":"


def main():
    for f in ("main.py", "dashboard_template.html", "FSN.ico"):
        if not (ROOT / f).exists():
            print(f"[오류] {f} 파일이 없습니다. build.bat과 같은 폴더에 있어야 합니다.")
            return 1
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onefile", "--noconsole",
        f"--icon={ROOT / 'FSN.ico'}",
        f"--name={NAME}",
        "--collect-submodules=selenium",  # 없으면 Edge WebDriver 모듈이 빠짐 (FCO 때 확인)
        f"--add-data={ROOT / 'dashboard_template.html'}{SEP}.",
        f"--add-data={ROOT / 'FSN.ico'}{SEP}.",
        str(ROOT / "main.py"),
    ]
    print(" ".join(cmd))
    r = subprocess.run(cmd, cwd=ROOT)
    if r.returncode != 0:
        print("[오류] PyInstaller 빌드 실패")
        return r.returncode
    exe = ROOT / "dist" / (NAME + (".exe" if sys.platform == "win32" else ""))
    if not exe.exists():
        print(f"[오류] {exe} 가 만들어지지 않았습니다.")
        return 1
    final = ROOT / "dist" / FINAL
    shutil.copy2(exe, final)
    print(f"\n완료: {final}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
