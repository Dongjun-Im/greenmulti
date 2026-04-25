"""릴리스 자산 파일명을 한글 접두사 포함 이름으로 정규화.

Windows GitHub Actions runner 에서 PowerShell·Inno Setup 이 한글 파일명을
ANSI(cp1252) 로 잘못 처리해 자산 이름에서 "초록멀티" 한글 접두사가 사라지는
문제가 있다. 이를 우회하기 위해 빌드 단계에서는 ASCII 안전한 이름으로
산출물을 만들고, 이 스크립트가 Python 의 os.rename(UTF-16 Win32 API) 로
한글 이름으로 리네임한다.

또 sha256 파일 본문에 들어 있는 ASCII 파일명도 한글 이름으로 교체한다.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys

# Windows runner 에서 stdout/stderr 가 cp1252 라 한글 print 가 깨진다.
# 강제로 UTF-8 로 재구성해 한글이 로그에 그대로 남도록 한다.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _rename(src: str, dst: str) -> None:
    if not os.path.exists(src):
        print(f"  skip (not found): {src}")
        return
    if os.path.exists(dst):
        try:
            os.remove(dst)
        except OSError:
            pass
    shutil.move(src, dst)
    print(f"  renamed: {os.path.basename(src)} -> {os.path.basename(dst)}")


def main() -> int:
    out_dir = "installer_out"
    if not os.path.isdir(out_dir):
        print(f"directory not found: {out_dir}")
        return 1

    # 1) 설치 파일 (.exe) — chorokmulti_v1.6_setup.exe -> 초록멀티_v1.6_setup.exe
    ascii_setup = os.path.join(out_dir, "chorokmulti_v1.6_setup.exe")
    final_setup = os.path.join(out_dir, "초록멀티_v1.6_setup.exe")
    _rename(ascii_setup, final_setup)

    # 2) 포터블 ZIP — 빌드 단계에서 만들어 둔 모든 .zip 후보 처리.
    #    워크플로는 dist 폴더 leaf 명을 그대로 ZIP 이름으로 쓰는데, leaf 명이
    #    인코딩 손상으로 "v1.6" 또는 " v1.6" 같이 떨어지는 경우가 있다.
    target_zip = os.path.join(out_dir, "초록멀티 v1.6.zip")
    if not os.path.exists(target_zip):
        for fname in os.listdir(out_dir):
            if not fname.lower().endswith(".zip"):
                continue
            if fname == os.path.basename(target_zip):
                continue
            _rename(os.path.join(out_dir, fname), target_zip)
            break

    # 3) sha256 파일 — 새 .exe 이름 기준으로 다시 생성 (이름·내용 모두 한글)
    if os.path.exists(final_setup):
        h = hashlib.sha256()
        with open(final_setup, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        sha_path = final_setup + ".sha256"
        with open(sha_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(f"{h.hexdigest().lower()}  {os.path.basename(final_setup)}\n")
        # 기존(ASCII 이름의) sha256 파일이 있으면 정리
        old_sha = ascii_setup + ".sha256"
        if os.path.exists(old_sha):
            try:
                os.remove(old_sha)
                print(f"  removed old: {os.path.basename(old_sha)}")
            except OSError:
                pass
        print(f"  rewrote sha256: {os.path.basename(sha_path)}")

    print()
    print("=== final installer_out contents ===")
    for fname in sorted(os.listdir(out_dir)):
        size = os.path.getsize(os.path.join(out_dir, fname))
        print(f"  {fname}  ({size} bytes)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
