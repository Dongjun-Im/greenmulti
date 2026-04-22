"""초록멀티 자동 업데이트.

GitHub Releases API를 조회해 최신 릴리스 정보를 가져오고
현재 실행 중인 버전과 비교한다. 설치 파일을 다운로드하는
기능도 제공한다.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Callable, Optional

import requests

from config import UPDATE_API_URL, UPDATE_LIST_API_URL


class DownloadCancelled(Exception):
    """사용자가 다운로드를 취소했을 때 발생."""


class ChecksumMismatch(Exception):
    """다운로드된 파일의 SHA256 이 릴리스의 체크섬과 다를 때 발생."""


@dataclass
class ReleaseInfo:
    """GitHub Release 정보."""
    version: str            # "1.5.0" (v prefix 제거된 순수 버전 문자열)
    tag_name: str           # "v1.5.0" 원본 태그
    name: str               # 릴리스 제목
    body: str               # 릴리스 노트 본문(markdown)
    html_url: str           # 릴리스 페이지
    installer_url: str = ""  # .exe 설치 파일 다운로드 URL (assets 중)
    installer_name: str = ""
    installer_size: int = 0
    checksum_url: str = ""   # .sha256 체크섬 파일 URL (있으면)
    checksum_name: str = ""
    # 포터블 업데이트용 zip
    zip_url: str = ""
    zip_name: str = ""
    zip_size: int = 0
    # 델타 업데이트용 manifest.json
    manifest_url: str = ""
    prerelease: bool = False
    assets_raw: list = field(default_factory=list)

    @property
    def version_tuple(self) -> tuple[int, ...]:
        return _parse_version(self.version)


def _parse_version(v: str) -> tuple[int, ...]:
    """'v1.5.0', '1.5.0', '1.5.0-beta' 등을 (1, 5, 0) 로 변환.

    비교 불가능한 문자는 0 으로 치환. 비교 편의를 위해 튜플 길이는
    4 자리로 맞춘다.
    """
    v = v.strip().lstrip("vV")
    # 버전 + suffix 분리 (예: "1.5.0-beta.2" → "1.5.0")
    core = re.split(r"[-+]", v, 1)[0]
    parts = []
    for seg in core.split("."):
        m = re.match(r"\d+", seg)
        parts.append(int(m.group()) if m else 0)
    while len(parts) < 4:
        parts.append(0)
    return tuple(parts[:4])


def is_newer(current: str, latest: str) -> bool:
    """latest 가 current 보다 최신 버전인지."""
    return _parse_version(latest) > _parse_version(current)


def _parse_release_json(data: dict) -> Optional[ReleaseInfo]:
    """GitHub Release API 응답 한 개를 ReleaseInfo 로 파싱."""
    tag = str(data.get("tag_name", "")).strip()
    if not tag:
        return None

    info = ReleaseInfo(
        version=tag.lstrip("vV"),
        tag_name=tag,
        name=str(data.get("name", tag)),
        body=str(data.get("body", "") or ""),
        html_url=str(data.get("html_url", "")),
        prerelease=bool(data.get("prerelease", False)),
        assets_raw=list(data.get("assets") or []),
    )
    return info


def check_latest_release(
    timeout: float = 5.0, channel: str = "stable",
) -> Optional[ReleaseInfo]:
    """GitHub Releases 최신 릴리스 조회.

    channel="beta" 이면 pre-release 를 포함한 모든 릴리스 중 최신 버전을 반환.
    네트워크 실패/비정상 응답은 모두 None (조용히 실패).
    """
    try:
        if channel == "beta":
            resp = requests.get(
                UPDATE_LIST_API_URL,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "greenmulti-updater",
                },
                params={"per_page": 10},
                timeout=timeout,
            )
            if resp.status_code != 200:
                return None
            releases = resp.json()
            if not isinstance(releases, list) or not releases:
                return None
            # draft 제외. 버전 기준 내림차순 정렬.
            candidates = [r for r in releases if not r.get("draft")]
            if not candidates:
                return None
            candidates.sort(
                key=lambda r: _parse_version(str(r.get("tag_name", ""))),
                reverse=True,
            )
            data = candidates[0]
        else:
            resp = requests.get(
                UPDATE_API_URL,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "greenmulti-updater",
                },
                timeout=timeout,
            )
            if resp.status_code != 200:
                return None
            data = resp.json()
    except (requests.RequestException, ValueError):
        return None

    info = _parse_release_json(data)
    if info is None:
        return None

    # 설치 파일(.exe) 자산 우선 탐색. 없으면 첫 번째 자산.
    assets = data.get("assets") or []
    exe_asset = None
    for asset in assets:
        name = str(asset.get("name", ""))
        if name.lower().endswith(".exe"):
            exe_asset = asset
            break
    if exe_asset is None and assets:
        exe_asset = assets[0]
    if exe_asset:
        info.installer_url = str(exe_asset.get("browser_download_url", ""))
        info.installer_name = str(exe_asset.get("name", ""))
        try:
            info.installer_size = int(exe_asset.get("size") or 0)
        except (TypeError, ValueError):
            info.installer_size = 0

    # 체크섬 자산 탐색: 설치 파일명 기준 "<name>.sha256" 우선,
    # 없으면 범용 SHA256SUMS / *.sha256 중 첫 번째.
    if exe_asset:
        expected = str(exe_asset.get("name", "")) + ".sha256"
        for asset in assets:
            if str(asset.get("name", "")).lower() == expected.lower():
                info.checksum_url = str(asset.get("browser_download_url", ""))
                info.checksum_name = str(asset.get("name", ""))
                break
    if not info.checksum_url:
        for asset in assets:
            name = str(asset.get("name", "")).lower()
            if name.endswith(".sha256") or name in ("sha256sums", "sha256sums.txt"):
                info.checksum_url = str(asset.get("browser_download_url", ""))
                info.checksum_name = str(asset.get("name", ""))
                break

    # 포터블 업데이트용 ZIP 탐색
    for asset in assets:
        name = str(asset.get("name", ""))
        if name.lower().endswith(".zip"):
            info.zip_url = str(asset.get("browser_download_url", ""))
            info.zip_name = name
            try:
                info.zip_size = int(asset.get("size") or 0)
            except (TypeError, ValueError):
                info.zip_size = 0
            break

    # 델타 업데이트용 manifest.json 탐색
    for asset in assets:
        name = str(asset.get("name", "")).lower()
        if name == "manifest.json":
            info.manifest_url = str(asset.get("browser_download_url", ""))
            break

    return info


def clean_release_notes(markdown: str, max_lines: int = 80) -> str:
    """릴리스 노트 markdown 을 화면 낭독에 적합한 텍스트로 정리.

    - 헤더 #/##/### 기호 제거
    - 리스트 마커 *, + 는 - 로 통일
    - 링크 [텍스트](URL) → 텍스트 (URL 제거는 길이만 늘림)
    - 코드 블록 ```...``` 제거
    - 연속 빈 줄 1줄로 축소
    - 최대 max_lines 줄까지만 유지
    """
    if not markdown:
        return ""
    # 코드 블록 제거
    text = re.sub(r"```[\s\S]*?```", "", markdown)
    # 인라인 코드 `x` → x
    text = re.sub(r"`([^`]+)`", r"\1", text)
    # 이미지 ![alt](url) 제거
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text)
    # 링크 [텍스트](url) → 텍스트
    text = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", text)

    out_lines: list[str] = []
    prev_blank = False
    for raw in text.splitlines():
        line = raw.rstrip()
        # 헤더 마커 제거
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            line = m.group(2)
        # 리스트 마커 통일 (* 또는 + → -)
        line = re.sub(r"^(\s*)[\*\+](\s+)", r"\1-\2", line)

        if not line.strip():
            if prev_blank:
                continue
            prev_blank = True
        else:
            prev_blank = False
        out_lines.append(line)

        if len(out_lines) >= max_lines:
            out_lines.append("...")
            break

    return "\n".join(out_lines).strip()


def sha256_of_file(path: str, chunk_size: int = 65536) -> str:
    """파일의 SHA256 해시(소문자 hex)."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def fetch_expected_checksum(url: str, installer_name: str, timeout: float = 10.0) -> Optional[str]:
    """체크섬 파일을 받아와 설치 파일에 해당하는 SHA256(hex) 반환.

    지원 포맷:
        1) "<64hex>\\n"  — 단일 해시만 있는 .sha256
        2) "<64hex>  filename\\n..."  — SHA256SUMS 형식 (여러 줄)
    매칭되지 않으면 None.
    """
    try:
        resp = requests.get(
            url,
            headers={"User-Agent": "greenmulti-updater"},
            timeout=timeout,
        )
        if resp.status_code != 200:
            return None
        text = resp.text.strip()
    except requests.RequestException:
        return None

    if not text:
        return None

    # 단일 해시만 있는 경우
    first_token = text.split()[0] if text.split() else ""
    if len(first_token) == 64 and re.fullmatch(r"[0-9a-fA-F]{64}", first_token):
        # SHA256SUMS 같은 다중 라인일 가능성도 함께 검사
        if "\n" in text or installer_name in text:
            for line in text.splitlines():
                parts = line.split()
                if len(parts) >= 2 and parts[0].lower() == first_token.lower():
                    name_in_line = parts[-1].lstrip("*")
                    if name_in_line == installer_name:
                        return parts[0].lower()
            for line in text.splitlines():
                parts = line.split()
                if len(parts) >= 2 and re.fullmatch(r"[0-9a-fA-F]{64}", parts[0]):
                    name_in_line = parts[-1].lstrip("*")
                    if name_in_line == installer_name:
                        return parts[0].lower()
        # 단일 해시 파일
        return first_token.lower()

    return None


def get_download_dir() -> str:
    """업데이트 설치 파일을 저장할 임시 디렉터리 (없으면 생성)."""
    path = os.path.join(tempfile.gettempdir(), "chorok_multi_update")
    os.makedirs(path, exist_ok=True)
    return path


def download_installer(
    url: str,
    dest_path: str,
    progress_cb: Optional[Callable[[int, int], bool]] = None,
    chunk_size: int = 65536,
    timeout: float = 30.0,
) -> str:
    """설치 파일을 dest_path 로 다운로드.

    progress_cb(downloaded, total) -> bool 을 매 청크마다 호출.
    False 를 반환하면 다운로드를 취소하고 DownloadCancelled 예외 발생.
    total 이 알려지지 않으면 0 을 넘김.

    반환: 실제 저장된 파일 경로 (dest_path).
    실패 시 requests 예외 또는 OSError 전파.
    """
    tmp_path = dest_path + ".part"

    with requests.get(
        url,
        stream=True,
        timeout=timeout,
        headers={"User-Agent": "greenmulti-updater"},
        allow_redirects=True,
    ) as resp:
        resp.raise_for_status()
        total = int(resp.headers.get("Content-Length") or 0)
        downloaded = 0

        try:
            with open(tmp_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=chunk_size):
                    if not chunk:
                        continue
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_cb is not None:
                        cont = progress_cb(downloaded, total)
                        if cont is False:
                            raise DownloadCancelled()
        except DownloadCancelled:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            raise
        except Exception:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            raise

    # 최종 파일로 이동 (기존 파일 있으면 덮어쓰기)
    if os.path.exists(dest_path):
        try:
            os.remove(dest_path)
        except OSError:
            pass
    os.replace(tmp_path, dest_path)
    return dest_path


# ─────────────────────────────────────────────────────────────
# 포터블 / 델타 업데이트
# ─────────────────────────────────────────────────────────────

def detect_installation_kind() -> str:
    """현재 실행 파일이 설치형인지 포터블인지 판단.

    판단 기준 (frozen 전제):
      - 폴더에 unins000.exe(Inno Setup 생성)가 있으면 "installed"
      - 그 외는 "portable"
    frozen 이 아닌 개발 실행은 "portable" 로 취급.
    """
    if not getattr(sys, "frozen", False):
        return "portable"
    exe_dir = os.path.dirname(sys.executable)
    if os.path.exists(os.path.join(exe_dir, "unins000.exe")):
        return "installed"
    if os.path.exists(os.path.join(exe_dir, "unins000.dat")):
        return "installed"
    return "portable"


def get_install_dir() -> str:
    """현재 실행 폴더 (포터블 업데이트의 대상 디렉터리)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def fetch_manifest(url: str, timeout: float = 10.0) -> Optional[dict]:
    """manifest.json 다운로드 후 파싱.

    기대 포맷:
        {
          "version": "1.5.0",
          "executable": "초록멀티 v1.5.exe",
          "files": [
            {"path": "초록멀티 v1.5.exe", "sha256": "..."},
            {"path": "_internal/base_library.zip", "sha256": "..."},
            ...
          ]
        }
    실패 시 None.
    """
    if not url:
        return None
    try:
        resp = requests.get(
            url,
            headers={"User-Agent": "greenmulti-updater"},
            timeout=timeout,
        )
        if resp.status_code != 200:
            return None
        data = resp.json()
    except (requests.RequestException, ValueError):
        return None
    if not isinstance(data, dict):
        return None
    if not isinstance(data.get("files"), list):
        return None
    return data


def compute_delta(install_dir: str, manifest: dict) -> list[dict]:
    """manifest 와 로컬 파일을 비교해 교체가 필요한 엔트리만 반환.

    반환 원소: {"path": str, "sha256": str}
    """
    out = []
    for entry in manifest.get("files") or []:
        if not isinstance(entry, dict):
            continue
        rel = entry.get("path")
        want_hash = str(entry.get("sha256", "")).lower()
        if not rel or not want_hash:
            continue
        local = os.path.join(install_dir, rel)
        if not os.path.exists(local):
            out.append({"path": rel, "sha256": want_hash})
            continue
        try:
            have = sha256_of_file(local).lower()
        except OSError:
            out.append({"path": rel, "sha256": want_hash})
            continue
        if have != want_hash:
            out.append({"path": rel, "sha256": want_hash})
    return out


def extract_zip(
    zip_path: str,
    dest_dir: str,
    only_paths: Optional[list[str]] = None,
    progress_cb: Optional[Callable[[int, int], bool]] = None,
) -> int:
    """zip 을 dest_dir 로 풀기. only_paths 가 주어지면 그 경로만 추출.

    반환: 추출한 항목 수. 사용자 취소 시 DownloadCancelled.
    """
    # ZIP 의 최상위 폴더(예: "초록멀티 v1.5/") 를 벗겨내고 풀도록 처리.
    # 일반적으로 PyInstaller 결과 폴더 이름으로 감싸서 배포되므로.
    with zipfile.ZipFile(zip_path, "r") as zf:
        names = zf.namelist()
        if not names:
            return 0

        # 최상위 공통 prefix 감지
        first = names[0].split("/")[0]
        has_common = all(
            n == first or n.startswith(first + "/") for n in names if n
        )
        strip_prefix = (first + "/") if has_common else ""

        targets: list[str] = []
        want_norm = None
        if only_paths is not None:
            want_norm = {p.replace("\\", "/").lstrip("/") for p in only_paths}
        for n in names:
            if n.endswith("/"):
                continue
            rel = n[len(strip_prefix):] if strip_prefix and n.startswith(strip_prefix) else n
            if want_norm is not None and rel not in want_norm:
                continue
            targets.append(n)

        total = len(targets)
        extracted = 0
        for member in targets:
            rel = member[len(strip_prefix):] if strip_prefix and member.startswith(strip_prefix) else member
            dest_path = os.path.join(dest_dir, rel.replace("/", os.sep))
            os.makedirs(os.path.dirname(dest_path) or dest_dir, exist_ok=True)
            with zf.open(member) as src, open(dest_path, "wb") as dst:
                shutil.copyfileobj(src, dst)
            extracted += 1
            if progress_cb is not None:
                cont = progress_cb(extracted, total)
                if cont is False:
                    raise DownloadCancelled()
        return extracted


def write_restart_script(
    staging_dir: str,
    install_dir: str,
    new_exe_name: str,
    old_exe_path: str,
) -> str:
    """구 exe 를 종료시키고 staging_dir 의 내용을 install_dir 에 복사한 뒤
    새 exe 를 실행하는 배치 스크립트를 임시로 생성해 경로 반환.

    배치는 실행 후 자동 삭제되도록 자기 자신을 지운다.
    """
    script_dir = tempfile.mkdtemp(prefix="chorok_multi_update_")
    script_path = os.path.join(script_dir, "apply_update.bat")
    new_exe_path = os.path.join(install_dir, new_exe_name)

    # 구 exe 이름이 바뀐 경우(예: "초록멀티 v1.4.exe" → "초록멀티 v1.5.exe")
    # 구 exe 를 삭제해서 폴더에 두 개가 공존하지 않게 한다.
    remove_old_line = ""
    if os.path.normcase(os.path.abspath(old_exe_path)) != os.path.normcase(
        os.path.abspath(new_exe_path)
    ):
        remove_old_line = f'del /f /q "{old_exe_path}" >nul 2>&1\r\n'

    # robocopy 로 staging_dir → install_dir 로 복사. /E 로 하위 폴더 포함.
    # /NFL /NDL /NJH /NJS /NC /NS /NP 로 출력 조용히.
    # robocopy 종료 코드 8 이상이 에러라서 /xo 로만으로도 대부분 정상 종료.
    lines = [
        "@echo off",
        "chcp 65001 >nul",
        # 구 exe 핸들이 풀릴 때까지 잠시 대기
        "ping -n 3 127.0.0.1 >nul",
        remove_old_line.rstrip("\r\n"),
        f'robocopy "{staging_dir}" "{install_dir}" /E /NFL /NDL /NJH /NJS /NC /NS /NP >nul',
        # robocopy 종료 코드가 0-7 은 정상
        "if errorlevel 8 (",
        '    echo 파일 복사 중 오류가 발생했습니다. 스테이징 폴더: "%~dp0"',
        "    pause",
        "    exit /b 1",
        ")",
        f'start "" "{new_exe_path}"',
        # 자기 자신 삭제
        'rmdir /s /q "%~dp0" >nul 2>&1',
    ]
    with open(script_path, "w", encoding="utf-8") as f:
        f.write("\r\n".join(l for l in lines if l is not None) + "\r\n")
    return script_path
