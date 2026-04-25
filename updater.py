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

from config import UPDATE_API_URL, UPDATE_LIST_API_URL, UPDATE_REPO, DATA_DIR


# 응답 캐시 — 단시간 내 반복 호출을 줄여 GitHub API rate-limit(시간당 60회)
# 에 걸릴 가능성을 낮춘다. 캐시는 단순 JSON 파일로 저장.
_CACHE_TTL_SEC = 600  # 10분
_CACHE_PATH = os.path.join(DATA_DIR, "update_check_cache.json")


def _load_cache(channel: str) -> Optional[dict]:
    """캐시가 유효(TTL 내, 같은 channel)하면 응답 dict 반환."""
    import time
    try:
        if not os.path.exists(_CACHE_PATH):
            return None
        with open(_CACHE_PATH, "r", encoding="utf-8") as f:
            cache = json.load(f)
        if cache.get("channel") != channel:
            return None
        ts = float(cache.get("timestamp", 0))
        if time.time() - ts > _CACHE_TTL_SEC:
            return None
        return cache.get("data")
    except Exception:
        return None


def _save_cache(channel: str, data: dict) -> None:
    """API 응답을 캐시에 저장."""
    import time
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump({
                "channel": channel,
                "timestamp": time.time(),
                "data": data,
            }, f, ensure_ascii=False)
    except OSError:
        pass


def _fetch_via_html(timeout: float = 10.0) -> Optional[dict]:
    """API 가 rate-limit 등으로 막혔을 때 GitHub 릴리스 HTML 페이지에서
    최신 태그를 추출하는 폴백.

    GitHub 의 https://github.com/<owner>/<repo>/releases/latest 는
    `/releases/tag/<tag>` 로 302 리다이렉트하므로 Location 헤더만
    읽으면 태그를 빠르게 알 수 있다. 자산 정보는 /releases/expanded_assets
    HTML 을 추가로 가져와 다운로드 URL 을 추출한다.

    실패하면 None 반환.
    """
    try:
        # 1) /releases/latest 의 리다이렉트 대상에서 태그 얻기
        url = f"https://github.com/{UPDATE_REPO}/releases/latest"
        resp = requests.get(
            url,
            headers={"User-Agent": "greenmulti-updater"},
            timeout=timeout,
            allow_redirects=False,
        )
        location = resp.headers.get("Location", "")
        m = re.search(r"/releases/tag/([^/?#]+)", location or resp.url)
        if not m:
            return None
        tag = m.group(1)
    except requests.RequestException:
        return None

    # 2) expanded_assets HTML 에서 자산 다운로드 URL 추출
    assets: list[dict] = []
    try:
        ea_url = f"https://github.com/{UPDATE_REPO}/releases/expanded_assets/{tag}"
        ea_resp = requests.get(
            ea_url,
            headers={"User-Agent": "greenmulti-updater"},
            timeout=timeout,
        )
        if ea_resp.status_code == 200:
            for m in re.finditer(
                rf'href="(/{re.escape(UPDATE_REPO)}/releases/download/'
                rf'{re.escape(tag)}/([^"]+))"',
                ea_resp.text,
            ):
                href = "https://github.com" + m.group(1)
                name = m.group(2)
                assets.append({
                    "name": name,
                    "browser_download_url": href,
                    "size": 0,
                })
    except requests.RequestException:
        pass

    return {
        "tag_name": tag,
        "name": tag,
        "body": "",
        "html_url": f"https://github.com/{UPDATE_REPO}/releases/tag/{tag}",
        "prerelease": "-" in tag,
        "assets": assets,
    }


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


_LAST_CHECK_ERROR: str = ""


def get_last_check_error() -> str:
    """가장 최근 check_latest_release 실패 원인 한 줄 요약."""
    return _LAST_CHECK_ERROR


def check_latest_release(
    timeout: float = 10.0, channel: str = "stable",
) -> Optional[ReleaseInfo]:
    """GitHub Releases 최신 릴리스 조회.

    channel="beta" 이면 pre-release 를 포함한 모든 릴리스 중 최신 버전을 반환.
    네트워크 실패/비정상 응답은 모두 None (조용히 실패). 실패 원인은
    get_last_check_error() 로 조회 가능.
    """
    global _LAST_CHECK_ERROR
    _LAST_CHECK_ERROR = ""
    url = UPDATE_LIST_API_URL if channel == "beta" else UPDATE_API_URL

    # 1) 캐시(10분 TTL) 우선 — 짧은 시간 내 반복 호출은 GitHub API
    #    시간당 60회(IP 단위) 제한에 걸릴 수 있어 캐시로 회피한다.
    cached = _load_cache(channel)
    if cached is not None:
        data = cached
        return _build_release_info(data)

    try:
        if channel == "beta":
            resp = requests.get(
                url,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "greenmulti-updater",
                },
                params={"per_page": 10},
                timeout=timeout,
            )
            if resp.status_code != 200:
                # rate limit 이거나 일시 장애일 때 HTML 폴백 시도
                if resp.status_code in (403, 429, 502, 503, 504):
                    fallback = _fetch_via_html(timeout=timeout)
                    if fallback is not None:
                        _save_cache(channel, fallback)
                        return _build_release_info(fallback)
                _LAST_CHECK_ERROR = (
                    f"HTTP {resp.status_code} ({url})"
                    + (" - 저장소가 비공개이거나 이름이 바뀌었을 수 있습니다."
                       if resp.status_code == 404 else "")
                    + (" - GitHub API 시간당 호출 한도(rate limit)에 걸렸습니다. "
                       "잠시 후 다시 시도해 주세요."
                       if resp.status_code == 403 else "")
                )
                return None
            releases = resp.json()
            if not isinstance(releases, list) or not releases:
                _LAST_CHECK_ERROR = "릴리스 목록이 비어 있습니다."
                return None
            candidates = [r for r in releases if not r.get("draft")]
            if not candidates:
                _LAST_CHECK_ERROR = "공개된 릴리스(draft 제외)가 없습니다."
                return None
            candidates.sort(
                key=lambda r: _parse_version(str(r.get("tag_name", ""))),
                reverse=True,
            )
            data = candidates[0]
        else:
            resp = requests.get(
                url,
                headers={
                    "Accept": "application/vnd.github+json",
                    "User-Agent": "greenmulti-updater",
                },
                timeout=timeout,
            )
            if resp.status_code != 200:
                if resp.status_code in (403, 429, 502, 503, 504):
                    fallback = _fetch_via_html(timeout=timeout)
                    if fallback is not None:
                        _save_cache(channel, fallback)
                        return _build_release_info(fallback)
                _LAST_CHECK_ERROR = (
                    f"HTTP {resp.status_code} ({url})"
                    + (" - 저장소가 비공개이거나 이름이 바뀌었을 수 있습니다."
                       if resp.status_code == 404 else "")
                    + (" - GitHub API 시간당 호출 한도(rate limit)에 걸렸습니다. "
                       "잠시 후 다시 시도해 주세요."
                       if resp.status_code == 403 else "")
                )
                return None
            data = resp.json()
    except requests.Timeout:
        # 타임아웃·연결 실패 시에도 HTML 폴백을 한 번 시도
        fb = _fetch_via_html(timeout=timeout)
        if fb is not None:
            _save_cache(channel, fb)
            return _build_release_info(fb)
        _LAST_CHECK_ERROR = f"응답 시간 초과 ({int(timeout)}초): {url}"
        return None
    except requests.ConnectionError as e:
        fb = _fetch_via_html(timeout=timeout)
        if fb is not None:
            _save_cache(channel, fb)
            return _build_release_info(fb)
        _LAST_CHECK_ERROR = f"연결 실패: {e.__class__.__name__}"
        return None
    except requests.RequestException as e:
        _LAST_CHECK_ERROR = f"요청 오류: {e.__class__.__name__}: {e}"
        return None
    except ValueError as e:
        _LAST_CHECK_ERROR = f"응답 JSON 파싱 실패: {e}"
        return None

    # 캐시에 저장 후, 자산 정보를 ReleaseInfo 로 정리해 반환.
    _save_cache(channel, data)
    return _build_release_info(data)


def _build_release_info(data: dict) -> Optional[ReleaseInfo]:
    """Release JSON(또는 캐시·HTML 폴백 dict) 한 개에서 ReleaseInfo 구성.

    설치 파일/체크섬/zip/manifest 자산 URL 까지 모두 채운다.
    """
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
    backup_exe_path: Optional[str] = None,
) -> str:
    """구 exe 종료 → staging_dir 내용을 install_dir 에 복사 → 새 exe 실행 하는
    PowerShell 스크립트를 임시로 생성해 경로 반환.

    두 가지 모드:
    - backup_exe_path 가 주어짐 (rename-then-cleanup 모드): 호출자가 이미
      실행 중 exe 를 .old 로 rename 해 둔 상태. PS 는 짧게 대기 → robocopy →
      새 exe 실행 → .old 정리 루프. 경쟁 조건 없음.
    - backup_exe_path 가 None (폴백 모드): rename 이 실패했거나 같은 이름
      업데이트. PS 는 긴 대기 + 구 exe 삭제 재시도 루프 → robocopy → 새 exe.
    """
    script_dir = tempfile.mkdtemp(prefix="chorok_multi_update_")
    script_path = os.path.join(script_dir, "apply_update.ps1")
    log_path = os.path.join(script_dir, "apply_update.log")
    new_exe_path = os.path.join(install_dir, new_exe_name)

    same_name = os.path.normcase(os.path.abspath(old_exe_path)) == os.path.normcase(
        os.path.abspath(new_exe_path)
    )

    def ps_quote(s: str) -> str:
        # 작은따옴표 문자열 안에서 ' 를 '' 로 이스케이프 (PowerShell 리터럴 규칙)
        return "'" + s.replace("'", "''") + "'"

    ps_new = ps_quote(new_exe_path)
    ps_staging = ps_quote(staging_dir)
    ps_install = ps_quote(install_dir)
    ps_log = ps_quote(log_path)
    ps_script_dir = ps_quote(script_dir)

    if backup_exe_path is not None:
        # rename-then-cleanup 모드
        ps_backup = ps_quote(backup_exe_path)
        mode_block = f"""
"[{{0}}] mode: rename-then-cleanup" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
"[{{0}}] backup path: {backup_exe_path}" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
# 부모 프로세스 완전 종료 대기 (rename 이 선행됐으므로 짧게)
Start-Sleep -Seconds 2
""".strip("\n")
        cleanup_block = f"""
# .old 정리 — 설치 폴더 전체의 *.exe.old 대상. 실패해도 치명적 아님.
Get-ChildItem -LiteralPath $install -Filter '*.exe.old' -ErrorAction SilentlyContinue | ForEach-Object {{
    $target = $_.FullName
    for ($i = 0; $i -lt 10; $i++) {{
        try {{
            Remove-Item -LiteralPath $target -Force -ErrorAction Stop
            "[{{0}}] removed .old: $target" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
            break
        }} catch {{
            Start-Sleep -Seconds 1
        }}
    }}
    if (Test-Path -LiteralPath $target) {{
        "[{{0}}] WARNING: .old still present: $target" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
    }}
}}
""".strip("\n")
    else:
        # 폴백 모드 — 기존 동작
        ps_old = ps_quote(old_exe_path)
        remove_block = ""
        if not same_name:
            remove_block = f"""
# 구 exe 삭제 — 잠금 해제까지 최대 15회 × 1초 재시도
$oldExe = {ps_old}
for ($i = 0; $i -lt 15; $i++) {{
    if (-not (Test-Path -LiteralPath $oldExe)) {{ break }}
    try {{
        Remove-Item -LiteralPath $oldExe -Force -ErrorAction Stop
        "[{{0}}] removed old exe: $oldExe" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
        break
    }} catch {{
        Start-Sleep -Seconds 1
    }}
}}
if (Test-Path -LiteralPath $oldExe) {{
    "[{{0}}] WARNING: old exe still exists after retries: $oldExe" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
}}
""".strip("\n")
        mode_block = f"""
"[{{0}}] mode: delete-retry (fallback)" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
# 구 exe 핸들이 풀릴 때까지 잠시 대기
Start-Sleep -Seconds 3

{remove_block}
""".strip("\n")
        cleanup_block = ""

    content = f"""$ErrorActionPreference = 'Continue'
$logPath = {ps_log}
"[{{0}}] apply_update start (PID={{1}})" -f (Get-Date -Format o), $PID | Set-Content -LiteralPath $logPath -Encoding UTF8

{mode_block}

$staging = {ps_staging}
$install = {ps_install}
$newExe = {ps_new}

# staging 내용 로깅 — 압축 해제 결과 검증
"[{{0}}] staging listing:" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
try {{
    Get-ChildItem -LiteralPath $staging -Recurse -ErrorAction Stop | ForEach-Object {{
        "  {{0}}  ({{1}} bytes)" -f $_.FullName, $_.Length | Add-Content -LiteralPath $logPath
    }}
}} catch {{
    "[{{0}}] ERROR listing staging: $_" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
}}

# staging_dir 의 내용을 install_dir 로 복사 (robocopy /E = 하위 폴더 포함)
# Start-Process 대신 call operator 사용 — CREATE_NO_WINDOW 환경에서 더 안정적.
"[{{0}}] robocopy invoke: $staging -> $install" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
$roboOut = & robocopy $staging $install /E /R:3 /W:1 /NFL /NDL /NJH /NJS /NC /NS /NP 2>&1
$roboExit = $LASTEXITCODE
"[{{0}}] robocopy exit=$roboExit" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
if ($roboOut) {{
    "[{{0}}] robocopy output:" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
    $roboOut | Out-String | Add-Content -LiteralPath $logPath
}}
# robocopy 종료 코드: 0-7 정상, 8+ 실패. 단 2+ 은 "추가 파일 있음" 등 정상 상태 포함.
if ($roboExit -ge 8) {{
    "[{{0}}] ERROR: robocopy failed with exit=$roboExit. Staging left for inspection." -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
    exit 1
}}

# install_dir 에 새 exe 가 실제로 생성됐는지 확인
if (Test-Path -LiteralPath $newExe) {{
    $newInfo = Get-Item -LiteralPath $newExe
    "[{{0}}] new exe present: $newExe ($($newInfo.Length) bytes)" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
}} else {{
    "[{{0}}] ERROR: new exe missing after robocopy: $newExe" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
    "[{{0}}] install_dir listing:" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
    try {{
        Get-ChildItem -LiteralPath $install -ErrorAction Stop | ForEach-Object {{
            "  {{0}}" -f $_.FullName | Add-Content -LiteralPath $logPath
        }}
    }} catch {{
        "[{{0}}] ERROR listing install: $_" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
    }}
    exit 2
}}

# 새 exe 실행 — -WorkingDirectory 명시 (PyInstaller onedir 는 cwd 기준 리소스 탐색)
"[{{0}}] launching new exe: $newExe" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
try {{
    Start-Process -FilePath $newExe -WorkingDirectory $install -ErrorAction Stop
    "[{{0}}] launched successfully" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
}} catch {{
    "[{{0}}] ERROR launching new exe: $_" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath
}}

{cleanup_block}

"[{{0}}] apply_update done" -f (Get-Date -Format o) | Add-Content -LiteralPath $logPath

# 자기 자신(임시 스크립트 폴더) 삭제
Start-Sleep -Seconds 1
Remove-Item -LiteralPath {ps_script_dir} -Recurse -Force -ErrorAction SilentlyContinue
"""

    # PowerShell 스크립트는 UTF-8 with BOM 으로 저장해야 한글 안전
    with open(script_path, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(content)
    return script_path
