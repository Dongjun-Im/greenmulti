"""초록등대 NAS WebDAV 자동 마운트.

- URL: https://webdav.kbugreenlight.synology.me:5006 (Synology 기본 WebDAV HTTPS 포트)
- Windows 'net use' 로 드라이브 매핑
- NAS 자격증명은 data/nas_credentials.ini 에 Fernet 암호화 저장 (소리샘 credentials 와 동일 패턴)
"""
import base64
import configparser
import hashlib
import os
import platform
import re
import string
import subprocess
from urllib.parse import urlparse

import wx
from cryptography.fernet import Fernet

from config import DATA_DIR, ENCRYPTION_SALT


NAS_URL = "https://webdav.kbugreenlight.synology.me:5006"
NAS_HOSTNAME = "webdav.kbugreenlight.synology.me"
NAS_PORT = 5006

# SSL 신뢰 문제를 호출자가 감지하기 위한 마커. mount() 에러 메시지 앞에 붙음.
SSL_UNTRUSTED_MARKER = "[SSL_UNTRUSTED]"
# Windows WebClient 내부 변환 형식. HTTPS는 @SSL@ 로 표기.
# net use 가 HTTPS URL 직접 입력을 제대로 처리 못하는 경우(오류 1244 등)의
# 폴백. \DavWWWRoot 는 WebDAV 루트 가상 공유 이름.
NAS_UNC_SSL = f"\\\\{NAS_HOSTNAME}@SSL@{NAS_PORT}\\DavWWWRoot"
# 탐색기에 고정 표시할 드라이브 이름
NAS_DRIVE_LABEL = "초록등대 자료실"
NAS_CREDENTIALS_FILE = os.path.join(DATA_DIR, "nas_credentials.ini")

# 이번 프로세스가 마운트한 드라이브 문자 ("Z:" 형식).
# unmount 때 이걸 기본값으로 사용.
_mounted_drive: str | None = None

# 콘솔 창 안 띄우고 net/sc 실행
_CREATE_NO_WINDOW = 0x08000000


# ─────────────────────────── 자격증명 저장 ───────────────────────────

def _get_fernet() -> Fernet:
    seed = (os.getlogin() + os.environ.get("COMPUTERNAME", "default")).encode()
    key_material = hashlib.pbkdf2_hmac("sha256", seed, ENCRYPTION_SALT, 100000)
    key = base64.urlsafe_b64encode(key_material[:32])
    return Fernet(key)


def save_nas_credentials(user_id: str, password: str) -> None:
    fernet = _get_fernet()
    config = configparser.ConfigParser()
    config["nas"] = {
        "user_id": fernet.encrypt(user_id.encode()).decode(),
        "password": fernet.encrypt(password.encode()).decode(),
    }
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(NAS_CREDENTIALS_FILE, "w", encoding="utf-8") as f:
        config.write(f)


def load_nas_credentials() -> tuple[str, str] | None:
    if not os.path.exists(NAS_CREDENTIALS_FILE):
        return None
    try:
        config = configparser.ConfigParser()
        config.read(NAS_CREDENTIALS_FILE, encoding="utf-8")
        if "nas" not in config:
            return None
        fernet = _get_fernet()
        user = fernet.decrypt(config["nas"]["user_id"].encode()).decode()
        pw = fernet.decrypt(config["nas"]["password"].encode()).decode()
        return user, pw
    except Exception:
        return None


def delete_nas_credentials() -> None:
    if os.path.exists(NAS_CREDENTIALS_FILE):
        try:
            os.remove(NAS_CREDENTIALS_FILE)
        except Exception:
            pass


# ─────────────────────────── 드라이브 탐지 / net 래퍼 ───────────────────────────

def _get_used_drive_letters() -> set[str]:
    """현재 사용 중인 드라이브 문자 집합. 'Z:' 형식."""
    try:
        import ctypes
        bits = ctypes.windll.kernel32.GetLogicalDrives()
        used = set()
        for i, letter in enumerate(string.ascii_uppercase):
            if bits & (1 << i):
                used.add(f"{letter}:")
        return used
    except Exception:
        return set()


def first_free_drive_letter() -> str | None:
    """Z부터 역순으로 훑어 첫 번째 빈 드라이브 문자 반환."""
    used = _get_used_drive_letters()
    for letter in reversed(string.ascii_uppercase):
        if letter < "D":
            break  # A/B/C 는 제외
        cand = f"{letter}:"
        if cand not in used:
            return cand
    return None


def _run_net(args: list[str], timeout: int = 30) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["net"] + args,
        capture_output=True, text=True, timeout=timeout,
        encoding="cp949", errors="replace",
        creationflags=_CREATE_NO_WINDOW,
    )


def _run_sc(args: list[str], timeout: int = 15) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["sc"] + args,
        capture_output=True, text=True, timeout=timeout,
        encoding="cp949", errors="replace",
        creationflags=_CREATE_NO_WINDOW,
    )


def _run_cmdkey(args: list[str], timeout: int = 10) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["cmdkey"] + args,
        capture_output=True, text=True, timeout=timeout,
        encoding="cp949", errors="replace",
        creationflags=_CREATE_NO_WINDOW,
    )


def _cmdkey_delete(target: str) -> None:
    """Credential Manager의 기존 target 자격증명 삭제 (실패 무시).
    Windows가 잘못된 캐시 자격증명으로 재시도해 1244 오류를 내는 상황 방지."""
    try:
        _run_cmdkey([f"/delete:{target}"], timeout=5)
    except Exception:
        pass


def _cmdkey_add(target: str, user: str, password: str) -> bool:
    """Credential Manager에 generic 자격증명 사전 등록."""
    try:
        result = _run_cmdkey(
            [f"/add:{target}", f"/user:{user}", f"/pass:{password}"],
        )
        return result.returncode == 0
    except Exception:
        return False


def _stage_webdav_credentials(user: str, password: str) -> None:
    """NAS WebDAV 요청 시 Windows WebClient가 참조할 수 있도록 가능한 모든
    target 형식으로 자격증명 사전 등록. 이전 캐시는 모두 삭제 후 등록."""
    targets = _credential_targets()
    for t in targets:
        _cmdkey_delete(t)
    for t in targets:
        _cmdkey_add(t, user, password)


def _clear_staged_credentials() -> None:
    """마운트 해제/정리 시 캐시된 자격증명 제거. 다음 연결에서 새 자격증명 사용."""
    for t in _credential_targets():
        _cmdkey_delete(t)


def _credential_targets() -> list[str]:
    """WebClient/WinInet가 조회할 만한 모든 target 형식."""
    return [
        NAS_HOSTNAME,
        f"{NAS_HOSTNAME}:{NAS_PORT}",
        f"{NAS_HOSTNAME}@SSL@{NAS_PORT}",
        f"\\\\{NAS_HOSTNAME}@SSL@{NAS_PORT}",
        NAS_URL,
        NAS_URL + "/",
    ]


# ─────────────────────────── 서버 사전 진단 (requests) ───────────────────────────

def _propfind_once(user: str, password: str, verify: bool):
    """내부 헬퍼: 한 번의 PROPFIND 요청. (resp, err) 중 하나 반환."""
    try:
        import requests
        # verify=False 시 urllib3 경고 억제
        if not verify:
            try:
                import urllib3
                urllib3.disable_warnings(
                    urllib3.exceptions.InsecureRequestWarning,
                )
            except Exception:
                pass
        resp = requests.request(
            "PROPFIND", NAS_URL + "/",
            auth=(user, password),
            headers={"Depth": "0", "Content-Length": "0"},
            timeout=15, verify=verify,
        )
        return resp, None
    except Exception as e:
        return None, e


def _classify_resp(resp) -> tuple[bool, str]:
    """HTTP 상태코드별 분석."""
    if resp.status_code in (200, 207, 301, 302):
        return True, f"HTTP {resp.status_code}"
    if resp.status_code == 401:
        www_auth = resp.headers.get("WWW-Authenticate", "(없음)")
        return False, f"인증 거부 (HTTP 401). 서버 요구 방식: {www_auth}"
    if resp.status_code == 403:
        return False, "권한 없음 (HTTP 403). 계정에 WebDAV 권한이 없을 수 있음."
    if resp.status_code == 404:
        return False, "경로 없음 (HTTP 404). WebDAV 서비스가 꺼져 있을 수 있음."
    return False, f"HTTP {resp.status_code}"


def preflight_auth(user: str, password: str) -> tuple[bool, str]:
    """Python 으로 PROPFIND 요청. SSL 검증 실패와 인증 실패를 분리해서 진단.

    결과 해석:
      - (True, "HTTP 207"): 자격증명·서버·SSL 모두 OK
      - (False, "SSL 신뢰 실패 ..."): SSL 문제, 인증 정보는 미확인
      - (True, "SSL 신뢰 실패지만 인증은 성공"): SSL 만 문제 — Windows net use 가
        같은 이유로 실패할 가능성 높음. 자격증명은 정상.
      - (False, "인증 거부 ..."): 아이디/비밀번호 오류
    """
    import requests

    # 1) 정상 경로: SSL 검증 ON
    resp, err = _propfind_once(user, password, verify=True)
    if err is None and resp is not None:
        return _classify_resp(resp)

    ssl_failed = isinstance(err, requests.exceptions.SSLError)

    if ssl_failed:
        # 2) SSL 검증 끄고 다시 시도 → 자격증명만 테스트
        resp2, err2 = _propfind_once(user, password, verify=False)
        if err2 is None and resp2 is not None:
            ok, msg = _classify_resp(resp2)
            if ok:
                # SSL 만 문제인 상황. 이 경우 Windows net use 도 동일 이유로 실패함.
                return True, (
                    f"자격증명 OK ({msg}) 이지만 SSL 인증서를 이 PC가 신뢰하지 "
                    "않습니다. Windows net use 는 같은 이유로 실패할 수 있습니다."
                )
            return False, f"SSL 미검증 재시도: {msg}"
        return False, (
            f"SSL 인증서 신뢰 실패 (원본 오류): {err}. "
            "NAS 인증서를 이 PC의 '신뢰할 수 있는 루트 인증 기관'에 설치해야 합니다."
        )

    # SSL 외 오류
    if isinstance(err, requests.exceptions.ConnectTimeout):
        return False, "연결 시간 초과"
    if isinstance(err, requests.exceptions.ConnectionError):
        return False, f"서버 연결 실패: {err}"
    return False, f"요청 실패: {err}"


# ─────────────────────────── Windows WebClient 레지스트리 튜닝 ───────────────────────────

def try_set_basic_auth_level(level: int = 2) -> bool:
    """HKLM 레지스트리의 BasicAuthLevel 을 조정. 관리자 권한이 필요.
    - 0: Basic 인증 비활성
    - 1: SSL 공유에만 Basic 허용 (기본)
    - 2: SSL·비SSL 모두 허용
    관리자 권한이 없으면 SetValueEx 가 실패하며 False 반환."""
    if platform.system() != "Windows":
        return False
    try:
        import winreg
        key = winreg.CreateKeyEx(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Services\WebClient\Parameters",
            0, winreg.KEY_SET_VALUE,
        )
        try:
            winreg.SetValueEx(
                key, "BasicAuthLevel", 0, winreg.REG_DWORD, level,
            )
        finally:
            winreg.CloseKey(key)
        return True
    except Exception:
        return False


def fetch_server_cert_der() -> bytes | None:
    """NAS 서버의 SSL 인증서(leaf)를 DER 형식 바이트로 가져온다.
    검증을 생략하므로 untrusted 서버에서도 동작."""
    import socket
    import ssl
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        with socket.create_connection((NAS_HOSTNAME, NAS_PORT), timeout=10) as sock:
            with ctx.wrap_socket(sock, server_hostname=NAS_HOSTNAME) as ssock:
                return ssock.getpeercert(binary_form=True)
    except Exception:
        return None


def _install_cert_via_crypto_api(cert_der: bytes,
                                 store_name: str = "ROOT") -> tuple[bool, str]:
    """Windows CryptoAPI 로 인증서를 저장소에 직접 추가.
    certutil 의 '루트가 아닌 인증서 거부' 검증을 우회하므로 leaf 인증서도
    ROOT 저장소에 등록 가능. 현재 사용자 저장소이므로 관리자 권한 불필요."""
    try:
        import ctypes
        from ctypes import wintypes

        crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)

        # Windows 상수
        CERT_STORE_PROV_SYSTEM_W = 10
        CERT_SYSTEM_STORE_CURRENT_USER = 0x00010000
        X509_ASN_ENCODING = 0x00000001
        PKCS_7_ASN_ENCODING = 0x00010000
        CERT_STORE_ADD_REPLACE_EXISTING = 3

        # 함수 시그니처 정의
        crypt32.CertOpenStore.restype = ctypes.c_void_p
        crypt32.CertOpenStore.argtypes = [
            ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p,
            wintypes.DWORD, ctypes.c_wchar_p,
        ]
        crypt32.CertAddEncodedCertificateToStore.restype = wintypes.BOOL
        crypt32.CertAddEncodedCertificateToStore.argtypes = [
            ctypes.c_void_p, wintypes.DWORD, ctypes.c_char_p,
            wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
        ]
        crypt32.CertCloseStore.restype = wintypes.BOOL
        crypt32.CertCloseStore.argtypes = [ctypes.c_void_p, wintypes.DWORD]

        # CERT_STORE_PROV_SYSTEM_W 는 sentinel (포인터 아닌 상수)이므로 캐스팅
        store = crypt32.CertOpenStore(
            ctypes.c_void_p(CERT_STORE_PROV_SYSTEM_W),
            0, None,
            CERT_SYSTEM_STORE_CURRENT_USER,
            store_name,
        )
        if not store:
            err = ctypes.get_last_error()
            return False, f"{store_name} 저장소 열기 실패 (0x{err:08x})"

        try:
            success = crypt32.CertAddEncodedCertificateToStore(
                store,
                X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
                cert_der, len(cert_der),
                CERT_STORE_ADD_REPLACE_EXISTING,
                None,
            )
            if not success:
                err = ctypes.get_last_error()
                return False, f"인증서 추가 실패 (0x{err:08x})"
        finally:
            crypt32.CertCloseStore(store, 0)

        return True, "인증서 설치 완료"
    except Exception as e:
        return False, f"CryptoAPI 호출 오류: {e}"


def install_server_cert_to_trust_store() -> tuple[bool, str]:
    """NAS 서버의 SSL 인증서를 '현재 사용자'의 '신뢰할 수 있는 루트 인증 기관'
    저장소에 추가. CryptoAPI 우선 시도 → 실패 시 certutil 폴백.
    반환: (성공여부, 메시지). 관리자 권한 불필요 (user 저장소)."""
    if platform.system() != "Windows":
        return False, "Windows에서만 지원됩니다."

    cert_der = fetch_server_cert_der()
    if not cert_der:
        return False, "서버 인증서를 가져올 수 없습니다 (연결 실패)."

    # 1) CryptoAPI 직접 호출 — leaf 인증서라도 ROOT 저장소에 추가 가능.
    ok, msg = _install_cert_via_crypto_api(cert_der, store_name="ROOT")
    if ok:
        return True, msg
    crypto_err = msg

    # 2) CryptoAPI 실패 시 certutil 폴백 (self-signed 루트 인증서 대응)
    import tempfile
    try:
        fd, cert_path = tempfile.mkstemp(suffix=".cer", prefix="greenmulti_nas_")
        os.close(fd)
        with open(cert_path, "wb") as f:
            f.write(cert_der)
    except Exception as e:
        return False, f"CryptoAPI: {crypto_err}\n임시 파일 생성 실패: {e}"

    try:
        result = subprocess.run(
            ["certutil", "-user", "-addstore", "ROOT", cert_path],
            capture_output=True, text=True, timeout=60,
            encoding="cp949", errors="replace",
            creationflags=_CREATE_NO_WINDOW,
        )
        if result.returncode == 0:
            return True, "인증서 설치 완료 (certutil)"
        err = (result.stderr or result.stdout or "").strip()
        if not err:
            err = f"certutil 실패 (코드 {result.returncode})"
        return False, f"CryptoAPI: {crypto_err}\ncertutil: {err}"
    except subprocess.TimeoutExpired:
        return False, f"CryptoAPI: {crypto_err}\ncertutil: 시간 초과"
    except Exception as e:
        return False, f"CryptoAPI: {crypto_err}\ncertutil: {e}"
    finally:
        try:
            os.remove(cert_path)
        except Exception:
            pass


def is_server_cert_already_trusted() -> bool:
    """SSL 인증서가 이미 신뢰되는지 확인. verify=True PROPFIND 가 SSLError 없이
    응답(모든 HTTP 상태코드 포함)을 받으면 True."""
    try:
        import requests
        try:
            requests.request(
                "PROPFIND", NAS_URL + "/",
                headers={"Depth": "0", "Content-Length": "0"},
                timeout=10, verify=True,
            )
            return True
        except requests.exceptions.SSLError:
            return False
        except Exception:
            return False
    except Exception:
        return False


def _run_elevated_wait(exe: str, args: str, timeout_sec: int = 90) -> int | None:
    """지정한 exe 를 관리자 권한(UAC)으로 실행하고 종료까지 대기한다.
    종료 코드를 반환, 실패 시 None."""
    import ctypes
    from ctypes import wintypes, byref, sizeof

    class SHELLEXECUTEINFOW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("fMask", wintypes.ULONG),
            ("hwnd", wintypes.HWND),
            ("lpVerb", wintypes.LPCWSTR),
            ("lpFile", wintypes.LPCWSTR),
            ("lpParameters", wintypes.LPCWSTR),
            ("lpDirectory", wintypes.LPCWSTR),
            ("nShow", ctypes.c_int),
            ("hInstApp", wintypes.HINSTANCE),
            ("lpIDList", ctypes.c_void_p),
            ("lpClass", wintypes.LPCWSTR),
            ("hkeyClass", wintypes.HKEY),
            ("dwHotKey", wintypes.DWORD),
            ("hIconOrMonitor", wintypes.HANDLE),
            ("hProcess", wintypes.HANDLE),
        ]

    SEE_MASK_NOCLOSEPROCESS = 0x00000040
    SW_HIDE = 0

    info = SHELLEXECUTEINFOW()
    info.cbSize = sizeof(info)
    info.fMask = SEE_MASK_NOCLOSEPROCESS
    info.lpVerb = "runas"
    info.lpFile = exe
    info.lpParameters = args
    info.nShow = SW_HIDE

    try:
        if not ctypes.windll.shell32.ShellExecuteExW(byref(info)):
            return None
        if not info.hProcess:
            return None
        WAIT_OBJECT_0 = 0
        rc = ctypes.windll.kernel32.WaitForSingleObject(
            info.hProcess, timeout_sec * 1000,
        )
        exit_code = wintypes.DWORD()
        ctypes.windll.kernel32.GetExitCodeProcess(
            info.hProcess, byref(exit_code),
        )
        ctypes.windll.kernel32.CloseHandle(info.hProcess)
        if rc != WAIT_OBJECT_0:
            return None
        return int(exit_code.value)
    except Exception:
        return None


def elevated_fix_webdav_trust() -> tuple[bool, str]:
    """UAC 상승해서 Windows WebClient 가 NAS WebDAV 에 Basic 인증을 쓸 수 있도록
    시스템 전역 설정을 한 번에 적용한다. 수행 작업:
      1) NAS 서버 인증서를 LocalMachine\\Root 저장소에 설치
      2) HKLM\\...\\WebClient\\Parameters\\BasicAuthLevel = 2 (DWord, Force)
      3) AuthForwardServerList 에 NAS URL 여러 형식 추가
      4) WebClient 서비스 Stop + Start (Restart 보다 확실)
      5) 각 단계 결과를 로그 파일에 기록 (문제 진단용)

    관리자 권한 불허(취소) 또는 스크립트 실패면 False 반환. 반환 메시지에는
    로그 파일 경로가 포함됨.
    """
    if platform.system() != "Windows":
        return False, "Windows에서만 지원됩니다."

    cert_der = fetch_server_cert_der()
    if not cert_der:
        return False, "서버 인증서를 가져올 수 없습니다."

    import tempfile
    try:
        fd, cert_path = tempfile.mkstemp(suffix=".cer", prefix="greenmulti_nas_")
        os.close(fd)
        with open(cert_path, "wb") as f:
            f.write(cert_der)
    except Exception as e:
        return False, f"임시 인증서 파일 생성 실패: {e}"

    # 로그 파일 (스크립트가 단계별로 기록)
    log_path = os.path.join(DATA_DIR, "nas_elevated_fix.log")
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except Exception:
        pass

    # AuthForwardServerList 에 넣을 URL 후보들 — 형식 민감성 때문에 여러 형식 등록
    url_variants = [
        NAS_URL,                    # https://host:5006
        NAS_URL + "/",              # https://host:5006/
        f"https://{NAS_HOSTNAME}",  # protocol+host (no port) 보조
    ]
    ps_urls = "@(" + ",".join(f"'{u}'" for u in url_variants) + ")"

    # 로그 경로는 PS 문자열 안전하게 포함 (이스케이프)
    ps_log = log_path.replace("'", "''")
    ps_cert = cert_path.replace("'", "''")

    ps_script = f"""\
$ErrorActionPreference = 'Continue'
$log = '{ps_log}'
"=== greenmulti nas elevated fix ===" | Out-File -FilePath $log -Encoding utf8
"start: $(Get-Date -Format o)" | Out-File -FilePath $log -Append -Encoding utf8
function LogStep($name, $block) {{
    try {{
        & $block
        "OK  $name" | Out-File -FilePath $log -Append -Encoding utf8
    }} catch {{
        "ERR $name : $_" | Out-File -FilePath $log -Append -Encoding utf8
    }}
}}
LogStep 'Import-Certificate' {{
    Import-Certificate -FilePath '{ps_cert}' -CertStoreLocation Cert:\\LocalMachine\\Root | Out-Null
}}
$key = 'HKLM:\\SYSTEM\\CurrentControlSet\\Services\\WebClient\\Parameters'
LogStep 'BasicAuthLevel=2' {{
    New-ItemProperty -Path $key -Name BasicAuthLevel -Value 2 -PropertyType DWord -Force | Out-Null
}}
LogStep 'FileAttributesLimitInBytes' {{
    New-ItemProperty -Path $key -Name FileAttributesLimitInBytes -Value 1000000 -PropertyType DWord -Force | Out-Null
}}
LogStep 'FileSizeLimitInBytes' {{
    New-ItemProperty -Path $key -Name FileSizeLimitInBytes -Value 4294967295 -PropertyType DWord -Force | Out-Null
}}
$urls = {ps_urls}
LogStep 'AuthForwardServerList' {{
    $existing = @((Get-ItemProperty -Path $key -Name AuthForwardServerList -ErrorAction SilentlyContinue).AuthForwardServerList)
    if (-not $existing) {{ $existing = @() }}
    $merged = @($existing)
    foreach ($u in $urls) {{
        if ($merged -notcontains $u) {{ $merged += $u }}
    }}
    # 기존 값이 다른 타입이면 삭제 후 재등록
    Remove-ItemProperty -Path $key -Name AuthForwardServerList -ErrorAction SilentlyContinue
    New-ItemProperty -Path $key -Name AuthForwardServerList -Value $merged -PropertyType MultiString -Force | Out-Null
}}
LogStep 'Stop WebClient' {{
    Stop-Service -Name WebClient -Force -ErrorAction Stop
    Start-Sleep -Seconds 1
}}
LogStep 'Start WebClient' {{
    Start-Service -Name WebClient -ErrorAction Stop
}}
# 최종 상태 기록
$props = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
"BasicAuthLevel final = $($props.BasicAuthLevel)" | Out-File -FilePath $log -Append -Encoding utf8
"AuthForwardServerList final = $($props.AuthForwardServerList -join '|')" | Out-File -FilePath $log -Append -Encoding utf8
$svc = Get-Service -Name WebClient
"WebClient service status = $($svc.Status)" | Out-File -FilePath $log -Append -Encoding utf8
"end: $(Get-Date -Format o)" | Out-File -FilePath $log -Append -Encoding utf8
exit 0
"""

    try:
        fd, ps_path = tempfile.mkstemp(suffix=".ps1", prefix="greenmulti_")
        os.close(fd)
        with open(ps_path, "w", encoding="utf-8-sig") as f:
            f.write(ps_script)
    except Exception as e:
        try:
            os.remove(cert_path)
        except Exception:
            pass
        return False, f"PowerShell 스크립트 생성 실패: {e}"

    ps_args = f'-ExecutionPolicy Bypass -NoProfile -File "{ps_path}"'
    exit_code = _run_elevated_wait("powershell.exe", ps_args, timeout_sec=120)

    # 임시 파일 정리 (로그는 남김)
    for p in (cert_path, ps_path):
        try:
            os.remove(p)
        except Exception:
            pass

    # 로그 읽어 결과 메시지에 포함
    log_content = ""
    try:
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as f:
                log_content = f.read()
    except Exception:
        pass

    if exit_code is None:
        return False, "관리자 권한을 얻지 못했거나 시간이 초과됐습니다."
    if exit_code != 0:
        return False, f"PowerShell 실행 실패 (종료 코드 {exit_code}).\n\n{log_content}"
    return True, f"Windows WebClient 설정 적용 완료\n\n{log_content}"


def restart_webclient() -> bool:
    """WebClient 서비스 재시작. 관리자 권한 필요.
    BasicAuthLevel 변경 후 적용하려면 재시작 필요."""
    if platform.system() != "Windows":
        return False
    try:
        _run_sc(["stop", "WebClient"])
    except Exception:
        pass
    try:
        _run_sc(["start", "WebClient"])
    except Exception:
        return False
    try:
        q = _run_sc(["query", "WebClient"])
        return q.returncode == 0 and "RUNNING" in (q.stdout or "")
    except Exception:
        return False


def ensure_webclient_running() -> bool:
    """Windows WebClient 서비스가 실행 중이 아니면 시작을 시도한다.
    관리자 권한이 없으면 시작에 실패할 수 있으나, 그때는 False 반환."""
    try:
        q = _run_sc(["query", "WebClient"])
        if q.returncode == 0 and "RUNNING" in (q.stdout or ""):
            return True
    except Exception:
        return False
    try:
        _run_sc(["start", "WebClient"])
    except Exception:
        pass
    try:
        q = _run_sc(["query", "WebClient"])
        return q.returncode == 0 and "RUNNING" in (q.stdout or "")
    except Exception:
        return False


def find_existing_mount() -> str | None:
    """이미 NAS가 마운트되어 있다면 드라이브 문자 반환, 아니면 None.
    'net use' 출력에서 NAS 호스트명이 포함된 행을 찾는다."""
    try:
        result = _run_net(["use"], timeout=10)
    except Exception:
        return None
    if result.returncode != 0:
        return None
    for line in (result.stdout or "").splitlines():
        if NAS_HOSTNAME in line:
            m = re.search(r"\b([A-Z]):", line)
            if m:
                return f"{m.group(1)}:"
    return None


# ─────────────────────────── mount / unmount ───────────────────────────

def _attempt_mount(letter: str, target: str,
                   user: str | None, password: str | None) -> tuple[bool, str]:
    """지정한 target(URL 또는 UNC)로 net use 1회 시도.
    user/password가 둘 다 None이면 자격증명 인자 없이 호출 → Windows가
    Credential Manager에 사전 등록된 자격증명을 사용.
    성공 시 (True, letter), 실패 시 (False, 에러메시지)."""
    args = ["use", letter, target]
    if password is not None:
        args.append(password)
    if user is not None:
        args.append(f"/user:{user}")
    args.append("/persistent:no")

    try:
        result = _run_net(args, timeout=60)
    except subprocess.TimeoutExpired:
        return False, "마운트 시도가 시간을 초과했습니다."
    except Exception as e:
        return False, f"net 명령 실행 실패: {e}"
    if result.returncode == 0:
        return True, letter
    err = (result.stderr or result.stdout or "").strip()
    if not err:
        err = f"마운트 실패 (코드 {result.returncode})"
    return False, err


def mount(user: str, password: str,
          drive_letter: str | None = None,
          skip_preflight: bool = False) -> tuple[bool, str]:
    """NAS 마운트. 성공 시 (True, 드라이브문자), 실패 시 (False, 메시지).

    시도 순서:
      1) HTTPS URL 형식: https://host:port
      2) 실패하면 Windows 내부 UNC 형식: \\\\host@SSL@port\\DavWWWRoot

    skip_preflight=True 면 Python requests 기반 사전 진단을 건너뛴다. Python 의
    certifi CA 번들은 Windows 인증서 저장소와 별개라, Windows 에 인증서를 막
    설치한 직후엔 Python 이 여전히 SSL 실패로 보고하므로 재시도 루프를 돌 수
    있기 때문.
    """
    global _mounted_drive
    if platform.system() != "Windows":
        return False, "Windows에서만 지원됩니다."
    if not user or not password:
        return False, "사용자 이름과 비밀번호가 필요합니다."

    # 이미 마운트되어 있다면 그걸 재사용
    existing = find_existing_mount()
    if existing:
        _mounted_drive = existing
        set_nas_drive_label()
        return True, existing

    # WebClient 서비스가 꺼져 있으면 net use WebDAV 가 실패하므로 먼저 시작 시도
    ensure_webclient_running()

    if skip_preflight:
        # Python 진단 스킵 — 바로 Windows net use 로 진행
        pre_ok, pre_msg = True, "skip_preflight"
    else:
        # 1단계: Python 으로 PROPFIND 해서 자격증명/서버/SSL 상태 사전 진단.
        pre_ok, pre_msg = preflight_auth(user, password)

    # SSL 신뢰 실패인데 인증은 성공한 특수 케이스를 구분해 처리
    ssl_only_issue = pre_ok and "SSL 인증서를 이 PC가 신뢰하지 않습니다" in pre_msg

    if not pre_ok:
        return False, (
            f"서버 인증 확인 실패: {pre_msg}\n\n"
            "Python 요청으로 직접 서버에 접속한 결과입니다. 이 단계에서 실패하면 "
            "Windows net use 도 성공할 수 없습니다.\n"
            "- 401 = 아이디/비밀번호 문제\n"
            "- 403 = NAS 계정에 WebDAV 권한 미부여\n"
            "- SSL 오류 = 이 PC가 NAS 인증서를 신뢰하지 않음\n"
            "- 연결 실패 = 방화벽/포트/URL 확인"
        )

    if ssl_only_issue:
        # Windows 도 같은 이유로 실패할 것이므로 net use 시도 자체를 스킵.
        # 호출자(main_frame)가 SSL_UNTRUSTED_MARKER 를 감지해서 자동 설치 플로우를 개시.
        return False, (
            f"{SSL_UNTRUSTED_MARKER}\n"
            "자격증명과 서버는 정상이지만 NAS 의 SSL 인증서를 이 PC가 "
            "신뢰하지 않습니다. Windows 드라이브 매핑도 같은 이유로 실패합니다."
        )

    # 2단계(베스트 에포트): BasicAuthLevel=2 로 설정. 관리자 권한이 있을 때만 성공.
    # 설정이 바뀌면 WebClient 를 재시작해야 적용됨.
    changed = try_set_basic_auth_level(2)
    if changed:
        restart_webclient()

    letter = drive_letter or first_free_drive_letter()
    if not letter:
        return False, "사용 가능한 드라이브 문자가 없습니다."

    # 오류 1244 예방: Windows Credential Manager의 stale 캐시 제거 + 새 자격증명 등록
    _stage_webdav_credentials(user, password)

    # 해당 드라이브 문자가 다른 용도로 쓰이면 먼저 해제 시도 (무시 가능)
    try:
        _run_net(["use", letter, "/delete", "/y"], timeout=10)
    except Exception:
        pass

    errs: list[str] = []

    # 사전 등록된 자격증명 + 인라인 자격증명 두 조합을 두 URL 형식에 교차 시도
    targets = [("HTTPS URL", NAS_URL), ("UNC(SSL)", NAS_UNC_SSL)]
    attempts = [
        # 1) 사전 등록된 자격증명 사용 (인라인 미제공)
        ("credential manager", None, None),
        # 2) 인라인 자격증명 병행 (폴백)
        ("inline", user, password),
    ]

    for target_name, target in targets:
        for attempt_label, u, pw in attempts:
            # 매 시도 전 드라이브 해제
            try:
                _run_net(["use", letter, "/delete", "/y"], timeout=10)
            except Exception:
                pass
            ok, info = _attempt_mount(letter, target, u, pw)
            if ok:
                _mounted_drive = letter
                set_nas_drive_label()
                return True, letter
            errs.append(f"{target_name}({attempt_label}): {info}")

    # 여기까지 왔다면 서버 PROPFIND 는 성공(preflight 통과)했는데 Windows net use
    # 만 실패한 상황 → Windows 측 Basic 인증 설정 문제가 거의 확실.
    admin_hint = ""
    if not changed:
        admin_hint = (
            "\n\n※ 관리자 권한으로 초록멀티를 실행하면 BasicAuthLevel 레지스트리를 "
            "자동 조정해서 이 문제를 해결할 수 있습니다. 또는 관리자 권한 CMD 에서:\n"
            "  reg add \"HKLM\\SYSTEM\\CurrentControlSet\\Services\\WebClient\\"
            "Parameters\" /v BasicAuthLevel /t REG_DWORD /d 2 /f\n"
            "  net stop WebClient & net start WebClient"
        )
    return False, (
        "서버 인증은 성공했으나 Windows 드라이브 매핑이 실패했습니다.\n"
        "Windows WebClient 의 Basic 인증 설정 문제로 보입니다.\n\n"
        + "\n".join(errs)
        + admin_hint
    )


def unmount(drive_letter: str | None = None) -> bool:
    """드라이브 해제. 인수 생략 시 이번 프로세스가 마운트한 드라이브 해제.
    그것도 없으면 NAS 호스트명 기반으로 탐지해 해제."""
    global _mounted_drive
    if platform.system() != "Windows":
        return False
    letter = drive_letter or _mounted_drive or find_existing_mount()
    if not letter:
        return False
    try:
        result = _run_net(["use", letter, "/delete", "/y"], timeout=15)
    except Exception:
        return False
    if result.returncode == 0:
        if letter == _mounted_drive:
            _mounted_drive = None
        return True
    return False


def get_mounted_drive() -> str | None:
    return _mounted_drive or find_existing_mount()


def set_nas_drive_label(label: str = NAS_DRIVE_LABEL) -> bool:
    """탐색기가 표시하는 네트워크 드라이브 이름을 지정한 라벨로 고정.

    Windows는 HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\
    MountPoints2\\<UNC를 ##/#로 치환한 키>\\_LabelFromReg REG_SZ 값을 보면
    네트워크 드라이브의 기본 이름 대신 이 값을 표시한다.
    UNC 경로 기반이므로 드라이브 문자가 달라져도 라벨은 유지된다."""
    if platform.system() != "Windows":
        return False
    try:
        import winreg
        # \\host@SSL@port\share → ##host@SSL@port#share
        mp2_key = NAS_UNC_SSL.replace("\\", "#")
        full_path = (
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
            f"\\{mp2_key}"
        )
        key = winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER, full_path, 0, winreg.KEY_SET_VALUE,
        )
        try:
            winreg.SetValueEx(key, "_LabelFromReg", 0, winreg.REG_SZ, label)
        finally:
            winreg.CloseKey(key)
        # 탐색기에 변경 알림 (열려 있는 창 갱신)
        try:
            import ctypes
            SHCNE_ASSOCCHANGED = 0x08000000
            ctypes.windll.shell32.SHChangeNotify(
                SHCNE_ASSOCCHANGED, 0, None, None,
            )
        except Exception:
            pass
        return True
    except Exception:
        return False


def open_in_explorer(drive_letter: str | None = None) -> bool:
    letter = drive_letter or get_mounted_drive()
    if not letter:
        return False
    try:
        os.startfile(letter + "\\")
        return True
    except Exception:
        return False


def open_nas_url_in_explorer() -> bool:
    """net use 대신 탐색기가 직접 NAS URL 을 WebDAV 로 열게 한다.
    탐색기는 WebClient 서비스와 독립적인 자체 WebDAV 핸들러를 사용하므로
    net use 가 1244/1790/59 로 실패하는 환경에서도 동작하는 경우가 많음."""
    try:
        os.startfile(NAS_URL)
        return True
    except Exception:
        return False


def create_desktop_url_shortcut(
    name: str = "초록등대 자료실",
) -> tuple[bool, str]:
    """바탕화면에 NAS 주소로 가는 인터넷 바로가기(.url)를 생성.
    더블클릭하면 기본 브라우저로 열려 Basic 인증 프롬프트가 뜸 → 로그인 시
    웹 UI 로 파일 브라우징·다운로드 가능. net use 가 막힌 환경의 실질적 대안."""
    if platform.system() != "Windows":
        return False, "Windows에서만 지원됩니다."

    desktop = os.environ.get("USERPROFILE", "")
    if desktop:
        desktop = os.path.join(desktop, "Desktop")
    if not desktop or not os.path.exists(desktop):
        return False, "바탕화면 경로를 찾을 수 없습니다."

    path = os.path.join(desktop, f"{name}.url")
    try:
        with open(path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(
                "[InternetShortcut]\r\n"
                f"URL={NAS_URL}\r\n"
                "IconIndex=0\r\n"
            )
        return True, path
    except Exception as e:
        return False, f"바로가기 생성 실패: {e}"


# ─────────────────────────── 자격증명 입력 대화상자 ───────────────────────────

class NasCredentialsDialog(wx.Dialog):
    """NAS ID/비밀번호 입력 (항상 암호화 저장)."""

    def __init__(self, parent, default_user: str = ""):
        super().__init__(
            parent, title="초록등대 자료실 NAS 로그인",
            style=wx.DEFAULT_DIALOG_STYLE, size=(440, 240),
        )
        panel = wx.Panel(self)

        lbl_info = wx.StaticText(
            panel,
            label=(
                "초록등대 NAS WebDAV에 연결합니다.\n"
                "입력한 아이디와 비밀번호는 이 컴퓨터에 암호화 저장되어\n"
                "다음 실행 시 자동 로그인에 사용됩니다."
            ),
        )

        lbl_user = wx.StaticText(panel, label="NAS 아이디(&U):")
        self.user_ctrl = wx.TextCtrl(panel, value=default_user, name="NAS 아이디")

        lbl_pw = wx.StaticText(panel, label="NAS 비밀번호(&P):")
        self.pw_ctrl = wx.TextCtrl(
            panel, style=wx.TE_PASSWORD, name="NAS 비밀번호",
        )

        ok_btn = wx.Button(panel, wx.ID_OK, "연결(&O)")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "취소(&X)")
        ok_btn.SetDefault()

        # 레이아웃
        form = wx.FlexGridSizer(rows=2, cols=2, hgap=8, vgap=8)
        form.AddGrowableCol(1, 1)
        form.Add(lbl_user, 0, wx.ALIGN_CENTER_VERTICAL)
        form.Add(self.user_ctrl, 1, wx.EXPAND)
        form.Add(lbl_pw, 0, wx.ALIGN_CENTER_VERTICAL)
        form.Add(self.pw_ctrl, 1, wx.EXPAND)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        btn_row.AddStretchSpacer()
        btn_row.Add(ok_btn, 0, wx.RIGHT, 5)
        btn_row.Add(cancel_btn, 0)

        outer = wx.BoxSizer(wx.VERTICAL)
        outer.Add(lbl_info, 0, wx.ALL, 10)
        outer.Add(form, 0, wx.EXPAND | wx.ALL, 10)
        outer.Add(btn_row, 0, wx.EXPAND | wx.ALL, 10)
        panel.SetSizer(outer)

        self.Centre()
        self.user_ctrl.SetFocus()

    def get_credentials(self) -> tuple[str, str]:
        return self.user_ctrl.GetValue().strip(), self.pw_ctrl.GetValue()


def prompt_and_mount(parent, speak_func=None) -> tuple[bool, str]:
    """NAS 자격증명 입력 대화상자 → 마운트.
    반환: (성공여부, 드라이브문자 또는 에러메시지).

    사용자가 대화상자에서 확인을 누른 시점에 **자격증명을 즉시 저장**한다.
    SSL 인증서 설치가 필요한 상황 등 mount() 가 실패해도 재시도 플로우에서
    저장된 자격증명을 다시 읽어 쓸 수 있도록 하기 위함.
    """
    saved = load_nas_credentials()
    default_user = saved[0] if saved else ""
    dlg = NasCredentialsDialog(parent, default_user=default_user)
    try:
        if dlg.ShowModal() != wx.ID_OK:
            return False, "취소되었습니다."
        user, pw = dlg.get_credentials()
    finally:
        dlg.Destroy()

    if not user or not pw:
        return False, "아이디와 비밀번호를 모두 입력해 주세요."

    # 확인 시점에 바로 저장 → mount 실패해도 재시도 시 다시 입력 불필요.
    # 다만 이후 서버가 401(인증 거부)를 반환하면 자격증명이 틀린 것이므로
    # 저장소에서 다시 지운다 (자동 마운트 루프 방지).
    save_nas_credentials(user, pw)

    ok, info = mount(user, pw)
    if ok:
        if speak_func:
            speak_func("초록등대 자료실에 연결되었습니다.")
        return ok, info

    # 실패 시 원인별 처리
    if info and _is_auth_error(info):
        delete_nas_credentials()
    return ok, info


def _is_auth_error(info: str) -> bool:
    """에러 메시지가 자격증명(401) 오류를 가리키는지 판단."""
    return (
        "HTTP 401" in info
        or "인증 거부" in info
        or "401" in info.split("HTTP ", 1)[-1][:10]
    )
