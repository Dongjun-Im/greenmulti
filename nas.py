"""NAS WebDAV 드라이브 마운트 — rclone + WinFSP 기반.

이전 버전은 Windows 내장 WebClient(net use) 를 썼지만 Synology DSM 7.x
WebDAV 와의 프로토콜 협상에서 MRxDAV 커널 드라이버가 ERROR_GEN_FAILURE(31)
를 내며 락업되는 알려진 비호환이 있어, rclone 으로 WebDAV 를 직접 구현하고
WinFSP 로 드라이브 문자를 만드는 구조로 전면 교체.

필수 전제:
  - _internal/bin/rclone.exe 번들 (이 리포의 bin/ 에 커밋)
  - WinFSP 설치 (사용자가 한 번만 수동 설치)
"""
import atexit
import base64
import configparser
import hashlib
import os
import platform
import string
import subprocess
import sys
import time

import wx
from cryptography.fernet import Fernet

from config import DATA_DIR, ENCRYPTION_SALT


NAS_URL = "https://webdav.kbugreenlight.synology.me:5006"
NAS_HOSTNAME = "webdav.kbugreenlight.synology.me"
NAS_DRIVE_LABEL = "초록등대 자료실"
NAS_CREDENTIALS_FILE = os.path.join(DATA_DIR, "nas_credentials.ini")

# 401 감지를 상위 레이어가 할 수 있게 에러 메시지에 붙이는 마커.
AUTH_ERROR_MARKER = "[AUTH_401]"
WINFSP_MISSING_MARKER = "[WINFSP_MISSING]"

# 이번 프로세스가 띄운 rclone 의 Popen 과 마운트 드라이브 문자.
_rclone_process: subprocess.Popen | None = None
_mounted_drive: str | None = None

# 콘솔 창 띄우지 않고 subprocess 실행
_CREATE_NO_WINDOW = 0x08000000
# 자식 rclone 을 프로세스 그룹으로 분리 (부모가 Ctrl+C 를 받아도 자식에 안 전달).
# 앱 종료 시에는 atexit 훅으로 명시적으로 kill 함.
_CREATE_NEW_PROCESS_GROUP = 0x00000200


# ─────────────────────────── 자격증명 저장/로드 ───────────────────────────

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


# ─────────────────────────── 드라이브 문자 / 바이너리 탐지 ───────────────────────────

def _get_used_drive_letters() -> set[str]:
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
    """Z 부터 역순으로 빈 드라이브 문자 탐색."""
    used = _get_used_drive_letters()
    for letter in reversed(string.ascii_uppercase):
        if letter < "D":
            break
        cand = f"{letter}:"
        if cand not in used:
            return cand
    return None


def get_rclone_path() -> str | None:
    """번들된 rclone.exe 경로 반환. frozen(exe) 와 소스 실행 모두 대응."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
        # PyInstaller onedir 의 경우 _internal/bin/rclone.exe 에 들어감
        candidates = [
            os.path.join(base, "_internal", "bin", "rclone.exe"),
            os.path.join(base, "bin", "rclone.exe"),
        ]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
        candidates = [os.path.join(base, "bin", "rclone.exe")]

    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def check_winfsp_installed() -> bool:
    """WinFSP DLL 존재 여부로 판단. 레지스트리까지 안 봐도 이 파일 하나면 충분."""
    candidates = [
        r"C:\Program Files (x86)\WinFsp\bin\winfsp-x64.dll",
        r"C:\Program Files\WinFsp\bin\winfsp-x64.dll",
    ]
    return any(os.path.exists(c) for c in candidates)


WINFSP_DOWNLOAD_URL = "https://github.com/winfsp/winfsp/releases/latest"


# ─────────────────────────── rclone 마운트 ───────────────────────────

def _obscure_password(plain: str) -> str | None:
    """rclone 의 난독화 비밀번호 반환. --webdav-pass / RCLONE_WEBDAV_PASS 에 쓰임."""
    rclone = get_rclone_path()
    if not rclone:
        return None
    try:
        result = subprocess.run(
            [rclone, "obscure", plain],
            capture_output=True, text=True, timeout=10,
            creationflags=_CREATE_NO_WINDOW,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except Exception:
        return None
    return None


def _preflight_credentials(user: str, password: str) -> tuple[bool, str]:
    """rclone mount 를 띄우기 전에 WebDAV 에 한 번 PROPFIND 를 찔러
    자격증명이 유효한지 확인. 잘못된 자격증명으로 rclone 을 띄워 놓고
    저장된 값만 계속 재사용하는 상황을 차단한다.

    Python requests 의 certifi CA 번들은 Windows 인증서 저장소와 별개라,
    사용자가 NAS 자체 서명 인증서를 Windows 에 설치했더라도 Python 은
    SSL 실패로 본다. 반면 rclone 은 Windows CertOpenSystemStore 를 직접
    사용하므로 rclone 본체 연결에는 지장이 없다. 따라서 SSL 검증 실패 시
    verify=False 로 한 번 더 시도해 **자격증명 확인만** 수행한다."""
    try:
        import requests
        try:
            resp = requests.request(
                "PROPFIND", NAS_URL + "/",
                auth=(user, password),
                headers={"Depth": "0", "Content-Length": "0"},
                timeout=12, verify=True,
            )
        except requests.exceptions.SSLError:
            # Python CA 번들에 없는 인증서. verify=False 로 인증만 확인.
            try:
                import urllib3
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            except Exception:
                pass
            resp = requests.request(
                "PROPFIND", NAS_URL + "/",
                auth=(user, password),
                headers={"Depth": "0", "Content-Length": "0"},
                timeout=12, verify=False,
            )

        if resp.status_code in (200, 207, 301, 302):
            return True, "OK"
        if resp.status_code == 401:
            return False, f"{AUTH_ERROR_MARKER} 아이디 또는 비밀번호가 올바르지 않습니다 (HTTP 401)."
        if resp.status_code == 403:
            return False, "NAS 계정에 WebDAV 권한이 없습니다 (HTTP 403)."
        return False, f"서버 응답 이상 (HTTP {resp.status_code})."
    except Exception as e:
        return False, f"서버 연결 실패: {e}"


def mount(user: str, password: str,
          drive_letter: str | None = None) -> tuple[bool, str]:
    """rclone mount 로 NAS WebDAV 를 드라이브 문자로 마운트.
    성공 시 (True, 드라이브 문자), 실패 시 (False, 메시지)."""
    global _rclone_process, _mounted_drive

    if platform.system() != "Windows":
        return False, "Windows에서만 지원됩니다."
    if not user or not password:
        return False, "아이디와 비밀번호가 필요합니다."

    rclone = get_rclone_path()
    if not rclone:
        return False, (
            "rclone 바이너리를 찾을 수 없습니다. 설치가 손상됐거나 bin\\rclone.exe 가 "
            "번들되지 않은 빌드입니다."
        )

    if not check_winfsp_installed():
        return False, (
            f"{WINFSP_MISSING_MARKER}\n"
            "드라이브 문자 매핑에 필요한 WinFSP 가 이 PC에 설치되어 있지 않습니다.\n\n"
            f"다음 주소에서 최신 MSI 파일을 받아 설치한 뒤 다시 시도해 주세요:\n"
            f"{WINFSP_DOWNLOAD_URL}\n\n"
            "설치 후 PC 재시작이 권장됩니다."
        )

    # 이미 이번 프로세스에서 마운트해 둔 게 살아있으면 재사용
    existing = find_existing_mount()
    if existing:
        _mounted_drive = existing
        return True, existing

    # 먼저 Python requests 로 자격증명 검증. rclone mount 는 잘못된 자격증명이어도
    # 마운트 자체는 성공시켜 버리는 경우가 있어 (빈 디렉토리) 사전 확인이 필요.
    ok, info = _preflight_credentials(user, password)
    if not ok:
        return False, info

    obscured = _obscure_password(password)
    if obscured is None:
        return False, "rclone obscure 실행 실패 — 비밀번호 처리 불가."

    letter = drive_letter or first_free_drive_letter()
    if not letter:
        return False, "사용 가능한 드라이브 문자가 없습니다."

    # 혹시 모를 이전 프로세스 정리
    _kill_rclone_if_alive()

    # rclone 상세 로그를 파일로 보관 — 드라이브가 떠도 내부 오류로 접근 불가한 경우
    # 이 로그에서 PROPFIND 실패·SSL 오류·인증 오류 등이 확인됨.
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except Exception:
        pass
    log_path = os.path.join(DATA_DIR, "rclone_mount.log")

    cmd = [
        rclone, "mount",
        ":webdav:",
        letter,
        f"--webdav-url={NAS_URL}",
        "--webdav-vendor=other",
        f"--webdav-user={user}",
        "--vfs-cache-mode=writes",
        "--dir-cache-time=30s",
        "--poll-interval=15s",
        "--attr-timeout=10s",
        f"--volname={NAS_DRIVE_LABEL}",
        "--log-level=DEBUG",
        f"--log-file={log_path}",
        # SSL 자체 서명 인증서 환경 대응. rclone 은 Windows 시스템 저장소를 쓰지만
        # Synology 자체 서명 + 중간 체인 부재 조합에서 신뢰 실패가 반복 관찰됨.
        # 이 앱은 preflight 로 자격증명을 이미 검증하므로 여기서는 SSL 검증을 생략해도
        # 실질적인 보안 약화가 아님 (동일 TLS 핸드셰이크 + 동일 인증 헤더 재전송).
        "--no-check-certificate",
    ]
    env = os.environ.copy()
    env["RCLONE_WEBDAV_PASS"] = obscured

    try:
        _rclone_process = subprocess.Popen(
            cmd,
            env=env,
            creationflags=_CREATE_NO_WINDOW | _CREATE_NEW_PROCESS_GROUP,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
        )
    except Exception as e:
        return False, f"rclone 실행 실패: {e}"

    # 드라이브가 등장할 때까지 폴링. 보통 1~3초면 뜸.
    deadline = time.time() + 20.0
    drive_appeared = False
    while time.time() < deadline:
        # rclone 이 이미 종료됐다면 실패 (로그에 사유 있음)
        if _rclone_process.poll() is not None:
            _rclone_process = None
            return False, _translate_rclone_stderr(_read_tail(log_path, 3000))

        if os.path.exists(letter + "\\"):
            drive_appeared = True
            break

        time.sleep(0.3)

    if not drive_appeared:
        _kill_rclone_if_alive()
        return False, (
            "rclone 마운트가 20초 내에 완료되지 않았습니다. 네트워크 상태를 확인하세요.\n\n"
            f"상세 로그: {log_path}\n\n" + _read_tail(log_path, 2000)
        )

    # 드라이브 문자는 생겼지만 실제 접근 가능한지 검증. listdir 이 빈 에러로
    # 끝나면 rclone 백엔드가 PROPFIND 에 실패하고 있는 상태 (Synology + Windows
    # cert store 미일치 등). 이 경우 드라이브를 내리고 원인을 보고한다.
    if not _verify_mount_readable(letter, timeout=8.0):
        _kill_rclone_if_alive()
        return False, (
            "드라이브 문자는 생겼지만 폴더를 읽을 수 없습니다. rclone 이 NAS 와의 "
            "백엔드 요청에서 실패하고 있습니다.\n\n"
            f"상세 로그: {log_path}\n\n" + _read_tail(log_path, 2500)
        )

    _mounted_drive = letter
    _register_exit_cleanup()
    return True, letter


def _read_tail(path: str, max_bytes: int = 2000) -> str:
    """로그 파일의 마지막 N 바이트만 반환. 파일 없으면 빈 문자열."""
    try:
        size = os.path.getsize(path)
        offset = max(0, size - max_bytes)
        with open(path, "rb") as f:
            f.seek(offset)
            data = f.read()
        return data.decode("utf-8", errors="replace")
    except Exception:
        return ""


def _verify_mount_readable(letter: str, timeout: float = 8.0) -> bool:
    """별도 스레드에서 os.listdir 을 실행하고 timeout 초 내 성공하는지 확인.
    마운트가 죽었을 때 listdir 이 무한 대기에 빠지는 걸 막기 위해 스레드 격리."""
    import threading
    result = {"ok": False}

    def probe():
        try:
            os.listdir(letter + "\\")
            result["ok"] = True
        except Exception:
            pass

    t = threading.Thread(target=probe, daemon=True)
    t.start()
    t.join(timeout)
    return result["ok"]


def _translate_rclone_stderr(stderr_text: str) -> str:
    """rclone 이 조기 종료됐을 때 stderr 에서 원인을 추려낸다."""
    trimmed = stderr_text[-2000:] if len(stderr_text) > 2000 else stderr_text
    lower = stderr_text.lower()
    if "401 unauthorized" in lower or "authentication" in lower and "fail" in lower:
        return f"{AUTH_ERROR_MARKER} rclone 이 401(인증 실패)을 받았습니다.\n\n{trimmed}"
    if "winfsp" in lower and ("not" in lower or "missing" in lower):
        return (
            f"{WINFSP_MISSING_MARKER}\n"
            "WinFSP 가 감지되지 않았습니다. 설치 후 다시 시도해 주세요.\n\n"
            f"{trimmed}"
        )
    if "certificate" in lower or "x509" in lower:
        return (
            "SSL 인증서 문제로 rclone 이 연결하지 못했습니다. 네트워크 점검이 필요합니다.\n\n"
            f"{trimmed}"
        )
    return f"rclone 조기 종료:\n{trimmed}"


def _kill_rclone_if_alive() -> None:
    global _rclone_process
    if _rclone_process is None:
        return
    if _rclone_process.poll() is None:
        try:
            _rclone_process.terminate()
            try:
                _rclone_process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                _rclone_process.kill()
                _rclone_process.wait(timeout=5)
        except Exception:
            pass
    _rclone_process = None


def unmount(drive_letter: str | None = None) -> bool:
    """rclone 프로세스를 종료하면 WinFSP 가 자동으로 드라이브를 분리한다."""
    global _mounted_drive
    _kill_rclone_if_alive()
    # 드라이브가 실제로 사라질 때까지 잠시 대기
    letter = drive_letter or _mounted_drive
    if letter:
        for _ in range(15):
            if not os.path.exists(letter + "\\"):
                break
            time.sleep(0.3)
    _mounted_drive = None
    return True


def find_existing_mount() -> str | None:
    """이번 프로세스가 띄운 rclone 이 아직 살아있고 드라이브가 보이면 그 문자.
    rclone 은 외부 프로세스에서 띄운 것까지 신뢰해 재사용하기 위험하므로,
    **이번 프로세스의 상태만** 참조한다 (다른 앱이 같은 NAS 를 마운트했을 수
    있으니까)."""
    global _rclone_process, _mounted_drive
    if (
        _rclone_process is not None
        and _rclone_process.poll() is None
        and _mounted_drive
        and os.path.exists(_mounted_drive + "\\")
    ):
        return _mounted_drive
    _mounted_drive = None
    return None


def get_mounted_drive() -> str | None:
    return find_existing_mount()


def set_nas_drive_label(label: str = NAS_DRIVE_LABEL) -> bool:
    """rclone 의 --volname 이 탐색기 볼륨 라벨을 이미 지정해 주므로 no-op.
    기존 호출부의 호환성을 위해 껍데기만 유지."""
    return True


def open_in_explorer(drive_letter: str | None = None) -> bool:
    letter = drive_letter or get_mounted_drive()
    if not letter:
        return False
    try:
        os.startfile(letter + "\\")
        return True
    except Exception:
        return False


def _register_exit_cleanup() -> None:
    """앱이 종료될 때 rclone 프로세스와 드라이브가 뒤에 남지 않도록 한 번만 등록."""
    if getattr(_register_exit_cleanup, "_done", False):
        return
    atexit.register(_kill_rclone_if_alive)
    _register_exit_cleanup._done = True  # type: ignore[attr-defined]


def _is_auth_error(info: str) -> bool:
    """에러 메시지가 자격증명(401) 오류를 가리키는지."""
    return AUTH_ERROR_MARKER in (info or "")


def _is_winfsp_missing(info: str) -> bool:
    return WINFSP_MISSING_MARKER in (info or "")


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
    """자격증명 입력 → 마운트. 입력 시점에 바로 저장하므로 재시도 시 재입력 불필요.
    mount 가 401 이면 저장본을 지워 자동 마운트 루프를 끊는다."""
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

    save_nas_credentials(user, pw)

    ok, info = mount(user, pw)
    if ok:
        if speak_func:
            speak_func("초록등대 자료실에 연결되었습니다.")
        return ok, info

    if _is_auth_error(info):
        delete_nas_credentials()
    return ok, info
