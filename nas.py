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
                   user: str, password: str) -> tuple[bool, str]:
    """지정한 target(URL 또는 UNC)로 net use 1회 시도.
    성공 시 (True, letter), 실패 시 (False, 에러메시지)."""
    try:
        result = _run_net(
            ["use", letter, target, password,
             f"/user:{user}", "/persistent:no"],
            timeout=60,
        )
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
          drive_letter: str | None = None) -> tuple[bool, str]:
    """NAS 마운트. 성공 시 (True, 드라이브문자), 실패 시 (False, 메시지).

    시도 순서:
      1) HTTPS URL 형식: https://host:port
      2) 실패하면 Windows 내부 UNC 형식: \\\\host@SSL@port\\DavWWWRoot
    두 형식 모두 동일한 WebDAV 엔드포인트를 가리키지만, 시스템/DSM 설정에 따라
    한쪽이 오류 1244(ERROR_NOT_AUTHENTICATED) 등으로 실패할 수 있음.
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

    letter = drive_letter or first_free_drive_letter()
    if not letter:
        return False, "사용 가능한 드라이브 문자가 없습니다."

    # 해당 드라이브 문자가 다른 용도로 쓰이면 먼저 해제 시도 (무시 가능)
    try:
        _run_net(["use", letter, "/delete", "/y"], timeout=10)
    except Exception:
        pass

    # 1) HTTPS URL 형식
    ok, info = _attempt_mount(letter, NAS_URL, user, password)
    if ok:
        _mounted_drive = letter
        set_nas_drive_label()
        return True, letter
    first_err = info

    # 2) UNC 폴백 — 드라이브 문자가 선점되어 있을 수 있으니 다시 해제
    try:
        _run_net(["use", letter, "/delete", "/y"], timeout=10)
    except Exception:
        pass
    ok, info = _attempt_mount(letter, NAS_UNC_SSL, user, password)
    if ok:
        _mounted_drive = letter
        set_nas_drive_label()
        return True, letter

    # 두 방법 모두 실패. 더 많은 맥락을 담아서 반환.
    return False, (
        f"HTTPS URL 방식: {first_err}\n"
        f"UNC(SSL) 방식: {info}\n"
        f"- Windows 'WebClient' 서비스가 실행 중인지 확인하세요.\n"
        f"- 아이디/비밀번호가 정확한지 재확인하세요."
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
    """NAS 자격증명 입력 대화상자 → 마운트 → 성공 시 저장.
    반환: (성공여부, 드라이브문자 또는 에러메시지)."""
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

    ok, info = mount(user, pw)
    if ok:
        save_nas_credentials(user, pw)
        if speak_func:
            speak_func("초록등대 자료실에 연결되었습니다.")
    return ok, info
