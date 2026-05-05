"""초록멀티 설정값"""
import os
import sys

# 앱 정보
APP_NAME = "초록멀티 v1.8"
APP_VERSION = "1.8.1"
APP_BUILD_DATE = "2026-05-05"
APP_AUTHOR = "임동준"
APP_EMAIL = "greenlightsori@gmail.com"
APP_ADMIN_EMAIL = "greenlightsori@gmail.com"
APP_COPYRIGHT = "Copyright\u24D2 2026 초록등대 동호회. All rights reserved."

# 경로
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# PyInstaller onedir 빌드에서 번들된 리소스는 _internal(=sys._MEIPASS)에 들어간다.
# 쓰기 전용 파일(credentials, 검색 히스토리 등)은 BASE_DIR에,
# 읽기 전용 번들 리소스(manual.txt 등)는 _RESOURCE_DIR에서 찾는다.
_RESOURCE_DIR = getattr(sys, "_MEIPASS", BASE_DIR)

def resource_path(*relpath: str) -> str:
    """번들된 리소스 파일 경로. _MEIPASS 우선, 없으면 BASE_DIR."""
    bundled = os.path.join(_RESOURCE_DIR, *relpath)
    if os.path.exists(bundled):
        return bundled
    return os.path.join(BASE_DIR, *relpath)

SOUNDS_DIR = os.path.join(BASE_DIR, "sounds")
DATA_DIR = os.path.join(BASE_DIR, "data")
CREDENTIALS_FILE = os.path.join(DATA_DIR, "credentials.ini")
MENU_LIST_FILE = os.path.join(DATA_DIR, "menu_list.json")
# 사용자가 직접 편집할 수 있는 텍스트 메뉴 파일. 존재하면 자동 감지 결과를
# 덮어쓴다. 포맷은 한 줄에 "이름 | URL | 타입" (타입은 선택).
MENU_LIST_TXT_FILE = os.path.join(DATA_DIR, "menu_list.txt")

# v1.7 — 즐겨찾기 / 답장 템플릿 / 게시판 구독
BOOKMARKS_FILE = os.path.join(DATA_DIR, "bookmarks.json")
REPLY_TEMPLATES_FILE = os.path.join(DATA_DIR, "reply_templates.txt")
SUBSCRIPTIONS_FILE = os.path.join(DATA_DIR, "subscriptions.json")

# 소리샘 URL
SORISEM_BASE_URL = "https://www.sorisem.net"
LOGIN_URL = f"{SORISEM_BASE_URL}/bbs/login_check.php"
LOGOUT_URL = f"{SORISEM_BASE_URL}/bbs/logout.php"
GREEN_CLUB_MEMBERS_URL = f"{SORISEM_BASE_URL}/plugin/ar.club/admin.member.php?cl=green"

# 쪽지 (소리샘 전용 ar.memo 플러그인)
MEMO_PLUGIN_BASE = f"{SORISEM_BASE_URL}/plugin/ar.memo"
MEMO_LIST_URL = f"{MEMO_PLUGIN_BASE}/memo.php"
MEMO_VIEW_URL = f"{MEMO_PLUGIN_BASE}/memo_view.php"
MEMO_FORM_URL = f"{MEMO_PLUGIN_BASE}/memo_form.php"
MEMO_FORM_UPDATE_URL = f"{MEMO_PLUGIN_BASE}/memo_form_update.php"
MEMO_DELETE_URL = f"{MEMO_PLUGIN_BASE}/memo_delete.php"
MEMO_LIST_UPDATE_URL = f"{MEMO_PLUGIN_BASE}/memo_list_update.php"
MEMO_CHECK_NEW_URL = f"{MEMO_PLUGIN_BASE}/memo_check_new.php"

# 쪽지 실시간 알림 설정
MEMO_NOTIFY_SETTINGS_FILE = os.path.join(DATA_DIR, "memo_notify_settings.json")
MEMO_NOTIFY_INTERVAL_SEC = 60  # 1분마다 폴링

# 메일(gnuboard5 formmail plugin) — 외부 이메일 발송
MAIL_FORM_URL = f"{SORISEM_BASE_URL}/bbs/formmail.php"
MAIL_SEND_URL = f"{SORISEM_BASE_URL}/bbs/formmail_send.php"

# 메일 수신함 (사이트 내 /message/ 플러그인) — gnuboard5 내부 메시지 시스템
MAIL_INBOX_BASE = f"{SORISEM_BASE_URL}/message"
MAIL_INBOX_URL = f"{MAIL_INBOX_BASE}/inbox.php"
MAIL_SENT_URL = f"{MAIL_INBOX_BASE}/sent.php"
MAIL_INBOX_VIEW_URL = f"{MAIL_INBOX_BASE}/inbox_view.php"
MAIL_SENT_VIEW_URL = f"{MAIL_INBOX_BASE}/sent_view.php"
MAIL_WRITE_URL = f"{MAIL_INBOX_BASE}/write.php"

# 다운로드 폴더 설정 파일
DOWNLOAD_DIR_FILE = os.path.join(DATA_DIR, "download_dir.txt")

# 검색 히스토리 파일
SEARCH_HISTORY_FILE = os.path.join(DATA_DIR, "search_history.json")
SEARCH_HISTORY_MAX = 10

# 자동 업데이트
UPDATE_REPO = "Dongjun-Im/greenmulti"
UPDATE_API_URL = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
UPDATE_LIST_API_URL = f"https://api.github.com/repos/{UPDATE_REPO}/releases"
UPDATE_RELEASES_PAGE = f"https://github.com/{UPDATE_REPO}/releases/latest"
UPDATE_SETTINGS_FILE = os.path.join(DATA_DIR, "update_settings.json")
# 릴리스 채널: "stable"(정식만) / "beta"(pre-release 포함)
UPDATE_CHANNELS = ("stable", "beta")

# 업데이트 자동 확인 주기. 키 → (표시 이름, 간격 시간(0=항상))
UPDATE_INTERVALS = (
    ("always",   "실행할 때마다", 0),
    ("weekly",   "1주에 한 번",    24 * 7),
    ("biweekly", "2주에 한 번",    24 * 14),
    ("monthly",  "1달에 한 번",    24 * 30),
)
UPDATE_INTERVAL_KEYS = tuple(k for k, _, _ in UPDATE_INTERVALS)

def get_download_dir() -> str:
    """다운로드 폴더 경로 반환"""
    if os.path.exists(DOWNLOAD_DIR_FILE):
        with open(DOWNLOAD_DIR_FILE, "r", encoding="utf-8") as f:
            d = f.read().strip()
            if d and os.path.exists(d):
                return d
    default = os.path.join(os.path.expanduser("~"), "Downloads")
    return default if os.path.exists(default) else os.path.expanduser("~")

def set_download_dir(path: str):
    """다운로드 폴더 경로 저장"""
    os.makedirs(os.path.dirname(DOWNLOAD_DIR_FILE), exist_ok=True)
    with open(DOWNLOAD_DIR_FILE, "w", encoding="utf-8") as f:
        f.write(path)

def load_search_history() -> list[dict]:
    """검색 히스토리 로드. 최신이 index 0.

    반환: [{"query": str, "type": str}, ...]
    """
    import json
    if not os.path.exists(SEARCH_HISTORY_FILE):
        return []
    try:
        with open(SEARCH_HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, list):
            return []
        cleaned = []
        for item in data:
            if isinstance(item, dict) and item.get("query"):
                cleaned.append({
                    "query": str(item["query"]),
                    "type": str(item.get("type", "")),
                })
        return cleaned[:SEARCH_HISTORY_MAX]
    except (OSError, ValueError):
        return []

def add_search_history(query: str, type_name: str):
    """검색 히스토리에 항목 추가. 중복 시 맨 앞으로, 최대 SEARCH_HISTORY_MAX개 유지."""
    import json
    query = (query or "").strip()
    if not query:
        return
    history = load_search_history()
    history = [h for h in history if h.get("query") != query]
    history.insert(0, {"query": query, "type": type_name or ""})
    history = history[:SEARCH_HISTORY_MAX]
    try:
        os.makedirs(os.path.dirname(SEARCH_HISTORY_FILE), exist_ok=True)
        with open(SEARCH_HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except OSError:
        pass

def load_update_settings() -> dict:
    """업데이트 설정 로드.

    반환 키:
        check_on_startup (bool): 시작 시 자동 확인 (기본 True)
        skip_version (str): "v1.5.0"처럼 사용자가 건너뛰기 선택한 버전
        last_check_iso (str): 마지막 자동 체크 성공 시각 ISO-8601 ("" 이면 체크한 적 없음)
    """
    import json
    defaults = {
        "check_on_startup": True,
        "skip_version": "",
        "last_check_iso": "",
        "channel": "stable",
        "check_interval": "weekly",
    }
    if not os.path.exists(UPDATE_SETTINGS_FILE):
        return defaults
    try:
        with open(UPDATE_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return defaults
        channel = str(data.get("channel", "stable"))
        if channel not in UPDATE_CHANNELS:
            channel = "stable"
        interval = str(data.get("check_interval", "weekly"))
        if interval not in UPDATE_INTERVAL_KEYS:
            interval = "weekly"
        defaults.update({
            "check_on_startup": bool(data.get("check_on_startup", True)),
            "skip_version": str(data.get("skip_version", "")),
            "last_check_iso": str(data.get("last_check_iso", "")),
            "channel": channel,
            "check_interval": interval,
        })
        return defaults
    except (OSError, ValueError):
        return defaults


def get_update_interval_hours(key: str) -> float:
    """UPDATE_INTERVALS 에서 해당 키의 시간(시간 단위) 반환. 미지의 키는 주 단위."""
    for k, _, hours in UPDATE_INTERVALS:
        if k == key:
            return float(hours)
    return 24.0 * 7

def save_update_settings(settings: dict):
    """업데이트 설정 저장."""
    import json
    try:
        os.makedirs(os.path.dirname(UPDATE_SETTINGS_FILE), exist_ok=True)
        with open(UPDATE_SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except OSError:
        pass

# 암호화 키 (Fernet 대칭키 생성용 시드)
ENCRYPTION_SALT = b"chorok_multi_2024_salt"
