"""초록멀티 설정값"""
import os
import sys

# 앱 정보
APP_NAME = "초록멀티 v1.4"
APP_VERSION = "1.4.0"
APP_BUILD_DATE = "2026-04-23"
APP_AUTHOR = "임동준"
APP_EMAIL = "d.june0503@gmail.com"
APP_ADMIN_EMAIL = "greenlightsori@gmail.com"
APP_COPYRIGHT = "\u00A9 2026 초록등대 동호회. All rights reserved."

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

# 소리샘 URL
SORISEM_BASE_URL = "https://www.sorisem.net"
LOGIN_URL = f"{SORISEM_BASE_URL}/bbs/login_check.php"
LOGOUT_URL = f"{SORISEM_BASE_URL}/bbs/logout.php"
GREEN_CLUB_MEMBERS_URL = f"{SORISEM_BASE_URL}/plugin/ar.club/admin.member.php?cl=green"

# 다운로드 폴더 설정 파일
DOWNLOAD_DIR_FILE = os.path.join(DATA_DIR, "download_dir.txt")

# 검색 히스토리 파일
SEARCH_HISTORY_FILE = os.path.join(DATA_DIR, "search_history.json")
SEARCH_HISTORY_MAX = 10

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

# 암호화 키 (Fernet 대칭키 생성용 시드)
ENCRYPTION_SALT = b"chorok_multi_2024_salt"
