"""초록멀티 설정값"""
import os
import sys

# 앱 정보
APP_NAME = "초록멀티 v1.2"
APP_VERSION = "1.2.0"
APP_BUILD_DATE = "2026-04-18"
APP_AUTHOR = "임동준"
APP_EMAIL = "d.june0503@gmail.com"
APP_ADMIN_EMAIL = "greenlightsori@gmail.com"
APP_COPYRIGHT = "\u00A9 2026 초록등대 동호회. All rights reserved."

# 경로
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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

# 암호화 키 (Fernet 대칭키 생성용 시드)
ENCRYPTION_SALT = b"chorok_multi_2024_salt"
