"""초록등대 인증 설정값"""
import os
import sys

# 앱 정보
AUTH_TITLE = "초록등대 회원 인증"

# 경로: 호출하는 프로그램의 위치 기준
if getattr(sys, "frozen", False):
    CALLER_DIR = os.path.dirname(sys.executable)
else:
    CALLER_DIR = os.getcwd()

# 패키지 내부 경로
PACKAGE_DIR = os.path.dirname(os.path.abspath(__file__))
SOUNDS_DIR = os.path.join(PACKAGE_DIR, "sounds")

# 자격 증명 저장 경로: 사용자 AppData에 저장
APP_DATA_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "GreenAuth",
)
CREDENTIALS_FILE = os.path.join(APP_DATA_DIR, "credentials.ini")

# 소리샘 URL
SORISEM_BASE_URL = "https://www.sorisem.net"
LOGIN_URL = f"{SORISEM_BASE_URL}/bbs/login_check.php"
LOGOUT_URL = f"{SORISEM_BASE_URL}/bbs/logout.php"
GREEN_CLUB_MEMBERS_URL = f"{SORISEM_BASE_URL}/plugin/ar.club/admin.member.php?cl=green"

# 암호화 솔트
ENCRYPTION_SALT = b"green_auth_2024_salt"
