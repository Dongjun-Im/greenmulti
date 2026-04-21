"""초록등대 동호회 인증 패키지

다른 프로그램에서 사용 예시:
    from green_auth import run_authentication

    # 인증 성공 시 session 객체 반환, 실패 시 None
    session = run_authentication()
    if session is None:
        sys.exit()
    # session을 사용하여 소리샘 페이지 접근 가능
"""

from green_auth.auth_app import run_authentication
from green_auth.authenticator import Authenticator, AuthResult
from green_auth.credentials import save_credentials, load_credentials, delete_credentials
from green_auth.screen_reader import speak

__version__ = "1.0.0"
__all__ = [
    "run_authentication",
    "Authenticator",
    "AuthResult",
    "save_credentials",
    "load_credentials",
    "delete_credentials",
    "speak",
]
