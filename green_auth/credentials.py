"""자격 증명 저장/로드 모듈 (암호화된 INI 파일)"""
import base64
import configparser
import hashlib
import os

from cryptography.fernet import Fernet

from green_auth.config import CREDENTIALS_FILE, ENCRYPTION_SALT


def _get_fernet() -> Fernet:
    """머신 고유 암호화 키 생성"""
    seed = (os.getlogin() + os.environ.get("COMPUTERNAME", "default")).encode()
    key_material = hashlib.pbkdf2_hmac("sha256", seed, ENCRYPTION_SALT, 100000)
    key = base64.urlsafe_b64encode(key_material[:32])
    return Fernet(key)


def save_credentials(user_id: str, password: str) -> None:
    """아이디와 비밀번호를 암호화하여 INI 파일에 저장"""
    fernet = _get_fernet()

    encrypted_id = fernet.encrypt(user_id.encode()).decode()
    encrypted_pw = fernet.encrypt(password.encode()).decode()

    config = configparser.ConfigParser()
    config["credentials"] = {
        "user_id": encrypted_id,
        "password": encrypted_pw,
    }

    os.makedirs(os.path.dirname(CREDENTIALS_FILE), exist_ok=True)
    with open(CREDENTIALS_FILE, "w", encoding="utf-8") as f:
        config.write(f)


def load_credentials() -> tuple[str, str] | None:
    """INI 파일에서 암호화된 자격 증명을 복호화하여 반환"""
    if not os.path.exists(CREDENTIALS_FILE):
        return None

    config = configparser.ConfigParser()
    config.read(CREDENTIALS_FILE, encoding="utf-8")

    if "credentials" not in config:
        return None

    try:
        fernet = _get_fernet()
        encrypted_id = config["credentials"]["user_id"]
        encrypted_pw = config["credentials"]["password"]

        user_id = fernet.decrypt(encrypted_id.encode()).decode()
        password = fernet.decrypt(encrypted_pw.encode()).decode()
        return user_id, password
    except Exception:
        delete_credentials()
        return None


def delete_credentials() -> None:
    """저장된 자격 증명 파일 삭제"""
    if os.path.exists(CREDENTIALS_FILE):
        os.remove(CREDENTIALS_FILE)
