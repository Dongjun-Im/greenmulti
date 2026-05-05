"""인증 실행 모듈 - 다른 프로그램에서 호출하는 진입점"""
import threading
import time

import wx

from green_auth.authenticator import Authenticator, AuthResult
from green_auth.credentials import save_credentials, load_credentials, delete_credentials
from green_auth.login_dialog import LoginDialog
from green_auth.progress import ProgressIndicator
from green_auth.screen_reader import speak


def _play_auth_success_beep():
    """인증 성공 비프음"""
    try:
        import winsound
        winsound.Beep(1000, 300)
    except Exception:
        pass


def _josa_eul_reul(name: str) -> str:
    """이름 끝 글자에 따라 목적격 조사 '을/를'을 반환."""
    if not name:
        return "을"
    last = name[-1]
    code = ord(last)
    if 0xAC00 <= code <= 0xD7A3:
        return "을" if (code - 0xAC00) % 28 != 0 else "를"
    return "를"


def _run_authenticator_in_thread(
    user_id: str, password: str
) -> tuple[AuthResult, Authenticator]:
    """별도 스레드에서 인증을 실행하고, 메인 스레드는 wx 이벤트를 처리하며 대기."""
    authenticator = Authenticator()
    holder: dict = {}

    def worker():
        try:
            holder["result"] = authenticator.authenticate(user_id, password)
        except Exception as e:
            holder["result"] = AuthResult(AuthResult.NETWORK_ERROR, str(e))

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    while thread.is_alive():
        wx.YieldIfNeeded()
        time.sleep(0.05)
    thread.join()
    return holder["result"], authenticator


def _do_authenticate(
    user_id: str,
    password: str,
    program_name: str,
    silent: bool = False,
) -> tuple[bool, Authenticator | None, str]:
    """
    실제 인증을 수행한다. 진행 중에는 비프음 + 스크린리더로 진행 알림.

    Returns:
        (성공 여부, 성공 시 Authenticator, AuthResult.status)
    """
    speak("인증 중입니다. 잠시만 기다려 주세요.")
    wx.SafeYield()

    progress = ProgressIndicator()
    progress.start()
    try:
        result, authenticator = _run_authenticator_in_thread(user_id, password)
    finally:
        progress.stop()

    if result.is_success:
        _play_auth_success_beep()
        josa = _josa_eul_reul(program_name)
        speak(f"인증이 완료되었습니다. {program_name}{josa} 실행합니다.")
        return True, authenticator, result.status

    if result.status == AuthResult.NETWORK_ERROR:
        msg = f"인증에 실패했습니다.\n{result.message}"
        speak(msg)
        wx.MessageBox(msg, "네트워크 오류", wx.OK | wx.ICON_ERROR)
    elif result.status == AuthResult.LOGIN_FAILED:
        msg = result.message
        speak(msg)
        if not silent:
            wx.MessageBox(msg, "로그인 실패", wx.OK | wx.ICON_ERROR)
    elif result.status == AuthResult.NOT_MEMBER:
        msg = result.message
        speak(msg)
        wx.MessageBox(msg, "인증 실패", wx.OK | wx.ICON_ERROR)

    return False, None, result.status


def run_authentication(program_name: str = "프로그램") -> Authenticator | None:
    """
    초록등대 동호회 인증을 실행한다.

    이 함수를 호출하기 전에 wx.App이 생성되어 있어야 한다.

    Args:
        program_name: 인증 성공 시 음성으로 안내할 프로그램명.
            예) "초록멀티" → "인증이 완료되었습니다. 초록멀티를 실행합니다."

    Returns:
        인증 성공 시 Authenticator 객체 (session, user_id, nickname, rank 포함),
        실패/취소 시 None.
    """
    saved = load_credentials()
    if saved:
        user_id, password = saved
        success, authenticator, status = _do_authenticate(
            user_id, password, program_name, silent=True,
        )
        if success:
            if authenticator is not None and not authenticator.user_id:
                authenticator.user_id = user_id
            return authenticator
        if status == AuthResult.LOGIN_FAILED:
            delete_credentials()
        else:
            return None

    dialog = LoginDialog()
    try:
        while True:
            result = dialog.ShowModal()
            if result == wx.ID_CANCEL:
                return None

            user_id, password = dialog.get_credentials()
            should_save = dialog.get_save_option()

            success, authenticator, status = _do_authenticate(
                user_id, password, program_name,
            )

            if success:
                if should_save:
                    save_credentials(user_id, password)
                if authenticator is not None and not authenticator.user_id:
                    authenticator.user_id = user_id
                return authenticator

            if status in (AuthResult.NOT_MEMBER, AuthResult.NETWORK_ERROR):
                return None
    finally:
        dialog.Destroy()
