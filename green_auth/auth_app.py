"""인증 실행 모듈 - 다른 프로그램에서 호출하는 진입점"""
import requests
import wx

from green_auth.authenticator import Authenticator, AuthResult
from green_auth.credentials import save_credentials, load_credentials, delete_credentials
from green_auth.login_dialog import LoginDialog
from green_auth.screen_reader import speak


def _play_auth_success_beep():
    """인증 성공 비프음"""
    try:
        import winsound
        winsound.Beep(1000, 300)
    except Exception:
        pass


def _do_authenticate(user_id: str, password: str, silent: bool = False) -> tuple[bool, Authenticator | None]:
    """
    실제 인증을 수행한다.
    Args:
        user_id: 소리샘 아이디
        password: 비밀번호
        silent: True이면 실패 시 대화상자를 표시하지 않음
    Returns:
        (성공 여부, 성공 시 Authenticator 객체)
    """
    authenticator = Authenticator()
    result = authenticator.authenticate(user_id, password)

    if result.is_success:
        _play_auth_success_beep()
        speak("초록등대 동호회 회원 인증에 성공했습니다.")
        wx.MessageBox(
            "초록등대 동호회 회원 인증에 성공했습니다.\n"
            "확인을 누르면 프로그램이 실행됩니다.",
            "인증 성공",
            wx.OK | wx.ICON_INFORMATION,
        )
        return True, authenticator

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

    return False, None


def run_authentication() -> requests.Session | None:
    """
    초록등대 동호회 인증을 실행한다.

    이 함수를 호출하기 전에 wx.App이 생성되어 있어야 한다.

    Returns:
        인증 성공 시 로그인된 requests.Session 객체,
        실패 시 None
    """
    # 1. 저장된 자격 증명으로 자동 인증 시도
    saved = load_credentials()
    if saved:
        user_id, password = saved
        speak("인증 중입니다. 잠시만 기다려 주세요.")
        wx.SafeYield()  # 음성 출력이 즉시 처리되도록
        success, authenticator = _do_authenticate(user_id, password, silent=True)
        if success:
            return authenticator.session
        # 자동 인증 실패 시 저장된 정보 삭제
        delete_credentials()

    # 2. 로그인 대화상자 표시
    dialog = LoginDialog()

    while True:
        result = dialog.ShowModal()

        if result == wx.ID_CANCEL:
            dialog.Destroy()
            return None

        user_id, password = dialog.get_credentials()
        should_save = dialog.get_save_option()

        speak("인증 중입니다. 잠시만 기다려 주세요.")
        wx.SafeYield()  # 음성 출력이 즉시 처리되도록
        success, authenticator = _do_authenticate(user_id, password)

        if success:
            if should_save:
                save_credentials(user_id, password)
            dialog.Destroy()
            return authenticator.session

    dialog.Destroy()
    return None
