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


def _estimate_speech_seconds(text: str) -> float:
    """한국어 발화 길이 추정. 한글 음절 ≈ 0.18s, 그 외 0.06s. 1.2~5.0s 클램프."""
    if not text:
        return 0.0
    syllables = sum(1 for ch in text if 0xAC00 <= ord(ch) <= 0xD7A3)
    others = sum(1 for ch in text if ch.strip() and not (0xAC00 <= ord(ch) <= 0xD7A3))
    secs = syllables * 0.18 + others * 0.06
    return max(1.2, min(5.0, secs))


def _wait_speech(text: str, gap: float = 0.25) -> None:
    """발화 길이 + gap 만큼 wx 이벤트는 처리하면서 대기.
    time.sleep 단독으로 막아 두면 일부 스크린리더가 발화를 정상 큐잉하지 못한다.
    gap (기본 0.25s) 은 발화가 완전히 끝난 뒤 다음 안내가 즉시 따라붙어 어색하게
    이어 들리는 것을 막아 자연스러운 텀을 두면서도 반응 속도를 유지한다."""
    end = time.monotonic() + _estimate_speech_seconds(text) + max(0.0, gap)
    while time.monotonic() < end:
        try:
            wx.SafeYield()
        except Exception:
            pass
        time.sleep(0.05)


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
    intro_msg = "인증 중입니다. 잠시만 기다려 주세요."
    speak(intro_msg)
    # 저장된 자격증명이 있으면 인증이 매우 빨리(<500ms) 끝나, 이 안내가
    # 다음 단계의 "인증이 완료되었습니다" 발화에 즉시 잘려 들리지 않게 된다.
    # 발화 길이 + 0.5s gap 만큼 메인 스레드 대기로 순서 보장.
    _wait_speech(intro_msg, gap=0.25)

    progress = ProgressIndicator()
    progress.start()
    try:
        result, authenticator = _run_authenticator_in_thread(user_id, password)
    finally:
        progress.stop()

    if result.is_success:
        _play_auth_success_beep()
        josa = _josa_eul_reul(program_name)
        completion_msg = (
            f"인증이 완료되었습니다. {program_name}{josa} 실행합니다."
        )
        speak(completion_msg)
        # 완료 안내가 끝나기 전에 main.py 의 _auto_detect_menus() 가 곧바로
        # "N개 메뉴를 불러왔습니다" 를 발화해 이 안내를 잘라먹는 문제 방지.
        # 발화 길이만큼 + 0.5s gap 만큼 메인 스레드를 대기 (wx 이벤트는 yield).
        _wait_speech(completion_msg, gap=0.25)
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
