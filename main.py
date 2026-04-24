"""초록멀티 - 메인 진입점"""
import os
import sys

import wx

from green_auth import run_authentication
from config import APP_NAME


class ChorokMultiApp(wx.App):
    """초록멀티 애플리케이션"""

    def OnInit(self):
        self.session = None
        self.user_id: str | None = None
        self.user_nickname: str | None = None
        self._play_startup_sound()
        self._cleanup_old_update_artifacts()

        # 인증 절차 실행
        authenticator = run_authentication()
        if authenticator is None:
            return False
        self.session = authenticator.session
        self.user_id = authenticator.user_id
        self.user_nickname = authenticator.nickname

        # 인증 성공: 메뉴 자동 감지 후 메인 윈도우 표시
        self._auto_detect_menus()

        from main_frame import MainFrame
        frame = MainFrame(
            self.session,
            current_user_id=self.user_id,
            current_user_nickname=self.user_nickname,
        )
        self.SetTopWindow(frame)
        return True

    def OnExit(self):
        self._play_shutdown_sound()
        return 0

    def _auto_detect_menus(self):
        """로그인 후 소리샘 메인 페이지에서 실제 메뉴 URL을 자동 감지"""
        try:
            from green_auth import speak
            from config import SORISEM_BASE_URL
            from menu_manager import MenuManager, MenuItem
            from bs4 import BeautifulSoup
            import re

            speak("메뉴를 불러오는 중입니다.")
            resp = self.session.get(SORISEM_BASE_URL, timeout=15)
            soup = BeautifulSoup(resp.text, "html.parser")

            menus = []
            seen_urls = set()

            # 초록등대 동호회 고정
            menus.append(MenuItem("초록등대 동호회", "/plugin/ar.club/?cl=green", "club"))
            seen_urls.add("/plugin/ar.club/?cl=green")

            # 페이지의 모든 링크에서 메뉴 후보 추출
            for a in soup.find_all("a", href=True):
                href = a.get("href", "").strip()
                text = a.get_text(strip=True)

                if not text or len(text) < 2 or len(text) > 40:
                    continue
                if href in ("#", "", "javascript:void(0)", "javascript:;"):
                    continue
                if any(k in href for k in ["login", "logout", "register",
                                            "memo.php", "formmail", "password"]):
                    continue
                if any(k in text for k in ["본문으로", "상단으로", "로그아웃",
                                            "개인정보", "이용약관", "메일",
                                            "쪽지", "돌아가기"]):
                    continue
                if re.match(r"^\d+$", text):
                    continue

                # URL 정규화
                if href.startswith("http") and SORISEM_BASE_URL in href:
                    href = href.replace(SORISEM_BASE_URL, "")

                if href in seen_urls:
                    continue
                seen_urls.add(href)

                # 타입 결정
                if href.startswith("http") and SORISEM_BASE_URL not in href:
                    menu_type = "external"
                elif "ar.club" in href:
                    menu_type = "club"
                elif "bo_table" in href:
                    menu_type = "board"
                else:
                    menu_type = "category"

                menus.append(MenuItem(text, href, menu_type))

            if len(menus) > 1:
                manager = MenuManager()
                manager.menus = menus
                manager.save()
                speak(f"{len(menus)}개 메뉴를 불러왔습니다.")
            else:
                speak("메뉴를 불러오지 못했습니다. 기존 메뉴를 사용합니다.")

        except Exception as e:
            try:
                from green_auth import speak
                speak("메뉴 감지에 실패했습니다. 기존 메뉴를 사용합니다.")
            except Exception:
                pass

    def _cleanup_old_update_artifacts(self):
        """직전 업데이트에서 남은 *.exe.old 파일들을 정리.

        PS 재시작 스크립트가 정리에 실패했더라도 다음 실행 시 여기서
        확실하게 지운다. 현재 실행 중인 새 exe 와 이름이 다르므로 잠금 없음.
        """
        if not getattr(sys, "frozen", False):
            return
        import glob
        exe_dir = os.path.dirname(sys.executable)
        for path in glob.glob(os.path.join(exe_dir, "*.exe.old")):
            try:
                os.remove(path)
            except OSError:
                # 아직 핸들이 잡혀 있으면 다음 실행 때 재시도
                pass

    def _play_startup_sound(self):
        """시작 사운드 재생 (사용자 설정 이벤트)."""
        try:
            from sound import play_event
            play_event("program_start")
        except Exception:
            pass

    def _play_shutdown_sound(self):
        """종료 사운드 재생 (사용자 설정 이벤트, 동기)."""
        try:
            from sound import play_event
            play_event("program_end", block=True)
        except Exception:
            pass


def main():
    app = ChorokMultiApp()
    app.MainLoop()


if __name__ == "__main__":
    main()
