"""초록멀티 - 메인 진입점"""
import os

import wx

from green_auth import run_authentication
from config import APP_NAME


class ChorokMultiApp(wx.App):
    """초록멀티 애플리케이션"""

    def OnInit(self):
        self.session = None
        self._play_startup_sound()

        # 인증 절차 실행
        self.session = run_authentication()
        if self.session is None:
            return False

        # 인증 성공: 메뉴 자동 감지 후 메인 윈도우 표시
        self._auto_detect_menus()

        from main_frame import MainFrame
        frame = MainFrame(self.session)
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

    def _play_startup_sound(self):
        """시작 사운드 재생"""
        try:
            import winsound
            sound_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "sounds", "startup.wav"
            )
            if os.path.exists(sound_path):
                winsound.PlaySound(sound_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception:
            pass

    def _play_shutdown_sound(self):
        """종료 사운드 재생"""
        try:
            import winsound
            sound_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "sounds", "shutdown.wav"
            )
            if os.path.exists(sound_path):
                winsound.PlaySound(sound_path, winsound.SND_FILENAME)
        except Exception:
            pass


def main():
    app = ChorokMultiApp()
    app.MainLoop()


if __name__ == "__main__":
    main()
