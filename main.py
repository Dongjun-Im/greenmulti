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

        # 단일 인스턴스 체크 — 같은 프로그램이 이미 실행 중이면 안내 후 종료.
        # 사용자 프로필마다 별도 파일을 쓰도록 이름에 사용자명을 포함.
        self._instance_checker = wx.SingleInstanceChecker(
            f"chorok_multi_{wx.GetUserId()}"
        )
        if self._instance_checker.IsAnotherRunning():
            wx.MessageBox(
                "초록멀티가 이미 실행 중입니다.\n"
                "작업 표시줄이나 알림 영역에서 실행 중인 창을 확인해 주세요.",
                "초록멀티 실행 중",
                wx.OK | wx.ICON_INFORMATION,
            )
            return False

        self._play_startup_sound()
        self._cleanup_old_update_artifacts()

        # 인증 절차 실행
        authenticator = run_authentication(APP_NAME)
        if authenticator is None:
            return False
        self.session = authenticator.session
        self.user_id = authenticator.user_id
        self.user_nickname = authenticator.nickname
        self.user_rank = getattr(authenticator, "rank", None)

        # 인증 성공: 메뉴 자동 감지 후 메인 윈도우 표시
        self._auto_detect_menus()

        from main_frame import MainFrame
        frame = MainFrame(
            self.session,
            current_user_id=self.user_id,
            current_user_nickname=self.user_nickname,
            current_user_rank=self.user_rank,
        )
        self.SetTopWindow(frame)
        return True

    def OnExit(self):
        self._play_shutdown_sound()
        return 0

    def _auto_detect_menus(self):
        """로그인 후 소리샘 메인 페이지에서 실제 메뉴 URL을 자동 감지.

        사용자가 data/menu_list.txt 를 편집해 두었다면 자동 감지를 건너뛰고
        해당 파일을 그대로 사용한다.
        """
        try:
            from green_auth import speak
            from config import SORISEM_BASE_URL
            from menu_manager import MenuManager, MenuItem
            from bs4 import BeautifulSoup
            import re

            # 사용자 편집 파일이 있으면 자동 감지 건너뛰기
            _probe_manager = MenuManager()
            if _probe_manager.has_user_override():
                speak("사용자 메뉴 파일을 사용합니다.")
                return

            # 안내 음성 ("메뉴를 불러오는 중입니다") 은 의도적으로 제거.
            # 직전에 인증 완료 직후 "초록멀티를 실행합니다" 가 발화되므로,
            # 곧바로 다른 메시지를 speak() 하면 PurgeBeforeSpeak 동작으로
            # 그 발화가 즉시 잘려나가 사용자가 듣지 못한다. 메뉴 감지 결과는
            # 끝부분의 "N개 메뉴를 불러왔습니다" 로 충분히 안내된다.
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

            # ── 메인 메뉴 정돈 ──
            # 1) 이전 버전에서 메인 메뉴에 잘못 추가된 초록등대 클럽 자료실
            #    (cl=green4) / 엔터테인먼트 자료실 (cl=green6) 엔트리를 제거.
            #    이 두 메뉴는 초록등대 동호회 하위 메뉴에만 존재하도록 한다.
            _stale_urls = {
                "/plugin/ar.club/?cl=green4",
                "/plugin/ar.club/?cl=green6",
            }
            menus = [m for m in menus if m.url not in _stale_urls]

            # 2) 소리샘 자료실 이름 보정: "자료실"(번호 접두사 허용) 을
            #    "소리샘 자료실" 로 표시해 메인 메뉴에서 어떤 자료실인지
            #    명확하게 드러낸다. URL에 mo=pds 가 포함된 항목이 대상.
            for m in menus:
                if "mo=pds" not in m.url:
                    continue
                m_num = re.match(r'^(\s*\d+[\.\)]\s*)', m.name or "")
                _prefix = m_num.group(1) if m_num else ""
                _core = (m.name or "")[len(_prefix):].strip()
                if _core == "자료실":
                    m.name = f"{_prefix}소리샘 자료실"

            if len(menus) > 1:
                manager = MenuManager()
                manager.menus = menus
                # 자료실 / 전자도서관 / 노원시각장애인학습지원센터 등 표준 메뉴
                # 자동 보충 — load() 경로 외에는 _ensure_forced_club_menus 가
                # 호출되지 않으므로 여기서 명시적으로 한 번 돌려준다.
                # (이걸 빠뜨리면 안내 음성에 보충 메뉴 개수가 누락되어
                # "13개" 처럼 실제 화면 표시(16개)보다 적게 나온다.)
                try:
                    manager._ensure_forced_club_menus()
                except Exception:
                    pass
                manager.save()
                # 사용자 편집용 텍스트 파일 seed. 이미 있으면 덮어쓰지 않음.
                try:
                    manager.export_to_txt()
                except Exception:
                    pass
                # 화면에 실제 표시되는 메뉴 수 기준으로 안내.
                try:
                    displayed_count = len(manager.get_display_names())
                except Exception:
                    displayed_count = len(menus)
                speak(f"{displayed_count}개 메뉴를 불러왔습니다.")
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
