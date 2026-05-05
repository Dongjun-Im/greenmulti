"""초록등대 회원 인증 대화상자"""
import wx

from green_auth.config import AUTH_TITLE
from green_auth.screen_reader import speak


class LoginDialog(wx.Dialog):
    """소리샘 아이디/비밀번호 입력 대화상자"""

    def __init__(self, parent=None):
        super().__init__(
            parent,
            title=AUTH_TITLE,
            style=wx.DEFAULT_DIALOG_STYLE,
        )

        # 비밀번호 본문 (실제 값). 화면 표시는 '*' 만, 실제 값은 여기 보관.
        # wx.TE_PASSWORD 의 OS 기본 마스크 문자(●/•) 가 일부 한국어 스크린리더
        # (센스리더 등) 에서 "괄호닫고" 처럼 깨진 발음으로 들리는 문제 회피.
        # EM_SETPASSWORDCHAR 로 강제 변경도 시도했으나 wx 가 RichEdit 등을
        # 쓰는 환경에서 무시되는 케이스가 있어 수동 마스킹으로 통일.
        self._real_password = ""

        # 첫 자동 포커스(다이얼로그 오픈 시 id_input 으로 들어가는 포커스) 만
        # 안내 발화를 건너뛰어 스크린리더가 다이얼로그 제목 ("초록등대 인증") 을
        # 끝까지 읽도록 한다. 이후 모든 Tab/Shift+Tab 포커스는 즉시 안내.
        self._initial_auto_focus_consumed = False

        self._create_controls()
        self._do_layout()
        self._bind_events()

        # 저시력 테마 적용 (호스트 앱에 theme 모듈이 있으면)
        try:
            from theme import apply_theme, make_font, load_font_size
            apply_theme(self, make_font(load_font_size()))
        except Exception:
            pass

        self.id_input.SetFocus()
        self.Centre()

    def _create_controls(self):
        self.panel = wx.Panel(self)

        self.id_label = wx.StaticText(
            self.panel, label="소리샘 아이디(&I):"
        )
        self.id_input = wx.TextCtrl(
            self.panel, name="소리샘 아이디",
            style=wx.TE_PROCESS_ENTER,
        )

        self.pw_label = wx.StaticText(
            self.panel, label="비밀번호(&P):"
        )
        # 수동 마스킹 — TE_PASSWORD 미사용. 입력은 EVT_CHAR/EVT_KEY_DOWN 으로
        # 가로채 self._real_password 에 누적, 화면에는 길이만큼 '*' 만 표시.
        self.pw_input = wx.TextCtrl(
            self.panel, name="비밀번호",
            style=wx.TE_PROCESS_ENTER,
        )

        self.save_check = wx.CheckBox(
            self.panel, label="아이디/비밀번호 저장(&S)",
            name="아이디/비밀번호 저장 체크상자",
            style=wx.CHK_2STATE,
        )
        self.save_check.SetValue(True)

        self.ok_btn = wx.Button(self.panel, wx.ID_OK, "확인(&O)")
        self.cancel_btn = wx.Button(self.panel, wx.ID_CANCEL, "취소")
        self.ok_btn.SetDefault()

    def _do_layout(self):
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        grid = wx.FlexGridSizer(cols=2, vgap=8, hgap=8)
        grid.AddGrowableCol(1, 1)

        grid.Add(self.id_label, 0, wx.ALIGN_CENTER_VERTICAL)
        grid.Add(self.id_input, 0, wx.EXPAND)
        grid.Add(self.pw_label, 0, wx.ALIGN_CENTER_VERTICAL)
        grid.Add(self.pw_input, 0, wx.EXPAND)

        main_sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 15)
        main_sizer.Add(self.save_check, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 15)
        main_sizer.Add(wx.StaticLine(self.panel), 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        btn_sizer = wx.StdDialogButtonSizer()
        btn_sizer.AddButton(self.ok_btn)
        btn_sizer.AddButton(self.cancel_btn)
        btn_sizer.Realize()
        main_sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 15)

        self.panel.SetSizer(main_sizer)
        main_sizer.Fit(self)

        self.SetMinSize(wx.Size(350, -1))
        self.Fit()

    def _bind_events(self):
        self.ok_btn.Bind(wx.EVT_BUTTON, self.on_ok)
        self.cancel_btn.Bind(wx.EVT_BUTTON, self.on_cancel)
        self.id_input.Bind(wx.EVT_TEXT_ENTER, self.on_ok)
        self.pw_input.Bind(wx.EVT_TEXT_ENTER, self.on_ok)
        self.save_check.Bind(wx.EVT_CHECKBOX, self.on_save_check)

        # 비밀번호 수동 마스킹 입력 핸들러
        self.pw_input.Bind(wx.EVT_CHAR, self._on_pw_char)
        self.pw_input.Bind(wx.EVT_KEY_DOWN, self._on_pw_keydown)

        # 모든 컨트롤이 포커스를 받자마자 즉시 안내 — 이전 200ms 지연을 제거해
        # Tab/Shift+Tab 연타 시 첫 안내가 다음 포커스에 묻히는 문제 해결.
        self.id_input.Bind(wx.EVT_SET_FOCUS, self.on_id_focus)
        self.pw_input.Bind(wx.EVT_SET_FOCUS, self.on_pw_focus)
        self.save_check.Bind(wx.EVT_SET_FOCUS, self.on_save_check_focus)

    # ── 비밀번호 수동 마스킹 ──

    def _refresh_pw_display(self):
        """self._real_password 길이만큼 '*' 로 표시 갱신 + 커서 끝으로."""
        masked = "*" * len(self._real_password)
        self.pw_input.ChangeValue(masked)  # ChangeValue: EVT_TEXT 발화 안 함
        self.pw_input.SetInsertionPointEnd()

    def _on_pw_keydown(self, event):
        """Backspace/Delete/Ctrl+V 처리. 그 외는 EVT_CHAR 로 흘려보낸다."""
        kc = event.GetKeyCode()
        ctrl = event.ControlDown()

        if kc == wx.WXK_BACK:
            if self._real_password:
                self._real_password = self._real_password[:-1]
                self._refresh_pw_display()
            return  # 기본 동작 차단 (이미 처리)
        if kc == wx.WXK_DELETE:
            # 비밀번호 필드에서 Delete 는 백스페이스와 동일하게 취급 (커서 위치
            # 무시). 사용자가 일반적으로 마지막 글자 지움을 기대.
            if self._real_password:
                self._real_password = self._real_password[:-1]
                self._refresh_pw_display()
            return

        # Ctrl+V — 클립보드 텍스트를 비밀번호에 추가
        if ctrl and kc in (ord("V"), ord("v")):
            try:
                if wx.TheClipboard.Open():
                    data = wx.TextDataObject()
                    if wx.TheClipboard.GetData(data):
                        self._real_password += data.GetText()
                        self._refresh_pw_display()
                    wx.TheClipboard.Close()
            except Exception:
                pass
            return

        # Ctrl+A 등 텍스트 컨트롤 단축키는 표시값(별표) 에 적용해도 의미 없으므로 차단
        if ctrl and kc in (ord("A"), ord("a"), ord("X"), ord("x"),
                            ord("C"), ord("c")):
            return

        event.Skip()

    def _on_pw_char(self, event):
        """일반 인쇄 가능 문자 입력 → real_password 에 추가, 표시 갱신."""
        kc = event.GetKeyCode()
        # Tab/Enter/탐색 키는 통과
        if kc in (
            wx.WXK_TAB, wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER, wx.WXK_ESCAPE,
            wx.WXK_LEFT, wx.WXK_RIGHT, wx.WXK_UP, wx.WXK_DOWN,
            wx.WXK_HOME, wx.WXK_END,
        ):
            event.Skip()
            return

        # 일반 인쇄 가능 ASCII 범위 — 한글 IME 입력은 EVT_CHAR 가 아닌 다른
        # 이벤트로 들어와 여기 안 잡히지만 소리샘 비밀번호는 ASCII 가 일반적이므로
        # 충분.
        if 32 <= kc < 127:
            ch = chr(kc)
            self._real_password += ch
            self._refresh_pw_display()
            return

        event.Skip()

    # ── 포커스 안내 ──

    def _maybe_skip_initial_auto_focus(self) -> bool:
        """다이얼로그 오픈 시의 첫 자동 포커스 한 번만 건너뛰고, 그 이후의
        모든 Tab/Shift+Tab 포커스는 즉시 안내하도록 해주는 가드.
        반환값: True 면 안내 건너뛰어야 함."""
        if not self._initial_auto_focus_consumed:
            self._initial_auto_focus_consumed = True
            return True
        return False

    def _speak_save_check_state(self):
        state = "체크됨" if self.save_check.GetValue() else "체크 해제됨"
        speak(f"아이디/비밀번호 저장 체크상자 {state}, 단축키 Alt S")

    def on_save_check(self, event):
        self._speak_save_check_state()

    def on_id_focus(self, event):
        event.Skip()
        if self._maybe_skip_initial_auto_focus():
            return
        speak("소리샘 아이디 편집창")

    def on_pw_focus(self, event):
        event.Skip()
        if self._maybe_skip_initial_auto_focus():
            return
        speak("비밀번호 편집창")

    def on_save_check_focus(self, event):
        event.Skip()
        if self._maybe_skip_initial_auto_focus():
            return
        self._speak_save_check_state()

    # ── 확인/취소 ──

    def on_ok(self, event):
        user_id = self.id_input.GetValue().strip()
        password = self._real_password

        if not user_id:
            wx.MessageBox(
                "소리샘 아이디를 입력해 주세요.",
                "입력 오류",
                wx.OK | wx.ICON_WARNING,
                self,
            )
            self.id_input.SetFocus()
            return

        if not password:
            wx.MessageBox(
                "비밀번호를 입력해 주세요.",
                "입력 오류",
                wx.OK | wx.ICON_WARNING,
                self,
            )
            self.pw_input.SetFocus()
            return

        self.EndModal(wx.ID_OK)

    def on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)

    def get_credentials(self) -> tuple[str, str]:
        return self.id_input.GetValue().strip(), self._real_password

    def get_save_option(self) -> bool:
        return self.save_check.GetValue()

    def set_credentials(self, user_id: str, password: str):
        self.id_input.SetValue(user_id)
        self._real_password = password or ""
        self._refresh_pw_display()
