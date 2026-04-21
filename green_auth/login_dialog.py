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
        self.pw_input = wx.TextCtrl(
            self.panel, name="비밀번호",
            style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER,
        )

        self.save_check = wx.CheckBox(
            self.panel, label="아이디와 비밀번호 저장(&S)",
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
        self.save_check.Bind(wx.EVT_SET_FOCUS, self.on_save_check_focus)

    def _speak_save_check_state(self):
        state = "체크됨" if self.save_check.GetValue() else "체크 해제됨"
        speak(f"아이디와 비밀번호 저장 확인란 {state}, 단축키 Alt S")

    def on_save_check(self, event):
        self._speak_save_check_state()

    def on_save_check_focus(self, event):
        event.Skip()
        wx.CallLater(200, self._speak_save_check_state)

    def on_ok(self, event):
        user_id = self.id_input.GetValue().strip()
        password = self.pw_input.GetValue()

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
        return self.id_input.GetValue().strip(), self.pw_input.GetValue()

    def get_save_option(self) -> bool:
        return self.save_check.GetValue()

    def set_credentials(self, user_id: str, password: str):
        self.id_input.SetValue(user_id)
        self.pw_input.SetValue(password)
