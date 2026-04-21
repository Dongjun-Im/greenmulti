"""초록등대 회원 인증 대화상자"""
import wx


class LoginDialog(wx.Dialog):
    """소리샘 아이디/비밀번호 입력 대화상자"""

    def __init__(self, parent=None):
        super().__init__(
            parent,
            title="초록등대 회원 인증",
            style=wx.DEFAULT_DIALOG_STYLE,
        )

        self._create_controls()
        self._do_layout()
        self._bind_events()

        # 스크린리더 접근성: 첫 입력란에 포커스
        self.id_input.SetFocus()

        self.Centre()

    def _create_controls(self):
        """컨트롤 생성"""
        self.panel = wx.Panel(self)

        # 아이디 입력
        self.id_label = wx.StaticText(
            self.panel, label="소리샘 아이디(&I):"
        )
        self.id_input = wx.TextCtrl(
            self.panel, name="소리샘 아이디",
            style=wx.TE_PROCESS_ENTER,
        )

        # 비밀번호 입력
        self.pw_label = wx.StaticText(
            self.panel, label="비밀번호(&P):"
        )
        self.pw_input = wx.TextCtrl(
            self.panel, name="비밀번호",
            style=wx.TE_PASSWORD | wx.TE_PROCESS_ENTER,
        )

        # 자격 증명 저장 체크박스
        self.save_check = wx.CheckBox(
            self.panel, label="아이디와 비밀번호 저장(&S)",
            name="아이디와 비밀번호 저장",
        )
        self.save_check.SetValue(True)

        # 버튼
        self.ok_btn = wx.Button(self.panel, wx.ID_OK, "확인(&O)")
        self.cancel_btn = wx.Button(self.panel, wx.ID_CANCEL, "취소")
        self.ok_btn.SetDefault()

    def _do_layout(self):
        """레이아웃 배치"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 입력 영역
        grid = wx.FlexGridSizer(cols=2, vgap=8, hgap=8)
        grid.AddGrowableCol(1, 1)

        grid.Add(self.id_label, 0, wx.ALIGN_CENTER_VERTICAL)
        grid.Add(self.id_input, 0, wx.EXPAND)
        grid.Add(self.pw_label, 0, wx.ALIGN_CENTER_VERTICAL)
        grid.Add(self.pw_input, 0, wx.EXPAND)

        main_sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 15)

        # 저장 체크박스
        main_sizer.Add(self.save_check, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 15)

        # 구분선
        main_sizer.Add(wx.StaticLine(self.panel), 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        # 버튼 영역
        btn_sizer = wx.StdDialogButtonSizer()
        btn_sizer.AddButton(self.ok_btn)
        btn_sizer.AddButton(self.cancel_btn)
        btn_sizer.Realize()
        main_sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 15)

        self.panel.SetSizer(main_sizer)
        main_sizer.Fit(self)

        # 최소 크기 설정
        self.SetMinSize(wx.Size(350, -1))
        self.Fit()

    def _bind_events(self):
        """이벤트 바인딩"""
        self.ok_btn.Bind(wx.EVT_BUTTON, self.on_ok)
        self.cancel_btn.Bind(wx.EVT_BUTTON, self.on_cancel)
        # Enter 키로 확인
        self.id_input.Bind(wx.EVT_TEXT_ENTER, self.on_ok)
        self.pw_input.Bind(wx.EVT_TEXT_ENTER, self.on_ok)

    def on_ok(self, event):
        """확인 버튼 클릭"""
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
        """취소 버튼 클릭"""
        self.EndModal(wx.ID_CANCEL)

    def get_credentials(self) -> tuple[str, str]:
        """입력된 아이디와 비밀번호 반환"""
        return self.id_input.GetValue().strip(), self.pw_input.GetValue()

    def get_save_option(self) -> bool:
        """자격 증명 저장 여부 반환"""
        return self.save_check.GetValue()

    def set_credentials(self, user_id: str, password: str):
        """저장된 자격 증명을 입력란에 설정"""
        self.id_input.SetValue(user_id)
        self.pw_input.SetValue(password)
