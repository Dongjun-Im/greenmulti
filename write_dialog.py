"""게시물 작성 대화상자"""
import os
import threading

import requests
import wx

from config import SORISEM_BASE_URL
from screen_reader import speak


class WriteDialog(wx.Dialog):
    """게시물 작성 대화상자"""

    def __init__(self, parent, session: requests.Session, bo_table: str,
                 existing_title: str = "", existing_body: str = ""):
        super().__init__(
            parent, title="게시물 작성",
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
            size=(600, 500),
        )

        self.session = session
        self.bo_table = bo_table
        self.attached_files: list[str] = []  # 첨부파일 경로 목록
        self._is_edit_mode = bool(existing_title or existing_body)

        self._create_controls()
        self._do_layout()
        self._bind_events()

        # 수정 모드: 기존 제목/본문 미리 채우고 제목에 포커스
        if self._is_edit_mode:
            self.title_input.SetValue(existing_title)
            self.body_input.SetValue(existing_body)
            self.SetTitle("게시물 수정")
            self.submit_btn.SetLabel("수정(&W)")

        # 저시력 테마 적용
        try:
            from theme import apply_theme, make_font, load_font_size
            apply_theme(self, make_font(load_font_size()))
        except Exception:
            pass

        self.Centre()
        self.title_input.SetFocus()
        # 제목 끝으로 커서 이동 (수정 모드에서 바로 추가 편집 가능)
        if self._is_edit_mode:
            self.title_input.SetInsertionPointEnd()
            speak(
                "게시물 수정. 기존 제목과 본문이 채워져 있습니다. "
                "제목과 본문을 모두 수정할 수 있습니다. "
                "탭 키로 본문으로 이동합니다."
            )
        else:
            speak("게시물 작성. 제목을 입력해 주세요.")

    def _create_controls(self):
        self.panel = wx.Panel(self)

        # 제목
        self.title_label = wx.StaticText(self.panel, label="제목(&T):")
        self.title_input = wx.TextCtrl(
            self.panel, name="제목",
        )

        # 본문
        self.body_label = wx.StaticText(self.panel, label="본문(&B):")
        self.body_input = wx.TextCtrl(
            self.panel, name="본문",
            style=wx.TE_MULTILINE | wx.TE_PROCESS_ENTER,
        )

        # 첨부파일 목록
        self.file_label = wx.StaticText(self.panel, label="첨부파일(&F):")
        self.file_list = wx.ListBox(
            self.panel, style=wx.LB_SINGLE, name="첨부파일 목록",
        )

        # 첨부파일 추가/제거 버튼
        self.add_file_btn = wx.Button(self.panel, label="파일 추가(&A)")
        self.remove_file_btn = wx.Button(self.panel, label="파일 제거(&R)")

        # 글쓰기/취소 버튼
        self.submit_btn = wx.Button(self.panel, wx.ID_OK, "글쓰기(&W)")
        self.cancel_btn = wx.Button(self.panel, wx.ID_CANCEL, "취소")

    def _do_layout(self):
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 제목
        main_sizer.Add(self.title_label, 0, wx.LEFT | wx.RIGHT | wx.TOP, 10)
        main_sizer.Add(self.title_input, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # 본문
        main_sizer.Add(self.body_label, 0, wx.LEFT | wx.RIGHT, 10)
        main_sizer.Add(self.body_input, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # 첨부파일
        main_sizer.Add(self.file_label, 0, wx.LEFT | wx.RIGHT, 10)
        main_sizer.Add(self.file_list, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        file_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        file_btn_sizer.Add(self.add_file_btn, 0, wx.RIGHT, 5)
        file_btn_sizer.Add(self.remove_file_btn, 0)
        main_sizer.Add(file_btn_sizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # 글쓰기/취소
        btn_sizer = wx.StdDialogButtonSizer()
        btn_sizer.AddButton(self.submit_btn)
        btn_sizer.AddButton(self.cancel_btn)
        btn_sizer.Realize()
        main_sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        self.panel.SetSizer(main_sizer)

    def _bind_events(self):
        self.add_file_btn.Bind(wx.EVT_BUTTON, self.on_add_file)
        self.remove_file_btn.Bind(wx.EVT_BUTTON, self.on_remove_file)
        self.submit_btn.Bind(wx.EVT_BUTTON, self.on_submit)
        self.cancel_btn.Bind(wx.EVT_BUTTON, self.on_cancel)
        # 본문 편집창에서 Tab으로 다음 컨트롤로 이동
        self.Bind(wx.EVT_CHAR_HOOK, self.on_char_hook)

    def on_char_hook(self, event):
        if event.GetKeyCode() == wx.WXK_TAB and self.FindFocus() == self.body_input:
            if event.ShiftDown():
                self.body_input.Navigate(wx.NavigationKeyEvent.IsBackward)
            else:
                self.body_input.Navigate(wx.NavigationKeyEvent.IsForward)
            return
        event.Skip()

    def on_add_file(self, event):
        """첨부파일 추가"""
        dlg = wx.FileDialog(
            self, "첨부할 파일 선택",
            style=wx.FD_OPEN | wx.FD_MULTIPLE,
        )
        if dlg.ShowModal() == wx.ID_OK:
            paths = dlg.GetPaths()
            for path in paths:
                if path not in self.attached_files:
                    self.attached_files.append(path)
            self._refresh_file_list()
            speak(f"파일 {len(paths)}개가 추가되었습니다.")
        dlg.Destroy()

    def on_remove_file(self, event):
        """선택된 첨부파일 제거"""
        sel = self.file_list.GetSelection()
        if sel == wx.NOT_FOUND:
            speak("제거할 파일을 선택해 주세요.")
            return
        removed = self.attached_files.pop(sel)
        self._refresh_file_list()
        speak(f"{os.path.basename(removed)} 파일이 제거되었습니다.")

    def _refresh_file_list(self):
        names = [os.path.basename(p) for p in self.attached_files]
        self.file_list.Set(names)
        if names:
            self.file_list.SetSelection(min(len(names) - 1, 0))

    def on_submit(self, event):
        """글쓰기"""
        title = self.title_input.GetValue().strip()
        body = self.body_input.GetValue().strip()

        if not title:
            speak("제목을 입력해 주세요.")
            wx.MessageBox("제목을 입력해 주세요.", "입력 오류",
                          wx.OK | wx.ICON_WARNING, self)
            self.title_input.SetFocus()
            return

        if not body:
            speak("본문을 입력해 주세요.")
            wx.MessageBox("본문을 입력해 주세요.", "입력 오류",
                          wx.OK | wx.ICON_WARNING, self)
            self.body_input.SetFocus()
            return

        # 수정/답변 모드 확인
        is_edit = hasattr(self, '_edit_wr_id') and self._edit_wr_id
        is_reply = hasattr(self, '_reply_wr_id') and self._reply_wr_id
        if is_edit:
            speak("게시물을 수정하는 중입니다.")
        elif is_reply:
            speak("답변을 등록하는 중입니다.")
        else:
            speak("게시물을 등록하는 중입니다.")

        def worker():
            try:
                url = f"{SORISEM_BASE_URL}/bbs/write_update.php"
                if is_edit:
                    w_mode = "u"
                elif is_reply:
                    w_mode = "r"
                else:
                    w_mode = ""
                data = {
                    "bo_table": self.bo_table,
                    "wr_subject": title,
                    "wr_content": body,
                    "w": w_mode,
                }
                if is_edit:
                    data["wr_id"] = self._edit_wr_id
                elif is_reply:
                    data["wr_id"] = self._reply_wr_id

                files_data = {}
                for i, path in enumerate(self.attached_files):
                    key = f"bf_file[{i}]"
                    files_data[key] = (
                        os.path.basename(path),
                        open(path, "rb"),
                        "application/octet-stream",
                    )

                if files_data:
                    resp = self.session.post(url, data=data, files=files_data, timeout=30)
                    # 파일 핸들 닫기
                    for key in files_data:
                        files_data[key][1].close()
                else:
                    resp = self.session.post(url, data=data, timeout=15)

                wx.CallAfter(self._submit_done, resp.status_code)
            except Exception as e:
                wx.CallAfter(self._submit_error, str(e))

        threading.Thread(target=worker, daemon=True).start()

    def _submit_done(self, status_code: int):
        if self._is_edit_mode:
            msg = "게시물이 수정되었습니다."
        else:
            msg = "게시물이 등록되었습니다."
        speak(msg)
        wx.MessageBox(msg, "완료", wx.OK | wx.ICON_INFORMATION, self)
        self.EndModal(wx.ID_OK)

    def _submit_error(self, error: str):
        speak(f"게시물 등록에 실패했습니다.")
        wx.MessageBox(f"게시물 등록에 실패했습니다.\n{error}",
                      "오류", wx.OK | wx.ICON_ERROR, self)

    def on_cancel(self, event):
        self.EndModal(wx.ID_CANCEL)
