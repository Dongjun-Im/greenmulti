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
                 existing_title: str = "", existing_body: str = "",
                 user_rank: str | None = None):
        super().__init__(
            parent, title="게시물 작성",
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
            size=(600, 500),
        )

        self.session = session
        self.bo_table = bo_table
        self.attached_files: list[str] = []  # 첨부파일 경로 목록
        self._is_edit_mode = bool(existing_title or existing_body)
        # 동호회 관리자 등급(또는 클럽 관리자) 인 경우에만 공지 체크박스 노출.
        # rank 텍스트는 green_auth.authenticator 가 회원 목록 페이지에서 추출한
        # 그대로(예: "동호회관리자", "동호회 관리자", "클럽관리자") 들어온다.
        self._is_admin = bool(
            user_rank and ("관리자" in user_rank)
        )

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

        # 공지 체크박스 — 동호회 관리자만 노출.
        if self._is_admin:
            self.notice_checkbox = wx.CheckBox(
                self.panel, label="공지로 등록(&O)", name="공지로 등록",
            )
        else:
            self.notice_checkbox = None

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

        # 공지 체크박스 — 관리자에게만
        if self.notice_checkbox is not None:
            main_sizer.Add(
                self.notice_checkbox, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 10,
            )

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

                # 공지 체크박스 처리.
                # gnuboard5 와 ar.club 플러그인은 버전·테마에 따라 필드명이 다른
                # 사례가 있다 — write 폼을 GET 해서 실제 form 안의 "notice" 류
                # input 이름들을 추출하고 모두 함께 POST 한다. 추출 실패 시에도
                # 알려진 표준 필드명들을 폴백으로 모두 실어 둔다.
                if (
                    self.notice_checkbox is not None
                    and self.notice_checkbox.GetValue()
                ):
                    self._apply_notice_fields(data, w_mode, is_reply)

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

    def _apply_notice_fields(self, data: dict, w_mode: str, is_reply: bool) -> None:
        """write 폼의 실제 notice 필드를 추출해 data 에 병합.

        먼저 write.php?bo_table=... 를 GET 해 form 의 모든 `<input>` 중 name 에
        "notice" 가 들어 있는 것의 이름을 수집한다. 이름별 적절한 값을 결정하고
        data 에 병합. 실패 시 gnuboard5 의 표준 후보 필드들을 폴백으로 모두 채운다.
        진단을 위해 첫 호출 시 폼 HTML 을 `data/write_form_<bo>.html` 로 dump.
        """
        try:
            from bs4 import BeautifulSoup
        except Exception:
            BeautifulSoup = None

        # 답변·수정 모드에 맞는 write.php URL.
        get_url = f"{SORISEM_BASE_URL}/bbs/write.php?bo_table={self.bo_table}"
        if w_mode == "r" and getattr(self, "_reply_wr_id", None):
            get_url += f"&wr_id={self._reply_wr_id}&w=r"
        elif w_mode == "u" and getattr(self, "_edit_wr_id", None):
            get_url += f"&wr_id={self._edit_wr_id}&w=u"

        notice_field_names: list[str] = []
        try:
            resp = self.session.get(get_url, timeout=15)
            html = resp.text or ""
            try:
                from config import DATA_DIR
                os.makedirs(DATA_DIR, exist_ok=True)
                safe = self.bo_table.replace("/", "_")[:40]
                with open(
                    os.path.join(DATA_DIR, f"write_form_{safe}.html"),
                    "w", encoding="utf-8",
                ) as _wf:
                    _wf.write(html)
            except Exception:
                pass
            if BeautifulSoup is not None and html:
                soup = BeautifulSoup(html, "html.parser")
                for inp in soup.find_all(["input", "select"]):
                    name = (inp.get("name") or "").strip()
                    name_low = name.lower()
                    if "notice" in name_low or "공지" in name:
                        if name and name not in notice_field_names:
                            notice_field_names.append(name)
        except Exception:
            pass

        # 실제 폼에서 발견된 필드는 그대로 채운다.
        # gnuboard 컨벤션:
        #   · "notice_check"  → value 는 bo_table (관리자 게시판에서 공지로 지정)
        #   · "notice"        → value 는 보통 "1" 또는 wr_id
        #   · "chk_notice"    → "1"
        #   · "wr_notice"     → "1"
        for name in notice_field_names:
            n_low = name.lower()
            if n_low in ("notice_check",):
                data[name] = self.bo_table
            else:
                data[name] = "1"

        # 폴백: 폼에서 못 찾았더라도 알려진 후보를 모두 실어 둔다.
        if not notice_field_names:
            data.setdefault("chk_notice", "1")
            data.setdefault("notice_check", self.bo_table)
            data.setdefault("wr_notice", "1")
            data.setdefault("notice", "1")

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
