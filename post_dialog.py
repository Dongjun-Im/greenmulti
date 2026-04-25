"""게시물 내용 대화상자"""
import os
import re
import threading
import webbrowser

import requests
import wx

from config import SORISEM_BASE_URL
from page_parser import PostContent, CommentItem
from screen_reader import speak


# URL 감지 (http/https)
URL_PATTERN = re.compile(r'https?://[^\s<>"\')\]]+')


class ContextMenuTextCtrl(wx.TextCtrl):
    """팝업 키(application key)·Shift+F10 로 커스텀 팝업 메뉴가 뜨도록 하는 TextCtrl.

    wx.TE_RICH2 스타일의 TextCtrl 은 내부적으로 Windows RICHEDIT 네이티브
    컨트롤을 쓰기 때문에, 키보드로 컨텍스트 메뉴를 호출하면 WM_CONTEXTMENU
    가 컨트롤 레벨에서 자체 편집 메뉴(잘라내기/복사/붙여넣기)를 띄우고
    wx.EVT_CONTEXT_MENU 이벤트를 발사하지 않는다.

    해결 방법: 키 입력(EVT_KEY_DOWN)에서 VK_APPS(메뉴 키)와 Shift+F10 을
    직접 감지해, 바깥에서 `bind_context_menu(handler)` 로 등록한 핸들러를
    호출한다. 마우스 오른쪽 클릭은 기존 EVT_CONTEXT_MENU 로 같은 핸들러를
    받도록 래퍼를 함께 바인딩한다.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._ctx_menu_handler = None
        self.Bind(wx.EVT_KEY_DOWN, self._on_ctx_key_down)
        self.Bind(wx.EVT_CONTEXT_MENU, self._on_ctx_mouse)

    def bind_context_menu(self, handler):
        """팝업 키/우클릭에서 호출할 커스텀 메뉴 표시 함수 등록."""
        self._ctx_menu_handler = handler

    def _on_ctx_mouse(self, event):
        if self._ctx_menu_handler:
            self._ctx_menu_handler(event)
            return
        event.Skip()

    def _on_ctx_key_down(self, event):
        k = event.GetKeyCode()
        # VK_APPS(팝업 키) = WXK_WINDOWS_MENU, 또는 Shift+F10
        if k == wx.WXK_WINDOWS_MENU or (k == wx.WXK_F10 and event.ShiftDown()):
            if self._ctx_menu_handler:
                self._ctx_menu_handler(event)
                return
        event.Skip()


class ItemTextCtrl(ContextMenuTextCtrl):
    """특정 키를 네이티브 레벨에서 차단하는 TextCtrl.
    댓글 목록처럼 한 줄만 표시하면서 좌/우·Ctrl+좌/우로 글자/단어 단위
    기본 낭독을 받기 위해 사용.

    ContextMenuTextCtrl 를 상속해 팝업 키에서도 커스텀 컨텍스트 메뉴가 뜬다.
    """

    def MSWHandleMessage(self, msg, wParam, lParam):
        WM_KEYDOWN = 0x0100
        WM_CHAR = 0x0102
        if msg in (WM_KEYDOWN, WM_CHAR):
            # Home/End/Up/Down/Return/Back/PageUp/PageDn/Escape/Delete
            blocked_vk = {0x24, 0x23, 0x26, 0x28, 0x0D, 0x08, 0x21, 0x22, 0x1B, 0x2E}
            if wParam in blocked_vk:
                return True, 0
        return super().MSWHandleMessage(msg, wParam, lParam)


def _extract_urls(text: str) -> list[tuple[int, int, str]]:
    """텍스트에서 모든 URL의 (시작, 끝, URL) 목록 반환."""
    if not text:
        return []
    results = []
    for m in URL_PATTERN.finditer(text):
        url = m.group(0).rstrip(".,;:!?")
        results.append((m.start(), m.start() + len(url), url))
    return results


def _get_url_at_cursor(textctrl) -> str | None:
    """커서 위치의 URL 반환. 없으면 None."""
    pos = textctrl.GetInsertionPoint()
    text = textctrl.GetValue()
    for start, end, url in _extract_urls(text):
        if start <= pos <= end:
            return url
    return None


class CommentDialog(wx.Dialog):
    """댓글 작성/수정 대화상자"""

    def __init__(self, parent, title: str = "댓글 작성",
                 initial_text: str = ""):
        super().__init__(parent, title=title, style=wx.DEFAULT_DIALOG_STYLE)

        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        label = wx.StaticText(panel, label="댓글 내용(&C):")
        self.comment_text = wx.TextCtrl(
            panel, value=initial_text,
            style=wx.TE_MULTILINE, name="댓글 내용",
            size=(400, 150),
        )

        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(panel, wx.ID_OK, "확인(&O)")
        cancel_btn = wx.Button(panel, wx.ID_CANCEL, "취소")
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        ok_btn.SetDefault()

        sizer.Add(label, 0, wx.ALL, 5)
        sizer.Add(self.comment_text, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        panel.SetSizer(sizer)
        sizer.Fit(self)

        # 저시력 테마 적용
        try:
            from theme import apply_theme, make_font, load_font_size
            apply_theme(self, make_font(load_font_size()))
        except Exception:
            pass

        self.comment_text.SetFocus()
        self.Centre()

        # Ctrl+Enter 로 확인 — 일반 Enter 는 TE_MULTILINE 에서 줄바꿈으로 쓰인다.
        # 다른 모든 키는 event.Skip() 으로 흘려 정상 입력되도록 한다.
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

    def _on_char_hook(self, event):
        if event.GetKeyCode() == wx.WXK_RETURN and event.ControlDown():
            if self.IsModal():
                self.EndModal(wx.ID_OK)
            return
        event.Skip()

    def get_text(self) -> str:
        return self.comment_text.GetValue().strip()


class PostDialog(wx.Dialog):
    """게시물 내용을 표시하는 대화상자"""

    def __init__(self, parent, content: PostContent, session: requests.Session,
                 current_user_id: str | None = None,
                 current_user_nickname: str | None = None):
        title = content.title if content.title else "게시물 내용"
        super().__init__(
            parent, title=title,
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
            size=(700, 550),
        )

        self.content = content
        self.session = session
        # 로그인한 사용자의 소리샘 아이디·닉네임. 본인 게시물 여부 검증에 사용.
        self.current_user_id = current_user_id
        self.current_user_nickname = current_user_nickname

        # 첨부파일/댓글 유무를 명확하게 판단
        self.has_files = isinstance(content.files, list) and len(content.files) > 0
        self.has_comments = isinstance(content.comments, list) and len(content.comments) > 0
        self.comment_reversed = False
        self.navigate_result = ""  # "prev" or "next"
        self._comment_index = 0
        self._comment_displays: list[str] = []

        self.panel = wx.Panel(self)
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)

        self._create_controls()
        self._do_layout()
        self._fill_content()

        self.panel.SetSizer(self.main_sizer)

        # 키보드 이벤트
        self.Bind(wx.EVT_CHAR_HOOK, self.on_char_hook)

        # 상황별 팝업 메뉴 — 본문/댓글에서 오른쪽 클릭 또는 메뉴 키(팝업 키)/
        # Shift+F10 으로 호출. ContextMenuTextCtrl.bind_context_menu 는 마우스
        # 이벤트와 키보드 이벤트 모두 같은 핸들러로 라우팅한다.
        self.body_text.bind_context_menu(self._on_body_context_menu)
        if self.has_comments:
            self.comment_ctrl.bind_context_menu(self._on_comment_context_menu)

        # 저시력 테마 적용
        try:
            from theme import apply_theme, make_font, load_font_size
            apply_theme(self, make_font(load_font_size()))
        except Exception:
            pass

        self.Centre()
        self.body_text.SetFocus()
        self.body_text.SetInsertionPoint(0)

        # 스크린리더 안내
        announce = f"게시물. 제목 {content.title}."
        if content.author:
            announce += f" 작성자 {content.author}."
        if content.date:
            announce += f" 작성 일시 {content.date}."
        if self.has_comments:
            announce += f" 댓글 {len(content.comments)}개."
        speak(announce)

    def _create_controls(self):
        # 제목
        self.title_label = wx.StaticText(self.panel, label="제목:")
        self.title_text = wx.TextCtrl(
            self.panel, style=wx.TE_READONLY, name="제목",
        )

        # 작성자/날짜
        self.info_label = wx.StaticText(self.panel, label="정보:")
        self.info_text = wx.TextCtrl(
            self.panel, style=wx.TE_READONLY, name="작성자 및 날짜",
        )

        # 본문
        self.body_label = wx.StaticText(self.panel, label="본문:")
        self.body_text = ContextMenuTextCtrl(
            self.panel,
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
            name="본문",
        )

        # 첨부파일 (있을 때만)
        if self.has_files:
            self.file_label = wx.StaticText(self.panel, label="첨부파일(&F):")
            self.file_list = wx.ListBox(
                self.panel, style=wx.LB_SINGLE, name="첨부파일 목록",
            )
            self.file_download_btn = wx.Button(self.panel, label="첨부파일 저장 Alt+S")
            self.file_download_btn.Bind(wx.EVT_BUTTON, self.on_download_file)

        # 댓글 (있을 때만)
        if self.has_comments:
            self.comment_label = wx.StaticText(
                self.panel,
                label=f"댓글 ({len(self.content.comments)}개):",
            )
            # TextCtrl — 현재 선택된 댓글만 표시. 기존 ListBox와 비슷한 시각적
            # 높이를 확보하기 위해 MULTILINE + 최소 높이. ItemTextCtrl이 Up/Down/
            # Home/End/PageUp/Dn을 Windows 메시지 레벨에서 차단하므로 커서가 내부
            # 줄 이동으로 빠질 걱정은 없다. 좌/우·Ctrl+좌/우만 TextCtrl 기본 동작 →
            # 스크린리더가 글자/단어 단위로 자동 낭독한다.
            self.comment_ctrl = ItemTextCtrl(
                self.panel,
                style=wx.TE_READONLY | wx.TE_MULTILINE | wx.TE_DONTWRAP,
                name="댓글 목록",
            )
            self.comment_ctrl.SetMinSize((-1, 120))
            self.comment_edit_btn = wx.Button(self.panel, label="댓글 수정(&M)")
            self.comment_edit_btn.Bind(wx.EVT_BUTTON, self.on_edit_comment)
            self.comment_delete_btn = wx.Button(self.panel, label="댓글 삭제 Alt+D")
            self.comment_delete_btn.Bind(wx.EVT_BUTTON, self.on_delete_comment)

        # 게시물 수정/삭제/답변 버튼
        self.post_edit_btn = wx.Button(self.panel, label="게시물 수정 Alt+M")
        self.post_delete_btn = wx.Button(self.panel, label="게시물 삭제 Alt+D")
        self.post_reply_btn = wx.Button(self.panel, label="게시물 답변 Alt+R")
        self.post_edit_btn.Enable(bool(self.content.edit_url))
        self.post_delete_btn.Enable(bool(self.content.delete_url))
        self.post_reply_btn.Enable(bool(self.content.reply_url))
        self.post_edit_btn.Bind(wx.EVT_BUTTON, self.on_post_edit)
        self.post_delete_btn.Bind(wx.EVT_BUTTON, self.on_post_delete)
        self.post_reply_btn.Bind(wx.EVT_BUTTON, self.on_post_reply)

        # 이전/다음 게시물 버튼
        self.prev_btn = wx.Button(self.panel, label="이전 게시물 Alt+B")
        self.next_btn = wx.Button(self.panel, label="다음 게시물 Alt+N")
        self.prev_btn.Enable(bool(self.content.prev_url))
        self.next_btn.Enable(bool(self.content.next_url))
        self.prev_btn.Bind(wx.EVT_BUTTON, self.on_prev_post)
        self.next_btn.Bind(wx.EVT_BUTTON, self.on_next_post)

        # 닫기
        self.close_btn = wx.Button(self.panel, wx.ID_CANCEL, "닫기(&X)")
        self.close_btn.Bind(wx.EVT_BUTTON, self.on_close)

    def _do_layout(self):
        s = self.main_sizer

        # 제목
        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(self.title_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        row.Add(self.title_text, 1, wx.EXPAND)
        s.Add(row, 0, wx.EXPAND | wx.ALL, 5)

        # 정보
        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(self.info_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        row.Add(self.info_text, 1, wx.EXPAND)
        s.Add(row, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # 본문
        s.Add(self.body_label, 0, wx.LEFT | wx.RIGHT, 5)
        s.Add(self.body_text, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # 첨부파일
        if self.has_files:
            s.Add(self.file_label, 0, wx.LEFT | wx.RIGHT, 5)
            s.Add(self.file_list, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)
            s.Add(self.file_download_btn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # 댓글
        if self.has_comments:
            s.Add(self.comment_label, 0, wx.LEFT | wx.RIGHT, 5)
            s.Add(self.comment_ctrl, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)
            cmt_btn_row = wx.BoxSizer(wx.HORIZONTAL)
            cmt_btn_row.Add(self.comment_edit_btn, 0, wx.RIGHT, 5)
            cmt_btn_row.Add(self.comment_delete_btn, 0)
            s.Add(cmt_btn_row, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # 게시물 수정/삭제/답변 버튼
        post_btn_row = wx.BoxSizer(wx.HORIZONTAL)
        post_btn_row.Add(self.post_edit_btn, 0, wx.RIGHT, 5)
        post_btn_row.Add(self.post_delete_btn, 0, wx.RIGHT, 5)
        post_btn_row.Add(self.post_reply_btn, 0)
        s.Add(post_btn_row, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)

        # 이전/다음/닫기 버튼
        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        btn_row.Add(self.prev_btn, 0, wx.RIGHT, 5)
        btn_row.Add(self.next_btn, 0, wx.RIGHT, 5)
        btn_row.AddStretchSpacer()
        btn_row.Add(self.close_btn, 0)
        s.Add(btn_row, 0, wx.EXPAND | wx.ALL, 5)

    def _fill_content(self):
        self.title_text.SetValue(self.content.title)

        # 작성자, 작성 날짜, 작성 시간을 분리
        import re as _re
        post_date = ""
        post_time = ""
        if self.content.date:
            m = _re.match(
                r'^\s*(\d{2,4}[-./]\d{1,2}[-./]\d{1,2})(?:\s+(\d{1,2}:\d{2}(?::\d{2})?))?\s*$',
                self.content.date,
            )
            if m:
                post_date = m.group(1)
                post_time = m.group(2) or ""
            else:
                post_date = self.content.date

        # 정보 필드 (요약)
        info = []
        if self.content.author:
            info.append(f"작성자: {self.content.author}")
        if post_date:
            info.append(f"작성 날짜: {post_date}")
        if post_time:
            info.append(f"작성 시간: {post_time}")
        self.info_text.SetValue("  |  ".join(info))

        # 본문 상단에 메타정보 헤더 추가 → 본문 영역에서도 확인 가능
        header_lines = [f"제목: {self.content.title}"]
        header_lines.append(f"작성자: {self.content.author or '알 수 없음'}")
        header_lines.append(f"작성 날짜: {post_date or '알 수 없음'}")
        header_lines.append(f"작성 시간: {post_time or '알 수 없음'}")
        header_lines.append("-" * 30)
        header_lines.append("")
        self.body_text.SetValue("\n".join(header_lines) + "\n" + self.content.body)

        if self.has_files:
            names = [f["name"] for f in self.content.files]
            self.file_list.Set(names)
            if names:
                self.file_list.SetSelection(0)

        if self.has_comments:
            self._refresh_comment_list()

    # ── 키보드 ──

    def on_char_hook(self, event):
        keycode = event.GetKeyCode()
        focused = self.FindFocus()
        alt = event.AltDown()
        ctrl = event.ControlDown()

        # 자식 모달 대화상자(CommentDialog 등)가 열려 있을 때는 이 핸들러의
        # 단축키(Alt+M, Alt+D, Ctrl+U, Alt+R, Alt+B, Alt+N 등)가 자식의 입력을
        # 가로채서 "마침표 → M/B/…" 같은 엉뚱한 입력 증상을 만들 수 있다.
        # 포커스가 이 대화상자 내부에 없으면 모든 단축키 처리를 건너뛰고
        # 이벤트를 그대로 흘려보낸다.
        if focused is not None and not self.IsDescendant(focused):
            event.Skip()
            return

        # Enter: 커서 위치의 URL을 브라우저에서 열기 (본문 영역에서만)
        if keycode == wx.WXK_RETURN and not alt and not ctrl:
            if focused == self.body_text:
                url = _get_url_at_cursor(self.body_text)
                if url:
                    speak("브라우저에서 엽니다.")
                    webbrowser.open(url)
                    return

        # Ctrl+U: 게시물 내 URL 목록 보기
        if ctrl and keycode in (ord("U"), ord("u")):
            self._show_url_list()
            return

        # Alt+M: 게시물 수정
        if keycode in (ord("M"), ord("m")) and alt:
            self.on_post_edit(None)
            return

        # Alt+D: 게시물 삭제
        if keycode in (ord("D"), ord("d")) and alt:
            self.on_post_delete(None)
            return

        # Alt+S: 첨부파일 저장
        if keycode in (ord("S"), ord("s")) and alt:
            if self.has_files:
                self.on_download_file(None)
            return

        # Alt+R: 게시물 답변
        if keycode in (ord("R"), ord("r")) and alt:
            self.on_post_reply(None)
            return

        # Alt+B: 이전 게시물
        if keycode in (ord("B"), ord("b")) and alt:
            if self.content.prev_url:
                self.navigate_result = "prev"
                self.EndModal(wx.ID_BACKWARD)
            else:
                speak("이전 게시물이 없습니다.")
            return

        # Alt+N: 다음 게시물
        if keycode in (ord("N"), ord("n")) and alt:
            if self.content.next_url:
                self.navigate_result = "next"
                self.EndModal(wx.ID_FORWARD)
            else:
                speak("다음 게시물이 없습니다.")
            return

        # B: 본문 txt 저장 (본문에 포커스, Alt 없이)
        if keycode in (ord("B"), ord("b")) and not alt and focused == self.body_text:
            self.on_save_body()
            return

        # C: 댓글 작성
        if keycode in (ord("C"), ord("c")) and not alt:
            if focused == self.body_text or (self.has_comments and focused == self.comment_ctrl):
                self.on_write_comment()
                return

        # Alt+D 또는 D: 댓글 삭제
        if keycode in (ord("D"), ord("d")):
            if self.has_comments and (alt or focused == self.comment_ctrl):
                self.on_delete_comment()
                return

        # N: 댓글 정렬 순서 변경 (댓글 목록에 포커스)
        if keycode in (ord("N"), ord("n")) and not alt:
            if self.has_comments and focused == self.comment_ctrl:
                self.on_toggle_comment_sort()
                return

        # M: 댓글 수정 (댓글 목록에 포커스, Alt 없이)
        if keycode in (ord("M"), ord("m")) and not alt and not ctrl:
            if self.has_comments and focused == self.comment_ctrl:
                self.on_edit_comment()
                return

        # 댓글 목록 TextCtrl: 위/아래 → 댓글 이동, Home/End → 처음/끝 댓글.
        # 좌/우·Ctrl+좌/우는 TextCtrl 기본 동작(글자/단어 이동)이 스크린리더 낭독을
        # 자동으로 발생시키므로 별도 처리하지 않고 event.Skip() 으로 흘려보낸다.
        if self.has_comments and focused == self.comment_ctrl and not alt and not ctrl:
            if keycode == wx.WXK_UP:
                self._jump_to_comment(self._comment_index - 1)
                return
            if keycode == wx.WXK_DOWN:
                self._jump_to_comment(self._comment_index + 1)
                return
            if keycode == wx.WXK_HOME:
                self._jump_to_comment(0)
                return
            if keycode == wx.WXK_END:
                self._jump_to_comment(len(self._comment_displays) - 1)
                return

        event.Skip()

    # ── 상황별 팝업 메뉴 ──

    def _on_body_context_menu(self, event):
        """본문 영역 팝업 메뉴 — 게시물/본문 관련 액션 모음.

        이전/다음 게시물, 글자 이동 같은 내비게이션 항목은 제외.
        """
        menu = wx.Menu()

        id_save_body = wx.NewIdRef()
        menu.Append(id_save_body, "본문 텍스트 저장(&B)\tB")
        self.Bind(wx.EVT_MENU, lambda e: self.on_save_body(), id=id_save_body)

        if self.has_files:
            id_save_file = wx.NewIdRef()
            menu.Append(id_save_file, "첨부파일 저장(&S)\tAlt+S")
            self.Bind(wx.EVT_MENU, lambda e: self.on_download_file(None), id=id_save_file)

        id_url_list = wx.NewIdRef()
        menu.Append(id_url_list, "URL 목록 보기(&U)\tCtrl+U")
        self.Bind(wx.EVT_MENU, lambda e: self._show_url_list(), id=id_url_list)

        menu.AppendSeparator()

        id_write_comment = wx.NewIdRef()
        menu.Append(id_write_comment, "댓글 작성(&C)\tC")
        self.Bind(wx.EVT_MENU, lambda e: self.on_write_comment(), id=id_write_comment)

        menu.AppendSeparator()

        if self.content.edit_url:
            id_edit = wx.NewIdRef()
            menu.Append(id_edit, "게시물 수정(&M)\tAlt+M")
            self.Bind(wx.EVT_MENU, lambda e: self.on_post_edit(None), id=id_edit)
        if self.content.delete_url:
            id_del = wx.NewIdRef()
            menu.Append(id_del, "게시물 삭제(&D)\tAlt+D")
            self.Bind(wx.EVT_MENU, lambda e: self.on_post_delete(None), id=id_del)
        if self.content.reply_url:
            id_reply = wx.NewIdRef()
            menu.Append(id_reply, "게시물 답변(&R)\tAlt+R")
            self.Bind(wx.EVT_MENU, lambda e: self.on_post_reply(None), id=id_reply)

        # 이전/다음 게시물 이동 — 팝업에서도 바로 이동할 수 있도록.
        if self.content.prev_url or self.content.next_url:
            menu.AppendSeparator()
            if self.content.prev_url:
                id_prev = wx.NewIdRef()
                menu.Append(id_prev, "이전 게시물(&B)\tAlt+B")
                self.Bind(wx.EVT_MENU, lambda e: self._goto_prev_post(), id=id_prev)
            if self.content.next_url:
                id_next = wx.NewIdRef()
                menu.Append(id_next, "다음 게시물(&N)\tAlt+N")
                self.Bind(wx.EVT_MENU, lambda e: self._goto_next_post(), id=id_next)

        self.PopupMenu(menu)
        menu.Destroy()

    def _goto_prev_post(self):
        if self.content.prev_url:
            self.navigate_result = "prev"
            self.EndModal(wx.ID_BACKWARD)

    def _goto_next_post(self):
        if self.content.next_url:
            self.navigate_result = "next"
            self.EndModal(wx.ID_FORWARD)

    def _on_comment_context_menu(self, event):
        """댓글 목록 팝업 메뉴 — 댓글 관련 액션 모음.

        댓글 이동(위/아래/처음/끝)은 내비게이션이라 제외.
        """
        menu = wx.Menu()

        id_write = wx.NewIdRef()
        menu.Append(id_write, "댓글 작성(&C)\tC")
        self.Bind(wx.EVT_MENU, lambda e: self.on_write_comment(), id=id_write)

        id_edit = wx.NewIdRef()
        menu.Append(id_edit, "댓글 수정(&M)\tM")
        self.Bind(wx.EVT_MENU, lambda e: self.on_edit_comment(), id=id_edit)

        id_del = wx.NewIdRef()
        menu.Append(id_del, "댓글 삭제(&D)\tAlt+D")
        self.Bind(wx.EVT_MENU, lambda e: self.on_delete_comment(), id=id_del)

        menu.AppendSeparator()

        id_sort = wx.NewIdRef()
        label = "댓글 최신순 보기(&N)" if not self.comment_reversed else "댓글 오래된순 보기(&N)"
        menu.Append(id_sort, f"{label}\tN")
        self.Bind(wx.EVT_MENU, lambda e: self.on_toggle_comment_sort(), id=id_sort)

        self.PopupMenu(menu)
        menu.Destroy()

    # ── 댓글 목록 이동 ──

    def _jump_to_comment(self, new_index: int):
        """댓글 목록에서 다른 댓글로 이동.

        댓글 본문이 여러 줄일 수 있어, SetValue 후 커서 위치 0 기준으로는
        스크린리더가 첫 글자만 읽는 문제가 있다. 메인 메뉴의 Home/End 와
        동일한 "비움 → speak → 복원" 패턴으로 전체 내용을 한 번에 낭독.
        Up/Down 에도 같은 방식을 적용해 일관된 낭독이 되도록 한다.
        """
        if not self._comment_displays:
            return
        n = len(self._comment_displays)
        new_index = max(0, min(new_index, n - 1))
        self._comment_index = new_index
        display = self._comment_displays[new_index]
        # 1) TextCtrl 을 비워 스크린리더가 커서 위치 글자를 읽지 않게 함
        self.comment_ctrl.ChangeValue("")
        # 2) 전체 텍스트를 직접 낭독.
        #    줄바꿈 문자를 ". " 로 바꾸어 스크린리더가 줄 끝마다 "빈줄" 을 읽지
        #    않도록 한다. (NVDA·센스리더 등 한글 스크린리더가 다줄 텍스트의
        #    각 \n 을 빈 줄로 인식해 "빈줄" 을 발화하는 문제 회피.)
        spoken = re.sub(r"\n+", ". ", display).strip()
        speak(spoken)
        # 3) 스크린리더의 키 처리가 끝난 뒤 화면 표시 복원
        wx.CallLater(80, self._restore_comment_display, display)
        self._update_comment_buttons()

    def _restore_comment_display(self, display: str):
        """_jump_to_comment 에서 비웠던 댓글 내용을 화면에 다시 표시.

        ChangeValue 를 써서 스크린리더 재낭독을 유발하지 않는다. 또 다줄
        텍스트를 한 줄로 합쳐(`\\n` → ` · `) Windows 다줄 edit 컨트롤이 SR 에
        끝빈줄("빈줄") 을 알리는 것을 막는다.
        """
        visual = re.sub(r"\n+", " · ", display).strip()
        self.comment_ctrl.ChangeValue(visual)
        self.comment_ctrl.SetInsertionPoint(0)

    def _update_comment_buttons(self):
        """현재 선택된 댓글의 수정/삭제 버튼 활성 상태 갱신."""
        if not self.has_comments:
            return
        comment = self._get_selected_comment_raw(self._comment_index)
        if comment:
            self.comment_edit_btn.Enable(bool(comment.edit_url))
            self.comment_delete_btn.Enable(bool(comment.delete_url))
        else:
            self.comment_edit_btn.Enable(False)
            self.comment_delete_btn.Enable(False)

    # ── 댓글 정렬 ──

    def on_toggle_comment_sort(self):
        self.comment_reversed = not self.comment_reversed
        self._refresh_comment_list()
        if self.comment_reversed:
            speak("댓글 역순")
        else:
            speak("댓글 등록순")

    def _refresh_comment_list(self):
        if not self.has_comments:
            return
        comments = self.content.comments
        if self.comment_reversed:
            comments = list(reversed(comments))

        display = []
        for c in comments:
            header = ""
            if c.author:
                header = f"{c.author}님"
            if c.date:
                header = f"{header} {c.date}" if header else c.date

            # 본문 중복 제거
            body = c.body or ""
            # 작성자/날짜가 본문 앞에 중복 포함되어 있으면 제거
            if c.author and body.startswith(c.author):
                body = body[len(c.author):].strip()
            if body.startswith("님"):
                body = body[1:].strip()
            if c.date and body.startswith(c.date):
                body = body[len(c.date):].strip()
            # 본문이 정확히 2번 반복되는 경우 (예: "감사합니다감사합니다")
            if len(body) >= 2 and len(body) % 2 == 0:
                half = len(body) // 2
                if body[:half] == body[half:]:
                    body = body[:half]
            # 본문 앞부분이 뒤에서 반복되는 경우도 체크
            elif len(body) >= 4:
                for cut in range(3, len(body) // 2 + 1):
                    if body[cut:].startswith(body[:cut]):
                        body = body[:cut] + body[cut + cut:]
                        break

            # 본문 양 끝의 공백·줄바꿈을 정리하고 내부 연속 빈 줄을 한 줄로
            # 압축한다 — 그러지 않으면 본문 끝 빈 줄을 스크린리더가 "빈줄"
            # 이라고 따로 읽어 댓글 낭독이 끊겨 보인다.
            body = (body or "").strip()
            body = re.sub(r"\n\s*\n+", "\n", body)

            if body:
                # 머리말(작성자·날짜)과 본문은 줄바꿈으로 구분해, 본문의 줄바꿈이
                # 그대로 보존되도록 한다. comment_ctrl은 TE_MULTILINE 이므로
                # 화면에도 여러 줄로 표시된다.
                raw = f"{header}\n{body}" if header else body
            elif header:
                raw = header
            else:
                raw = "(빈 댓글)"
            # 줄바꿈 정규화만 수행 (탭은 공백으로). 본문의 \n 은 보존해서 다줄 표시.
            raw = raw.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ").rstrip()
            display.append(raw)

        total = len(display)
        # 위치 정보는 마지막 줄에 별도 표시 — 본문 글자 낭독을 방해하지 않도록.
        # text.rstrip() 로 본문과 위치 사이에 빈 줄이 끼지 않게 한다.
        display_with_index = [
            f"{text.rstrip()}\n({i + 1}/{total})" for i, text in enumerate(display)
        ]

        self._comment_displays = display_with_index
        if display_with_index:
            # 현재 인덱스를 범위 내로 보정 (정렬 토글 후에도 동일 위치 유지)
            self._comment_index = min(self._comment_index, len(display_with_index) - 1)
            # SetValue 로 설정해 포커스가 comment_ctrl 에 들어왔을 때 스크린리더가
            # 내용을 인식하도록 한다. (ChangeValue 는 EVT_TEXT 를 발사하지 않아
            # 스크린리더가 변경을 놓칠 수 있음.)
            # 다줄 텍스트를 한 줄로 합쳐 SR 가 끝빈줄을 "빈줄"로 발화하는 문제 회피.
            visual = re.sub(
                r"\n+", " · ",
                display_with_index[self._comment_index],
            ).strip()
            self.comment_ctrl.SetValue(visual)
            self.comment_ctrl.SetInsertionPoint(0)
            self._update_comment_buttons()
        else:
            self._comment_index = 0
            self.comment_ctrl.ChangeValue("")
            self.comment_edit_btn.Enable(False)
            self.comment_delete_btn.Enable(False)

    def _get_selected_comment_raw(self, sel: int) -> CommentItem | None:
        """인덱스로 댓글 반환 (정렬 상태 반영)"""
        comments = self.content.comments
        if self.comment_reversed:
            comments = list(reversed(comments))
        if 0 <= sel < len(comments):
            return comments[sel]
        return None

    def _get_selected_comment(self) -> CommentItem | None:
        """현재 선택된 댓글 (정렬 상태 반영)"""
        if not self.has_comments:
            return None
        return self._get_selected_comment_raw(self._comment_index)

    # ── 본문 저장 ──

    def on_save_body(self):
        safe_title = self.content.title
        for ch in r'\/:*?"<>|':
            safe_title = safe_title.replace(ch, "_")

        dlg = wx.FileDialog(
            self, "본문 저장", defaultFile=f"{safe_title}.txt",
            wildcard="텍스트 파일 (*.txt)|*.txt",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
        )
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return
        path = dlg.GetPath()
        dlg.Destroy()

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"제목: {self.content.title}\n")
                if self.content.author:
                    f.write(f"작성자: {self.content.author}\n")
                if self.content.date:
                    f.write(f"날짜: {self.content.date}\n")
                f.write("─" * 40 + "\n\n")
                f.write(self.content.body + "\n")
            speak("본문 저장이 완료되었습니다.")
            wx.MessageBox(f"본문이 저장되었습니다.\n{path}",
                          "저장 완료", wx.OK | wx.ICON_INFORMATION, self)
        except Exception as e:
            speak(f"저장에 실패했습니다.")
            wx.MessageBox(f"저장에 실패했습니다.\n{e}",
                          "오류", wx.OK | wx.ICON_ERROR, self)

    # ── 첨부파일 다운로드 ──

    @staticmethod
    def _clean_filename(name: str) -> str:
        """파일명에서 용량 정보 제거"""
        import re
        # 괄호 포함: "(378byte)", "(21.6KB)", "(1.2M)" 등
        cleaned = re.sub(r'\s*\(\d+[\.\d]*\s*[BbKkMmGg][Bb]?[Yy]?[Tt]?[Ee]?[Ss]?\)\s*$', '', name).strip()
        # 괄호 없이: "378byte", "21.6k", "1.2MB" 등
        cleaned = re.sub(r'\s+\d+[\.\d]*\s*[BbKkMmGg][Bb]?[Yy]?[Tt]?[Ee]?[Ss]?\s*$', '', cleaned).strip()
        return cleaned if cleaned else name

    def on_download_file(self, event):
        """모든 첨부파일을 다운로드 폴더에 자동 저장"""
        if not self.has_files:
            speak("첨부파일이 없습니다.")
            return

        from config import get_download_dir
        from main_frame import download_list
        try:
            from sound import play_event
        except Exception:
            play_event = None

        download_dir = get_download_dir()
        total = len(self.content.files)
        speak(f"첨부파일 다운로드를 시작합니다. {total}개 파일")
        if play_event:
            try:
                play_event("download_start")
            except Exception:
                pass

        def _beep(freq):
            try:
                import winsound
                winsound.Beep(freq, 100)
            except Exception:
                pass

        def worker():
            success = 0
            fail = 0
            for fi_idx, fi in enumerate(self.content.files):
                url = fi["url"]
                raw_name = fi["name"]
                clean_name = self._clean_filename(raw_name)
                if not url.startswith("http"):
                    url = f"{SORISEM_BASE_URL}{url}"
                save_path = os.path.join(download_dir, clean_name)

                dl_entry = {
                    "name": clean_name, "size": 0,
                    "downloaded": 0, "status": "다운로드 중",
                }
                download_list.append(dl_entry)

                try:
                    resp = self.session.get(url, stream=True, timeout=30)
                    total_size = int(resp.headers.get("content-length", 0))
                    dl_entry["size"] = total_size
                    downloaded = 0
                    last_pct = 0
                    with open(save_path, "wb") as f:
                        for chunk in resp.iter_content(8192):
                            f.write(chunk)
                            downloaded += len(chunk)
                            dl_entry["downloaded"] = downloaded

                            # 10% 단위로 비프음 (주파수 점점 올라감)
                            if total_size > 0:
                                pct = int(downloaded / total_size * 100)
                                if pct >= last_pct + 10:
                                    last_pct = (pct // 10) * 10
                                    freq = 400 + last_pct * 6  # 400Hz ~ 1000Hz
                                    _beep(freq)

                    dl_entry["status"] = "완료"
                    _beep(1200)  # 완료음
                    success += 1
                except Exception:
                    dl_entry["status"] = "실패"
                    fail += 1

            if fail == 0:
                wx.CallAfter(speak, f"첨부파일 다운로드가 완료되었습니다. {success}개 파일")
            else:
                wx.CallAfter(speak, f"다운로드 완료: 성공 {success}개, 실패 {fail}개")
            if play_event:
                try:
                    play_event("download_complete")
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()

    # ── 댓글 작성 ──

    def on_write_comment(self):
        if not self.content.bo_table or not self.content.wr_id:
            speak("이 게시물에는 댓글을 작성할 수 없습니다.")
            return

        dlg = CommentDialog(self, "댓글 작성")
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return
        text = dlg.get_text()
        dlg.Destroy()
        if not text:
            return

        speak("댓글을 등록하는 중입니다.")

        def worker():
            try:
                self.session.post(
                    f"{SORISEM_BASE_URL}/bbs/write_comment_update.php",
                    data={
                        "bo_table": self.content.bo_table,
                        "wr_id": self.content.wr_id,
                        "wr_content": text, "w": "c",
                    }, timeout=15)
                wx.CallAfter(speak, "댓글이 등록되었습니다.")
                wx.CallAfter(wx.MessageBox, "댓글이 등록되었습니다.",
                             "완료", wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.CallAfter(speak, f"댓글 등록 실패. {e}")
                wx.CallAfter(wx.MessageBox, f"댓글 등록 실패.\n{e}",
                             "오류", wx.OK | wx.ICON_ERROR)

        threading.Thread(target=worker, daemon=True).start()

    # ── 댓글 수정 ──

    def on_edit_comment(self, event=None):
        comment = self._get_selected_comment()
        if not comment:
            speak("수정할 댓글을 선택해 주세요.")
            return
        if not comment.edit_url and not (self.content.bo_table and comment.comment_id):
            speak("이 댓글은 수정할 수 없습니다.")
            return

        dlg = CommentDialog(self, "댓글 수정", comment.body)
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return
        new_text = dlg.get_text()
        dlg.Destroy()
        if not new_text:
            return

        speak("댓글을 수정하는 중입니다.")

        def worker():
            try:
                self.session.post(
                    f"{SORISEM_BASE_URL}/bbs/write_comment_update.php",
                    data={
                        "bo_table": self.content.bo_table,
                        "wr_id": self.content.wr_id,
                        "comment_id": comment.comment_id,
                        "wr_content": new_text, "w": "cu",
                    }, timeout=15)
                comment.body = new_text
                wx.CallAfter(self._refresh_comment_list)
                wx.CallAfter(speak, "댓글이 수정되었습니다.")
                wx.CallAfter(wx.MessageBox, "댓글이 수정되었습니다.",
                             "완료", wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.CallAfter(speak, f"댓글 수정 실패. {e}")
                wx.CallAfter(wx.MessageBox, f"댓글 수정 실패.\n{e}",
                             "오류", wx.OK | wx.ICON_ERROR)

        threading.Thread(target=worker, daemon=True).start()

    # ── 댓글 삭제 ──

    def _verify_comment_deleted(self, comment) -> bool:
        """게시물을 다시 로드해서 해당 comment_id가 사라졌는지 확인.
        서버가 응답 alert을 명확히 주지 않아도 실제 삭제 여부를 판정하기 위한 체크."""
        if not (self.content.bo_table and self.content.wr_id and comment.comment_id):
            return False
        try:
            vurl = (
                f"{SORISEM_BASE_URL}/bbs/board.php"
                f"?bo_table={self.content.bo_table}"
                f"&wr_id={self.content.wr_id}"
            )
            vresp = self.session.get(vurl, timeout=15)
            html = vresp.text or ""
        except Exception:
            return False
        cid = comment.comment_id
        # id="c_123" 형식의 article 속성만 검사 (본문 등에 숫자 ID가 우연히 들어가는
        # 오탐 방지). 작은 따옴표 / 큰 따옴표 모두 대응.
        if f'id="{cid}"' in html or f"id='{cid}'" in html:
            return False
        return True

    def _attempt_comment_delete(self, comment, delete_url: str,
                                allow_token_retry: bool) -> None:
        """실제 댓글 삭제 요청 수행. alert 감지 → 실패 시 1회 재시도."""
        import html as _html
        import re as _re
        from urllib.parse import urljoin, urlsplit, parse_qsl

        if self.content.bo_table and self.content.wr_id:
            referer = (
                f"{SORISEM_BASE_URL}/bbs/board.php"
                f"?bo_table={self.content.bo_table}"
                f"&wr_id={self.content.wr_id}"
            )
        else:
            referer = SORISEM_BASE_URL + "/"

        raw = _html.unescape(delete_url or "")
        if raw.startswith("//"):
            url = "https:" + raw
        elif raw.startswith(("http://", "https://")):
            url = raw
        else:
            url = urljoin(referer, raw)

        # URL이 write_comment_update.php 같은 공용 작성·수정 엔드포인트를 가리킨다면
        # GET 요청이 "댓글을 입력하여 주십시오"로 반려되므로, 쿼리 파라미터를 그대로
        # POST body로 옮겨서 재전송한다. (w=cd 같은 삭제 플래그가 URL에 들어 있음)
        path_lower = urlsplit(url).path.lower()
        is_write_endpoint = (
            "write_comment_update.php" in path_lower
            or "write_comment.php" in path_lower
            or path_lower.endswith("/write.php")
        )

        try:
            if is_write_endpoint:
                parts = urlsplit(url)
                params = dict(parse_qsl(parts.query, keep_blank_values=True))
                # 최소한의 식별 정보 보강
                if self.content.bo_table:
                    params.setdefault("bo_table", self.content.bo_table)
                if self.content.wr_id:
                    params.setdefault("wr_id", str(self.content.wr_id))
                if comment.comment_id:
                    cid = comment.comment_id.replace("c_", "")
                    params.setdefault("comment_id", cid)
                post_url = f"{parts.scheme}://{parts.netloc}{parts.path}"
                resp = self.session.post(
                    post_url, data=params,
                    headers={"Referer": referer},
                    timeout=15, allow_redirects=True,
                )
            else:
                resp = self.session.get(
                    url, headers={"Referer": referer},
                    timeout=15, allow_redirects=True,
                )
        except Exception as e:
            wx.CallAfter(speak, f"댓글 삭제 실패. {e}")
            wx.CallAfter(wx.MessageBox, f"댓글 삭제 실패.\n{e}",
                         "오류", wx.OK | wx.ICON_ERROR)
            return

        body = resp.text or ""
        alert_match = _re.search(r"""alert\(\s*['"]([^'"]+)['"]\s*\)""", body)

        # 1. 명확한 성공 alert → 바로 완료 처리
        if alert_match:
            msg = alert_match.group(1)
            if "삭제" in msg and ("되었" in msg or "완료" in msg):
                self._finish_comment_delete(comment)
                return

        # 2. 서버 쪽 실제 삭제 여부를 확인 (alert이 모호하거나 HTTP 에러여도
        #    POST/GET 과정에서 실제 삭제가 이루어졌을 수 있음).
        if self._verify_comment_deleted(comment):
            self._finish_comment_delete(comment)
            return

        # 3. 서버에서 아직 댓글이 남아 있음 → 실제 실패. 첫 시도면 재시도.
        if allow_token_retry:
            self._retry_comment_delete_with_fresh_token(comment)
            return

        # 4. 최종 실패 → 원인 표시
        if alert_match:
            msg = alert_match.group(1)
            wx.CallAfter(speak, f"댓글 삭제 실패. {msg}")
            wx.CallAfter(
                wx.MessageBox,
                f"댓글 삭제 실패.\n{msg}\n\n"
                f"원본 링크: {delete_url}\n요청 URL: {url}",
                "오류", wx.OK | wx.ICON_ERROR,
            )
            return
        if resp.status_code >= 400:
            wx.CallAfter(speak, f"댓글 삭제 실패. HTTP {resp.status_code}")
            wx.CallAfter(
                wx.MessageBox,
                f"댓글 삭제 실패 (HTTP {resp.status_code})\n\n"
                f"원본 링크: {delete_url}\n요청 URL: {url}",
                "오류", wx.OK | wx.ICON_ERROR,
            )
            return
        wx.CallAfter(speak, "댓글 삭제 실패.")
        wx.CallAfter(
            wx.MessageBox,
            "서버에서 댓글이 삭제되지 않았습니다.\n"
            f"원본 링크: {delete_url}\n요청 URL: {url}",
            "오류", wx.OK | wx.ICON_ERROR,
        )

    def _finish_comment_delete(self, comment):
        """삭제 후 목록/UI 갱신."""
        try:
            self.content.comments.remove(comment)
        except ValueError:
            pass
        wx.CallAfter(self._refresh_comment_list)
        wx.CallAfter(speak, "댓글이 삭제되었습니다.")
        wx.CallAfter(wx.MessageBox, "댓글이 삭제되었습니다.",
                     "완료", wx.OK | wx.ICON_INFORMATION)

    def _retry_comment_delete_with_fresh_token(self, comment):
        """토큰 에러 시 게시물을 재로드해서 이 댓글(comment_id)의 최신 delete URL을
        뽑아낸 뒤 1회 재시도."""
        speak("토큰을 갱신하여 다시 시도합니다.")

        def worker():
            import re as _re
            import html as _html_mod
            from page_parser import _unwrap_js_url

            try:
                post_url = (
                    f"{SORISEM_BASE_URL}/bbs/board.php"
                    f"?bo_table={self.content.bo_table}"
                    f"&wr_id={self.content.wr_id}"
                )
                resp_page = self.session.get(post_url, timeout=15)
                html = resp_page.text or ""

                cid = comment.comment_id or ""
                fresh_url = ""

                # 1) comment_id 블록 범위에서 delete_comment 링크 탐색
                if cid:
                    block_match = _re.search(
                        r'id=["\']' + _re.escape(cid) + r'["\'].*?</article>',
                        html, _re.DOTALL,
                    )
                    if block_match:
                        block = block_match.group(0)
                        m = _re.search(
                            r'href=["\']([^"\']*(?:delete_comment\.php|delete\.php)[^"\']*)["\']',
                            block,
                        )
                        if m:
                            fresh_url = _unwrap_js_url(_html_mod.unescape(m.group(1)))

                # 2) fallback: delete_comment.php URL 중 comment_id 일치하는 것
                if not fresh_url and cid:
                    cid_num = cid.replace("c_", "")
                    for m in _re.finditer(
                        r'href=["\']([^"\']*delete_comment\.php[^"\']*)["\']',
                        html,
                    ):
                        candidate = _html_mod.unescape(m.group(1))
                        if f"comment_id={cid_num}" in candidate or f"comment_id={cid}" in candidate:
                            fresh_url = _unwrap_js_url(candidate)
                            break

                if not fresh_url:
                    # 페이지에서 delete_comment.php 링크를 못 찾았다면 이미 서버에서
                    # 삭제되어 해당 comment_id가 사라졌을 가능성이 크다. 검증 후
                    # 삭제됐으면 성공 처리, 아니면 오류 표시.
                    if comment.comment_id and f'id="{comment.comment_id}"' not in html \
                            and f"id='{comment.comment_id}'" not in html:
                        self._finish_comment_delete(comment)
                        return
                    wx.CallAfter(speak, "삭제 URL을 찾을 수 없습니다.")
                    wx.CallAfter(
                        wx.MessageBox,
                        "페이지에서 삭제 링크를 찾지 못했습니다.",
                        "오류", wx.OK | wx.ICON_ERROR,
                    )
                    return

                self._attempt_comment_delete(
                    comment, fresh_url, allow_token_retry=False,
                )
            except Exception as e:
                wx.CallAfter(speak, f"댓글 삭제 실패. {e}")
                wx.CallAfter(wx.MessageBox, f"댓글 삭제 실패.\n{e}",
                             "오류", wx.OK | wx.ICON_ERROR)

        threading.Thread(target=worker, daemon=True).start()

    def on_delete_comment(self):
        comment = self._get_selected_comment()
        if not comment:
            speak("삭제할 댓글을 선택해 주세요.")
            return

        r = wx.MessageBox(
            f"'{comment.author}'님의 댓글을 삭제하시겠습니까?\n\n{comment.body}",
            "댓글 삭제", wx.YES_NO | wx.ICON_QUESTION, self)
        if r != wx.YES:
            return

        if not comment.delete_url:
            speak("이 댓글은 삭제할 수 없습니다.")
            return

        speak("댓글을 삭제하는 중입니다.")

        def worker():
            self._attempt_comment_delete(
                comment, comment.delete_url, allow_token_retry=True,
            )

        threading.Thread(target=worker, daemon=True).start()

    # ── 게시물 수정/삭제/답변 ──

    def _names_match(self, a: str | None, b: str | None) -> bool:
        """닉네임 두 개가 같은 사람을 가리키는지 판단. 대소문자·공백·존칭 정규화."""
        def norm(s: str | None) -> str:
            if not s:
                return ""
            s = re.sub(r"<[^>]+>", "", s).strip()
            for suf in ("님의 정보", "님 정보", "님", "씨"):
                if s.endswith(suf):
                    s = s[: -len(suf)].strip()
            s = re.sub(r"\s+", " ", s)
            return s.lower()
        na, nb = norm(a), norm(b)
        if not na or not nb:
            return False
        if na == nb:
            return True
        if na in nb or nb in na:
            return True
        return False

    def _verify_own_post(self) -> tuple[bool, str]:
        """이 게시물이 현재 로그인 사용자의 글인지 확인.

        판단 순서 (main_frame._verify_post_ownership 과 동일한 정책):
          1) 로그인 닉네임과 게시물 작성자 닉네임 비교 (HTTP 호출 없음, 가장 빠름)
          2) 본문 페이지에서 작성자 mb_id 추출 후 current_user_id 와 비교
          3) 두 방법 모두 판단 불가 → 서버 기본 권한 체크에 맡기기 위해 허용.

        반환: (True, "") 본인 / (False, 사유) 타인.
        """
        if not self.current_user_id:
            return False, "로그인 사용자 정보를 확인할 수 없어 수정·삭제를 중단합니다."
        if not self.content.bo_table or not self.content.wr_id:
            return False, "게시물 식별자를 확인할 수 없어 수정·삭제를 중단합니다."

        # 1) 닉네임 비교
        if self.current_user_nickname and self.content.author:
            if self._names_match(self.current_user_nickname, self.content.author):
                return True, ""
            return False, "본인이 작성한 게시물만 수정·삭제할 수 있습니다."

        # 2) mb_id 추출 비교
        try:
            post_url = (
                f"{SORISEM_BASE_URL}/bbs/board.php?"
                f"bo_table={self.content.bo_table}&wr_id={self.content.wr_id}"
            )
            resp = self.session.get(post_url, timeout=15)
        except Exception:
            return True, ""

        from page_parser import extract_post_author_id
        author_id = extract_post_author_id(resp.text)
        if author_id:
            if author_id.strip().lower() == self.current_user_id.strip().lower():
                return True, ""
            return False, "본인이 작성한 게시물만 수정·삭제할 수 있습니다."

        # 3) 판단 불가 — 서버 기본 권한 체크에 맡김
        return True, ""

    def on_post_edit(self, event):
        """게시물 수정"""
        if not self.content.edit_url:
            speak("이 게시물은 수정할 수 없습니다. 본인이 작성한 글만 수정할 수 있습니다.")
            return

        # 클라이언트측 작성자 검증: 서버가 관리자에게 수정 URL 을 제공해도
        # 본인 글이 아니면 여기서 차단.
        def verify_then_fetch():
            owned, msg = self._verify_own_post()
            if not owned:
                wx.CallAfter(speak, msg)
                wx.CallAfter(
                    wx.MessageBox, msg, "수정 불가",
                    wx.OK | wx.ICON_WARNING, self,
                )
                return
            wx.CallAfter(self._do_post_edit)

        speak("수정 페이지를 불러오는 중입니다.")
        threading.Thread(target=verify_then_fetch, daemon=True).start()

    def _do_post_edit(self):
        """검증 통과 후 실제 수정 페이지 로드."""

        def worker():
            try:
                url = self.content.edit_url
                if not url.startswith("http"):
                    url = f"{SORISEM_BASE_URL}{url}"
                resp = self.session.get(url, timeout=15)

                from bs4 import BeautifulSoup as _BS
                soup = _BS(resp.text, "html.parser")
                title_input = soup.find("input", {"name": "wr_subject"})
                body_area = soup.find("textarea", {"name": "wr_content"})

                # 편집 폼이 없으면 권한 없음/에러 → alert 메시지 확인
                if not title_input and not body_area:
                    import re as _re
                    alert_match = _re.search(
                        r'alert\(["\'](.+?)["\']\)', resp.text
                    )
                    if alert_match:
                        wx.CallAfter(speak, f"수정 불가: {alert_match.group(1)}")
                    else:
                        wx.CallAfter(
                            speak,
                            "수정할 수 없습니다. 본인이 작성한 글만 수정 가능합니다.",
                        )
                    return

                old_title = title_input.get("value", "") if title_input else ""
                old_body = body_area.get_text() if body_area else ""

                wx.CallAfter(self._show_edit_dialog, old_title, old_body)
            except Exception as e:
                wx.CallAfter(speak, f"수정 페이지를 불러올 수 없습니다. {e}")

        threading.Thread(target=worker, daemon=True).start()

    def _show_edit_dialog(self, old_title: str, old_body: str):
        from write_dialog import WriteDialog
        dialog = WriteDialog(
            self, self.session, self.content.bo_table,
            existing_title=old_title, existing_body=old_body,
        )
        dialog._edit_wr_id = self.content.wr_id

        result = dialog.ShowModal()
        dialog.Destroy()

        if result == wx.ID_OK:
            speak("게시물이 수정되었습니다.")
            self.navigate_result = "refresh"
            self.EndModal(wx.ID_OK)

    def on_post_delete(self, event):
        """게시물 삭제"""
        if not self.content.delete_url:
            speak("이 게시물은 삭제할 수 없습니다. 본인이 작성한 글만 삭제할 수 있습니다.")
            return

        # 클라이언트측 작성자 검증. 확인 대화상자 이전에 수행해야 "확인 → 본인
        # 아님" 순서의 어색한 흐름을 피할 수 있다. 네트워크 요청이 필요하므로
        # 별도 스레드에서 돌린 뒤 메인 스레드로 복귀해 대화상자를 띄운다.
        def verify_then_confirm():
            owned, msg = self._verify_own_post()
            if not owned:
                wx.CallAfter(speak, msg)
                wx.CallAfter(
                    wx.MessageBox, msg, "삭제 불가",
                    wx.OK | wx.ICON_WARNING, self,
                )
                return
            wx.CallAfter(self._confirm_and_delete_post)

        threading.Thread(target=verify_then_confirm, daemon=True).start()

    def _confirm_and_delete_post(self):
        """작성자 검증 통과 후 확인 대화상자 → 실제 삭제."""
        result = wx.MessageBox(
            f"'{self.content.title}' 게시물을 삭제하시겠습니까?\n\n"
            "삭제하면 복구할 수 없습니다.",
            "게시물 삭제", wx.YES_NO | wx.ICON_WARNING, self,
        )
        if result != wx.YES:
            return

        speak("게시물을 삭제하는 중입니다.")

        delete_url_saved = self.content.delete_url

        def worker():
            try:
                import re as _re

                # delete_url을 직접 사용 (BS4가 이미 &amp; → & 변환)
                url = delete_url_saved
                if not url.startswith("http"):
                    url = f"{SORISEM_BASE_URL}{url}"

                # Referer 헤더 추가
                referer = url.replace("/bbs/delete.php", "/bbs/board.php")
                referer = _re.sub(r'[&?]token=[^&]*', '', referer)
                resp = self.session.get(url, headers={"Referer": referer}, timeout=15)

                alert_match = _re.search(r'alert\(["\'](.+?)["\']\)', resp.text)
                if alert_match:
                    msg = alert_match.group(1)
                    if "검색어" in msg:
                        # 삭제와 무관한 alert → 삭제 성공으로 처리
                        wx.CallAfter(self._post_delete_done)
                    elif "토큰" in msg or "token" in msg.lower():
                        wx.CallAfter(self._retry_delete_with_fresh_token)
                    else:
                        wx.CallAfter(speak, f"삭제 불가: {msg}")
                        wx.CallAfter(wx.MessageBox,
                                     f"삭제할 수 없습니다.\n{msg}",
                                     "삭제 불가", wx.OK | wx.ICON_WARNING)
                else:
                    wx.CallAfter(self._post_delete_done)
            except Exception as e:
                wx.CallAfter(speak, f"삭제에 실패했습니다. {e}")

        threading.Thread(target=worker, daemon=True).start()

    def _retry_delete_with_fresh_token(self):
        """토큰 에러 시 페이지를 다시 로드하여 최신 토큰으로 재시도"""
        speak("토큰을 갱신하여 다시 시도합니다.")

        def worker():
            try:
                import re as _re
                import html as _html_mod

                # 게시물 페이지 다시 로드
                post_url = (
                    f"{SORISEM_BASE_URL}/bbs/board.php"
                    f"?bo_table={self.content.bo_table}&wr_id={self.content.wr_id}"
                )
                resp_page = self.session.get(post_url, timeout=15)

                # 최신 삭제 URL 추출
                delete_match = _re.search(
                    r'href=["\']([^"\']*delete\.php[^"\']*)["\']',
                    resp_page.text
                )
                if not delete_match:
                    wx.CallAfter(speak, "삭제할 수 없습니다.")
                    return

                fresh_url = _html_mod.unescape(delete_match.group(1))
                if not fresh_url.startswith("http"):
                    fresh_url = f"{SORISEM_BASE_URL}{fresh_url}"

                resp = self.session.get(
                    fresh_url,
                    headers={"Referer": post_url},
                    timeout=15,
                )

                alert_match = _re.search(r'alert\(["\'](.+?)["\']\)', resp.text)
                if alert_match and "검색어" not in alert_match.group(1):
                    wx.CallAfter(speak, f"삭제 불가: {alert_match.group(1)}")
                    wx.CallAfter(wx.MessageBox,
                                 f"삭제할 수 없습니다.\n{alert_match.group(1)}",
                                 "삭제 불가", wx.OK | wx.ICON_WARNING)
                else:
                    wx.CallAfter(self._post_delete_done)
            except Exception as e:
                wx.CallAfter(speak, f"삭제에 실패했습니다. {e}")

        threading.Thread(target=worker, daemon=True).start()

    def _post_delete_done(self):
        speak("게시물이 삭제되었습니다.")
        wx.MessageBox("게시물이 삭제되었습니다.", "완료",
                      wx.OK | wx.ICON_INFORMATION, self)
        self.navigate_result = "refresh"
        self.EndModal(wx.ID_OK)

    def on_post_reply(self, event):
        """게시물 답변"""
        if not self.content.reply_url:
            speak("이 게시물에는 답변할 수 없습니다.")
            return

        from write_dialog import WriteDialog
        dialog = WriteDialog(self, self.session, self.content.bo_table)
        dialog.SetTitle("게시물 답변")
        dialog.submit_btn.SetLabel("답변 등록(&W)")
        # 답변 모드: w=r, wr_id 설정
        dialog._reply_wr_id = self.content.wr_id

        result = dialog.ShowModal()
        dialog.Destroy()

        if result == wx.ID_OK:
            speak("답변이 등록되었습니다.")
            self.navigate_result = "refresh"
            self.EndModal(wx.ID_OK)

    # ── 이전/다음 게시물 ──

    def on_prev_post(self, event=None):
        if self.content.prev_url:
            self.navigate_result = "prev"
            self.EndModal(wx.ID_BACKWARD)
        else:
            speak("이전 게시물이 없습니다.")

    def on_next_post(self, event=None):
        if self.content.next_url:
            self.navigate_result = "next"
            self.EndModal(wx.ID_FORWARD)
        else:
            speak("다음 게시물이 없습니다.")

    def on_close(self, event):
        self.EndModal(wx.ID_CANCEL)

    def _show_url_list(self):
        """게시물 본문과 댓글에서 URL을 추출해 목록으로 보여준다."""
        urls = []
        seen = set()
        # 본문
        for _, _, u in _extract_urls(self.content.body or ""):
            if u not in seen:
                seen.add(u)
                urls.append(u)
        # 댓글
        for c in (self.content.comments or []):
            for _, _, u in _extract_urls(c.body or ""):
                if u not in seen:
                    seen.add(u)
                    urls.append(u)

        if not urls:
            speak("이 게시물에는 URL이 없습니다.")
            return

        dlg = wx.SingleChoiceDialog(
            self,
            f"원하는 URL을 선택하고 확인을 누르면 브라우저에서 열립니다.\n총 {len(urls)}개의 URL이 있습니다.",
            "URL 목록 (Ctrl+U)",
            urls,
        )
        dlg.SetSelection(0)

        # 저시력 테마 적용
        try:
            from theme import apply_theme, make_font, load_font_size
            apply_theme(dlg, make_font(load_font_size()))
        except Exception:
            pass

        if dlg.ShowModal() == wx.ID_OK:
            selected = urls[dlg.GetSelection()]
            speak("브라우저에서 엽니다.")
            webbrowser.open(selected)
        dlg.Destroy()
