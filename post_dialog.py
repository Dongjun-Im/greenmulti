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

    def get_text(self) -> str:
        return self.comment_text.GetValue().strip()


class PostDialog(wx.Dialog):
    """게시물 내용을 표시하는 대화상자"""

    def __init__(self, parent, content: PostContent, session: requests.Session):
        title = content.title if content.title else "게시물 내용"
        super().__init__(
            parent, title=title,
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
            size=(700, 550),
        )

        self.content = content
        self.session = session

        # 첨부파일/댓글 유무를 명확하게 판단
        self.has_files = isinstance(content.files, list) and len(content.files) > 0
        self.has_comments = isinstance(content.comments, list) and len(content.comments) > 0
        self.comment_reversed = False
        self.navigate_result = ""  # "prev" or "next"

        self.panel = wx.Panel(self)
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)

        self._create_controls()
        self._do_layout()
        self._fill_content()

        self.panel.SetSizer(self.main_sizer)

        # 키보드 이벤트
        self.Bind(wx.EVT_CHAR_HOOK, self.on_char_hook)

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
        self.body_text = wx.TextCtrl(
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
            self.comment_list = wx.ListBox(
                self.panel, style=wx.LB_SINGLE, name="댓글 목록",
            )
            self.comment_list.Bind(wx.EVT_LISTBOX, self._on_comment_selected)
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
            s.Add(self.comment_list, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 5)
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
            if focused == self.body_text or (self.has_comments and focused == self.comment_list):
                self.on_write_comment()
                return

        # Alt+D 또는 D: 댓글 삭제
        if keycode in (ord("D"), ord("d")):
            if self.has_comments and (alt or focused == self.comment_list):
                self.on_delete_comment()
                return

        # N: 댓글 정렬 순서 변경 (댓글 목록에 포커스)
        if keycode in (ord("N"), ord("n")) and not alt:
            if self.has_comments and focused == self.comment_list:
                self.on_toggle_comment_sort()
                return

        event.Skip()

    # ── 댓글 정렬 ──

    def on_toggle_comment_sort(self):
        self.comment_reversed = not self.comment_reversed
        self._refresh_comment_list()
        if self.comment_reversed:
            speak("댓글 역순")
        else:
            speak("댓글 등록순")

    def _on_comment_selected(self, event):
        """댓글 선택 시 수정/삭제 버튼 활성화 상태 업데이트"""
        comment = self._get_selected_comment()
        if comment:
            # 수정/삭제 URL이 있으면 본인 댓글로 판단
            has_edit = bool(comment.edit_url)
            has_delete = bool(comment.delete_url)
            self.comment_edit_btn.Enable(has_edit)
            self.comment_delete_btn.Enable(has_delete)
        else:
            self.comment_edit_btn.Enable(False)
            self.comment_delete_btn.Enable(False)
        event.Skip()

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

            if body:
                display.append(f"{header} : {body}" if header else body)
            elif header:
                display.append(header)
            else:
                display.append("(빈 댓글)")
        self.comment_list.Set(display)
        if display:
            self.comment_list.SetSelection(0)
            # 첫 댓글 선택에 따른 버튼 상태 업데이트
            comment = self._get_selected_comment_raw(0)
            if comment:
                self.comment_edit_btn.Enable(bool(comment.edit_url))
                self.comment_delete_btn.Enable(bool(comment.delete_url))
            else:
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
        sel = self.comment_list.GetSelection()
        if sel == wx.NOT_FOUND:
            return None
        comments = self.content.comments
        if self.comment_reversed:
            comments = list(reversed(comments))
        if sel < len(comments):
            return comments[sel]
        return None

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

        download_dir = get_download_dir()
        total = len(self.content.files)
        speak(f"첨부파일 다운로드를 시작합니다. {total}개 파일")

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
            try:
                url = comment.delete_url
                if not url.startswith("http"):
                    url = f"{SORISEM_BASE_URL}{url}"
                self.session.get(url, timeout=15)
                self.content.comments.remove(comment)
                wx.CallAfter(self._refresh_comment_list)
                wx.CallAfter(speak, "댓글이 삭제되었습니다.")
                wx.CallAfter(wx.MessageBox, "댓글이 삭제되었습니다.",
                             "완료", wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.CallAfter(speak, f"댓글 삭제 실패. {e}")
                wx.CallAfter(wx.MessageBox, f"댓글 삭제 실패.\n{e}",
                             "오류", wx.OK | wx.ICON_ERROR)

        threading.Thread(target=worker, daemon=True).start()

    # ── 게시물 수정/삭제/답변 ──

    def on_post_edit(self, event):
        """게시물 수정"""
        if not self.content.edit_url:
            speak("이 게시물은 수정할 수 없습니다. 본인이 작성한 글만 수정할 수 있습니다.")
            return

        speak("수정 페이지를 불러오는 중입니다.")

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
