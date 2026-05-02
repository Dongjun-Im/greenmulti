"""메일(formmail) API + 대화상자.

gnuboard5 의 formmail 플러그인은 기본적으로 **발송 전용** (외부 이메일로
보내는 기능). 사이트에 별도 메일함이 없으면 받은 메일 조회 기능은
제공 불가 — 그런 경우 이 모듈은 발송 대화상자만 제공하고 수신함은
"지원되지 않음" 안내로 폴백한다.

엔드포인트:
- GET  /bbs/formmail.php?mb_id=USER   발송 폼 (token 등 hidden 필드 획득)
- POST /bbs/formmail_send.php         발송
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional
from urllib.parse import urljoin

import requests
import wx
from bs4 import BeautifulSoup

from config import (
    SORISEM_BASE_URL,
    MAIL_FORM_URL, MAIL_SEND_URL, APP_ADMIN_EMAIL,
    MAIL_INBOX_URL, MAIL_SENT_URL, MAIL_INBOX_VIEW_URL,
    MAIL_SENT_VIEW_URL, MAIL_WRITE_URL, MAIL_INBOX_BASE,
)
from screen_reader import speak
from theme import apply_theme, make_font, load_font_size
from post_dialog import ContextMenuTextCtrl


@dataclass
class MailItem:
    """메일 수신함 한 줄."""
    mail_id: str
    sender: str
    subject: str
    date: str = ""
    is_read: bool = True


def _is_login_redirect(resp: requests.Response) -> bool:
    """실제 로그인 페이지인지 정확히 판단. memo.py 의 같은 함수와 동일한 로직."""
    final_url = resp.url or ""
    if re.search(r"/bbs/login\.php(\?|$)", final_url):
        return True
    head = resp.text[:4000]
    has_login_title = "<title>소리샘 로그인</title>" in head
    has_login_form = ('name="flogin"' in head or "flogin_submit" in head)
    has_password_input = 'name="mb_password"' in head
    return (has_login_title and has_password_input) or (has_login_form and has_password_input)


def send_mail(session: requests.Session,
              recipient: str,
              subject: str,
              body: str,
              recipient_is_userid: bool = True,
              attachments: list[str] | None = None) -> tuple[bool, str]:
    """메일 발송. gnuboard5 formmail → formmail_send 흐름.

    recipient_is_userid=True 이면 소리샘 회원 ID(mb_id). False 이면 이메일 주소 직접.
    attachments: 첨부할 파일 경로 리스트. 폼이 file input 을 가지고 있으면
        multipart/form-data 로 함께 전송한다. 없으면 첨부 무시.
    """
    # 1. 폼 GET — mb_id 파라미터가 있으면 수신자 이메일이 자동 채워짐
    params = {"mb_id": recipient} if recipient_is_userid else {"email": recipient}
    try:
        form_resp = session.get(MAIL_FORM_URL, params=params, timeout=15)
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if form_resp.status_code != 200:
        return False, f"HTTP {form_resp.status_code}"
    if _is_login_redirect(form_resp):
        return False, "로그인 세션이 만료되었습니다."

    soup = BeautifulSoup(form_resp.text, "lxml")
    form = (soup.find("form", {"name": "fformmail"})
            or soup.find("form", id="fformmail")
            or soup.find("form"))
    if form is None:
        return False, "메일 작성 폼 구조를 인식하지 못했습니다."

    data = {}
    file_field_names: list[str] = []
    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        itype = (inp.get("type") or "").lower()
        if itype == "submit":
            continue
        if itype == "file":
            file_field_names.append(name)
            continue
        data[name] = inp.get("value", "")
    for ta in form.find_all("textarea"):
        name = ta.get("name")
        if name:
            data[name] = ta.get_text() or ""

    # gnuboard5 formmail 필드명: subject, content, email, name 등
    data["subject"] = subject
    data["content"] = body
    # 수신자가 이메일 주소면 email 필드로
    if not recipient_is_userid:
        data["email"] = recipient
    # 수신자가 userid 인데 폼이 email 필드만 기대하면 form GET 시 mb_id 로 채워져 있음

    action = form.get("action") or MAIL_SEND_URL
    post_url = urljoin(form_resp.url, action)

    # 첨부파일 처리: 폼에 file input 이 있으면 multipart 로 같이 전송
    files_payload = []
    open_handles = []
    if attachments:
        # file 필드명이 없는 경우의 폴백 후보 (gnuboard5 자주 쓰는 이름)
        candidates = file_field_names or ["bf_file[]", "bf_file", "files[]", "file"]
        for idx, path in enumerate(attachments):
            try:
                fh = open(path, "rb")
            except OSError:
                continue
            open_handles.append(fh)
            field_name = candidates[idx] if idx < len(candidates) else candidates[-1]
            import os as _os
            files_payload.append(
                (field_name, (_os.path.basename(path), fh, "application/octet-stream"))
            )

    try:
        if files_payload:
            resp = session.post(
                post_url, data=data, files=files_payload,
                timeout=120, allow_redirects=True,
            )
        else:
            resp = session.post(post_url, data=data, timeout=30, allow_redirects=True)
    except requests.RequestException as e:
        return False, f"전송 실패: {e}"
    finally:
        for fh in open_handles:
            try:
                fh.close()
            except Exception:
                pass

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"

    m = re.search(r"alert\(\s*['\"]([^'\"]+)['\"]", resp.text)
    if m:
        msg = m.group(1)
        if any(w in msg for w in ["발송", "보냈", "전송", "성공"]):
            return True, msg
        return False, msg
    return True, ""


class MailWriteDialog(wx.Dialog):
    """메일 작성 / 관리자 메일 / 일반 메일 모두 이 한 대화상자로 처리."""

    MODE_ADMIN = "admin"       # 관리자에게 메일 (수신자 고정)
    MODE_GENERAL = "general"   # 일반 메일 (수신자 입력)

    def __init__(self, parent, session: requests.Session,
                 mode: str = "general",
                 default_recipient: str = "",
                 default_subject: str = "",
                 default_body: str = ""):
        title = "관리자에게 메일 보내기" if mode == self.MODE_ADMIN else "메일 보내기"
        super().__init__(parent, title=title, size=(720, 600),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.mode = mode
        self.attachment_paths: list[str] = []

        # 관리자 모드면 수신자 고정(표시만, 편집 불가)
        if mode == self.MODE_ADMIN:
            default_recipient = APP_ADMIN_EMAIL

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 수신자
        if mode == self.MODE_ADMIN:
            lbl_r = wx.StaticText(panel, label="받는 사람 (관리자, 변경 불가)")
        else:
            lbl_r = wx.StaticText(panel, label="받는 사람 (소리샘 아이디 또는 이메일 주소)")
        vbox.Add(lbl_r, 0, wx.TOP | wx.LEFT | wx.RIGHT, 8)
        self.recipient_ctrl = wx.TextCtrl(panel, value=default_recipient)
        if mode == self.MODE_ADMIN:
            self.recipient_ctrl.SetEditable(False)
        vbox.Add(self.recipient_ctrl, 0, wx.ALL | wx.EXPAND, 8)

        # 제목
        lbl_s = wx.StaticText(panel, label="제목")
        vbox.Add(lbl_s, 0, wx.LEFT | wx.RIGHT, 8)
        self.subject_ctrl = wx.TextCtrl(panel, value=default_subject)
        vbox.Add(self.subject_ctrl, 0, wx.ALL | wx.EXPAND, 8)

        # 본문
        lbl_b = wx.StaticText(panel, label="메일 내용")
        vbox.Add(lbl_b, 0, wx.LEFT | wx.RIGHT, 8)
        self.body_ctrl = wx.TextCtrl(panel, value=default_body,
                                     style=wx.TE_MULTILINE | wx.TE_RICH2)
        vbox.Add(self.body_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 첨부파일 영역 — MailComposeDialog 와 동일한 UI
        atc_label = wx.StaticText(panel, label="첨부파일 (Alt+F 추가 / Del 삭제)")
        vbox.Add(atc_label, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)
        self.attach_list = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE)
        self.attach_list.SetMinSize((-1, 80))
        vbox.Add(self.attach_list, 0, wx.ALL | wx.EXPAND, 8)
        atc_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.add_file_btn = wx.Button(panel, label="파일 추가(&F)")
        self.add_file_btn.Bind(wx.EVT_BUTTON, self.on_add_files)
        atc_btn_sizer.Add(self.add_file_btn, 0, wx.RIGHT, 8)
        self.remove_file_btn = wx.Button(panel, label="선택 파일 제거(&R)")
        self.remove_file_btn.Bind(wx.EVT_BUTTON, self.on_remove_file)
        atc_btn_sizer.Add(self.remove_file_btn, 0)
        vbox.Add(atc_btn_sizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        # 전송/취소 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.send_btn = wx.Button(panel, label="보내기(&S)")
        self.send_btn.Bind(wx.EVT_BUTTON, self.on_send)
        btn_sizer.Add(self.send_btn, 0, wx.RIGHT, 8)
        self.cancel_btn = wx.Button(panel, wx.ID_CANCEL, label="취소")
        btn_sizer.Add(self.cancel_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))

        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

        # 포커스
        if mode == self.MODE_ADMIN or default_recipient:
            self.subject_ctrl.SetFocus() if not default_subject else self.body_ctrl.SetFocus()
        else:
            self.recipient_ctrl.SetFocus()

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        if key == ord("S") and event.ControlDown() and not event.AltDown():
            self.on_send(None)
            return
        # Alt+F — 파일 추가
        if key == ord("F") and event.AltDown() and not event.ControlDown():
            self.on_add_files(None)
            return
        # 첨부 리스트박스에서 Del/D 로 선택 파일 제거
        focused = self.FindFocus()
        if focused is self.attach_list and not mods:
            if key in (wx.WXK_DELETE, ord("D")):
                self.on_remove_file(None)
                return
        event.Skip()

    def on_add_files(self, event):
        """wx.FileDialog 로 파일 선택 (다중 선택 허용) → 첨부 목록에 추가."""
        dlg = wx.FileDialog(
            self, "첨부할 파일 선택",
            wildcard="모든 파일 (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE,
        )
        try:
            if dlg.ShowModal() != wx.ID_OK:
                return
            import os as _os
            paths = dlg.GetPaths()
            for p in paths:
                if p in self.attachment_paths:
                    continue
                self.attachment_paths.append(p)
                size = ""
                try:
                    n = _os.path.getsize(p)
                    size = f" ({_format_size(n)})"
                except OSError:
                    pass
                self.attach_list.Append(f"{_os.path.basename(p)}{size}")
            speak(f"첨부파일 {len(paths)}개 추가됨. 총 {len(self.attachment_paths)}개.")
        finally:
            dlg.Destroy()

    def on_remove_file(self, event):
        sel = self.attach_list.GetSelection()
        if sel == wx.NOT_FOUND:
            speak("제거할 첨부파일을 선택해 주세요.")
            return
        del self.attachment_paths[sel]
        self.attach_list.Delete(sel)
        speak("첨부파일 제거됨.")
        if self.attach_list.GetCount() > 0:
            self.attach_list.SetSelection(min(sel, self.attach_list.GetCount() - 1))

    def on_send(self, event):
        recipient = self.recipient_ctrl.GetValue().strip()
        subject = self.subject_ctrl.GetValue().strip()
        body = self.body_ctrl.GetValue().strip()

        if not recipient:
            wx.MessageBox("받는 사람을 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.recipient_ctrl.SetFocus()
            return
        if not subject:
            wx.MessageBox("제목을 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.subject_ctrl.SetFocus()
            return
        if not body:
            wx.MessageBox("메일 내용을 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.body_ctrl.SetFocus()
            return

        is_email = "@" in recipient

        # 외부 이메일(@포함) → 소리샘은 외부 메일 발송을 막아 두었기 때문에
        # 사용자의 기본 메일 프로그램(mailto:) 으로 열어 보낸다.
        # 본문은 클립보드에도 복사해 길이 제한·인코딩 문제를 피하고, 첨부파일이
        # 있으면 mailto: 가 자동 첨부를 지원하지 않으므로 안내 후 경로를 같이
        # 클립보드에 넣어 사용자가 직접 첨부할 수 있도록 한다.
        if is_email:
            self._send_via_default_mail_client(recipient, subject, body)
            return

        # 소리샘 내부 회원 ID(@없음) — formmail 흐름으로 정상 전송
        if self.attachment_paths:
            speak(f"첨부파일 {len(self.attachment_paths)}개와 함께 메일을 보냅니다.")
        else:
            speak("메일을 보내는 중입니다.")
        self.send_btn.Disable()
        ok, msg = send_mail(
            self.session, recipient, subject, body,
            recipient_is_userid=True,
            attachments=self.attachment_paths or None,
        )
        self.send_btn.Enable()

        if ok:
            speak("메일을 보냈습니다.")
            wx.MessageBox(msg or "메일을 보냈습니다.", "전송 완료",
                          wx.OK | wx.ICON_INFORMATION, self)
            self.EndModal(wx.ID_OK)
        else:
            speak("메일 전송에 실패했습니다.")
            wx.MessageBox(f"메일 전송에 실패했습니다.\n{msg}",
                          "전송 실패", wx.OK | wx.ICON_ERROR, self)

    def _send_via_default_mail_client(
        self, recipient: str, subject: str, body: str,
    ) -> None:
        """기본 메일 프로그램을 열어 사용자가 직접 발송하도록 안내.

        소리샘은 외부 이메일 발송을 차단하므로 외부 주소(예: 관리자 gmail) 로
        메일을 보내려면 사용자의 PC 메일 프로그램을 거쳐야 한다.

        mailto: 의 ?subject=...&body=... 쿼리는 일부 한국어 메일 클라이언트가
        잘못 해석해 받는사람 칸에 query 문자열이 그대로 들어가는 문제가
        있어서, mailto 에는 받는사람만 넣고 제목·본문·첨부 경로는 클립보드로
        전달해 사용자가 붙여넣어 작성하도록 한다.
        """
        import os as _os, sys as _sys

        # 클립보드에 제목·본문·첨부 경로를 정리해서 한 번에 복사
        clip_lines = [f"제목: {subject}", "", body]
        if self.attachment_paths:
            clip_lines.append("")
            clip_lines.append("--- 첨부할 파일 경로 ---")
            clip_lines.extend(self.attachment_paths)
        try:
            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(wx.TextDataObject("\n".join(clip_lines)))
                wx.TheClipboard.Close()
        except Exception:
            pass

        # 기본 메일 프로그램 실행 — 받는 사람만 mailto: 에 포함
        # Windows 는 os.startfile 로 ShellExecute 호출 (등록된 mailto 핸들러).
        # 다른 OS 는 webbrowser.open 으로 폴백.
        opened = False
        mailto_url = f"mailto:{recipient}"
        try:
            if _sys.platform.startswith("win"):
                _os.startfile(mailto_url)
                opened = True
            else:
                import webbrowser
                webbrowser.open(mailto_url)
                opened = True
        except OSError as e:
            try:
                import webbrowser
                webbrowser.open(mailto_url)
                opened = True
            except Exception:
                speak("메일 프로그램을 열 수 없습니다.")
                wx.MessageBox(
                    f"기본 메일 프로그램을 열 수 없습니다.\n{e}\n\n"
                    "받는 사람·제목·본문·첨부 경로가 모두 클립보드에 복사되어 "
                    "있으니, 사용하시는 메일 서비스에서 직접 붙여 넣어 보내주세요.\n\n"
                    f"받는 사람: {recipient}",
                    "메일 프로그램 실행 실패", wx.OK | wx.ICON_ERROR, self,
                )
                return
        except Exception as e:
            speak("메일 프로그램을 열 수 없습니다.")
            wx.MessageBox(
                f"기본 메일 프로그램을 열 수 없습니다.\n{e}\n\n"
                f"받는 사람: {recipient}\n"
                "제목·본문·첨부 경로는 클립보드에 복사되어 있습니다.",
                "메일 프로그램 실행 실패", wx.OK | wx.ICON_ERROR, self,
            )
            return

        # 안내 메시지 — 사용자에게 무엇을 해야 하는지 명확히
        if self.attachment_paths:
            tip = (
                f"기본 메일 프로그램을 열었습니다. 받는 사람은 {recipient} 으로 "
                "자동 입력되어 있습니다.\n\n"
                "소리샘은 외부 이메일 발송을 막아 두었기 때문에, 메일 프로그램에서 "
                "직접 보내 주셔야 전달됩니다.\n\n"
                "제목·본문·첨부 경로가 클립보드에 한꺼번에 복사되어 있으니, "
                "메일 프로그램에서 붙여넣기(Ctrl+V) 로 작성하시고, 첨부할 파일은 "
                f"{len(self.attachment_paths)}개의 경로를 보고 첨부 버튼으로 "
                "직접 추가해 주세요."
            )
        else:
            tip = (
                f"기본 메일 프로그램을 열었습니다. 받는 사람은 {recipient} 으로 "
                "자동 입력되어 있습니다.\n\n"
                "소리샘은 외부 이메일 발송을 막아 두었기 때문에, 메일 프로그램에서 "
                "직접 보내 주셔야 전달됩니다.\n\n"
                "제목과 본문이 클립보드에 복사되어 있으니, 메일 프로그램에서 "
                "붙여넣기(Ctrl+V) 로 작성하시면 됩니다."
            )
        speak("메일 프로그램을 열었습니다. 직접 보내 주세요.")
        wx.MessageBox(tip, "기본 메일 프로그램으로 보내기",
                      wx.OK | wx.ICON_INFORMATION, self)
        self.EndModal(wx.ID_OK)


def _dump_mail_debug(html: str, tag: str) -> str:
    """메일 관련 디버그 HTML 덤프 (data/mail_debug_<tag>.html)."""
    import os
    from config import DATA_DIR
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        path = os.path.join(DATA_DIR, f"mail_debug_{tag}.html")
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        return path
    except OSError:
        return ""


@dataclass
class MailAttachment:
    """메일 첨부 파일."""
    filename: str
    url: str
    size: str = ""


from dataclasses import field as _field


@dataclass
class MailContent:
    """개별 메일 내용."""
    mail_id: str
    sender: str = ""
    recipient: str = ""
    subject: str = ""
    date: str = ""
    body: str = ""
    kind: str = "recv"  # recv / send
    attachments: list = _field(default_factory=list)
    body_download_url: str = ""  # /message/download.php?from=...&mr_id=...


def _mail_is_login_redirect(resp) -> bool:
    return _is_login_redirect(resp)


def _looks_like_mail_error_page(html: str) -> bool:
    """소리샘 /message/ 가 돌려주는 에러 페이지 패턴 감지."""
    patterns = [
        "오류안내 페이지",
        "값을 넘겨주세요",
        "잘못된 접근",
        "잘못된 접근입니다",
        "Fatal error",
        "Uncaught exception",
        "id=\"validation_check\"",
    ]
    return any(p in html for p in patterns)


def _extract_immediate_alert(html: str) -> str:
    """실제로 실행되는 alert() 의 메시지만 추출.

    함수 정의 안에 포함된 alert() 는 호출되지 않는 한 실행되지 않음.
    gnuboard 의 목록 페이지 같은 경우 JS 함수 내부에 alert("삭제할 메일을
    하나 이상 선택...") 같은 문자열이 코드로 있는데, 이건 실제로 호출된 게
    아니므로 무시해야 한다.

    휴리스틱:
    - 각 <script>...</script> 블록 검사
    - 해당 블록에 'function' 키워드가 있으면 함수 정의로 간주 → alert 무시
      (script 짧은 'alert+redirect' 패턴은 function 없음 → 즉시 실행으로 간주)
    - 첫 번째 매치 반환
    """
    for script_match in re.finditer(
        r"<script[^>]*>(.*?)</script>", html, flags=re.DOTALL | re.I
    ):
        script = script_match.group(1)
        # function 정의가 있는 스크립트 블록은 함수 내부에 선언된 alert 일 가능성
        if re.search(r"function\s+\w+\s*\(", script) or re.search(
            r"=\s*function\s*\(", script
        ):
            continue
        m = re.search(r"alert\(\s*['\"]([^'\"]+)['\"]", script)
        if m:
            return m.group(1)
    return ""


def _classify_mail_alert(text: str, final_url: str = "") -> tuple[bool, str]:
    """메일 API 응답의 alert() 해석. 함수 내부에 숨은 alert 는 무시."""
    success_kw = ["성공", "전송되었", "보냈", "발송되었", "전달되었",
                  "전달하였", "보내졌", "완료", "삭제되었"]
    failure_kw = ["실패", "오류", "잘못", "권한", "차단",
                  "거부", "이미 삭제", "존재하지 않"]
    msg = _extract_immediate_alert(text)
    if msg:
        if any(w in msg for w in success_kw):
            return True, msg
        if any(w in msg for w in failure_kw):
            return False, msg
        # 판정 애매 — URL 리다이렉트로 성공 여부 보조 판정
        if "sent.php" in final_url or "inbox.php" in final_url:
            return True, ""
        return True, msg
    # alert 없음 → URL 로 판정
    if "sent.php" in final_url or "inbox.php" in final_url or "/message/" in final_url:
        return True, ""
    return True, ""


def fetch_mail_list(session: requests.Session, kind: str = "recv",
                    page: int = 1) -> tuple[bool, "list[MailItem] | str"]:
    """사이트 내 메일 목록 조회. kind='recv' 받은함 / 'send' 보낸함. page=1 부터."""
    url = MAIL_SENT_URL if kind == "send" else MAIL_INBOX_URL
    params = {"page": page} if page > 1 else None
    try:
        resp = session.get(url, params=params, timeout=15)
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"
    if _is_login_redirect(resp):
        return False, "로그인 세션이 만료되었습니다."

    soup = BeautifulSoup(resp.text, "lxml")
    items: list[MailItem] = []

    # 테이블 기반 목록 — 여러 셀렉터 시도
    table = None
    for sel in [
        "#message_list table",
        "#mail_list table",
        ".tbl_head01 table",
        "#fboardlist table",
        "table",
    ]:
        table = soup.select_one(sel)
        if table:
            break
    if not table:
        _dump_mail_debug(resp.text, "inbox_no_table")
        return True, items

    tbody = table.find("tbody") or table
    # 열 헤더 순서를 미리 스캔해서 인덱스 → 역할 맵 구성
    thead_row = None
    thead = table.find("thead")
    if thead:
        thead_row = thead.find("tr")
    if thead_row is None:
        # 첫 tr 이 <th> 만 있으면 그게 헤더
        first_tr = tbody.find("tr")
        if first_tr and first_tr.find("th") and not first_tr.find("td"):
            thead_row = first_tr

    col_role: dict[int, str] = {}  # idx → 'sender'|'recipient'|'subject'|'date'|'read'|'chk'
    if thead_row is not None:
        for i, th in enumerate(thead_row.find_all(["th", "td"])):
            text = th.get_text(" ", strip=True)
            if not text:
                col_role[i] = "chk"
            elif "보낸" in text or "발신" in text or "발송" in text:
                col_role[i] = "sender"
            elif "받는" in text or "수신" in text:
                col_role[i] = "recipient"
            elif "제목" in text:
                col_role[i] = "subject"
            elif any(k in text for k in ["시간", "날짜", "작성", "일시"]):
                col_role[i] = "date"
            elif "읽" in text:
                col_role[i] = "read"
            elif "번호" in text or "no" in text.lower():
                col_role[i] = "num"

    for tr in tbody.find_all("tr", recursive=True):
        if tr.find("th") and not tr.find("td"):
            continue
        empty_td = tr.find("td", class_="empty_table")
        if empty_td:
            continue
        cells = tr.find_all(["td", "th"], recursive=False)
        if len(cells) < 2:
            continue

        # mail_id 추출 — 폭넓게 (chk_no, chk_id, chk_ms_id, chk_mail_id 등)
        mail_id = ""
        # 체크박스 — 숫자 값을 가진 chk_ 시작 input
        for inp in tr.find_all("input"):
            nm = (inp.get("name") or "")
            if not nm.startswith("chk"):
                continue
            val = inp.get("value") or ""
            if val.isdigit():
                mail_id = val
                break
        # view 링크에서 no=/id=/mail_id= 등 어떤 숫자 파라미터든 추출
        if not mail_id:
            for a in tr.find_all("a", href=True):
                href = a.get("href", "")
                m = re.search(r"(?:^|[?&])(?:no|id|mail_id|ms_id|mg_id|idx)=(\d+)", href)
                if m:
                    mail_id = m.group(1)
                    break
        if not mail_id:
            continue

        def cell_text(idx):
            if idx >= len(cells):
                return ""
            return cells[idx].get_text(" ", strip=True)

        def cell_subject(idx):
            """제목 셀 — td_subject 안에는 아이콘 + 링크(제목) + 모바일용 span(발신인+날짜)이
            섞여 있으므로 링크 텍스트만 우선 추출."""
            if idx >= len(cells):
                return ""
            cell = cells[idx]
            a = cell.find("a")
            if a:
                return a.get_text(" ", strip=True)
            return cell.get_text(" ", strip=True)

        def cell_name(idx):
            """발신인/수신인 셀 — <span title="id">id</span> 내부 텍스트 우선."""
            if idx >= len(cells):
                return ""
            cell = cells[idx]
            span = cell.find("span")
            if span:
                t = span.get_text(strip=True)
                if t:
                    return t
            return cell.get_text(" ", strip=True)

        sender = ""
        recipient = ""
        subject = ""
        date = ""
        read_cell = ""

        if col_role:
            # 헤더 기반 매핑 (정확)
            for idx, role in col_role.items():
                if idx >= len(cells):
                    continue
                if role == "subject":
                    subject = cell_subject(idx)
                elif role == "sender":
                    sender = cell_name(idx)
                elif role == "recipient":
                    recipient = cell_name(idx)
                elif role == "date":
                    date = cell_text(idx)
                elif role == "read":
                    read_cell = cell_text(idx)
        else:
            # 폴백: 셀 수에 따른 추측
            if len(cells) >= 5:
                sender = cell_text(1); subject = cell_text(2)
                date = cell_text(3);   read_cell = cell_text(4)
            elif len(cells) == 4:
                sender = cell_text(1); subject = cell_text(2); date = cell_text(3)
            elif len(cells) == 3:
                sender = cell_text(0); subject = cell_text(1); date = cell_text(2)

        # 제목 링크가 있으면 덮어쓰기 (a 태그 안의 텍스트가 보통 제목)
        a = tr.find("a", href=re.compile(r"view"))
        if a:
            link_text = a.get_text(strip=True)
            if link_text and link_text != sender and len(link_text) > 1:
                subject = link_text

        # 보낸 메일함에서는 sender 가 비어있고 recipient 만 있을 수 있음.
        # MailItem 은 sender 필드만 가지므로 recipient 를 sender 로 대체 (표시용).
        display_sender = sender or recipient or "(알 수 없음)"

        # 읽음 여부 판정 — 행 전체 HTML 에서 안 읽음 단서를 폭넓게 검색한다.
        # gnuboard5 계열은 안 읽은 메일을 텍스트가 아니라 작은 아이콘 이미지나
        # CSS 클래스로만 표시하는 경우가 흔해, 단순 텍스트 검사로는 놓친다.
        row_html_lower = str(tr).lower()
        is_read = True
        # 텍스트 기반 단서
        if re.search(
            r"읽지\s*않|안\s*읽|안읽음|미열람|미확인|미읽|읽음\s*아니|아직",
            row_html_lower,
        ):
            is_read = False
        # 이미지·클래스·src 등 마크업 기반 단서
        elif re.search(
            r'class\s*=\s*["\'][^"\']*\b(new|unread|notread|hot)\b'
            r'|src\s*=\s*["\'][^"\']*(?:new|unread|hot)[^"\']*\.(?:gif|png|jpg|svg|webp)'
            r'|alt\s*=\s*["\'][^"\']*(?:new|unread|안\s*읽|미\s*열람|미\s*확인|미\s*읽|새|hot)'
            r'|ico_new|icon[-_]new|new[-_]?icon|newmsg|memo[-_]?new|mail[-_]?new',
            row_html_lower,
        ):
            is_read = False
        # 단독 X/-/N 마커 (텍스트만 들어 있는 셀)
        elif read_cell.strip() in ("X", "x", "-", "N", "n", "○", "●"):
            is_read = False

        items.append(MailItem(
            mail_id=mail_id,
            sender=display_sender,
            subject=subject or "(제목 없음)",
            date=date or "",
            is_read=is_read,
        ))

    # 목록 HTML 은 항상 덤프 (구조 변경 감지용)
    _dump_mail_debug(resp.text, f"list_{kind}")
    if not items:
        _dump_mail_debug(resp.text, f"list_{kind}_no_items")
    return True, items


# 레거시 호환 — 기존 호출자
def fetch_mail_inbox(session):
    return fetch_mail_list(session, kind="recv")


def fetch_mail_list_up_to(session, kind: str = "recv", target_count: int = 10,
                          max_pages: int = 20) -> tuple[bool, "list[MailItem] | str"]:
    """여러 페이지를 순차 조회해서 target_count 개까지 모으기.

    서버가 페이지당 돌려주는 개수가 미지수라 페이지를 넘기며 누적한다.
    연속 페이지에서 같은 mail_id 가 반복되거나 빈 페이지가 오면 종료.
    """
    all_items: list = []
    seen_ids: set[str] = set()
    page = 1
    while len(all_items) < target_count and page <= max_pages:
        ok, items = fetch_mail_list(session, kind=kind, page=page)
        if not ok:
            if page == 1:
                return False, items
            break
        if not items:
            break
        new_any = False
        for it in items:
            if it.mail_id in seen_ids:
                continue
            seen_ids.add(it.mail_id)
            all_items.append(it)
            new_any = True
            if len(all_items) >= target_count:
                break
        if not new_any:
            break
        page += 1
    return True, all_items[:target_count]


def fetch_mail_content(session: requests.Session, mail_id: str, kind: str = "recv") -> tuple[bool, "MailContent | str"]:
    """개별 메일 읽기. kind='recv' → inbox_view.php, 'send' → sent_view.php.
    파라미터명은 일반 gnuboard 패턴 여러 개 순차 시도.
    """
    url = MAIL_SENT_VIEW_URL if kind == "send" else MAIL_INBOX_VIEW_URL
    # 실측 확인 (mail_debug_list_recv.html): 소리샘은
    # 받은함 → `mr_id=` (message_receive), 보낸함 → `ms_id=` (message_send) 사용
    if kind == "send":
        attempts = [
            {"ms_id": mail_id},
            {"mr_id": mail_id},
            {"id": mail_id},
            {"no": mail_id},
            {"idx": mail_id},
        ]
    else:
        attempts = [
            {"mr_id": mail_id},
            {"ms_id": mail_id},
            {"id": mail_id},
            {"no": mail_id},
            {"idx": mail_id},
        ]
    last_text = ""
    last_url = ""
    last_err = ""
    for params in attempts:
        try:
            resp = session.get(url, params=params, timeout=15)
        except requests.RequestException as e:
            last_err = f"서버 연결 실패: {e}"
            continue
        last_text = resp.text
        last_url = resp.url
        if resp.status_code != 200:
            last_err = f"HTTP {resp.status_code}"
            continue
        if _mail_is_login_redirect(resp):
            last_err = "로그인 세션 만료"
            continue
        # 서버측 에러 페이지 감지 — 여러 패턴
        if _looks_like_mail_error_page(resp.text):
            last_err = "서버측 에러 응답 (잘못된 접근/파라미터 오류)"
            continue
        content = _parse_mail_content(resp.text, mail_id, kind)
        if content is not None:
            # 성공 시에도 구조 변경 감지용 덤프 유지
            _dump_mail_debug(resp.text, f"view_{mail_id}")
            return True, content
        last_err = "파싱 실패"

    path = _dump_mail_debug(last_text, f"view_{mail_id}")
    hint = f"  (디버그: {path})" if path else ""
    return False, f"메일 내용을 파싱하지 못했습니다. 최종 오류: {last_err}{hint}"


def _parse_mail_content(html: str, mail_id: str, kind: str) -> "MailContent | None":
    """메일 본문 파싱 — 다중 구조 지원.

    소리샘 /message/ 페이지는 사이트 스킨에 따라 여러 구조를 가질 수 있으므로
    우선순위로 여러 추출 경로 시도:
      1) <th>/<td> 페어 (표준 gnuboard)
      2) <dl>/<dt>/<dd> 페어 (일부 테마)
      3) ar.memo 같은 <li class="*_view_li"> <span class="*_subj">라벨</span> <strong>값</strong>
      4) 일반 라벨-값 패턴 (class 나 영문 필드 이름 탐색)
      5) 페이지 전체 텍스트에서 정규식
      6) 본문 — 광범위 셀렉터 (.view_content, article, #message_view 등)
    """
    soup = BeautifulSoup(html, "lxml")
    content = MailContent(mail_id=mail_id, kind=kind)

    def _label_matches(label: str, keywords: list[str]) -> bool:
        return any(k in label for k in keywords)

    def _set_from_label(label: str, value: str):
        """라벨 텍스트 보고 content 의 적절한 필드에 값 설정. 이미 값 있으면 무시."""
        if not value:
            return
        if _label_matches(label, ["보낸", "발신", "발송자"]) and "시간" not in label:
            content.sender = content.sender or value
        elif _label_matches(label, ["받는", "수신자", "수신인"]):
            content.recipient = content.recipient or value
        elif "제목" in label:
            content.subject = content.subject or value
        elif _label_matches(label, ["시간", "작성", "날짜", "일시", "일자", "보낸날짜", "받은날짜"]):
            content.date = content.date or value
        elif _label_matches(label, ["내용", "본문", "메시지"]):
            content.body = content.body or value

    def _clean_value(el) -> str:
        """요소 내 <br> 를 개행으로 변환하고 텍스트 추출."""
        import copy as _copy
        c = _copy.copy(el)
        for br in c.find_all("br"):
            br.replace_with("\n")
        return c.get_text("\n", strip=True)

    # ── 경로 1: th/td ──
    for tr in soup.find_all("tr"):
        th = tr.find("th")
        td = tr.find("td")
        if not th or not td:
            continue
        label = th.get_text(" ", strip=True)
        value = _clean_value(td)
        _set_from_label(label, value)

    # ── 경로 2: dl/dt/dd ──
    for dl in soup.find_all("dl"):
        dts = dl.find_all("dt")
        dds = dl.find_all("dd")
        for dt, dd in zip(dts, dds):
            label = dt.get_text(" ", strip=True)
            value = _clean_value(dd)
            _set_from_label(label, value)

    # ── 경로 3: ar.memo 식 — <li>/<span class="*subj">라벨</span> 값 ──
    for li in soup.find_all("li"):
        subj = li.find(class_=re.compile(r"(subj|label|title_lbl|view_subj)"))
        if not subj:
            continue
        label = subj.get_text(" ", strip=True)
        # 값 — <strong>, <em>, <span.value> 등 우선, 아니면 li 전체에서 label 제거 후
        value = ""
        for tag_name, class_re in (("strong", None), ("em", None),
                                    ("span", re.compile(r"(value|content)"))):
            tag = li.find(tag_name, class_=class_re) if class_re else li.find(tag_name)
            if tag:
                value = _clean_value(tag)
                break
        if not value:
            full = _clean_value(li)
            value = full.replace(label, "", 1).strip()
        _set_from_label(label, value)

    # ── 경로 4: 일반적인 라벨-값 (span.*subject/span.*value, etc.) ──
    # <label>받는 사람</label><span>xxx</span> 같은 인접 패턴
    if not (content.sender or content.recipient or content.subject or content.date):
        labels = soup.find_all(["label", "span", "strong", "b"],
                                class_=re.compile(r"(label|subj|field)"))
        for lbl in labels:
            label_text = lbl.get_text(" ", strip=True)
            if not label_text:
                continue
            # 인접 형제 값 탐색
            sibling = lbl.find_next_sibling()
            if sibling:
                value = _clean_value(sibling)
                _set_from_label(label_text, value)

    # ── 경로 5: 제목 — gnuboard 표준 selector 우선 (독립 단계) ──
    if not content.subject:
        # gnuboard5 메시지/게시판 뷰의 표준 제목 ID
        for sel in ["#bo_v_title", ".bo_v_tit", "#message_subject",
                    "#mail_subject", "header h1", "header h2",
                    "article header h1", "article h1"]:
            el = soup.select_one(sel)
            if el:
                t = el.get_text(" ", strip=True)
                if t and t not in ("받은 메일", "보낸 메일", "받은 메일함", "보낸 메일함",
                                    "메일 읽기", "메일 보기", "메일", "본문",
                                    "받은메일", "보낸메일", "받은메일함", "보낸메일함"):
                    content.subject = t
                    break
        if not content.subject:
            for h in soup.find_all(["h1", "h2", "h3"]):
                t = h.get_text(" ", strip=True)
                if t and t not in ("받은 메일", "보낸 메일", "받은 메일함", "보낸 메일함",
                                    "메일 읽기", "메일 보기", "메일", "본문",
                                    "받은메일", "보낸메일", "받은메일함", "보낸메일함"):
                    content.subject = t
                    break

    # ── 경로 6: 페이지 전체 텍스트에서 정규식 (메타 최후 폴백) ──
    if not (content.sender or content.recipient or content.date):
        full_text = soup.get_text(" ", strip=True)
        if not content.sender:
            m = re.search(r"(?:보낸\s*사람|발신인|보낸이|발신자)\s*[:：]?\s*([^\s|]+)", full_text)
            if m:
                content.sender = m.group(1)
        if not content.recipient:
            m = re.search(r"(?:받는\s*사람|수신인|받는이)\s*[:：]?\s*([^\s|]+)", full_text)
            if m:
                content.recipient = m.group(1)
        if not content.date:
            m = re.search(r"(\d{2,4}[-.]\d{1,2}[-.]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)",
                          full_text)
            if m:
                content.date = m.group(1)

    # ── 경로 6: 본문 ──
    if not content.body:
        body_selectors = [
            "#message_view_content", "#message_view_contents",
            "#mail_view_content", "#message_body", "#mail_body",
            "#message_view .content", "#mail_view .content",
            ".message_content", ".mail_content", ".view_content",
            "#bo_v_con", "article.message_view", "article.mail_view",
            "article", "#memo_view_contents",
        ]
        for sel in body_selectors:
            el = soup.select_one(sel)
            if not el:
                continue
            import copy as _copy
            clean = _copy.copy(el)
            # 네비·버튼 영역 제거
            for rem in clean.select(
                ".btn, .button, ._win_btn, .btn_confirm, .win_ul, "
                "ul.win_ul, form, .nav, .board_navi, #board_navi_info, "
                "script, style, header, h1, h2, h3"
            ):
                rem.decompose()
            for br in clean.find_all("br"):
                br.replace_with("\n")
            text = clean.get_text("\n", strip=True)
            # 메타 라벨 라인(보낸사람·날짜·제목·수신인) 는 본문에서 제거
            body_lines = []
            for line in text.split("\n"):
                s = line.strip()
                if not s:
                    continue
                if re.match(r"^(보낸|받는|발신|수신)\s*(사람|이|인|자|분)?\s*[:：]", s):
                    continue
                if re.match(r"^(제목|날짜|시간|일시|작성)\s*[:：]", s):
                    continue
                body_lines.append(s)
            candidate = "\n".join(body_lines).strip()
            if candidate:
                content.body = candidate
                break

    # ── 본문을 <p> 블록에서 찾기 (최후 폴백) ──
    if not content.body:
        ps = soup.find_all("p")
        for p in ps:
            t = _clean_value(p)
            if t and len(t) > 3:
                content.body = t
                break

    # ── 경로 7: 첨부파일 + 본문 다운로드 URL 추출 ──
    # 소리샘 구조 (view_159774.html 실측):
    # <section id="bo_v_atc"> 안에 <!-- 첨부파일 --> 영역
    # <div id="bo_v_bot"> 안에 <a href=".../download.php?from=inbox&mr_id=N">본문 다운로드</a>
    def _abs_url(href: str) -> str:
        """상대/절대 href 를 절대 URL 로 변환. 중복 /message/ 생성 방지."""
        if href.startswith("http"):
            return href
        return urljoin(SORISEM_BASE_URL, href)

    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        text = a.get_text(" ", strip=True)
        # 본문 다운로드 링크
        if "download.php" in href and "본문" in text:
            content.body_download_url = _abs_url(href)
            continue

    # 첨부파일 섹션 탐색 — 여러 후보
    atc_containers = []
    for sel in ["#bo_v_atc", ".bo_v_atc", "#bo_v_file", ".bo_v_file",
                ".view_file", "#mail_atc", ".attach", ".attachment"]:
        el = soup.select_one(sel)
        if el:
            atc_containers.append(el)
    # 첨부파일이 section 안의 section 안에 있는 경우 (실측 소리샘 구조),
    # #bo_v_atc 가 본문+첨부 둘 다 포함하므로 추가 셀렉터 없이 그대로 사용.

    seen_urls: set[str] = set()
    if content.body_download_url:
        seen_urls.add(content.body_download_url)
    for cont_el in atc_containers:
        for a in cont_el.find_all("a", href=True):
            href = a.get("href", "")
            # download_attachment.php / download.php?file=... / filedown.php 등 폭넓게
            if not re.search(r"(download_attachment|download|filedown)", href, re.I):
                continue
            text = a.get_text(" ", strip=True)
            # sound_only 클래스의 스크린리더 보조 텍스트 제거
            so = a.find(class_="sound_only")
            if so:
                text = text.replace(so.get_text(" ", strip=True), "", 1).strip()
            if not text or "본문" in text:
                continue
            full_url = _abs_url(href)
            if full_url in seen_urls:
                continue
            seen_urls.add(full_url)
            # 크기 정보 — 링크 직후 형제 span 또는 부모 텍스트
            size = ""
            nxt = a.find_next_sibling()
            if nxt:
                nxt_text = nxt.get_text(" ", strip=True)
                m_size = re.search(r"\(?(\d+(?:[.,]\d+)?\s*[KMGkmg]?B)\)?", nxt_text)
                if m_size:
                    size = m_size.group(1)
            if not size:
                parent_text = a.parent.get_text(" ", strip=True) if a.parent else ""
                m_size = re.search(r"\(?(\d+(?:[.,]\d+)?\s*[KMGkmg]?B)\)?", parent_text)
                if m_size:
                    size = m_size.group(1)
            content.attachments.append(MailAttachment(
                filename=text, url=full_url, size=size,
            ))

    # 전부 비면 실패
    if (not content.body and not content.sender and not content.recipient
            and not content.subject and not content.date):
        return None
    if not content.body:
        content.body = "(본문 없음)"
    return content


def send_mail_message(session: requests.Session, recipient: str,
                      subject: str, body: str,
                      attachments: "list[str] | None" = None) -> tuple[bool, str]:
    """사이트 내 메일 발송. /message/write.php → 폼 POST.

    attachments: 첨부할 로컬 파일 경로 리스트. None 이면 첨부 없음.
    파일 업로드가 있으면 multipart/form-data 로 전송.
    """
    import os as _os
    try:
        form_resp = session.get(MAIL_WRITE_URL, timeout=15)
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if form_resp.status_code != 200:
        return False, f"HTTP {form_resp.status_code}"
    if _mail_is_login_redirect(form_resp):
        return False, "로그인 세션 만료"

    # 진단 용 덤프 (첫 실행 시 form 구조 확인)
    _dump_mail_debug(form_resp.text, "write_form")

    soup = BeautifulSoup(form_resp.text, "lxml")
    form = (soup.find("form", {"name": re.compile(r"mail|message", re.I)})
            or soup.find("form", id=re.compile(r"mail|message|fwrite|write", re.I))
            or soup.find("form", attrs={"enctype": "multipart/form-data"})
            or soup.find("form"))
    if form is None:
        _dump_mail_debug(form_resp.text, "write_no_form")
        return False, "메일 작성 폼을 찾지 못했습니다."

    data = {}
    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name or inp.get("type") in ("submit", "file", "button"):
            continue
        inp_type = (inp.get("type") or "").lower()
        # 체크박스/라디오는 checked 된 경우만 전송 (브라우저 동작과 동일)
        if inp_type in ("checkbox", "radio"):
            if inp.has_attr("checked"):
                data[name] = inp.get("value", "1")
            continue
        data[name] = inp.get("value", "")
    for ta in form.find_all("textarea"):
        name = ta.get("name")
        if name:
            data[name] = ta.get_text() or ""
    # select 는 기본 선택 값 유지
    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name or name in data:
            continue
        selected = sel.find("option", selected=True)
        if selected is not None:
            data[name] = selected.get("value", "")
        else:
            data[name] = ""

    # 필드명 매핑 — 소리샘 /message/write.php 실측 (mail_debug_write_form.html):
    #   수신인: receivers (text input)
    #   제목:   ms_subject
    #   내용:   ms_content (textarea)
    #   파일:   ms_file[] (위에서 자동 탐지)
    # 호환성을 위해 구 gnuboard/다른 스킨 후보도 뒤에 열거.
    def set_first_matching(candidates: list[str], value: str):
        for c in candidates:
            if c in data:
                data[c] = value
                return True
        data[candidates[0]] = value
        return False

    set_first_matching(
        ["receivers", "mb_id", "recv_mb_id", "to_id", "to_mb_id",
         "to", "recipient"],
        recipient,
    )
    set_first_matching(
        ["ms_subject", "subject", "title", "wr_subject"],
        subject,
    )
    set_first_matching(
        ["ms_content", "content", "message", "body", "wr_content", "memo"],
        body,
    )

    # 파일 input 필드명 탐지 (폼 HTML 실측 기반)
    file_field_name = None
    for fi in form.find_all("input", {"type": "file"}):
        nm = fi.get("name")
        if nm:
            file_field_name = nm
            break
    if not file_field_name:
        # 소리샘 /message/write.php 실측 기준 ms_file[], 폴백으로 bf_file[]
        file_field_name = "ms_file[]"

    action = form.get("action") or MAIL_WRITE_URL
    post_url = urljoin(form_resp.url, action)

    try:
        if attachments:
            # multipart/form-data 업로드
            files_payload = []
            open_handles = []
            try:
                for path in attachments:
                    if not _os.path.isfile(path):
                        continue
                    fh = open(path, "rb")
                    open_handles.append(fh)
                    filename = _os.path.basename(path)
                    files_payload.append((file_field_name, (filename, fh)))
                resp = session.post(
                    post_url, data=data, files=files_payload,
                    timeout=120, allow_redirects=True,
                )
            finally:
                for fh in open_handles:
                    try: fh.close()
                    except Exception: pass
        else:
            resp = session.post(post_url, data=data, timeout=30, allow_redirects=True)
    except requests.RequestException as e:
        return False, f"전송 실패: {e}"
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"
    return _classify_mail_alert(resp.text, resp.url)


def delete_mail(session: requests.Session, mail_id: str, kind: str = "recv") -> tuple[bool, str]:
    """메일 삭제.

    실제 사이트 form 동작 (mail_debug_list_recv.html 참조):
    <form name="delete_message" method="post" action="/message/inbox.php">
        <input type="hidden" name="check_all" value="0">
        <input type="checkbox" name="chk_mr_id[]" value="...">
    → check_all=0 + chk_mr_id[]=<id> 로 본인 메일함에 POST.
    """
    primary_url = MAIL_INBOX_URL if kind == "recv" else MAIL_SENT_URL
    id_field = "chk_mr_id[]" if kind == "recv" else "chk_ms_id[]"
    try:
        resp = session.post(
            primary_url,
            data={"check_all": "0", id_field: mail_id},
            timeout=15,
            allow_redirects=True,
        )
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if resp.status_code == 200 and not _mail_is_login_redirect(resp) \
            and "오류안내" not in resp.text and "값을 넘겨주세요" not in resp.text:
        return _classify_mail_alert(resp.text, resp.url)

    # 폴백 — 옛 엔드포인트 후보
    base = MAIL_INBOX_BASE
    candidates = [
        (f"{base}/inbox_delete.php", {"id": mail_id, "kind": kind}),
        (f"{base}/sent_delete.php", {"id": mail_id, "kind": kind}),
        (f"{base}/delete.php", {"id": mail_id, "kind": kind}),
    ]
    for url, params in candidates:
        try:
            resp = session.get(url, params=params, timeout=15, allow_redirects=True)
        except requests.RequestException:
            continue
        if resp.status_code == 404:
            continue
        if _mail_is_login_redirect(resp):
            return False, "로그인 세션 만료"
        if "오류안내" in resp.text or "값을 넘겨주세요" in resp.text:
            continue
        ok, msg = _classify_mail_alert(resp.text, resp.url)
        if ok or msg:
            return ok, msg
        return True, ""
    return False, "삭제 엔드포인트를 찾지 못했습니다."


def delete_all_mails(session: requests.Session, kind: str = "recv") -> tuple[bool, str]:
    """메일함 전체 비우기.

    mail_debug_list_recv.html 의 delete_mail(1) JS 동작:
    - 같은 form (action=/message/inbox.php) 에 check_all=1 로 POST
    - 서버는 chk_mr_id[] 체크 여부를 무시하고 해당 메일함 전체 삭제
    """
    primary_url = MAIL_INBOX_URL if kind == "recv" else MAIL_SENT_URL
    try:
        resp = session.post(
            primary_url,
            data={"check_all": "1"},
            timeout=30,
            allow_redirects=True,
        )
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}  (URL: {resp.url})"
    if _mail_is_login_redirect(resp):
        return False, "로그인 세션이 만료되었습니다."
    return _classify_mail_alert(resp.text, resp.url)


class MailNotifier:
    """백그라운드에서 주기적으로 새 메일 확인. MemoNotifier 와 동일한 구조.

    알림 센터(notification.NotificationCenter)에 unread mail 을 등록.
    """

    def __init__(self, parent_frame, session):
        import wx as _wx
        self.frame = parent_frame
        self.session = session
        self.seen_ids: set[str] = set()
        self._in_flight = False
        self._initial_done = False
        # 주기마다 서버의 현재 안 읽은 메일 수를 UI 에 알려주는 콜백(선택).
        self.on_unread_count = None
        # Mail 은 별도 타이머 쓰지 않고 MemoNotifier 와 공통 타이머 공유 가능.
        # 여기서는 단순화 — 외부에서 poll_once() 직접 호출.

    def start_initial_fill(self):
        """앱 시작 시 메일함 초기화.

        "읽은" 메일만 seen 으로 등록. 안 읽은 메일은 의도적으로 빼두어
        첫 polling tick 에서 신규 항목으로 감지되어 알림이 발사된다.
        현재 안 읽은 총개수는 on_unread_count 콜백을 통해 제목 표시줄에 반영.
        """
        import threading, wx as _wx
        def worker():
            try:
                ok, items = fetch_mail_inbox(self.session)
                if ok and isinstance(items, list):
                    for it in items:
                        if getattr(it, "is_read", True):
                            self.seen_ids.add(it.mail_id)
                    unread_count = sum(
                        1 for it in items if not getattr(it, "is_read", True)
                    )
                    if self.on_unread_count is not None:
                        _wx.CallAfter(self.on_unread_count, unread_count)
            except Exception:
                pass
            finally:
                self._initial_done = True
        threading.Thread(target=worker, daemon=True).start()

    def poll_once_async(self, on_new_items, on_done=None):
        """1회 폴링. 새 메일 있으면 on_new_items(list[MailItem]) UI 스레드 호출.
        on_done 은 항상 호출됨 (성공/실패 관계없이).
        """
        import threading, wx as _wx
        if self._in_flight:
            if on_done:
                _wx.CallAfter(on_done)
            return
        self._in_flight = True
        def worker():
            try:
                if not self._initial_done:
                    return
                ok, items = fetch_mail_inbox(self.session)
                if not ok or not isinstance(items, list):
                    return
                # 현재 서버 기준 안 읽은 메일 총개수를 UI(제목 표시줄)에 전달.
                unread_count = sum(
                    1 for it in items if not getattr(it, "is_read", True)
                )
                if self.on_unread_count is not None:
                    _wx.CallAfter(self.on_unread_count, unread_count)
                new = [it for it in items if it.mail_id not in self.seen_ids]
                if not new:
                    return
                for it in new:
                    self.seen_ids.add(it.mail_id)
                _wx.CallAfter(on_new_items, new)
            except Exception:
                pass
            finally:
                self._in_flight = False
                if on_done:
                    _wx.CallAfter(on_done)
        threading.Thread(target=worker, daemon=True).start()

    def mark_all_as_seen(self):
        """현재 받은함의 모든 mail_id 를 seen 으로 갱신.

        메일함을 사용자가 직접 본 뒤(읽음·삭제 등) 호출. 다음 polling tick 에서
        삭제로 인해 페이지에 새로 드러난 옛 메일을 "신규 도착"으로 오인해
        팝업이 뜨는 문제를 막는다.
        """
        import threading
        def worker():
            try:
                ok, items = fetch_mail_inbox(self.session)
                if ok and isinstance(items, list):
                    for it in items:
                        self.seen_ids.add(it.mail_id)
            except Exception:
                pass
        threading.Thread(target=worker, daemon=True).start()

    def get_all_unread_async(self, on_result):
        """수동 체크용: 현재 메일함의 is_read=False 항목 전체를 UI 스레드로 전달."""
        import threading, wx as _wx
        def worker():
            try:
                ok, items = fetch_mail_inbox(self.session)
                if not ok or not isinstance(items, list):
                    _wx.CallAfter(on_result, [])
                    return
                unread = [it for it in items if not it.is_read]
                _wx.CallAfter(on_result, unread)
            except Exception:
                _wx.CallAfter(on_result, [])
        threading.Thread(target=worker, daemon=True).start()


def show_mail_inbox_unavailable(parent):
    """레거시 — 현재는 fetch_mail_list 로 실제 조회 가능. 유지는 호환성."""
    wx.MessageBox(
        "메일 받은함 조회를 진행합니다. 응답이 비어있거나 파싱에 실패하면 "
        "data/mail_debug_*.html 디버그 파일이 생성됩니다.",
        "메일 받은함", wx.OK | wx.ICON_INFORMATION, parent,
    )


# ─────────────────────────────────────────────────────────────
# 다운로드 API — 본문·첨부파일 저장
# ─────────────────────────────────────────────────────────────

def _sanitize_filename(name: str) -> str:
    """Windows 에서 금지된 문자 제거 + 길이 제한."""
    if not name:
        return "file"
    # Windows 금지 문자: < > : " / \ | ? *
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.strip(" .") or "file"
    return name[:200]


def _decode_korean_filename(fn: str) -> str:
    """requests 가 latin-1 로 디코딩한 헤더 문자열에서 원본 한글 복원.

    HTTP 헤더는 RFC 2616 에 따라 latin-1 로 읽히는데, 실제 서버는 UTF-8 이나
    EUC-KR(CP949) 로 인코딩한 바이트를 보낸다. latin-1 로 먼저 인코딩해서
    원본 바이트를 되살린 뒤 UTF-8 → CP949 순으로 재디코딩.
    """
    if not fn:
        return fn
    # 이미 올바른 한글이면 encode("latin-1") 에서 예외
    try:
        raw = fn.encode("latin-1")
    except UnicodeEncodeError:
        return fn  # 이미 유니코드로 제대로 디코딩된 상태
    # UTF-8 우선
    try:
        decoded = raw.decode("utf-8")
        # ASCII 가 그대로 통과했다면 원본과 같음 — 그 경우에도 OK
        return decoded
    except UnicodeDecodeError:
        pass
    # EUC-KR / CP949 시도 (구 gnuboard·한국 사이트 관행)
    try:
        return raw.decode("cp949")
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode("euc-kr")
    except UnicodeDecodeError:
        pass
    return fn


def _extract_filename_from_headers(resp, fallback: str) -> str:
    """Content-Disposition 헤더에서 filename 추출.

    지원 포맷:
    - RFC 5987: filename*=UTF-8''<percent-encoded>
    - filename="<bytes>" — requests 의 latin-1 디코딩 되돌리기 (한글 복원)
    - filename="%XX%YY..." — URL 퍼센트 인코딩
    """
    from urllib.parse import unquote
    cd = resp.headers.get("content-disposition") or ""
    if not cd:
        return _sanitize_filename(fallback)

    # 1) RFC 5987 filename*=CHARSET''encoded
    m = re.search(r"filename\*\s*=\s*([^;]+)", cd, re.I)
    if m:
        value = m.group(1).strip().strip('"')
        enc_m = re.match(r"([\w-]+)''(.+)", value)
        if enc_m:
            charset = enc_m.group(1).lower() or "utf-8"
            encoded = enc_m.group(2)
            # 일부 인코딩 이름 정규화
            if charset in ("euc-kr", "euckr"):
                charset = "euc-kr"
            if charset in ("ks_c_5601-1987", "ks_c_5601"):
                charset = "cp949"
            try:
                return _sanitize_filename(unquote(encoded, encoding=charset))
            except (LookupError, Exception):
                try:
                    return _sanitize_filename(unquote(encoded))
                except Exception:
                    pass

    # 2) filename="..."
    m = re.search(r'filename\s*=\s*"?([^";]+)"?', cd, re.I)
    if m:
        fn = m.group(1).strip().strip('"')
        # 2a) 퍼센트 인코딩인 경우
        if "%" in fn:
            for enc_try in ("utf-8", "cp949", "euc-kr"):
                try:
                    decoded = unquote(fn, encoding=enc_try)
                    # 디코딩이 의미있게 진행됐는지 확인 (%기호 줄었는지)
                    if decoded != fn and decoded.count("%") < fn.count("%"):
                        return _sanitize_filename(decoded)
                except Exception:
                    continue
        # 2b) latin-1 디코딩된 문자열을 올바른 한글로 재디코딩
        decoded = _decode_korean_filename(fn)
        return _sanitize_filename(decoded)

    return _sanitize_filename(fallback)


def download_mail_body(session: requests.Session, content: "MailContent",
                       save_dir: str) -> tuple[bool, str]:
    """메일 본문을 텍스트 파일로 저장 (download.php 엔드포인트 사용).

    download_url 이 content 에 있으면 그걸 쓰고, 없으면 추정 URL 구성.
    반환: (성공, 저장경로 또는 에러메시지)
    """
    import os as _os
    url = content.body_download_url
    if not url:
        from_val = "inbox" if content.kind == "recv" else "sent"
        param_key = "mr_id" if content.kind == "recv" else "ms_id"
        url = f"{MAIL_INBOX_BASE}/download.php"
        try:
            resp = session.get(url, params={"from": from_val, param_key: content.mail_id},
                               timeout=30, stream=True)
        except requests.RequestException as e:
            return False, f"서버 연결 실패: {e}"
    else:
        try:
            resp = session.get(url, timeout=30, stream=True)
        except requests.RequestException as e:
            return False, f"서버 연결 실패: {e}"

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"
    if _mail_is_login_redirect(resp):
        return False, "로그인 세션 만료"

    # 파일명 결정
    default_name = f"메일_{content.mail_id}"
    if content.subject:
        default_name = f"{_sanitize_filename(content.subject)}_{content.mail_id}"
    filename = _extract_filename_from_headers(resp, default_name + ".txt")
    if not _os.path.splitext(filename)[1]:
        filename += ".txt"

    try:
        _os.makedirs(save_dir, exist_ok=True)
        save_path = _os.path.join(save_dir, filename)
        # 같은 이름 파일 충돌 방지
        base, ext = _os.path.splitext(save_path)
        n = 1
        while _os.path.exists(save_path):
            save_path = f"{base} ({n}){ext}"
            n += 1
        with open(save_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    except OSError as e:
        return False, f"파일 저장 실패: {e}"
    return True, save_path


def download_attachment(session: requests.Session, attachment: "MailAttachment",
                        save_dir: str) -> tuple[bool, str]:
    """개별 첨부파일 다운로드."""
    import os as _os
    try:
        resp = session.get(attachment.url, timeout=60, stream=True,
                           allow_redirects=True)
    except requests.RequestException as e:
        return False, f"서버 연결 실패: {e}"
    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"
    if _mail_is_login_redirect(resp):
        return False, "로그인 세션 만료"

    filename = _extract_filename_from_headers(resp, attachment.filename or "attachment")
    try:
        _os.makedirs(save_dir, exist_ok=True)
        save_path = _os.path.join(save_dir, filename)
        base, ext = _os.path.splitext(save_path)
        n = 1
        while _os.path.exists(save_path):
            save_path = f"{base} ({n}){ext}"
            n += 1
        with open(save_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    except OSError as e:
        return False, f"파일 저장 실패: {e}"
    return True, save_path


# ─────────────────────────────────────────────────────────────
# 첨부파일 목록 대화상자
# ─────────────────────────────────────────────────────────────

class AttachmentListDialog(wx.Dialog):
    """첨부파일을 목록(ListBox) 형태로 보여주고 선택 다운로드·전체 다운로드 제공."""

    def __init__(self, parent, session, attachments: list, save_dir: str):
        super().__init__(parent, title="첨부파일 목록", size=(560, 420),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.attachments = attachments
        self.save_dir = save_dir

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 상단 안내
        header = wx.StaticText(
            panel,
            label=f"첨부파일 {len(attachments)}개 — Enter: 선택 다운로드 · A: 전체 · Esc: 닫기",
        )
        vbox.Add(header, 0, wx.ALL, 8)

        # ListBox
        self.list_ctrl = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE)
        for att in attachments:
            size_str = f"  ({att.size})" if att.size else ""
            self.list_ctrl.Append(f"{att.filename}{size_str}")
        if attachments:
            self.list_ctrl.SetSelection(0)
        self.list_ctrl.Bind(wx.EVT_LISTBOX_DCLICK, lambda e: self._download_selected())
        vbox.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.download_btn = wx.Button(panel, label="다운로드(&D)")
        self.download_btn.Bind(wx.EVT_BUTTON, lambda e: self._download_selected())
        btn_sizer.Add(self.download_btn, 0, wx.RIGHT, 8)
        self.all_btn = wx.Button(panel, label="전체 다운로드(&A)")
        self.all_btn.Bind(wx.EVT_BUTTON, lambda e: self._download_all())
        btn_sizer.Add(self.all_btn, 0, wx.RIGHT, 8)
        self.close_btn = wx.Button(panel, wx.ID_CLOSE, label="닫기(&C)")
        self.close_btn.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(wx.ID_CLOSE))
        btn_sizer.Add(self.close_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)
        self.list_ctrl.SetFocus()

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()
        focused = self.FindFocus()
        if isinstance(focused, wx.Button):
            event.Skip()
            return
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CLOSE)
            return
        if key == wx.WXK_RETURN and not mods:
            self._download_selected()
            return
        if key == ord("A") and not mods:
            self._download_all()
            return
        event.Skip()

    def _selected_index(self) -> int:
        sel = self.list_ctrl.GetSelection()
        return sel if sel != wx.NOT_FOUND else -1

    def _download_selected(self):
        idx = self._selected_index()
        if idx < 0 or idx >= len(self.attachments):
            return
        att = self.attachments[idx]
        speak(f"{att.filename} 내려받는 중입니다.")
        ok, result = download_attachment(self.session, att, self.save_dir)
        if ok:
            speak("다운로드 완료.")
            wx.MessageBox(f"다운로드 완료.\n\n{result}",
                          "다운로드 완료", wx.OK | wx.ICON_INFORMATION, self)
        else:
            speak("다운로드 실패.")
            wx.MessageBox(f"다운로드에 실패했습니다.\n{result}",
                          "다운로드 실패", wx.OK | wx.ICON_ERROR, self)

    def _download_all(self):
        if not self.attachments:
            return
        speak(f"첨부파일 {len(self.attachments)}개를 내려받습니다.")
        success = []
        failed = []
        for att in self.attachments:
            ok, result = download_attachment(self.session, att, self.save_dir)
            if ok:
                success.append(result)
            else:
                failed.append(f"{att.filename}: {result}")
        if success and not failed:
            speak(f"{len(success)}개 모두 내려받았습니다.")
            wx.MessageBox(
                f"{len(success)}개 첨부파일을 모두 저장했습니다.\n\n저장 폴더: {self.save_dir}",
                "다운로드 완료", wx.OK | wx.ICON_INFORMATION, self,
            )
        elif success and failed:
            speak(f"{len(success)}개 성공 {len(failed)}개 실패.")
            wx.MessageBox(
                f"성공 {len(success)}개 / 실패 {len(failed)}개\n\n실패 목록:\n"
                + "\n".join(failed),
                "다운로드 부분 성공", wx.OK | wx.ICON_WARNING, self,
            )
        else:
            speak("전체 다운로드에 실패했습니다.")
            wx.MessageBox(
                "전체 다운로드에 실패했습니다.\n\n" + "\n".join(failed),
                "다운로드 실패", wx.OK | wx.ICON_ERROR, self,
            )


# ─────────────────────────────────────────────────────────────
# 사이트 내 메일 UI — 쪽지와 동일한 구조
# ─────────────────────────────────────────────────────────────

class MailViewDialog(wx.Dialog):
    """개별 메일 보기. 메모의 MemoViewDialog 와 동일한 레이아웃·단축키."""

    def __init__(self, parent, session, content: "MailContent",
                 items: list | None = None, index: int = 0,
                 kind: str | None = None):
        kind_label = "받은 메일" if content.kind == "recv" else "보낸 메일"
        super().__init__(parent, title=kind_label, size=(720, 540),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.content = content
        self.items = items or []
        self.index = index
        self.kind = kind or content.kind

        panel = wx.Panel(self)
        self._panel = panel
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 메타+본문 통합 TextCtrl — 팝업 키로 커스텀 컨텍스트 메뉴가 뜨도록 서브클래스 사용
        self.body_ctrl = ContextMenuTextCtrl(
            panel, value="",
            style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
        )
        vbox.Add(self.body_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.prev_btn = None
        self.next_btn = None
        if self.items:
            self.prev_btn = wx.Button(panel, label="이전(&P)")
            self.prev_btn.Bind(wx.EVT_BUTTON, lambda e: self._navigate(-1))
            btn_sizer.Add(self.prev_btn, 0, wx.RIGHT, 8)
            self.next_btn = wx.Button(panel, label="다음(&N)")
            self.next_btn.Bind(wx.EVT_BUTTON, lambda e: self._navigate(+1))
            btn_sizer.Add(self.next_btn, 0, wx.RIGHT, 8)
        if self.kind == "recv":
            self.reply_btn = wx.Button(panel, label="답장(&R)")
            self.reply_btn.Bind(wx.EVT_BUTTON, self.on_reply)
            btn_sizer.Add(self.reply_btn, 0, wx.RIGHT, 8)
        self.delete_btn = wx.Button(panel, label="삭제(&D)")
        self.delete_btn.Bind(wx.EVT_BUTTON, self.on_delete)
        btn_sizer.Add(self.delete_btn, 0, wx.RIGHT, 8)
        self.close_btn = wx.Button(panel, wx.ID_CLOSE, label="닫기(&C)")
        self.close_btn.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(wx.ID_CLOSE))
        btn_sizer.Add(self.close_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)
        self.body_ctrl.bind_context_menu(self._on_body_context_menu)

        self._refresh_display(announce=False)
        self.body_ctrl.SetFocus()
        self.body_ctrl.SetInsertionPoint(0)

    def _on_body_context_menu(self, event):
        """메일 보기 팝업 메뉴 — 답장·삭제·본문 저장 등 메일 액션.

        이전/다음 메일 이동은 내비게이션이므로 제외.
        """
        menu = wx.Menu()

        if self.kind == "recv":
            id_reply = wx.NewIdRef()
            menu.Append(id_reply, "답장(&R)")
            self.Bind(wx.EVT_MENU, self.on_reply, id=id_reply)

        id_del = wx.NewIdRef()
        menu.Append(id_del, "삭제(&D)")
        self.Bind(wx.EVT_MENU, self.on_delete, id=id_del)

        if getattr(self.content, "attachments", None):
            menu.AppendSeparator()
            id_att = wx.NewIdRef()
            menu.Append(id_att, "첨부파일 전체 저장(&S)")
            self.Bind(wx.EVT_MENU, self.on_download_all_attachments, id=id_att)

        menu.AppendSeparator()
        id_save_body = wx.NewIdRef()
        menu.Append(id_save_body, "본문 텍스트 저장(&B)")
        self.Bind(wx.EVT_MENU, self.on_download_body, id=id_save_body)

        self.PopupMenu(menu)
        menu.Destroy()

    def _refresh_display(self, announce: bool = True):
        kind_label = "받은 메일" if self.content.kind == "recv" else "보낸 메일"
        total = len(self.items)
        subject = (self.content.subject or "").strip() or "(제목 없음)"
        if total:
            self.SetTitle(f"{subject} — {kind_label} ({self.index + 1}/{total})")
        else:
            self.SetTitle(f"{subject} — {kind_label}")

        # 날짜/시간 분리
        date_part = time_part = ""
        if self.content.date:
            parts = self.content.date.strip().split(None, 1)
            if len(parts) >= 2:
                date_part, time_part = parts
            else:
                if ":" in self.content.date:
                    time_part = self.content.date
                else:
                    date_part = self.content.date

        if self.content.kind == "recv":
            person_label = "보낸 사람"
            person_value = self.content.sender or "(알 수 없음)"
        else:
            person_label = "받는 사람"
            person_value = self.content.recipient or "(알 수 없음)"

        lines = [
            f"작성 날짜: {date_part or '(정보 없음)'}",
            f"보낸 시간: {time_part or '(정보 없음)'}",
            f"{person_label}: {person_value}",
            f"제목: {self.content.subject or '(제목 없음)'}",
        ]
        # 첨부파일 요약 (자세한 목록은 Alt+S 로 별도 대화상자 오픈)
        attachments = self.content.attachments or []
        if attachments:
            lines.append(
                f"첨부파일: {len(attachments)}개  (Alt+S: 목록 열기 / Alt+Shift+S: 전체 다운로드)"
            )
        else:
            lines.append("첨부파일: 없음")
        lines.extend([
            "",
            "내용: (B 로 텍스트 저장)",
            self.content.body or "(본문 없음)",
        ])
        self.body_ctrl.SetValue("\n".join(lines))
        self.body_ctrl.SetInsertionPoint(0)

        if self.prev_btn:
            self.prev_btn.Enable(self.index > 0)
        if self.next_btn:
            self.next_btn.Enable(self.index < len(self.items) - 1)

        if hasattr(self, "_panel"):
            self._panel.Layout()

        if announce and total:
            speak(f"{self.index + 1} / {total} · {person_value}")

    def _navigate(self, direction: int):
        if not self.items:
            return
        new_idx = self.index + direction
        if new_idx < 0:
            speak("첫 번째 메일입니다.")
            return
        if new_idx >= len(self.items):
            speak("마지막 메일입니다.")
            return
        item = self.items[new_idx]
        speak("메일을 불러옵니다.")
        ok, result = fetch_mail_content(self.session, item.mail_id, kind=self.kind)
        if not ok:
            speak("메일을 불러오지 못했습니다.")
            wx.MessageBox(f"메일을 불러오지 못했습니다.\n{result}",
                          "오류", wx.OK | wx.ICON_ERROR, self)
            return
        self.index = new_idx
        self.content = result
        dir_msg = "다음 메일입니다." if direction > 0 else "이전 메일입니다."
        self._refresh_display(announce=False)
        who = self.content.sender or self.content.recipient or "알 수 없음"
        speak(f"{dir_msg} {self.index + 1} / {len(self.items)} · {who}")
        self.body_ctrl.SetFocus()
        self.body_ctrl.SetInsertionPoint(0)

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()
        alt_only = event.AltDown() and not event.ControlDown() and not event.ShiftDown()
        alt_shift = event.AltDown() and event.ShiftDown() and not event.ControlDown()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CLOSE)
            return
        # B — 본문 다운로드 (텍스트 저장)
        if key == ord("B") and not mods:
            self.on_download_body(None)
            return
        # Alt+Shift+S — 첨부파일 전체 다운로드
        if key == ord("S") and alt_shift:
            self.on_download_all_attachments(None)
            return
        # Alt+S — 첨부파일 선택 다운로드
        if key == ord("S") and alt_only:
            self.on_download_attachment(None)
            return
        if (key == ord("D") or key == wx.WXK_DELETE) and not mods:
            self.on_delete(None)
            return
        if key == ord("R") and not mods and self.kind == "recv":
            self.on_reply(None)
            return
        if key == wx.WXK_PAGEUP and not mods:
            self._navigate(-1); return
        if key == wx.WXK_PAGEDOWN and not mods:
            self._navigate(+1); return
        if key == ord("P") and alt_only:
            self._navigate(-1); return
        if key == ord("N") and alt_only:
            self._navigate(+1); return
        event.Skip()

    # ── 다운로드 핸들러 ──

    def on_download_body(self, event):
        """B — 본문 텍스트 파일로 저장."""
        from config import get_download_dir
        save_dir = get_download_dir()
        speak("본문을 내려받습니다.")
        ok, result = download_mail_body(self.session, self.content, save_dir)
        if ok:
            speak("본문을 저장했습니다.")
            wx.MessageBox(f"본문을 저장했습니다.\n\n{result}",
                          "본문 저장 완료", wx.OK | wx.ICON_INFORMATION, self)
        else:
            speak("본문 저장에 실패했습니다.")
            wx.MessageBox(f"본문 저장에 실패했습니다.\n{result}",
                          "본문 저장 실패", wx.OK | wx.ICON_ERROR, self)

    def on_download_attachment(self, event):
        """Alt+S — 첨부파일 목록 대화상자 (ListBox 기반).

        첨부가 없으면 안내 후 종료.
        있으면 첨부파일이 1개든 여러 개든 AttachmentListDialog 를 띄워서
        목록 UI 로 사용자가 선택 · 다운로드할 수 있게 한다.
        """
        attachments = self.content.attachments or []
        if not attachments:
            speak("첨부파일이 없습니다.")
            wx.MessageBox("이 메일에는 첨부파일이 없습니다.",
                          "첨부파일 없음", wx.OK | wx.ICON_INFORMATION, self)
            return
        from config import get_download_dir
        save_dir = get_download_dir()
        dlg = AttachmentListDialog(self, self.session, attachments, save_dir)
        try:
            dlg.ShowModal()
        finally:
            dlg.Destroy()

    def on_download_all_attachments(self, event):
        """Alt+Shift+S — 모든 첨부파일 다운로드."""
        attachments = self.content.attachments or []
        if not attachments:
            speak("첨부파일이 없습니다.")
            wx.MessageBox("이 메일에는 첨부파일이 없습니다.",
                          "첨부파일 없음", wx.OK | wx.ICON_INFORMATION, self)
            return
        from config import get_download_dir
        save_dir = get_download_dir()
        speak(f"첨부파일 {len(attachments)}개를 내려받습니다.")
        success_paths = []
        failed = []
        for att in attachments:
            ok, result = download_attachment(self.session, att, save_dir)
            if ok:
                success_paths.append(result)
            else:
                failed.append(f"{att.filename}: {result}")
        if success_paths and not failed:
            speak(f"{len(success_paths)}개 모두 내려받았습니다.")
            msg = (f"{len(success_paths)}개 첨부파일을 모두 저장했습니다.\n\n"
                   f"저장 폴더: {save_dir}\n\n"
                   + "\n".join(f"· {p.split(chr(92))[-1]}" for p in success_paths))
            wx.MessageBox(msg, "다운로드 완료",
                          wx.OK | wx.ICON_INFORMATION, self)
        elif success_paths and failed:
            speak(f"{len(success_paths)}개 성공 {len(failed)}개 실패.")
            wx.MessageBox(
                f"일부 첨부파일 다운로드에 실패했습니다.\n\n"
                f"성공 {len(success_paths)}개 / 실패 {len(failed)}개\n\n"
                f"실패 목록:\n" + "\n".join(failed),
                "다운로드 부분 성공", wx.OK | wx.ICON_WARNING, self,
            )
        else:
            speak("첨부파일 다운로드에 모두 실패했습니다.")
            wx.MessageBox(
                "첨부파일 다운로드에 모두 실패했습니다.\n\n" + "\n".join(failed),
                "다운로드 실패", wx.OK | wx.ICON_ERROR, self,
            )

    def _do_download_one(self, attachment):
        from config import get_download_dir
        save_dir = get_download_dir()
        speak(f"{attachment.filename} 내려받는 중입니다.")
        ok, result = download_attachment(self.session, attachment, save_dir)
        if ok:
            speak("다운로드 완료.")
            wx.MessageBox(f"다운로드 완료.\n\n{result}",
                          "다운로드 완료", wx.OK | wx.ICON_INFORMATION, self)
        else:
            speak("다운로드 실패.")
            wx.MessageBox(f"다운로드에 실패했습니다.\n{result}",
                          "다운로드 실패", wx.OK | wx.ICON_ERROR, self)

    def on_reply(self, event):
        dlg = MailComposeDialog(self, self.session,
                                default_recipient=self.content.sender,
                                default_subject=f"Re: {self.content.subject}",
                                default_body=f"\n\n--- 원본 메일 ---\n{self.content.body}")
        dlg.ShowModal()
        dlg.Destroy()

    def on_delete(self, event):
        ans = wx.MessageBox("이 메일을 삭제하시겠습니까?",
                            "메일 삭제", wx.YES_NO | wx.ICON_QUESTION, self)
        if ans != wx.YES:
            return
        ok, msg = delete_mail(self.session, self.content.mail_id, kind=self.content.kind)
        if ok:
            speak("메일을 삭제했습니다.")
            if self.items:
                del self.items[self.index]
                if not self.items:
                    wx.MessageBox("마지막 메일이었습니다. 메일함으로 돌아갑니다.",
                                  "삭제 완료", wx.OK | wx.ICON_INFORMATION, self)
                    self.EndModal(wx.ID_OK)
                    return
                if self.index >= len(self.items):
                    self.index = len(self.items) - 1
                item = self.items[self.index]
                ok2, result = fetch_mail_content(self.session, item.mail_id, kind=self.kind)
                if ok2:
                    self.content = result
                    self._refresh_display(announce=True)
                    self.body_ctrl.SetFocus()
                    self.body_ctrl.SetInsertionPoint(0)
                    return
                self.EndModal(wx.ID_OK)
                return
            wx.MessageBox("메일을 삭제했습니다.", "삭제 완료",
                          wx.OK | wx.ICON_INFORMATION, self)
            self.EndModal(wx.ID_OK)
        else:
            speak("메일 삭제 실패.")
            wx.MessageBox(f"메일 삭제에 실패했습니다.\n{msg}",
                          "삭제 실패", wx.OK | wx.ICON_ERROR, self)


class MailComposeDialog(wx.Dialog):
    """사이트 내 메일 작성 (/message/write.php). 첨부파일 지원."""

    def __init__(self, parent, session, default_recipient="",
                 default_subject="", default_body=""):
        super().__init__(parent, title="메일 쓰기", size=(720, 600),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.attachment_paths: list[str] = []

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        lbl_r = wx.StaticText(panel, label="받는 사람 아이디")
        vbox.Add(lbl_r, 0, wx.TOP | wx.LEFT | wx.RIGHT, 8)
        self.recipient_ctrl = wx.TextCtrl(panel, value=default_recipient)
        vbox.Add(self.recipient_ctrl, 0, wx.ALL | wx.EXPAND, 8)

        lbl_s = wx.StaticText(panel, label="제목")
        vbox.Add(lbl_s, 0, wx.LEFT | wx.RIGHT, 8)
        self.subject_ctrl = wx.TextCtrl(panel, value=default_subject)
        vbox.Add(self.subject_ctrl, 0, wx.ALL | wx.EXPAND, 8)

        lbl_b = wx.StaticText(panel, label="내용")
        vbox.Add(lbl_b, 0, wx.LEFT | wx.RIGHT, 8)
        self.body_ctrl = wx.TextCtrl(panel, value=default_body,
                                     style=wx.TE_MULTILINE | wx.TE_RICH2)
        vbox.Add(self.body_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 첨부파일 영역
        atc_label = wx.StaticText(panel, label="첨부파일 (Alt+F 추가 / Del 삭제)")
        vbox.Add(atc_label, 0, wx.LEFT | wx.RIGHT | wx.TOP, 8)
        self.attach_list = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE)
        self.attach_list.SetMinSize((-1, 80))
        vbox.Add(self.attach_list, 0, wx.ALL | wx.EXPAND, 8)

        atc_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.add_file_btn = wx.Button(panel, label="파일 추가(&F)")
        self.add_file_btn.Bind(wx.EVT_BUTTON, self.on_add_files)
        atc_btn_sizer.Add(self.add_file_btn, 0, wx.RIGHT, 8)
        self.remove_file_btn = wx.Button(panel, label="선택 파일 제거(&R)")
        self.remove_file_btn.Bind(wx.EVT_BUTTON, self.on_remove_file)
        atc_btn_sizer.Add(self.remove_file_btn, 0)
        vbox.Add(atc_btn_sizer, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        # 전송/취소 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.send_btn = wx.Button(panel, label="보내기(&S)")
        self.send_btn.Bind(wx.EVT_BUTTON, self.on_send)
        btn_sizer.Add(self.send_btn, 0, wx.RIGHT, 8)
        self.cancel_btn = wx.Button(panel, wx.ID_CANCEL, label="취소")
        btn_sizer.Add(self.cancel_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

        if default_recipient and default_subject:
            self.body_ctrl.SetFocus()
        elif default_recipient:
            self.subject_ctrl.SetFocus()
        else:
            self.recipient_ctrl.SetFocus()

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        if key == ord("S") and event.ControlDown() and not event.AltDown():
            self.on_send(None)
            return
        # Alt+F — 파일 추가 (어느 필드에 포커스 있어도)
        if key == ord("F") and event.AltDown() and not event.ControlDown():
            self.on_add_files(None)
            return
        # 포커스가 첨부 리스트박스일 때 Del/D 로 선택 파일 제거
        focused = self.FindFocus()
        if focused is self.attach_list and not mods:
            if key in (wx.WXK_DELETE, ord("D")):
                self.on_remove_file(None)
                return
        event.Skip()

    def on_add_files(self, event):
        """wx.FileDialog 로 파일 선택 (다중 선택 허용) → 첨부 목록에 추가."""
        dlg = wx.FileDialog(
            self, "첨부할 파일 선택",
            wildcard="모든 파일 (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE,
        )
        try:
            if dlg.ShowModal() != wx.ID_OK:
                return
            import os as _os
            paths = dlg.GetPaths()
            for p in paths:
                if p in self.attachment_paths:
                    continue
                self.attachment_paths.append(p)
                size = ""
                try:
                    n = _os.path.getsize(p)
                    size = f" ({_format_size(n)})"
                except OSError:
                    pass
                self.attach_list.Append(f"{_os.path.basename(p)}{size}")
            speak(f"첨부파일 {len(paths)}개 추가됨. 총 {len(self.attachment_paths)}개.")
        finally:
            dlg.Destroy()

    def on_remove_file(self, event):
        sel = self.attach_list.GetSelection()
        if sel == wx.NOT_FOUND:
            speak("제거할 첨부파일을 선택해 주세요.")
            return
        removed_name = self.attach_list.GetString(sel)
        del self.attachment_paths[sel]
        self.attach_list.Delete(sel)
        speak("첨부파일 제거됨.")
        if self.attach_list.GetCount() > 0:
            self.attach_list.SetSelection(min(sel, self.attach_list.GetCount() - 1))

    def on_send(self, event):
        recipient = self.recipient_ctrl.GetValue().strip()
        subject = self.subject_ctrl.GetValue().strip()
        body = self.body_ctrl.GetValue().strip()
        if not recipient:
            wx.MessageBox("받는 사람 아이디를 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.recipient_ctrl.SetFocus()
            return
        if not subject:
            wx.MessageBox("제목을 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.subject_ctrl.SetFocus()
            return
        if not body:
            wx.MessageBox("내용을 입력해 주세요.", "입력 필요",
                          wx.OK | wx.ICON_WARNING, self)
            self.body_ctrl.SetFocus()
            return
        if self.attachment_paths:
            speak(f"첨부파일 {len(self.attachment_paths)}개와 함께 메일을 전송합니다.")
        else:
            speak("메일을 전송하는 중입니다.")
        self.send_btn.Disable()
        ok, msg = send_mail_message(
            self.session, recipient, subject, body,
            attachments=self.attachment_paths or None,
        )
        self.send_btn.Enable()
        if ok:
            speak("메일을 보냈습니다.")
            wx.MessageBox(msg or "메일을 보냈습니다.", "전송 완료",
                          wx.OK | wx.ICON_INFORMATION, self)
            self.EndModal(wx.ID_OK)
        else:
            speak("메일 전송 실패.")
            wx.MessageBox(f"메일 전송에 실패했습니다.\n{msg}", "전송 실패",
                          wx.OK | wx.ICON_ERROR, self)


def _format_size(n: int) -> str:
    """바이트를 사람이 읽기 좋게 포매팅."""
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f} {unit}" if unit != "B" else f"{n} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


class MailInboxDialog(wx.Dialog):
    """메일함 — 받은/보낸 전환 + 목록 (쪽지함과 동일한 구조)."""

    def __init__(self, parent, session):
        super().__init__(parent, title="메일함", size=(760, 560),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.kind = "recv"
        self.items: list = []
        self._index = 0
        # Shift+좌/우 필드 순회용 인덱스. 항목이 바뀌면 0 으로 초기화.
        self._field_index = 0

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 상단 버튼
        top = wx.BoxSizer(wx.HORIZONTAL)
        self.recv_btn = wx.Button(panel, label="받은 메일함(&R)")
        self.recv_btn.Bind(wx.EVT_BUTTON, lambda e: self._switch("recv"))
        self.send_btn = wx.Button(panel, label="보낸 메일함(&S)")
        self.send_btn.Bind(wx.EVT_BUTTON, lambda e: self._switch("send"))
        self.compose_btn = wx.Button(panel, label="새 메일(&N)")
        self.compose_btn.Bind(wx.EVT_BUTTON, self.on_compose)
        self.refresh_btn = wx.Button(panel, label="새로고침(&F)")
        self.refresh_btn.Bind(wx.EVT_BUTTON, lambda e: self.reload())
        self.delete_all_btn = wx.Button(panel, label="모든 메일 삭제(&A)")
        self.delete_all_btn.Bind(wx.EVT_BUTTON, self.on_delete_all)
        top.Add(self.recv_btn, 0, wx.RIGHT, 4)
        top.Add(self.send_btn, 0, wx.RIGHT, 4)
        top.Add(self.compose_btn, 0, wx.RIGHT, 4)
        top.Add(self.refresh_btn, 0, wx.RIGHT, 4)
        top.Add(self.delete_all_btn, 0)
        vbox.Add(top, 0, wx.ALL, 8)

        self.status_label = wx.StaticText(panel, label="불러오는 중...")
        vbox.Add(self.status_label, 0, wx.LEFT | wx.RIGHT, 8)

        self.list_ctrl = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE)
        self.list_ctrl.Bind(wx.EVT_LISTBOX_DCLICK, lambda e: self._open_current())
        vbox.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        hint = wx.StaticText(
            panel,
            label="↑↓ 이동 · PgUp/PgDn 더 불러오기 · Enter 읽기 · D/Del 삭제 · Shift+Del 전체 삭제 · R 답장 · N 새 메일 · Esc 닫기",
        )
        vbox.Add(hint, 0, wx.ALL, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)
        self.list_ctrl.Bind(wx.EVT_CONTEXT_MENU, self._on_list_context_menu)
        self.list_ctrl.SetFocus()
        wx.CallAfter(self.reload)

    def _on_list_context_menu(self, event):
        """메일함 팝업 메뉴 — 새 메일·답장·삭제 등 메일 관련 액션 모음.

        ↑↓·PgUp/PgDn 이동은 내비게이션이므로 제외.
        """
        menu = wx.Menu()

        id_compose = wx.NewIdRef()
        menu.Append(id_compose, "새 메일 쓰기(&N)")
        self.Bind(wx.EVT_MENU, self.on_compose, id=id_compose)

        if self.items:
            id_open = wx.NewIdRef()
            menu.Append(id_open, "선택한 메일 읽기(&O)")
            self.Bind(wx.EVT_MENU, lambda e: self._open_current(), id=id_open)

            id_reply = wx.NewIdRef()
            menu.Append(id_reply, "답장(&R)")
            self.Bind(wx.EVT_MENU, lambda e: self._reply_current(), id=id_reply)

            id_del = wx.NewIdRef()
            menu.Append(id_del, "선택한 메일 삭제(&D)")
            self.Bind(wx.EVT_MENU, lambda e: self._delete_current(), id=id_del)

        menu.AppendSeparator()

        id_refresh = wx.NewIdRef()
        menu.Append(id_refresh, "새로고침(&F)")
        self.Bind(wx.EVT_MENU, lambda e: self.reload(), id=id_refresh)

        if self.items:
            id_del_all = wx.NewIdRef()
            menu.Append(id_del_all, "모든 메일 삭제(&A)")
            self.Bind(wx.EVT_MENU, self.on_delete_all, id=id_del_all)

        self.PopupMenu(menu)
        menu.Destroy()

    def _switch(self, kind: str):
        if kind == self.kind:
            return
        self.kind = kind
        self.reload()

    def reload(self):
        label = "받은" if self.kind == "recv" else "보낸"
        # 메일함을 새로 그릴 때마다 알림 봇이 가지고 있는 seen_ids 도 함께
        # 갱신해, 삭제로 인해 페이지 안에 새로 드러난 옛 메일이 다음 polling
        # tick 에서 "새 메일"로 오인되지 않게 한다.
        if self.kind == "recv":
            try:
                top = wx.GetTopLevelParent(self)
                notifier = getattr(top, "_mail_notifier", None)
                if notifier is not None:
                    notifier.mark_all_as_seen()
            except Exception:
                pass
        self.status_label.SetLabel(f"{label} 메일함 불러오는 중...")
        try:
            from settings_dialog import load_notify_settings
            target = int(load_notify_settings().get("list_page_size", 10))
        except Exception:
            target = 10
        self._target_per_reload = target
        self._loaded_pages = 1
        ok, result = fetch_mail_list_up_to(self.session, kind=self.kind, target_count=target)
        if not ok:
            self.items = []
            self.list_ctrl.Clear()
            self.status_label.SetLabel(f"오류: {result}")
            speak("메일함을 불러오지 못했습니다.")
            wx.MessageBox(f"메일함을 불러오지 못했습니다.\n\n원인: {result}",
                          "메일함 불러오기 실패", wx.OK | wx.ICON_ERROR, self)
            return
        self.items = result
        self._index = 0
        if not self.items:
            self.list_ctrl.Clear()
            self.status_label.SetLabel(f"{label} 메일이 없습니다.")
            speak(f"{label} 메일이 없습니다.")
            return
        self.status_label.SetLabel(f"{label} 메일 {len(self.items)}개")
        lines = [self._format_item(i, it) for i, it in enumerate(self.items)]
        self.list_ctrl.Set(lines)
        self.list_ctrl.SetSelection(0)
        speak(f"{label} 메일 {len(self.items)}개.")

    def _format_item(self, i: int, item) -> str:
        who_label = "보낸이" if self.kind == "recv" else "받는이"
        # 받은함은 안읽음/읽음 상태를 항상 맨 앞에 표시해 스크린리더가 가장
        # 먼저 읽도록 한다. 보낸함은 읽음 개념이 다르므로 표시하지 않는다.
        if self.kind == "recv":
            status = "안 읽음" if not getattr(item, "is_read", True) else "읽음"
            prefix = f"{status} · "
        else:
            prefix = ""
        return (
            f"{prefix}{i+1}/{len(self.items)} · {who_label}: {item.sender} · "
            f"{item.date} · {item.subject}"
        )

    def _sync_index(self):
        sel = self.list_ctrl.GetSelection()
        if sel != wx.NOT_FOUND:
            if sel != self._index:
                # 항목이 바뀌면 필드 인덱스 초기화 — 새 항목의 첫 필드부터.
                self._field_index = 0
            self._index = sel

    def _get_fields(self, item) -> list[tuple[str, str]]:
        """현재 메일 항목에서 Shift+좌/우 로 순회할 필드 목록.

        받은함은 읽음 상태도 포함. 보낸함은 받는이로 라벨만 다름.
        """
        fields: list[tuple[str, str]] = []
        if self.kind == "recv":
            status = "안 읽음" if not getattr(item, "is_read", True) else "읽음"
            fields.append(("상태", status))
            fields.append(("보낸이", item.sender or ""))
        else:
            fields.append(("받는이", item.sender or ""))
        fields.append(("날짜", item.date or ""))
        fields.append(("제목", item.subject or ""))
        return fields

    def _read_field(self, direction: int):
        """현재 선택된 메일의 필드를 한 칸 이동해서 음성으로 안내."""
        if not self.items:
            return
        item = self.items[self._index]
        fields = self._get_fields(item)
        if not fields:
            return
        self._field_index += direction
        if self._field_index < 0:
            self._field_index = 0
            speak("첫 번째 필드")
            return
        if self._field_index >= len(fields):
            self._field_index = len(fields) - 1
            speak("마지막 필드")
            return
        name, value = fields[self._field_index]
        speak(f"{name} {value}" if value else name)

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()
        focused = self.FindFocus()
        if isinstance(focused, wx.Button):
            event.Skip()
            return
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CLOSE)
            return
        if key == wx.WXK_RETURN and not mods:
            self._sync_index(); self._open_current(); return
        # Shift+좌/우: 필드 단위로 이동(보낸이/날짜/제목/안읽음 등) — 스크린리더가
        # 각 필드를 따로 읽어 사용자가 정보를 한 조각씩 확인할 수 있다.
        if key == wx.WXK_LEFT and event.ShiftDown() and not event.ControlDown():
            self._sync_index(); self._read_field(-1); return
        if key == wx.WXK_RIGHT and event.ShiftDown() and not event.ControlDown():
            self._sync_index(); self._read_field(1); return
        # Shift+Del 은 전체 삭제 (D/Del 단독보다 먼저 검사)
        if key == wx.WXK_DELETE and event.ShiftDown() and not event.ControlDown() and not event.AltDown():
            self.on_delete_all(None); return
        if (key == ord("D") or key == wx.WXK_DELETE) and not mods:
            self._sync_index(); self._delete_current(); return
        if key == ord("R") and not mods and self.kind == "recv":
            self._sync_index(); self._reply_current(); return
        if key == ord("N") and not mods:
            self.on_compose(None); return
        if key == ord("F") and not mods:
            self.reload(); return
        # PageUp/PageDown — 다음 페이지를 누적 로드
        if key in (wx.WXK_PAGEDOWN, wx.WXK_PAGEUP) and not mods:
            self._load_more(); return
        event.Skip()

    def _load_more(self):
        """다음 페이지 메일을 현재 목록에 append."""
        if not hasattr(self, "_loaded_pages"):
            self._loaded_pages = 1

        next_page = self._loaded_pages + 1
        speak("다음 메일을 불러옵니다.")
        ok, new_items = fetch_mail_list(self.session, kind=self.kind, page=next_page)
        if not ok:
            speak("불러오기에 실패했습니다.")
            return
        if not new_items:
            speak("더 이상 불러올 메일이 없습니다.")
            return

        existing_ids = {it.mail_id for it in self.items}
        added = []
        for it in new_items:
            if it.mail_id in existing_ids:
                continue
            existing_ids.add(it.mail_id)
            self.items.append(it)
            added.append(it)

        if not added:
            speak("더 이상 불러올 메일이 없습니다.")
            return

        self._loaded_pages = next_page

        sel = self.list_ctrl.GetSelection()
        lines = [self._format_item(i, it) for i, it in enumerate(self.items)]
        self.list_ctrl.Set(lines)
        if sel == wx.NOT_FOUND:
            self.list_ctrl.SetSelection(len(self.items) - len(added))
        else:
            self.list_ctrl.SetSelection(sel)

        label = "받은" if self.kind == "recv" else "보낸"
        self.status_label.SetLabel(
            f"{label} 메일 {len(self.items)}개 (페이지 {self._loaded_pages}까지)"
        )
        speak(f"{len(added)}개 추가로 불러왔습니다. 총 {len(self.items)}개.")

    def _open_current(self):
        if not self.items:
            return
        item = self.items[self._index]
        speak("메일을 불러옵니다.")
        ok, result = fetch_mail_content(self.session, item.mail_id, kind=self.kind)
        if not ok:
            speak("메일을 불러오지 못했습니다.")
            wx.MessageBox(f"메일을 불러오지 못했습니다.\n{result}",
                          "오류", wx.OK | wx.ICON_ERROR, self)
            return
        dlg = MailViewDialog(self, self.session, result,
                             items=self.items, index=self._index, kind=self.kind)
        code = dlg.ShowModal()
        final_index = dlg.index
        dlg.Destroy()
        if code == wx.ID_OK:
            self.reload()
        else:
            if self.items and 0 <= final_index < len(self.items):
                self._index = final_index
                self.list_ctrl.SetSelection(self._index)

    def _delete_current(self):
        if not self.items:
            return
        item = self.items[self._index]
        ans = wx.MessageBox(f"선택한 메일을 삭제하시겠습니까?\n\n{item.subject}",
                            "메일 삭제", wx.YES_NO | wx.ICON_QUESTION, self)
        if ans != wx.YES:
            return
        ok, msg = delete_mail(self.session, item.mail_id, kind=self.kind)
        if ok:
            speak("메일을 삭제했습니다.")
            self.reload()
        else:
            speak("메일 삭제 실패.")
            wx.MessageBox(f"삭제에 실패했습니다.\n{msg}", "오류",
                          wx.OK | wx.ICON_ERROR, self)

    def _reply_current(self):
        if not self.items:
            return
        item = self.items[self._index]
        dlg = MailComposeDialog(self, self.session,
                                default_recipient=item.sender,
                                default_subject=f"Re: {item.subject}")
        dlg.ShowModal()
        dlg.Destroy()

    def on_compose(self, event):
        dlg = MailComposeDialog(self, self.session)
        code = dlg.ShowModal()
        dlg.Destroy()
        if code == wx.ID_OK and self.kind == "send":
            self.reload()

    def on_delete_all(self, event):
        """현재 메일함(받은/보낸)의 모든 메일을 일괄 삭제."""
        label = "받은" if self.kind == "recv" else "보낸"
        if not self.items:
            speak(f"{label} 메일함이 이미 비어 있습니다.")
            wx.MessageBox(f"{label} 메일함에 삭제할 메일이 없습니다.",
                          "알림", wx.OK | wx.ICON_INFORMATION, self)
            return
        ans = wx.MessageBox(
            f"{label} 메일함의 전체 메일 {len(self.items)}개를 모두 삭제하시겠습니까?\n"
            "삭제한 메일은 복구할 수 없습니다.",
            "전체 메일 삭제",
            wx.YES_NO | wx.ICON_WARNING | wx.NO_DEFAULT, self,
        )
        if ans != wx.YES:
            return
        speak("전체 메일을 삭제하는 중입니다.")
        ok, msg = delete_all_mails(self.session, kind=self.kind)
        if ok:
            speak(f"{label} 메일함을 비웠습니다.")
            wx.MessageBox(msg or f"{label} 메일함의 모든 메일을 삭제했습니다.",
                          "삭제 완료", wx.OK | wx.ICON_INFORMATION, self)
            self.reload()
        else:
            speak("전체 메일 삭제 실패.")
            wx.MessageBox(f"전체 메일 삭제에 실패했습니다.\n{msg}",
                          "삭제 실패", wx.OK | wx.ICON_ERROR, self)
