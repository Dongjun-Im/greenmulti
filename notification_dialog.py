"""알림 센터 대화상자.

쪽지·메일 알림을 통합 목록(ListBox)으로 표시. 각 항목 앞에 [쪽지]/[메일]
태그로 종류 구분. Enter 로 원본 열기, D/Del 로 개별 삭제, A 로 전체 삭제.
"""
from __future__ import annotations

import wx

from notification import NotificationItem, get_center
from screen_reader import speak
from theme import apply_theme, make_font, load_font_size


class NotificationCenterDialog(wx.Dialog):
    """알림 센터."""

    def __init__(self, parent, session, on_open_memo=None, on_open_mail=None):
        """
        on_open_memo: memo NotificationItem 을 받아 쪽지 보기 대화상자를 여는 콜백
        on_open_mail: mail NotificationItem 을 받아 메일 보기 대화상자를 여는 콜백
        """
        super().__init__(parent, title="알림 센터", size=(680, 480),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.session = session
        self.on_open_memo = on_open_memo
        self.on_open_mail = on_open_mail
        self.center = get_center()

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 상태 라벨
        self.status_label = wx.StaticText(panel, label="")
        vbox.Add(self.status_label, 0, wx.TOP | wx.LEFT | wx.RIGHT, 8)

        # 목록
        self.list_ctrl = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE)
        self.list_ctrl.Bind(wx.EVT_LISTBOX_DCLICK, lambda e: self._open_current())
        vbox.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)

        # 버튼
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.open_btn = wx.Button(panel, label="열기(&O)")
        self.open_btn.Bind(wx.EVT_BUTTON, lambda e: self._open_current())
        btn_sizer.Add(self.open_btn, 0, wx.RIGHT, 8)
        self.delete_btn = wx.Button(panel, label="알림 지우기(&D)")
        self.delete_btn.Bind(wx.EVT_BUTTON, lambda e: self._delete_current())
        btn_sizer.Add(self.delete_btn, 0, wx.RIGHT, 8)
        self.clear_all_btn = wx.Button(panel, label="모든 알림 지우기(&A)")
        self.clear_all_btn.Bind(wx.EVT_BUTTON, lambda e: self._clear_all())
        btn_sizer.Add(self.clear_all_btn, 0, wx.RIGHT, 8)
        self.refresh_btn = wx.Button(panel, label="새로고침(&F)")
        self.refresh_btn.Bind(wx.EVT_BUTTON, lambda e: self._refresh())
        btn_sizer.Add(self.refresh_btn, 0, wx.RIGHT, 8)
        self.close_btn = wx.Button(panel, wx.ID_CLOSE, label="닫기(&C)")
        self.close_btn.Bind(wx.EVT_BUTTON, lambda e: self.EndModal(wx.ID_CLOSE))
        btn_sizer.Add(self.close_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        # 키 안내
        hint = wx.StaticText(
            panel,
            label="↑↓ 이동 · Enter 열기 · D/Del 알림 지우기 · A 모두 지우기 · F 새로고침 · Esc 닫기",
        )
        vbox.Add(hint, 0, wx.ALL, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))

        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)
        self.list_ctrl.SetFocus()

        self._refresh()

    # ── 표시 갱신 ──

    def _format_item(self, idx: int, it: NotificationItem) -> str:
        tag = "[쪽지]" if it.type == "memo" else "[메일]" if it.type == "mail" else f"[{it.type}]"
        time_part = it.timestamp or it.received_at.strftime("%m-%d %H:%M")
        return f"{tag} {it.sender} · {it.summary} · {time_part}"

    def _refresh(self):
        items = self.center.items()
        lines = [self._format_item(i, it) for i, it in enumerate(items)]
        self.list_ctrl.Set(lines)
        if items:
            self.list_ctrl.SetSelection(0)
        memo_n = self.center.count_by_type("memo")
        mail_n = self.center.count_by_type("mail")
        total = len(items)
        if total == 0:
            self.status_label.SetLabel("알림이 없습니다.")
        else:
            self.status_label.SetLabel(
                f"총 {total}개 알림 (쪽지 {memo_n} · 메일 {mail_n})"
            )

    def _selected_index(self) -> int:
        sel = self.list_ctrl.GetSelection()
        return sel if sel != wx.NOT_FOUND else -1

    # ── 액션 ──

    def _open_current(self):
        idx = self._selected_index()
        items = self.center.items()
        if idx < 0 or idx >= len(items):
            return
        it = items[idx]
        if it.type == "memo" and self.on_open_memo:
            # 열기 후 해당 알림은 지움 (읽었으므로)
            self.on_open_memo(it)
            self.center.remove(it.type, it.item_id)
            self._refresh()
            self.list_ctrl.SetSelection(min(idx, self.list_ctrl.GetCount() - 1))
        elif it.type == "mail" and self.on_open_mail:
            self.on_open_mail(it)
            self.center.remove(it.type, it.item_id)
            self._refresh()
            self.list_ctrl.SetSelection(min(idx, self.list_ctrl.GetCount() - 1))
        else:
            wx.MessageBox(
                f"{it.type} 타입 알림은 현재 바로 열 수 없습니다.",
                "안내", wx.OK | wx.ICON_INFORMATION, self,
            )

    def _delete_current(self):
        idx = self._selected_index()
        items = self.center.items()
        if idx < 0 or idx >= len(items):
            return
        it = items[idx]
        ans = wx.MessageBox(
            f"이 알림을 지우시겠습니까?\n\n{self._format_item(idx, it)}",
            "알림 지우기", wx.YES_NO | wx.ICON_QUESTION, self,
        )
        if ans != wx.YES:
            return
        self.center.remove(it.type, it.item_id)
        self._refresh()
        if self.list_ctrl.GetCount() > 0:
            self.list_ctrl.SetSelection(min(idx, self.list_ctrl.GetCount() - 1))
        speak("알림을 지웠습니다.")

    def _clear_all(self):
        if self.center.count() == 0:
            speak("지울 알림이 없습니다.")
            return
        ans = wx.MessageBox(
            f"모든 알림 {self.center.count()}개를 지우시겠습니까?",
            "모든 알림 지우기",
            wx.YES_NO | wx.ICON_WARNING | wx.NO_DEFAULT, self,
        )
        if ans != wx.YES:
            return
        count = self.center.clear_all()
        self._refresh()
        speak(f"알림 {count}개를 모두 지웠습니다.")

    # ── 키 ──

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
            self._open_current()
            return
        if (key == ord("D") or key == wx.WXK_DELETE) and not mods:
            self._delete_current()
            return
        if key == ord("A") and not mods:
            self._clear_all()
            return
        if key == ord("F") and not mods:
            self._refresh()
            return
        event.Skip()
