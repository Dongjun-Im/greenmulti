"""즐겨찾기 관리 / 선택 대화상자 (v1.7).

키:
  · ↑/↓: 항목 이동
  · Enter: 선택 항목 열기
  · Delete / D: 선택 항목 삭제
  · F2: 이름 변경
  · Esc: 닫기
"""

from __future__ import annotations

import wx

from bookmark_manager import BookmarkManager
from screen_reader import speak
from theme import apply_theme, make_font, load_font_size


class BookmarkDialog(wx.Dialog):
    """즐겨찾기 목록 대화상자."""

    def __init__(self, parent, manager: BookmarkManager):
        super().__init__(parent, title="즐겨찾기", size=(560, 480),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.manager = manager
        # 결과: 호출자가 dlg.selected_url 로 확인. None 이면 닫기/취소.
        self.selected_url: str | None = None
        self.selected_name: str = ""

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        info = wx.StaticText(
            panel,
            label="↑↓ 이동 · Enter 열기 · D/Del 삭제 · F2 이름 변경 · Esc 닫기",
        )
        vbox.Add(info, 0, wx.ALL, 8)

        self.list_ctrl = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE)
        vbox.Add(self.list_ctrl, 1, wx.ALL | wx.EXPAND, 8)
        self.list_ctrl.Bind(wx.EVT_LISTBOX_DCLICK, lambda e: self._on_open())

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.open_btn = wx.Button(panel, label="열기(&O)")
        self.open_btn.Bind(wx.EVT_BUTTON, lambda e: self._on_open())
        btn_sizer.Add(self.open_btn, 0, wx.RIGHT, 6)

        self.rename_btn = wx.Button(panel, label="이름 변경(&R)")
        self.rename_btn.Bind(wx.EVT_BUTTON, lambda e: self._on_rename())
        btn_sizer.Add(self.rename_btn, 0, wx.RIGHT, 6)

        self.delete_btn = wx.Button(panel, label="삭제(&D)")
        self.delete_btn.Bind(wx.EVT_BUTTON, lambda e: self._on_delete())
        btn_sizer.Add(self.delete_btn, 0, wx.RIGHT, 6)

        self.up_btn = wx.Button(panel, label="위로(&U)")
        self.up_btn.Bind(wx.EVT_BUTTON, lambda e: self._on_move(-1))
        btn_sizer.Add(self.up_btn, 0, wx.RIGHT, 6)

        self.down_btn = wx.Button(panel, label="아래로(&N)")
        self.down_btn.Bind(wx.EVT_BUTTON, lambda e: self._on_move(1))
        btn_sizer.Add(self.down_btn, 0, wx.RIGHT, 6)

        self.close_btn = wx.Button(panel, wx.ID_CANCEL, label="닫기")
        btn_sizer.Add(self.close_btn, 0)
        vbox.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 8)

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

        self._refresh_list()
        self.list_ctrl.SetFocus()

    def _refresh_list(self) -> None:
        items = [f"{i+1}. {b.name}  ({b.url})"
                 for i, b in enumerate(self.manager.items)]
        if not items:
            items = ["(즐겨찾기가 비어 있습니다 — 게시판에서 Ctrl+D 로 추가)"]
            self.list_ctrl.Set(items)
            self.list_ctrl.Enable(False)
            self.open_btn.Enable(False)
            self.rename_btn.Enable(False)
            self.delete_btn.Enable(False)
            self.up_btn.Enable(False)
            self.down_btn.Enable(False)
        else:
            self.list_ctrl.Enable(True)
            self.open_btn.Enable(True)
            self.rename_btn.Enable(True)
            self.delete_btn.Enable(True)
            self.up_btn.Enable(True)
            self.down_btn.Enable(True)
            self.list_ctrl.Set(items)
            self.list_ctrl.SetSelection(0)

    def _selected_index(self) -> int:
        sel = self.list_ctrl.GetSelection()
        return sel if sel != wx.NOT_FOUND else -1

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        mods = event.HasModifiers()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        if key == wx.WXK_RETURN and not mods:
            self._on_open()
            return
        if key in (wx.WXK_DELETE, ord("D")) and not mods:
            self._on_delete()
            return
        if key == wx.WXK_F2 and not mods:
            self._on_rename()
            return
        event.Skip()

    def _on_open(self) -> None:
        idx = self._selected_index()
        item = self.manager.get(idx)
        if not item:
            return
        self.selected_url = item.url
        self.selected_name = item.name
        self.EndModal(wx.ID_OK)

    def _on_delete(self) -> None:
        idx = self._selected_index()
        item = self.manager.get(idx)
        if not item:
            return
        ans = wx.MessageBox(
            f"즐겨찾기에서 다음 항목을 삭제할까요?\n\n{item.name}",
            "삭제 확인", wx.YES_NO | wx.ICON_QUESTION, self,
        )
        if ans != wx.YES:
            return
        if self.manager.remove(idx):
            speak("즐겨찾기를 삭제했습니다.")
            self._refresh_list()
            new_sel = min(idx, len(self.manager.items) - 1)
            if new_sel >= 0:
                self.list_ctrl.SetSelection(new_sel)

    def _on_rename(self) -> None:
        idx = self._selected_index()
        item = self.manager.get(idx)
        if not item:
            return
        dlg = wx.TextEntryDialog(
            self, "새 이름을 입력하세요.", "즐겨찾기 이름 변경",
            value=item.name,
        )
        try:
            if dlg.ShowModal() != wx.ID_OK:
                return
            new_name = dlg.GetValue().strip()
        finally:
            dlg.Destroy()
        if not new_name:
            return
        item.name = new_name
        self.manager.save()
        self._refresh_list()
        self.list_ctrl.SetSelection(idx)
        speak("이름을 변경했습니다.")

    def _on_move(self, direction: int) -> None:
        idx = self._selected_index()
        if idx < 0:
            return
        new_idx = idx + direction
        if not (0 <= new_idx < len(self.manager.items)):
            return
        if self.manager.reorder(idx, new_idx):
            self._refresh_list()
            self.list_ctrl.SetSelection(new_idx)
            speak("순서를 바꿨습니다.")
