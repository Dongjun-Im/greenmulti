"""명령 도구 모음 (v1.7).

Ctrl+P 로 띄워 키워드를 입력해 모든 기능을 검색·실행한다.
입력 텍스트가 명령 이름·설명·단축키 어디든 포함되면 매칭.
↑/↓ 로 결과 이동, Enter 로 실행, Esc 로 닫기.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

import wx

from screen_reader import speak
from theme import apply_theme, make_font, load_font_size


@dataclass
class Command:
    name: str
    description: str
    shortcut: str
    callback: Callable
    when: Callable[[], bool] | None = None  # 사용 가능 조건 (None 이면 항상 가능)


class CommandPaletteDialog(wx.Dialog):
    """검색 입력 + 결과 목록으로 구성된 명령 도구 모음."""

    def __init__(self, parent, commands: list[Command]):
        super().__init__(parent, title="명령 도구 모음 (Ctrl+P)",
                         size=(560, 460),
                         style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
        self.all_commands = [c for c in commands if (c.when is None or c.when())]
        self.filtered: list[Command] = list(self.all_commands)

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        info = wx.StaticText(
            panel,
            label="키워드를 입력해 명령을 찾으세요. ↑↓로 이동, Enter로 실행, Esc로 닫기.",
        )
        vbox.Add(info, 0, wx.ALL, 8)

        self.search_ctrl = wx.TextCtrl(panel, value="", name="명령 검색")
        vbox.Add(self.search_ctrl, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 8)
        self.search_ctrl.Bind(wx.EVT_TEXT, lambda e: self._refresh())

        self.list_ctrl = wx.ListBox(panel, choices=[], style=wx.LB_SINGLE,
                                    name="명령 목록")
        vbox.Add(self.list_ctrl, 1, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 8)
        self.list_ctrl.Bind(wx.EVT_LISTBOX_DCLICK, lambda e: self._on_run())

        panel.SetSizer(vbox)
        apply_theme(self, make_font(load_font_size()))
        self.Bind(wx.EVT_CHAR_HOOK, self._on_char_hook)

        self._refresh()
        self.search_ctrl.SetFocus()

    def _refresh(self) -> None:
        q = self.search_ctrl.GetValue().strip().lower()
        if q:
            self.filtered = [
                c for c in self.all_commands
                if q in c.name.lower()
                or q in c.description.lower()
                or q in c.shortcut.lower()
            ]
        else:
            self.filtered = list(self.all_commands)
        items = [
            f"{c.name}    [{c.shortcut}]    — {c.description}"
            if c.shortcut else f"{c.name}    — {c.description}"
            for c in self.filtered
        ]
        if not items:
            self.list_ctrl.Set(["(일치하는 명령이 없습니다)"])
            self.list_ctrl.Enable(False)
        else:
            self.list_ctrl.Enable(True)
            self.list_ctrl.Set(items)
            self.list_ctrl.SetSelection(0)

    def _on_char_hook(self, event):
        key = event.GetKeyCode()
        focused = self.FindFocus()
        if key == wx.WXK_ESCAPE:
            self.EndModal(wx.ID_CANCEL)
            return
        if key == wx.WXK_RETURN:
            self._on_run()
            return
        # 검색창에 포커스가 있어도 ↑↓ 로 결과 목록 이동
        if key in (wx.WXK_UP, wx.WXK_DOWN) and focused is self.search_ctrl:
            n = self.list_ctrl.GetCount()
            if n == 0:
                return
            sel = self.list_ctrl.GetSelection()
            if sel == wx.NOT_FOUND:
                sel = 0
            new = sel + (-1 if key == wx.WXK_UP else 1)
            new = max(0, min(n - 1, new))
            self.list_ctrl.SetSelection(new)
            return
        event.Skip()

    def _on_run(self) -> None:
        if not self.filtered:
            return
        sel = self.list_ctrl.GetSelection()
        if sel == wx.NOT_FOUND or sel >= len(self.filtered):
            return
        cmd = self.filtered[sel]
        self.EndModal(wx.ID_OK)
        # 호출자(MainFrame) 가 dlg.selected_command 로 받아 실행
        self.selected_command = cmd

    selected_command: Command | None = None
