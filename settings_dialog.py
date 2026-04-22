"""통합 설정 대화상자 (F7).
- 왼쪽 목록(ListBox)에서 위/아래 방향키로 카테고리 선택
- 오른쪽 패널(Simplebook)이 해당 카테고리 설정을 표시
  · 화면 테마: 8가지 저시력 프리셋
  · 사운드: 마스터 스위치 + 이벤트별 on/off + WAV 파일 지정
"""
import os

import wx

from config import (
    load_update_settings, save_update_settings, APP_VERSION,
    UPDATE_INTERVALS, UPDATE_INTERVAL_KEYS,
)
from screen_reader import speak
from sound import (
    SOUND_EVENTS, load_sound_settings, save_sound_settings,
    play_file, resolve_event_path,
)
from theme import (
    THEME_PRESETS, THEME_ORDER, apply_theme, make_font, load_font_size,
    load_theme_key, set_current_theme, save_font_size,
    MIN_FONT_SIZE, MAX_FONT_SIZE, DEFAULT_FONT_SIZE,
)


# ─────────────────────────── 화면 테마 페이지 ───────────────────────────

class ThemePage(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self.original_key = load_theme_key()
        self.original_font_size = load_font_size()

        label = wx.StaticText(self, label="화면 테마를 선택하세요:")
        names = [THEME_PRESETS[k]["name"] for k in THEME_ORDER]
        try:
            cur_idx = THEME_ORDER.index(self.original_key)
        except ValueError:
            cur_idx = 0

        self.listbox = wx.ListBox(
            self, choices=names, style=wx.LB_SINGLE, name="화면 테마 목록",
        )
        self.listbox.SetSelection(cur_idx)
        self.listbox.Bind(wx.EVT_LISTBOX, self._on_pick)

        # 글꼴 크기 콤보
        font_label = wx.StaticText(self, label="글꼴 크기(&F):")
        font_choices = [str(s) for s in range(MIN_FONT_SIZE, MAX_FONT_SIZE + 1)]
        self.font_combo = wx.ComboBox(
            self, choices=font_choices,
            value=str(self.original_font_size),
            style=wx.CB_READONLY,
            name="글꼴 크기",
        )
        self.font_combo.Bind(wx.EVT_COMBOBOX, self._on_font_size_changed)

        hint = wx.StaticText(
            self,
            label=(
                "위/아래 방향키로 테마를 이동하면 즉시 미리보기됩니다.\n"
                "글꼴 크기도 선택 즉시 전체 화면에 적용됩니다.\n"
                "취소 시 이전 테마와 글꼴 크기로 되돌아갑니다."
            ),
        )

        font_row = wx.BoxSizer(wx.HORIZONTAL)
        font_row.Add(font_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        font_row.Add(self.font_combo, 0)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(label, 0, wx.ALL, 8)
        sizer.Add(self.listbox, 1, wx.EXPAND | wx.ALL, 8)
        sizer.Add(font_row, 0, wx.ALL, 8)
        sizer.Add(hint, 0, wx.ALL, 8)
        self.SetSizer(sizer)

    def _apply_font_size(self, new_size: int):
        """글꼴 크기를 저장하고 최상위 프레임에 즉시 적용."""
        save_font_size(new_size)
        top = wx.GetTopLevelParent(self)
        if hasattr(top, "current_font_size"):
            top.current_font_size = new_size
        apply = getattr(top, "_apply_full_theme", None)
        if callable(apply):
            apply()
        # 대화상자 자신도 새 글꼴로 재적용
        dlg = self.GetParent()
        while dlg and not isinstance(dlg, wx.Dialog):
            dlg = dlg.GetParent()
        if dlg:
            try:
                apply_theme(dlg, make_font(new_size))
            except Exception:
                pass

    def _on_font_size_changed(self, event):
        try:
            size = int(self.font_combo.GetValue())
        except (TypeError, ValueError):
            return
        self._apply_font_size(size)
        speak(f"글꼴 크기 {size}")

    def _on_pick(self, event):
        sel = self.listbox.GetSelection()
        if sel == wx.NOT_FOUND:
            return
        key = THEME_ORDER[sel]
        set_current_theme(key)
        top = wx.GetTopLevelParent(self)
        apply = getattr(top, "_apply_full_theme", None)
        if callable(apply):
            apply()
        speak(THEME_PRESETS[key]["name"])

    def focus_default(self):
        self.listbox.SetFocus()

    def revert_to_original(self):
        set_current_theme(self.original_key)
        if self.original_font_size != load_font_size():
            self._apply_font_size(self.original_font_size)
            return
        top = wx.GetTopLevelParent(self)
        apply = getattr(top, "_apply_full_theme", None)
        if callable(apply):
            apply()


# ─────────────────────────── 사운드 페이지 ───────────────────────────

class SoundPage(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)

        self.settings = load_sound_settings()
        self.enable_cbs: dict[str, wx.CheckBox] = {}
        self.path_ctrls: dict[str, wx.TextCtrl] = {}

        outer = wx.BoxSizer(wx.VERTICAL)

        # 마스터 스위치
        self.master_cb = wx.CheckBox(self, label="사운드 사용 (전체)(&A)")
        self.master_cb.SetValue(bool(self.settings.get("enabled", True)))
        outer.Add(self.master_cb, 0, wx.ALL, 8)

        master_hint = wx.StaticText(
            self,
            label=(
                "'사운드 사용 (전체)'를 끄면 아래 개별 설정과 무관하게 모든\n"
                "사운드가 재생되지 않습니다."
            ),
        )
        outer.Add(master_hint, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        # 이벤트별 박스를 세로로 쌓을 ScrolledWindow
        self.scroll = wx.ScrolledWindow(self, style=wx.VSCROLL)
        self.scroll.SetScrollRate(0, 20)
        scroll_sizer = wx.BoxSizer(wx.VERTICAL)

        events_map = self.settings.get("events") or {}
        event_enabled_map = self.settings.get("event_enabled") or {}

        for key, label in SOUND_EVENTS:
            box = wx.StaticBox(self.scroll, label=label)
            box_sizer = wx.StaticBoxSizer(box, wx.VERTICAL)

            cb = wx.CheckBox(box, label=f"{label} 사운드 재생(&P)")
            cb.SetValue(bool(event_enabled_map.get(key, True)))
            self.enable_cbs[key] = cb

            path_row = wx.BoxSizer(wx.HORIZONTAL)
            path_lbl = wx.StaticText(box, label="파일:")
            path_ctrl = wx.TextCtrl(
                box, value=events_map.get(key, ""),
                style=wx.TE_READONLY,
                name=f"{label} 사운드 파일 경로",
            )
            path_row.Add(path_lbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
            path_row.Add(path_ctrl, 1, wx.EXPAND)
            self.path_ctrls[key] = path_ctrl

            btn_row = wx.BoxSizer(wx.HORIZONTAL)
            browse_btn = wx.Button(box, label="찾아보기(&B)")
            reset_btn = wx.Button(box, label="초기화(&R)")
            test_btn = wx.Button(box, label="듣기(&L)")
            browse_btn.Bind(
                wx.EVT_BUTTON, lambda evt, k=key, lb=label: self._on_browse(k, lb),
            )
            reset_btn.Bind(
                wx.EVT_BUTTON, lambda evt, k=key: self._on_reset(k),
            )
            test_btn.Bind(
                wx.EVT_BUTTON, lambda evt, k=key: self._on_test(k),
            )
            btn_row.Add(browse_btn, 0, wx.RIGHT, 5)
            btn_row.Add(reset_btn, 0, wx.RIGHT, 5)
            btn_row.Add(test_btn, 0)

            box_sizer.Add(cb, 0, wx.ALL, 5)
            box_sizer.Add(path_row, 0, wx.EXPAND | wx.ALL, 5)
            box_sizer.Add(btn_row, 0, wx.ALL, 5)

            scroll_sizer.Add(
                box_sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8,
            )

        self.scroll.SetSizer(scroll_sizer)
        outer.Add(self.scroll, 1, wx.EXPAND | wx.ALL, 0)

        self.SetSizer(outer)

    def focus_default(self):
        self.master_cb.SetFocus()

    def _on_browse(self, event_key: str, label: str):
        cur = self.path_ctrls[event_key].GetValue().strip()
        start_dir = os.path.dirname(cur) if cur and os.path.exists(cur) else ""
        dlg = wx.FileDialog(
            self, f"{label} 사운드 파일 선택", defaultDir=start_dir,
            wildcard="WAV 사운드 파일 (*.wav)|*.wav|모든 파일 (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        )
        if dlg.ShowModal() == wx.ID_OK:
            self.path_ctrls[event_key].SetValue(dlg.GetPath())
            speak(f"{label} 사운드를 선택했습니다.")
        dlg.Destroy()

    def _on_reset(self, event_key: str):
        self.path_ctrls[event_key].SetValue("")

    def _on_test(self, event_key: str):
        custom = self.path_ctrls[event_key].GetValue().strip()
        if custom and os.path.exists(custom):
            path = custom
        else:
            path = resolve_event_path(
                event_key,
                {"enabled": True, "events": {event_key: custom}},
            )
        if not path:
            speak("재생할 사운드 파일이 없습니다.")
            return
        play_file(path)

    def collect(self) -> dict:
        events = {
            k: ctrl.GetValue().strip() for k, ctrl in self.path_ctrls.items()
        }
        event_enabled = {
            k: bool(cb.GetValue()) for k, cb in self.enable_cbs.items()
        }
        return {
            "enabled": bool(self.master_cb.GetValue()),
            "events": events,
            "event_enabled": event_enabled,
        }


# ─────────────────────────── 대화상자 ───────────────────────────

CATEGORIES: list[tuple[str, str]] = [
    ("theme", "화면 테마"),
    ("sound", "사운드"),
    ("update", "업데이트"),
]


class UpdatePage(wx.Panel):
    """자동 업데이트 설정 페이지."""

    def __init__(self, parent):
        super().__init__(parent)
        self.settings = load_update_settings()

        version_label = wx.StaticText(
            self, label=f"현재 설치된 버전: {APP_VERSION}",
        )

        self.check_cb = wx.CheckBox(
            self, label="프로그램 시작 시 자동으로 업데이트 확인(&S)",
        )
        self.check_cb.SetValue(bool(self.settings.get("check_on_startup", True)))

        # 릴리스 채널 라디오
        channel_box = wx.StaticBox(self, label="릴리스 채널")
        channel_sizer = wx.StaticBoxSizer(channel_box, wx.VERTICAL)
        self.rb_stable = wx.RadioButton(
            channel_box, label="정식 버전만(&A) — 안정적인 정식 릴리스만 받음",
            style=wx.RB_GROUP,
        )
        self.rb_beta = wx.RadioButton(
            channel_box, label="베타 포함(&B) — 정식 + 프리릴리스(테스트 버전) 모두 받음",
        )
        channel = self.settings.get("channel", "stable")
        if channel == "beta":
            self.rb_beta.SetValue(True)
        else:
            self.rb_stable.SetValue(True)
        channel_sizer.Add(self.rb_stable, 0, wx.ALL, 4)
        channel_sizer.Add(self.rb_beta, 0, wx.ALL, 4)

        # 자동 확인 주기 — RadioBox (접근성 좋음: 키보드 탐색 + 그룹 낭독)
        self._interval_keys = list(UPDATE_INTERVAL_KEYS)
        interval_choices = [label for _, label, _ in UPDATE_INTERVALS]
        self.interval_rb = wx.RadioBox(
            self, label="자동 확인 주기(&I)",
            choices=interval_choices,
            majorDimension=1, style=wx.RA_SPECIFY_COLS,
        )
        cur_interval = self.settings.get("check_interval", "weekly")
        try:
            cur_idx = self._interval_keys.index(cur_interval)
        except ValueError:
            cur_idx = self._interval_keys.index("weekly")
        self.interval_rb.SetSelection(cur_idx)

        hint = wx.StaticText(
            self,
            label=(
                "• 선택한 주기에 따라 자동 확인이 실행됩니다.\n"
                "• 네트워크가 없으면 조용히 넘어갑니다.\n"
                "• '지금 확인' 버튼으로 언제든지 바로 확인할 수 있습니다."
            ),
        )

        # 건너뛴 버전 표시 / 초기화
        skipped = self.settings.get("skip_version", "")
        self._skip_label_text = (
            f"건너뛴 버전: {skipped}" if skipped else "건너뛴 버전: (없음)"
        )
        self.skip_label = wx.StaticText(self, label=self._skip_label_text)
        self.reset_skip_btn = wx.Button(self, label="건너뛴 버전 다시 알림(&R)")
        self.reset_skip_btn.Enable(bool(skipped))
        self.reset_skip_btn.Bind(wx.EVT_BUTTON, self._on_reset_skip)

        # 지금 확인
        self.check_now_btn = wx.Button(self, label="지금 업데이트 확인(&C)")
        self.check_now_btn.Bind(wx.EVT_BUTTON, self._on_check_now)

        # 마지막 체크 시각 표시
        last = self.settings.get("last_check_iso", "")
        self.last_check_label = wx.StaticText(
            self,
            label=f"마지막 확인: {last}" if last else "마지막 확인: (아직 없음)",
        )

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(version_label, 0, wx.ALL, 8)
        sizer.Add(self.check_cb, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)
        sizer.Add(self.interval_rb, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)
        sizer.Add(channel_sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)
        sizer.Add(hint, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)
        sizer.AddSpacer(8)
        sizer.Add(self.skip_label, 0, wx.LEFT | wx.RIGHT, 8)
        sizer.Add(self.reset_skip_btn, 0, wx.ALL, 8)
        sizer.AddSpacer(8)
        sizer.Add(self.last_check_label, 0, wx.LEFT | wx.RIGHT, 8)
        sizer.Add(self.check_now_btn, 0, wx.ALL, 8)
        self.SetSizer(sizer)

    def focus_default(self):
        self.check_cb.SetFocus()

    def _on_reset_skip(self, event):
        s = load_update_settings()
        s["skip_version"] = ""
        save_update_settings(s)
        self.settings = s
        self.skip_label.SetLabel("건너뛴 버전: (없음)")
        self.reset_skip_btn.Enable(False)
        speak("건너뛴 버전을 해제했습니다.")

    def _on_check_now(self, event):
        # 메인 프레임의 수동 체크 경로를 재사용.
        top = wx.GetTopLevelParent(self)
        # SettingsDialog 의 부모 프레임을 찾아 거슬러 올라간다.
        frame = top.GetParent() if isinstance(top, wx.Dialog) else top
        runner = getattr(frame, "on_manual_update_check", None)
        if callable(runner):
            speak("업데이트를 확인합니다.")
            runner(None)
        else:
            speak("업데이트 확인 기능을 찾을 수 없습니다.")

    def collect(self) -> dict:
        s = load_update_settings()
        s["check_on_startup"] = bool(self.check_cb.GetValue())
        s["channel"] = "beta" if self.rb_beta.GetValue() else "stable"
        idx = self.interval_rb.GetSelection()
        if 0 <= idx < len(self._interval_keys):
            s["check_interval"] = self._interval_keys[idx]
        return s


class SettingsDialog(wx.Dialog):
    """왼쪽 목록에서 카테고리 선택, 오른쪽에 해당 설정 페이지 표시."""

    def __init__(self, parent):
        super().__init__(
            parent, title="설정",
            style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
            size=(820, 560),
        )

        # 왼쪽: 카테고리 목록
        self.category_list = wx.ListBox(
            self,
            choices=[name for _, name in CATEGORIES],
            style=wx.LB_SINGLE, name="설정 카테고리",
        )
        self.category_list.SetSelection(0)
        self.category_list.Bind(wx.EVT_LISTBOX, self._on_category)

        # 오른쪽: Simplebook (탭 없는 Notebook, 프로그래매틱 전환)
        self.book = wx.Simplebook(self)
        self.theme_page = ThemePage(self.book)
        self.sound_page = SoundPage(self.book)
        self.update_page = UpdatePage(self.book)
        self.book.AddPage(self.theme_page, "화면 테마")
        self.book.AddPage(self.sound_page, "사운드")
        self.book.AddPage(self.update_page, "업데이트")
        self.book.SetSelection(0)

        # 하단 버튼
        ok_btn = wx.Button(self, wx.ID_OK, "확인(&O)")
        cancel_btn = wx.Button(self, wx.ID_CANCEL, "취소(&X)")
        ok_btn.Bind(wx.EVT_BUTTON, self._on_ok)
        cancel_btn.Bind(wx.EVT_BUTTON, self._on_cancel)
        ok_btn.SetDefault()

        # 레이아웃
        split = wx.BoxSizer(wx.HORIZONTAL)
        split.Add(self.category_list, 0, wx.EXPAND | wx.ALL, 8)
        split.Add(self.book, 1, wx.EXPAND | wx.ALL, 8)

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        btn_row.AddStretchSpacer()
        btn_row.Add(ok_btn, 0, wx.RIGHT, 5)
        btn_row.Add(cancel_btn, 0)

        outer = wx.BoxSizer(wx.VERTICAL)
        outer.Add(split, 1, wx.EXPAND)
        outer.Add(btn_row, 0, wx.EXPAND | wx.ALL, 8)
        self.SetSizer(outer)

        # 카테고리 목록 너비를 적절히 고정
        self.category_list.SetMinSize((180, -1))
        self.Layout()

        # 현재 테마/글꼴 적용
        try:
            apply_theme(self, make_font(load_font_size()))
        except Exception:
            pass

        self.Centre()
        self.category_list.SetFocus()
        total = len(CATEGORIES)
        speak(
            "설정 대화상자. 위 아래 방향키로 카테고리를 선택하세요. "
            f"화면 테마 1/{total}"
        )

    def _on_category(self, event):
        sel = self.category_list.GetSelection()
        if sel == wx.NOT_FOUND:
            return
        self.book.SetSelection(sel)
        _, name = CATEGORIES[sel]
        total = len(CATEGORIES)
        # 위치 정보를 명시적으로 음성 안내해서 스크린리더가 자동으로 읽는
        # "사운드 전체 재생 체크상자" 같은 낭독을 이 메시지가 즉시 덮어쓰도록 한다.
        speak(f"{name} {sel + 1}/{total}")
        # 포커스가 오른쪽 페이지로 새지 않도록 카테고리 목록에 유지
        self.category_list.SetFocus()

    def _on_ok(self, event):
        # 테마는 ThemePage 선택 즉시 저장+적용됨. 여기선 사운드·업데이트만 저장.
        try:
            save_sound_settings(self.sound_page.collect())
        except Exception as e:
            speak(f"사운드 설정 저장 실패. {e}")
        try:
            save_update_settings(self.update_page.collect())
        except Exception as e:
            speak(f"업데이트 설정 저장 실패. {e}")
        self.EndModal(wx.ID_OK)

    def _on_cancel(self, event):
        self.theme_page.revert_to_original()
        self.EndModal(wx.ID_CANCEL)
